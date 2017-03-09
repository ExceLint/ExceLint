namespace ExceLint
    open System.Collections.Generic
    open System.Collections
    open System
    open Utils
    open ConfUtils
    open Pipeline

        module ModelBuilder =
            let private nop = Depends.Progress.NOPProgress()

            let private toDict(arr: ('a*'b)[]) : Dict<'a,'b> =
                // assumes that 'a is unique
                let d = new Dict<'a,'b>(arr.Length)
                Array.iter (fun (a,b) ->
                    d.Add(a,b)
                ) arr
                d

            // _analysis_base specifies which cells should be ranked:
            // 1. allCells means all cells in the spreadsheet
            // 2. onlyFormulas means only formulas
            let analysisBase(config: FeatureConf)(d: Depends.DAG) : AST.Address[] =
                if config.IsEnabled("AnalyzeOnlyFormulas") then
                    d.getAllFormulaAddrs()
                else
                    d.allCells()

            let private transitiveInputs(faddr: AST.Address)(dag : Depends.DAG) : AST.Address[] =
                let rec tf(addr: AST.Address) : AST.Address list =
                    if (dag.isFormula addr) then
                        // find all of the inputs (local dependencies) for formula
                        let refs_single = dag.getFormulaSingleCellInputs addr |> List.ofSeq
                        let refs_vector = dag.getFormulaInputVectors addr |>
                                                List.ofSeq |>
                                                List.map (fun rng -> rng.Addresses() |> Array.toList) |>
                                                List.concat
                        let refs = refs_single @ refs_vector
                        // prepend addr & recursively call this function
                        addr :: (List.map tf refs |> List.concat)
                    else
                        [addr]
    
                tf faddr |> List.toArray

            let private refCount(dag: Depends.DAG) : Dict<AST.Address,int> =
                // for each input in the dependence graph, count how many formulas transitively refer to it
                let refcounts = Array.map (fun i -> i,(dag.getFormulasThatRefCell i).Length) (dag.allCells()) |> adict

                // if an input was not available at the time of dependence graph construction,
                // it will not be in dag.allCells() but formulas may refer to it; this
                // adds what refcount information we can discern from the visible parts
                // of the dependence graph
                for f in (dag.getAllFormulaAddrs()) do
                    let inputs = transitiveInputs f dag
                    for i in inputs do
                        if not (refcounts.ContainsKey i) then
                            refcounts.Add(i, 1)
                        else
                            refcounts.[i] <- refcounts.[i] + 1

                refcounts

            let private intrinsicAnomalousnessWeights(analysis_base: Depends.DAG -> AST.Address[])(dag: Depends.DAG) : Weights =
                // get the set of cells to be analyzed
                let cells = analysis_base(dag)

                // determine how many formulas refer to each input
                let refcounts = refCount dag

                // for each cell, compute cumulative reference count. the insight here
                // is that summary rows are counting things that are counted by
                // subcomputations; thus, we should inflate their ranks by how much
                // they over-count.
                // this really only makes sense for formulas, but in case the user
                // asked for a ranking of all cells, we compute refcounts here even
                // for non-formulas
                let weights = Array.map (fun f ->
                                  let inputs = transitiveInputs f dag
                                  let weight = double (Array.sumBy (fun i -> refcounts.[i]) inputs)
                                  f,weight
                              ) cells

                weights |> dict

            let private noWeights(analysis_base: Depends.DAG -> AST.Address[])(dag: Depends.DAG) : Weights =
                // get the set of cells to be analyzed
                let cells = analysis_base(dag)

                // for each cell, compute cumulative reference count. the insight here
                // is that summary rows are counting things that are counted by
                // subcomputations; thus, we should inflate their ranks by how much
                // they over-count.
                // this really only makes sense for formulas, but in case the user
                // asked for a ranking of all cells, we compute refcounts here even
                // for non-formulas
                let weights = Array.map (fun f ->
                                  f,1.0
                              ) cells

                weights |> dict

            let weights(input: Input)(analysis: Analysis) : AnalysisOutcome =
                // compute weights
                let weights = if input.config.IsEnabled "WeightByIntrinsicAnomalousness" then
                                  intrinsicAnomalousnessWeights (analysisBase input.config) input.dag
                              else
                                  noWeights (analysisBase input.config) input.dag

                match analysis with
                | Histogram h -> Success(Histogram({ h with weights = weights }))
                | COF c -> Success(COF({ c with weights = weights }))
                | Cluster c -> Success(Cluster({ c with weights = weights }))

            // sanity check: are scores monotonically increasing?
            let monotonicallyIncreasing(r: Ranking) : bool =
                let mutable last = 0.0
                let mutable outcome = true
                for kvp in r do
                    if kvp.Value >= last then
                        last <- kvp.Value
                    else
                        outcome <- false
                outcome

            let private shouldNotHaveZeros(r: Ranking) : bool =
                Array.TrueForAll (r, fun (kvp: KeyValuePair<AST.Address,double>) -> kvp.Value > 0.0)

            // "AngleMin" algorithm
            let private dderiv(y: double[]) : int =
                let mutable anglemin = 1.0
                let mutable angleminindex = 0
                for index in 0..(y.Length - 3) do
                    let angle = y.[index] + y.[index + 2] - 2.0 * y.[index + 1]
                    if angle < anglemin then
                        anglemin <- angle
                        angleminindex <- index
                angleminindex

            let private equivalenceClasses(ranking: Ranking) : Dict<AST.Address,int> =
                let rankgrps = Array.groupBy (fun (kvp: KeyValuePair<AST.Address,double>) -> kvp.Value) ranking

                let grpids = Array.mapi (fun i (hb,_) -> hb,i) rankgrps |> adict

                let output = Array.map (fun (kvp: KeyValuePair<AST.Address,double>) -> kvp.Key, grpids.[kvp.Value]) ranking |> adict

                output

            let private seekEquivalenceBoundary(ranking: Ranking)(cut_idx: int) : int =
                // if we have no anomalies, because either the cut index excludes
                // all cells or because the ranking is zero-length, just return -1
                if cut_idx = -1 || ranking.Length = 0 then
                    -1
                else
                    // a map from AST.Address to equivalence class number
                    let ecs = equivalenceClasses ranking

                    // get the equivalence class of the element at the cut index (inclusive)
                    let cutEC = ecs.[ranking.[cut_idx].Key]

                    // the last index in the ranking is
                    let lastidx = ranking.Length - 1

                    // if there is a "next element" in the ranking after the cut,
                    // is it in the same equivalence class?
                    if lastidx >= cut_idx + 1 && cutEC = ecs.[ranking.[cut_idx + 1].Key] then
                        // find the first index that is different by scanning backward
                        if cut_idx <= 0 then
                            -1
                        else
                            // find the index in the ranking where the equivalence class changes
                            let mutable seek_idx = cut_idx - 1
                            while seek_idx >= 0 && ecs.[ranking.[seek_idx].Key] = cutEC do
                                seek_idx <- seek_idx - 1
                            seek_idx
                    else
                        cut_idx

            // returns the index of the last element to KEEP
            // returns -1 if you should keep nothing
            let private findCutIndex(ranking: Ranking)(thresh_idx: int): int =
                if ranking.Length = 0 then
                    -1
                else
                    // extract totally-ordered score vector
                    let rank_nums = Array.map (fun (kvp: KeyValuePair<AST.Address,double>) -> kvp.Value) ranking

                    // cut it at thresh (inclusive)
                    let rank_nums' = rank_nums.[..thresh_idx]

                    // find the index of the "knee"
                    dderiv(rank_nums')

            let private kneeIndexOpt(input: Input)(analysis: Analysis) : AnalysisOutcome =
                let ranking = match analysis with 
                              | Histogram h -> h.ranking 
                              | COF c -> c.ranking 
                              | Cluster c -> c.ranking

                let sig_threshold_idx =
                    match analysis with 
                    | Histogram h -> h.sig_threshold_idx 
                    | COF c -> c.sig_threshold_idx 
                    | Cluster c -> c.sig_threshold_idx

                let idx =
                    if input.config.IsEnabledSpectralRanking then
                        // compute knee cutoff
                        findCutIndex ranking sig_threshold_idx
                    else
                        // stick with total %
                        sig_threshold_idx
                // does the cut index straddle an equivalence class?
                let ce = seekEquivalenceBoundary ranking idx

                match analysis with
                | Histogram h -> Success(Histogram({ h with cutoff_idx = ce }))
                | COF c -> Success(COF({ c with cutoff_idx = ce }))
                | Cluster c -> Success(Cluster({ c with cutoff_idx = ce }))

            let private cutoffIndex(input: Input)(analysis: Analysis) : AnalysisOutcome =
                let ranking = match analysis with | Histogram h -> h.ranking | COF c -> c.ranking | Cluster c -> c.ranking

                // compute total mass of distribution
                let total_mass = double ranking.Length
                // compute the index of the maximum element
                let idx = int (Math.Floor(total_mass * input.alpha))

                match analysis with
                | Histogram h -> Success(Histogram({ h with sig_threshold_idx = idx }))
                | COF c -> Success(COF({ c with sig_threshold_idx = idx }))
                | Cluster c -> Success(Cluster({ c with sig_threshold_idx = idx }))

            let private canonicalSort(input: Input)(analysis: Analysis) : AnalysisOutcome =
                let r = match analysis with | Histogram h -> h.ranking | COF c -> c.ranking | Cluster c -> c.ranking
                let arr = Array.sortWith (fun (a: KeyValuePair<AST.Address,double>)(b: KeyValuePair<AST.Address,double>) ->
                                              let a_addr: AST.Address = a.Key
                                              let a_score: double = a.Value
                                              let b_addr: AST.Address = b.Key
                                              let b_score: double = b.Value

                                              if a_score < b_score then
                                                  -1
                                              elif a_score = b_score then
                                                  if a_addr.Col < b_addr.Col then
                                                      -1
                                                  elif a_addr.Col = b_addr.Col then
                                                      if a_addr.Row < b_addr.Row then
                                                          -1
                                                      elif a_addr.Row = b_addr.Row then
                                                          0
                                                      else
                                                          1
                                                  else
                                                      1
                                              else
                                                  1
                                         ) r
                assert (monotonicallyIncreasing arr)

                match analysis with
                | Histogram h -> Success(Histogram({ h with ranking = arr }))
                | COF c -> Success(COF({ c with ranking = Array.rev arr }))
                | Cluster c -> Success(Cluster({ c with ranking = Array.rev arr }))

            let private reweightRanking(input: Input)(analysis: Analysis) : AnalysisOutcome =
                let ranking = match analysis with | Histogram h -> h.ranking | COF c -> c.ranking | Cluster c -> c.ranking
                let weights = match analysis with | Histogram h -> h.weights | COF c -> c.weights | Cluster c -> c.weights

                let ranking' = 
                    ranking
                    |> Array.map (fun (kvp: KeyValuePair<AST.Address,double>) ->
                        let addr = kvp.Key
                        let score = kvp.Value
                        new KeyValuePair<AST.Address,double>(addr, weights.[addr] * score)
                      )

                match analysis with
                | Histogram h -> Success(Histogram({ h with ranking = ranking' }))
                | COF c -> Success(COF({ c with ranking = ranking' }))
                | Cluster c -> Success(Cluster({ c with ranking = ranking' }))

            let private getChangeSetAddresses(cs: ChangeSet) : AST.Address[] =
                Array.map (fun (kvp: KeyValuePair<AST.Address,string>) ->
                    kvp.Key
                ) cs.mutants

            let private mutateDAG(cs: ChangeSet)(dag: Depends.DAG)(app: Microsoft.Office.Interop.Excel.Application)(p: Depends.Progress) : Depends.DAG =
                dag.CopyWithUpdatedFormulas(cs.mutants, app, true, p)

            let private makeFlatScoreTable(scores: ScoreTable) : FlatScoreTable =
                let mutable max = 0
                for arr in scores do
                    if arr.Value.Length > max then
                        max <- arr.Value.Length

                let d = new Dict<string*AST.Address,Countable>(max * scores.Count)
                
                Seq.iter (fun (kvp: KeyValuePair<string,(AST.Address*Countable)[]>) ->
                    let fname = kvp.Key
                    let arr = kvp.Value
                    Array.iter (fun (addr,score) ->
                        d.Add((fname,addr), score)
                    ) arr
                ) scores

                d

            let private sizeOfConditioningSet(addr: AST.Address)(sel: Scope.Selector)(selcache: Scope.SelectorCache)(sidcache: Scope.SelectIDCache)(dag: Depends.DAG) : int =
                // get selector ID
                let sID = sel.id addr dag selcache

                // get number of cells with matching sID
                if sidcache.ContainsKey sID then
                    sidcache.[sID].Count
                else
                    failwith ("sID cache is missing sID " + sID.ToString())
             
            let private conditioningSetWeight(addr: AST.Address)(sel: Scope.Selector)(csstable: ConditioningSetSizeTable)(config: FeatureConf) : double =
                if config.IsEnabled "WeightByConditioningSetSize" then
                    1.0 / double (csstable.[sel].[addr])
                else
                    1.0

            let private getCause(addr: AST.Address)(cells: AST.Address[])(sel: Scope.Selector)(ftable: FreqTable)(scores: FlatScoreTable)(csstable: ConditioningSetSizeTable)(selcache: Scope.SelectorCache)(config: FeatureConf)(dag: Depends.DAG) : (HistoBin*int*double)[] =
                Array.map (fun fname -> 
                    // get selector ID
                    let sID = sel.id addr dag selcache
                    // get feature score
                    let fscore = scores.[fname,addr]
                    // get score count
                    let count = ftable.[(fname,sID,fscore)]
                    // get weight coefficient (beta)
                    let beta = conditioningSetWeight addr sel csstable config
                    (fname,sID,fscore),count,beta
                ) (config.EnabledFeatures)

            let private causes(cells: AST.Address[])(ftable: FreqTable)(scores: ScoreTable)(csstable: ConditioningSetSizeTable)(selcache: Scope.SelectorCache)(config: FeatureConf)(progress: Depends.Progress)(dag: Depends.DAG) : Causes =
                let fscores = makeFlatScoreTable scores

                // get histogram bin heights for every given cell
                // and for every enabled scope and feature
                let carr =
                    Array.map (fun addr ->
                        let causes = Array.map (fun sel ->
                                        if progress.IsCancelled() then
                                            raise AnalysisCancelled

                                        getCause addr cells sel ftable fscores csstable selcache config dag

                                     ) (config.EnabledScopes) |> Array.concat
                        (addr, causes)
                    ) cells

                carr |> adict

            // sum the count of the appropriate feature bin of every feature
            // for the given address
            let private sumFeatureCounts(addr: AST.Address)(sel: Scope.Selector)(ftable: FreqTable)(scores: FlatScoreTable)(selcache: Scope.SelectorCache)(config: FeatureConf)(dag: Depends.DAG) : int =
                Array.sumBy (fun fname -> 
                    // get selector ID
                    let sID = sel.id addr dag selcache
                    // get feature score
                    let fscore = scores.[fname,addr]
                    // get score count
                    ftable.[(fname,sID,fscore)]
                ) (config.EnabledFeatures)

            // for every cell and for every requested conditional,
            // find the bin height for the cell, then sum all
            // of these bin heights to produce a total ranking score
            // THIS IS WHERE ALL THE ACTION HAPPENS, FOLKS
            let private totalHistoSums(cells: AST.Address[])(ftable: FreqTable)(scores: ScoreTable)(csstable: ConditioningSetSizeTable)(selcache: Scope.SelectorCache)(config: FeatureConf)(progress: Depends.Progress)(dag: Depends.DAG) : Ranking*HypothesizedFixes option =
                let fscores = makeFlatScoreTable scores

                // get sums for every given cell
                // and for every enabled scope
                let addrSums: (AST.Address*double)[] =
                    Array.map (fun addr ->
                        if progress.IsCancelled() then
                                raise AnalysisCancelled

                        let sum = Array.sumBy (fun sel ->
                                      // compute conditioning set weight
                                      let beta_sel_i = conditioningSetWeight addr sel csstable config

                                      // this is the conditional count
                                      let count = sumFeatureCounts addr sel ftable fscores selcache config dag

                                      // weigh
                                      beta_sel_i * (double count)
                                  ) (config.EnabledScopes)
                        addr, sum
                    ) cells

                // return KeyValuePairs
                let ranking = Array.map (fun (addr,sum) -> new KeyValuePair<AST.Address,double>(addr,double sum)) addrSums

                ranking, None

            let private countBuckets(ftable: FreqTable) : int =
                // get total number of non-zero buckets in the entire table
                Seq.filter (fun (elem: KeyValuePair<HistoBin,int>) -> elem.Value > 0) ftable
                |> Seq.length

            let private crappy_argmin(f: 'a -> int)(xs: 'a[]) : int =
                let fx = Array.map f xs

                Array.mapi (fun i res -> (i, res)) fx |>
                Array.fold (fun arg (i, res) ->
                    if arg = -1 || res < fx.[arg] then
                        i
                    else
                        arg
                ) -1 

            let private also_crappy_argmin(f: 'a -> double)(xs: 'a[]) : 'a =
                let fxs = Array.map f xs

                let idx = Array.mapi (fun i fx -> (i, fx)) fxs |>
                          Array.fold (fun arg (i, fx) ->
                              if arg = -1 || fx < fxs.[arg] then
                                  i
                              else
                                  arg
                          ) -1

                xs.[idx]

            let private argwhatever(f: 'a -> double)(xs: seq<'a>)(whatev: double -> double -> bool) : 'a =
                Seq.reduce (fun arg x -> if whatev (f arg) (f x) then arg else x) xs
            let private argmax(f: 'a -> double)(xs: seq<'a>) : 'a =
                argwhatever f xs (fun a b -> a > b)
            let private argmin(f: 'a -> double)(xs: seq<'a>) : 'a =
                argwhatever f xs (fun a b -> a < b)

            let private transpose(mat: 'a[][]) : 'a[][] =
                // assumes that all subarrays are the same length
                Array.map (fun i ->
                    Array.map (fun j ->
                        mat.[j].[i]
                    ) [| 0..mat.Length - 1 |]
                ) [| 0..(mat.[0]).Length - 1 |]

            let private runEnabledFeatures(cells: AST.Address[])(dag: Depends.DAG)(config: FeatureConf)(progress: Depends.Progress) =
                config.EnabledFeatures |>
                Array.map (fun fname ->
                    // get feature lambda
                    let feature = config.FeatureByName fname

                    let fvals =
                        Array.map (fun cell ->
                            if progress.IsCancelled() then
                                raise AnalysisCancelled

                            progress.IncrementCounter()
                            cell, feature cell dag
                        ) cells
                    
                    fname, fvals
                ) |> adict

            let private buildFrequencyTable(scoretable: ScoreTable)(selcache: Scope.SelectorCache)(progress: Depends.Progress)(dag: Depends.DAG)(config: FeatureConf): FreqTable*Scope.SelectIDCache =
                // as a side-effect, maintain a selectID cache to make
                // conditioning set size lookup fast
                let s = new Scope.SelectIDCache()

                let d = new Dict<HistoBin,int>()
                Array.iter (fun fname ->
                    Array.iter (fun (sel: Scope.Selector) ->
                        Array.iter (fun (addr: AST.Address, score: Countable) ->
                            if progress.IsCancelled() then
                                raise AnalysisCancelled

                            // fetch SelectID for this selector and address
                            let sID = sel.id addr dag selcache

                            // update SelectIDCache if necessary
                            if not (s.ContainsKey(sID)) then
                                s.Add(sID, set [addr])
                            else if not (s.[sID].Contains(addr)) then
                                s.[sID] <- s.[sID].Add(addr)

                            if d.ContainsKey (fname,sID,score) then
                                let freq = d.[(fname,sID,score)]
                                d.[(fname,sID,score)] <- freq + 1
                            else
                                d.Add((fname,sID,score), 1)
                            progress.IncrementCounter()
                        ) (scoretable.[fname])
                    ) (config.EnabledScopes)
                ) (config.EnabledFeatures)
                d,s

            let private buildCSSTable(cells: AST.Address[])(progress: Depends.Progress)(dag: Depends.DAG)(selcache: Scope.SelectorCache)(sidcache: Scope.SelectIDCache)(config: FeatureConf): ConditioningSetSizeTable =
                let d = new ConditioningSetSizeTable()
                Array.iter (fun (sel: Scope.Selector) ->
                    Array.iter (fun cell ->
                        if progress.IsCancelled() then
                            raise AnalysisCancelled

                        progress.IncrementCounter()

                        // size of cell's conditioning set when conditioned on sel
                        let n = sizeOfConditioningSet cell sel selcache sidcache dag

                        // initialize nested storage, if necessary
                        if not (d.ContainsKey sel) then
                            d.Add(sel, new Dict<AST.Address,int>())

                        // add to dictionary
                        d.[sel].Add(cell, n)
                    ) cells
                ) (config.EnabledScopes)
                d

            let private genChanges(cell: AST.Address)(refs: AST.Address[])(dag: Depends.DAG)(config: FeatureConf)(progress: Depends.Progress)(app: Microsoft.Office.Interop.Excel.Application) : ChangeSet[] =
                // for each referencing formula, systematically generate all ref variants
                let fs' = Array.mapi (fun i f ->
                            // get AST
                            let ast = dag.getASTofFormulaAt f

                            let mutator = ASTMutator.mutateExpr ast cell

                            let cabs_rabs = mutator AST.AddressMode.Absolute AST.AddressMode.Absolute
                            let cabs_rrel = mutator AST.AddressMode.Absolute AST.AddressMode.Relative
                            let crel_rabs = mutator AST.AddressMode.Relative AST.AddressMode.Absolute
                            let crel_rrel = mutator AST.AddressMode.Relative AST.AddressMode.Relative

                            [|(f, cabs_rabs); (f, cabs_rrel); (f, crel_rabs); (f, crel_rrel); |]
                          ) refs

                // make the first index the mode, the second index the formula
                let fsT = transpose fs'

                Array.map (fun (addrs_exprs: (AST.Address*AST.Expression)[]) ->
                    // generate formulas for each expr AST
                    let mutants = Array.map (fun (addr, expr: AST.Expression) ->
                                    new KeyValuePair<AST.Address,string>(addr, expr.WellFormedFormula)
                                    ) addrs_exprs

                    // get new DAG
                    let dag' = dag.CopyWithUpdatedFormulas(mutants, app, true, progress)

                    // new selector cache
                    let selcache = Scope.SelectorCache()

                    // get the set of buckets
                    let mutBuckets = runEnabledFeatures (
                                        Array.map (fun (kvp: KeyValuePair<AST.Address,string>) ->
                                            kvp.Key
                                        ) mutants
                                     ) dag' config nop

                    // compute frequency tables
                    let (mutFtable,mutSIDCache) = buildFrequencyTable mutBuckets selcache nop dag' config
                    
                    { mutants = mutants; scores = mutBuckets; freqtable = mutFtable; selcache = selcache; sidcache = mutSIDCache }
                ) fsT


            let private chooseLikelyAddressMode(input: AST.Address)(refs: AST.Address[])(dag: Depends.DAG)(config: FeatureConf)(progress: Depends.Progress)(app: Microsoft.Office.Interop.Excel.Application) : ChangeSet =
                // generate all variants for the formulas that refer to this cell
                let css = genChanges input refs dag config progress app

                // count the buckets for the default
                let ref_fs = Array.map (fun (ref: AST.Address) ->
                                new KeyValuePair<AST.Address,string>(ref,dag.getFormulaAtAddress(ref))
                             ) refs
                let def_buckets = runEnabledFeatures (
                                        Array.map (fun (kvp: KeyValuePair<AST.Address,string>) ->
                                            kvp.Key
                                        ) ref_fs
                                     ) dag config nop
                let def_selcache = Scope.SelectorCache()
                let (def_freq, def_sidcache) = buildFrequencyTable def_buckets def_selcache nop dag config
                let def_count = countBuckets def_freq

                // find the variants that minimize the bucket count
                let mutant_counts = Array.map (fun mutant ->
                                        countBuckets mutant.freqtable
                                    ) css

                let mode_idx = crappy_argmin (fun mutant ->
                                   // count histogram buckets
                                   countBuckets mutant.freqtable
                               ) css

                if mutant_counts.[mode_idx] < def_count then
                    css.[mode_idx]
                else
                    { mutants = ref_fs; scores = def_buckets; freqtable = def_freq; selcache = def_selcache; sidcache = def_sidcache }

            let private inferAddressModes(input: Input)(analysis: Analysis) : AnalysisOutcome =
                match analysis with
                | Histogram h ->
                    if input.config.IsEnabled "InferAddressModes" then
                        try
                            let cells = analysisBase input.config input.dag

                            // convert ranking into map
                            let rankmap = h.ranking
                                          |> Array.map (fun (pair: KeyValuePair<AST.Address,double>) -> (pair.Key,pair.Value))
                                          |> toDict

                            let rank_positions = h.ranking
                                                 |> Array.mapi (fun i (pair: KeyValuePair<AST.Address,double>) -> (pair.Key, i))
                                                 |> toDict

                            // get all the formulas that ref each cell
                            let refss = Array.map (fun i -> i, input.dag.getFormulasThatRefCell i) cells |> toDict

                            // rank inputs by their impact on the ranking
                            let crank = Array.map (fun input_addr ->
                                            let anomalous_formulas = Array.filter (
                                                                        fun formula_addr ->
                                                                            let pos = try
                                                                                        rank_positions.[formula_addr]
                                                                                      with
                                                                                      | e ->
                                                                                         let is_formula = input.dag.isFormula formula_addr
                                                                                         h.cutoff_idx + 1
                                                                            pos <= h.cutoff_idx
                                                                     ) (refss.[input_addr])

                                            if anomalous_formulas.Length > 0 then
                                                let sum = Array.sumBy (fun formula ->
                                                                rankmap.[formula]
                                                            ) anomalous_formulas
                                                let average_score = sum / double (anomalous_formulas.Length)
                                                Some(input_addr, average_score)
                                            else
                                                None
                                        ) cells |>
                                        Array.choose id |>
                                        Array.sortBy (fun (input,score) -> score) |>
                                        Array.map (fun (input,score) -> input)

                            // for each input cell, try changing all refs to either abs or rel;
                            // if anomalousness drops, keep new interpretation
                            let dag' = Array.fold (fun accdag i ->
                                           // get referring formulas
                                           let refs = refss.[i]

                                           if refs.Length <> 0 then
                                               // run inference
                                               let cs = chooseLikelyAddressMode i refs accdag input.config input.progress input.app

                                               // update DAG
                                               mutateDAG cs accdag input.app input.progress
                                           else
                                               accdag
                                       ) input.dag crank

                            // initialize selector cache
                            let selcache = Scope.SelectorCache()

                            // score
                            let scores = runEnabledFeatures cells dag' input.config input.progress

                            // count freqs
                            let (freqs,sidcache) = buildFrequencyTable scores selcache input.progress dag' input.config

                            // compute conditioning set size
                            let csstable = buildCSSTable cells input.progress dag' selcache sidcache input.config

                            // rerank
                            let (ranking,_) = totalHistoSums cells freqs scores csstable selcache input.config input.progress input.dag

                            // get causes
                            let causes = causes cells freqs scores csstable selcache input.config input.progress input.dag

                            Success(Histogram { h with scores = scores; ftable = freqs; ranking = ranking; causes = causes; })
                        with
                        | AnalysisCancelled -> Cancellation
                        | e ->
                            // for breakpoint-friendliness
                            raise e
                    else
                        Success(Histogram h)
                | COF c ->
                    // do nothing for now
                    Success(COF c)
                | Cluster c ->
                    // do nothing for now
                    Success(Cluster c)

            let private cancellableWait(input: Input)(analysis: Analysis) : AnalysisOutcome =
                let mutable timer = 10
                while not (input.progress.IsCancelled()) && timer > 0 do
                    System.Threading.Thread.Sleep(1000)
                    timer <- timer - 1
                Success(analysis)

            let private toRawCoords(cells: Set<AST.Address>) : Set<double*double> =
                cells
                // convert to raw coords
                |> Set.map (fun cell -> (double cell.X, double cell.Y))

            // the mean
            let private binCentroid(cells: Countable[]) : Countable =
                assert (cells.Length > 0)
                let n = double (cells.Length)

                cells
                |> Array.fold (fun (acc: Countable)(c: Countable) -> acc.Add c) (cells.[0].Zero)
                |> (fun c -> c.ScalarDivide n)

            let euclideanDistance(cell1: Countable)(cell2: Countable) : double =
                // note that it is possible for the two cells to be the same;
                // like when the cell happens to be the exact centroid
                Math.Sqrt((cell1.Negate.Add cell2).VectorMultiply(cell1.Negate.Add cell2))

            let sameSheet(P: Distribution)(feature: Feature)(scope: Scope.SelectID)(other_hash: Countable)(anom_hash: Countable) =
                let acells = P.[feature].[scope].[anom_hash]
                let ocells = P.[feature].[scope].[other_hash]
                let bothcells = Set.union acells ocells |> Set.toList
                let (allsame,_) = List.fold (fun (is_same: bool, ws_opt: string option)(addr: AST.Address) ->
                                      match ws_opt with
                                      | Some(ws) ->
                                         let wssame = ws = addr.A1Worksheet()
                                         (is_same && wssame, Some ws)
                                      | None -> (is_same, Some(addr.A1Worksheet()))
                                  ) (true,None) bothcells
                allsame

            let private earthMoversDistance(P: Distribution)(feature: AST.Address -> Countable)(fname: Feature)(scope: Scope.SelectID)(other_hash: Countable)(anom_hash: Countable): double =
                assert (other_hash <> anom_hash)
                assert (sameSheet P fname scope other_hash anom_hash)

                // get every cell in the named anomalous bin in the sheets conditional table only
                let dirt = P.[fname].[scope].[anom_hash] |> Set.toArray |> Array.map (fun a -> feature a)

                // get every cell in the named other bin(s) across all conditional tables
                let other_dirt = P.[fname].[scope].[other_hash] |> Set.toArray |> Array.map (fun a -> feature a)

                // compute the centroid of the other_hash
                let centroid = binCentroid other_dirt

                // compute work required to move dirt
                // amount * distance
                let dist = Array.sumBy (fun coord -> euclideanDistance coord centroid) dirt

                dist

            let private binsBelowCutoff(h: HistoAnalysis)(cutoff: int) : Set<HistoBin> =
                // find the set of (by definition small) bins that contribute to the highest-ranked cells
                Array.map (fun (pair: KeyValuePair<AST.Address,double>) ->
                                    Array.map (fun (bin,count,weight) -> bin) (h.causes.[pair.Key])
                                ) (h.ranking.[..cutoff]) |>
                                Array.concat |>
                                Set.ofArray

            let private cartesianProductByX(xset: Set<'a>)(yset: Set<'a>) : ('a*'a[]) list =
                // cartesian product, grouped by the first element,
                // excluding the element itself
                Set.map (fun x -> x, (Set.difference yset (Set.ofList [x])) |> Set.toArray) xset |> Set.toList

            let private justHashes(P: Distribution)(feature: Feature)(scope: Scope.SelectID) : Set<Countable> =
                let mutable s = set[]
                for hash_row in P.[feature].[scope] do
                    let hash = hash_row.Key
                    let addr = hash_row.Value |> Set.toList |> List.head
                    let other_addresses = Seq.map (fun (pair: KeyValuePair<Countable,Set<AST.Address>>) -> pair.Value |> Set.toList) (P.[feature].[scope]) |> Seq.toList |> List.concat |> List.distinct
                    let allsame = List.forall (fun (other_addr: AST.Address) -> other_addr.A1Worksheet() = addr.A1Worksheet()) other_addresses
                    assert allsame
                    s <- s.Add(hash)
                s

            let private justScopes(P: Distribution)(feature: Feature) : Set<Scope.SelectID> =
                let mutable s = set[]
                for hash_row in P.[feature] do
                    s <- s.Add(hash_row.Key)
                s

            let private EMDsbyFeature(feature: AST.Address -> Countable)(fname: Feature)(causes: Causes) : Dict<Countable,Countable*double> =
                // the initial distribution
                let P = ErrorModel.toDistribution causes

                // get all sheet scopes 
                let scopes = justScopes P fname
                assert (Set.forall (fun (scope: Scope.SelectID) -> scope.IsKind = Scope.SameSheet) scopes)

                // find the distance to move every cell with a given hash to the nearest hash of a larger bin
                List.map (fun scope -> 
                    // get set of feature hashes from distribution
                    let hashes = justHashes P fname scope

                    List.map (fun (a: Countable, hs: Countable[]) ->
                        // do not consider smaller bins
                        let a_count = P.[fname].[scope].[a].Count
                        let bigger = Array.filter (fun h -> P.[fname].[scope].[h].Count > a_count) hs

                        if bigger.Length > 0 then
                            let f = (fun h -> earthMoversDistance P feature fname scope h a) 
                            let min_hash = also_crappy_argmin f bigger
                            let min_distance = f min_hash
                            assert (sameSheet P fname scope a min_hash)
                            a, (min_hash, min_distance)
                        else
                            // If there is no closest hash, then we say that
                            // the closest hash is itself with a distance
                            // of positive infinity in order to put it at the
                            // end of the ranking.  Seems a bit like a hack.
                            a, (a, Double.PositiveInfinity)
                    ) (cartesianProductByX hashes hashes)
                ) (scopes |> Set.toList)
                |> List.concat
                |> adict

            let private rankByEMD(cells: AST.Address[])(input: Input)(causes: Causes) : Ranking*HypothesizedFixes option =
                let emds = Array.map (fun fname ->
                               // get feature function
                               let feature = (fun addr -> (input.config.FeatureByName fname) addr input.dag)
                               // return (fname, distance) pair
                               fname, EMDsbyFeature feature fname causes
                           ) (input.config.EnabledFeatures) |> adict

                let ranking' = Array.map (fun (cell: AST.Address) ->
                                   let sum = Array.sumBy (fun fname ->
                                                 // get feature function
                                                 let feature = input.config.FeatureByName fname
                                                 // compute the hash of the cell using feature
                                                 let hash = feature cell input.dag
                                                 // find the EMD to the closest hash
                                                 let (min_hash,min_dist) = emds.[fname].[hash]
                                                 min_dist
                                             ) (input.config.EnabledFeatures)
                                   cell, sum
                               ) cells
                               |> Array.sortBy (fun (cell,sum) -> sum)
                               |> Array.map (fun (cell,sum) -> new KeyValuePair<AST.Address,double>(cell,sum))

                let hf = new HypothesizedFixes()
                Array.iter (fun (cell: AST.Address) ->
                    Array.iter (fun fname ->
                        // get feature function
                        let feature = input.config.FeatureByName fname
                        // compute the hash of the cell using feature
                        let hash = feature cell input.dag
                        // find the closest hash
                        let (min_hash,_) = emds.[fname].[hash]

                        // add entry for cell, if necessary
                        if not (hf.ContainsKey(cell)) then
                            hf.Add(cell, new Dict<Feature,Countable>())

                        // add closest hash
                        hf.[cell].Add(fname, min_hash)

                    ) (input.config.EnabledFeatures)
                ) cells

                (ranking', Some hf)

            let private rankCOFScores(scores: ScoreTable) : Ranking =
                let d = new Dictionary<AST.Address,double>()
                
                for kvp in scores do
                    for (addr,score) in kvp.Value do
                        let dscore: double = match score with
                                             | Num n -> n
                                             | _ -> failwith "Wrong Countable type for COF."
                        if not (d.ContainsKey addr) then
                            d.Add(addr, dscore)
                        else
                            d.[addr] <- d.[addr] + dscore

                Seq.sortBy (fun (kvp: KeyValuePair<AST.Address,double>) -> kvp.Value) d
                |> Seq.toArray

            let private HsUnion<'a>(h1: HashSet<'a>)(h2: HashSet<'a>) : HashSet<'a> =
                let h3 = new HashSet<'a>(h1)
                h3.UnionWith h2
                h3

            let private getCOFFixes(scores: ScoreTable)(dag: Depends.DAG)(normalizeRefSpace: bool)(normalizeSSSpace: bool)(BDD: Dictionary<string,Dictionary<AST.Address,Vector.SquareVector>>)(DDs: Dictionary<Vector.WorksheetName,Vector.DistDict>) : Dictionary<AST.Address,HashSet<AST.Address>> =
                let d = new Dictionary<AST.Address,HashSet<AST.Address>>()

                for kvp in scores do
                    for (addr,score) in kvp.Value do
                        let k = Vector.COFk addr dag normalizeRefSpace normalizeSSSpace
                        let DD = DDs.[addr.WorksheetName]
                        let nb = Vector.Nk_cells addr k dag normalizeRefSpace normalizeSSSpace BDD DD
                        if not (d.ContainsKey addr) then
                            d.Add(addr, nb)
                        else
                            d.[addr] <- HsUnion d.[addr] nb

                d

            let private normalizeScores(ss: ScoreTable) : ScoreTable =
                let d = new Dict<string, (AST.Address*Countable) list>()

                // for each feature
                for feat in ss.Keys do
                    // initialize storage for feature
                    d.Add(feat, [])

                    // get all values for feature
                    let vs = ss.[feat]

                    // group by path, wb, sheet
                    let vs_s = Array.groupBy (fun (a: AST.Address, c: Countable) ->
                                   a.A1Path() + a.A1Workbook() + a.A1Worksheet()
                               ) vs |> toDict

                    // for each sheet
                    for kvp in vs_s do
                        let sheet = kvp.Key
                        // get all (addr,countable) pairs
                        let cells = kvp.Value
                        // normalize
                        let ncount = cells |> Array.map (fun (a,c) -> c) |> Countable.Normalize 
                        // recombine with addrs
                        let cells' = cells |> Array.mapi (fun i (a,_) -> (a,ncount.[i])) |> Array.toList

                        d.[feat] <- cells' @ d.[feat]

                d |> Seq.map (fun kvp -> kvp.Key, kvp.Value |> List.toArray) |> Seq.toArray |> toDict

            let private convertToLFR(ss: ScoreTable) : ScoreTable =
                let d = new ScoreTable()

                for kvp in ss do
                    let feat = kvp.Key
                    let cells = kvp.Value
                    let cells' = cells |> Array.map (fun (a,c) -> a, c.ToCVectorResultant)
                    d.[feat] <- cells'

                d 

            // Bin addresses by HistoBin
            let private binCountables(scoretable: ScoreTable)(selcache: Scope.SelectorCache)(dag: Depends.DAG)(config: FeatureConf) : ClusterTable =
                let t = new ClusterTable()

                Array.iter (fun fname ->
                    Array.iter (fun (sel: Scope.Selector) ->
                        Array.iter (fun (addr: AST.Address, score: Countable) ->

                            // fetch SelectID for this selector and address
                            let sID = sel.id addr dag selcache

                            let binname = fname,sID,score

                            if t.ContainsKey binname then
                                let bin = t.[binname]
                                t.[binname] <- addr :: bin
                            else
                                t.Add(binname, [addr])
                        ) (scoretable.[fname])
                    ) (config.EnabledScopes)
                ) (config.EnabledFeatures)
                t

            let private initialClustering(scoretable: ScoreTable)(dag: Depends.DAG)(config: FeatureConf) : Clustering =
                let t: Clustering = new Clustering()

                Array.iter (fun fname ->
                    Array.iter (fun (addr: AST.Address, score: Countable) ->
                        t.Add(new HashSet<AST.Address>([addr])) |> ignore
                    ) (scoretable.[fname])
                ) (config.EnabledFeatures)
                t

            type InvertedHistogram = Dict<AST.Address,HistoBin>

            let private invertedHistogram(scoretable: ScoreTable)(selcache: Scope.SelectorCache)(dag: Depends.DAG)(config: FeatureConf) : InvertedHistogram =
                let d = new InvertedHistogram()

                assert (config.EnabledScopes.Length = 1 && config.EnabledFeatures.Length = 1)

                Array.iter (fun fname ->
                    Array.iter (fun (sel: Scope.Selector) ->
                        Array.iter (fun (addr: AST.Address, score: Countable) ->

                            // fetch SelectID for this selector and address
                            let sID = sel.id addr dag selcache

                            // get binname
                            let binname = fname,sID,score

                            d.Add(addr,binname)
                            
                        ) (scoretable.[fname])
                    ) (config.EnabledScopes)
                ) (config.EnabledFeatures)

                d

            let private mergeBins(source: HistoBin)(target: HistoBin)(ct: ClusterTable) : ClusterTable =
                let ct' = new ClusterTable(ct)
                ct'.Remove source |> ignore
                ct'.[target] <- ct.[source] @ ct.[target]
                ct'

            let private cartesianProduct(xs: seq<'a>)(ys: seq<'b>) : seq<'a*'b> =
                xs |> Seq.collect (fun x -> ys |> Seq.map (fun y -> x, y))

            let private induceCompleteGraph(xs: seq<'a>)(excluding: Set<'a>) : seq<'a*'a> =
                cartesianProduct xs xs
                |> Seq.filter (fun (source,target) ->
                                    source <> target &&                 // no self edges
                                    not (Set.contains target excluding) // filter disallowed targets
                              )

            type HBDistance = HistoBin -> HistoBin -> ClusterTable -> FlatScoreTable -> double
            type Distance = HashSet<AST.Address> -> HashSet<AST.Address> -> double

            let private SSE(lfr1: HistoBin)(lfr2: HistoBin)(addrs_by_lfr: ClusterTable)(nlfrs: FlatScoreTable) : double =
                // get cluster 1's points
                let (lfr1_feature,_,lfr1c) = lfr1
                let pts1 = addrs_by_lfr.[lfr1] |> List.map (fun addr -> nlfrs.[lfr1_feature,addr])

                // get cluster 2's points
                let (lfr2_feature,_,_) = lfr2
                let pts2 = addrs_by_lfr.[lfr2] |> List.map (fun addr -> nlfrs.[lfr2_feature,addr])

                // compute delta SSE
                let pts1_mean = pts1 |> List.reduce (fun acc p -> acc.Add p) |> (fun p -> p.ScalarDivide (double (pts1.Length)))
                let pts1_sse = pts1 |> List.map (fun p -> (p.Sub pts1_mean).VectorMultiply (p.Sub pts1_mean)) |> List.sum

                let pts_merged = pts1 @ pts2
                let pts_merged_mean = pts_merged |> List.reduce (fun acc p -> acc.Add p) |> (fun p -> p.ScalarDivide (double (pts_merged.Length)))
                let pts_merged_sse = pts_merged |> List.map (fun p -> (p.Sub pts_merged_mean).VectorMultiply (p.Sub pts_merged_mean)) |> List.sum

                Math.Abs(pts1_sse - pts_merged_sse)

            let private maxEuclid(lfr1: HistoBin)(lfr2: HistoBin)(addrs_by_lfr: ClusterTable)(nlfrs: FlatScoreTable) : double =
                // get cluster 1's points
                let (lfr1_feature,_,_) = lfr1
                let pts1 = addrs_by_lfr.[lfr1] |> List.map (fun addr -> nlfrs.[lfr1_feature,addr])

                // get cluster 2's points
                let (lfr2_feature,_,_) = lfr2
                let pts2 = addrs_by_lfr.[lfr2] |> List.map (fun addr -> nlfrs.[lfr2_feature,addr])

                // get cartesian product
                let p1p2 = cartesianProduct pts1 pts2

                // find maximum distance between any two points
                p1p2
                |> Seq.map (fun (p1,p2) -> p1.EuclideanDistance p2)
                |> Seq.max

            let private pairwiseDistances(g: (HistoBin*HistoBin) list)(ct: ClusterTable)(fst: FlatScoreTable)(d: HBDistance) : Dict<HistoBin*HistoBin,double> =
                g
                |> List.map (fun (source,target) -> (source,target),d source target ct fst)
                |> List.toArray
                |> toDict

            let private centroid(c: HashSet<AST.Address>)(ih: InvertedHistogram) : Countable =
                c
                |> Seq.map (fun a -> ih.[a])    // get histobin for address
                |> Seq.map (fun (_,_,c) -> c)   // get countable from bin
                |> Seq.toArray                  // convert to array
                |> Countable.Mean               // get mean

            let private pairwiseClusterCentroidDistances(C: Clustering)(ih: InvertedHistogram)(d: Distance) : Dict<HashSet<AST.Address>*HashSet<AST.Address>,double> =
                let dists = new Dict<HashSet<AST.Address>*HashSet<AST.Address>, double>()
                let centroids = new Dict<Countable,HashSet<AST.Address>>()
                let pairs = C |> Seq.map (fun c -> centroid c ih) |> (fun s -> induceCompleteGraph s (set []))
                pairs |> Seq.iter (fun (s,t) -> dists.Add((centroids.[s], centroids.[t]), d centroids.[s] centroids.[t]))
                dists

            // do not propose fixes where the target cluster is smaller than
            // the source cluster
            let noLargeToSmallMerges(g: (HistoBin*HistoBin) list)(ct: ClusterTable) : (HistoBin*HistoBin) list =
                g
                |> List.filter (fun (source,target) -> ct.[target].Length >= ct.[source].Length)

            let private runOldClusterModel(input: Input) : AnalysisOutcome =
                try
                    // initialize selector cache
                    let selcache = Scope.SelectorCache()

                    // determine the set of cells to be analyzed
                    let cells = (analysisBase input.config input.dag)

                    // get all NLFRs for every formula cell
                    let _runf = fun () -> runEnabledFeatures cells input.dag input.config input.progress
                    let (nlfrs: ScoreTable,score_time: int64) = PerfUtils.runMillis _runf ()

                    // get LFRs
                    let lfrs = convertToLFR nlfrs

                    // initially cluster points by location-free representation (LFR)
                    let _freqf = fun () -> buildFrequencyTable lfrs selcache input.progress input.dag input.config
                    let (ftable,sidcache),ftable_time = PerfUtils.runMillis _freqf ()

                    // bin by LFR
                    let lfr_bins = binCountables lfrs selcache input.dag input.config

                    // make flat score table for NLFRs
                    let nlfr_fst = makeFlatScoreTable nlfrs

                    // compute initial graph using LFRs
                    let vs = ftable.Keys |> Seq.toList
                    let g = induceCompleteGraph vs (set []) |> Seq.toList
                    let g' = noLargeToSmallMerges g lfr_bins

                    // compute pairwise distances
                    let dists_max_euclid = pairwiseDistances g' lfr_bins nlfr_fst maxEuclid
                    let dists_sse = pairwiseDistances g' lfr_bins nlfr_fst SSE
                    
                    // choose the cluster merge that minimizes distance
                    // and that merges a small bin into a big bin

                    failwith "nerp"

                    // 
                with
                | AnalysisCancelled -> Cancellation

            let private runClusterModel(input: Input) : AnalysisOutcome =
                try
                    // initialize selector cache
                    let selcache = Scope.SelectorCache()

                    // determine the set of cells to be analyzed
                    let cells = (analysisBase input.config input.dag)

                    // get all NLFRs for every formula cell
                    let _runf = fun () -> runEnabledFeatures cells input.dag input.config input.progress
                    let (nlfrs: ScoreTable,score_time: int64) = PerfUtils.runMillis _runf ()

                    // convert to flat table
                    let nflrs_flat = makeFlatScoreTable nlfrs

                    // make HistoBin lookup by address
                    let hb_inv = invertedHistogram nlfrs selcache input.dag input.config

                    // initially assign every cell to its own cluster
                    let clusters = initialClustering nlfrs input.dag input.config

                    // define distance
                    let distance = (fun (source: HashSet<AST.Address>)(target: HashSet<AST.Address>) ->
                                        let s_centroid = centroid source hb_inv
                                        let t_centroid = centroid target hb_inv
                                        let dist = s_centroid.EuclideanDistance t_centroid
                                        dist * (double source.Count)
                                   )

                    let pp(c: HashSet<AST.Address>) : string =
                        c
                        |> Seq.map (fun a -> a.A1Local())
                        |> (fun addrs -> "[" + String.Join(",", addrs) + "] with centroid " + (centroid c hb_inv).ToString())

                    let mutable log = []

                    // while there is more than 1 cluster, merge the two clusters with the smallest distance
                    while clusters.Count > 1 do
                        // get pairwise distances
                        let dists = pairwiseClusterCentroidDistances clusters hb_inv distance

                        // get the two clusters that minimize distance
                        let (source,target) = argmin (fun pair -> dists.[pair]) dists.Keys

                        // merge them
                        source |> Seq.iter (fun addr -> target.Add(addr) |> ignore)
                        clusters.Remove(source) |> ignore

                        // record merge in log
                        log <- (pp source, pp target) :: log

                    failwith "nerp"
                with
                | AnalysisCancelled -> Cancellation

            let private runCOFModel(input: Input) : AnalysisOutcome =
                try
                    let cells = (analysisBase input.config input.dag)

                    let _runf = fun () -> runEnabledFeatures cells input.dag input.config input.progress

                    // get COF scores for each feature: featurename -> (address, score)[]
                    let (scores: ScoreTable,score_time: int64) = PerfUtils.runMillis _runf ()

                    // rank
                    let _rankf = fun () -> rankCOFScores scores
                    let (ranking,ranking_time) = PerfUtils.runMillis _rankf ()

                    // get fixes
                    let _fixf = fun () -> getCOFFixes scores input.dag input.config.NormalizeRefs input.config.NormalizeSS (input.config.BDD input.dag) (input.config.DD input.dag)
                    let (fixes,fixes_time) = PerfUtils.runMillis _fixf ()

                    Success(COF
                        {
                            scores = scores;
                            ranking = ranking;
                            fixes = fixes;
                            fixes_time = fixes_time;
                            score_time = score_time;
                            ranking_time = ranking_time;
                            sig_threshold_idx = 0;
                            cutoff_idx = 0;
                            weights = new Dictionary<AST.Address,double>();
                        }
                    )
                with
                | AnalysisCancelled -> Cancellation

            let private runModel(input: Input) : AnalysisOutcome =
                try
                    // initialize selector cache
                    let selcache = Scope.SelectorCache()

                    let cells = (analysisBase input.config input.dag)

                    let _runf = fun () -> runEnabledFeatures cells input.dag input.config input.progress

                    // get scores for each feature: featurename -> (address, score)[]
                    let (scores: ScoreTable,score_time: int64) = PerfUtils.runMillis _runf ()

                    // build frequency table: (featurename, selector, score) -> freq
                    let _freqf = fun () -> buildFrequencyTable scores selcache input.progress input.dag input.config
                    let (ftable,sidcache),ftable_time = PerfUtils.runMillis _freqf ()

                    // build conditioning set size table
                    let _cssf = fun () -> buildCSSTable cells input.progress input.dag selcache sidcache input.config
                    let csstable,csstable_time = PerfUtils.runMillis _cssf ()

                    // save causes
                    let _causef = fun () -> causes cells ftable scores csstable selcache input.config input.progress input.dag
                    let causes,causes_time = PerfUtils.runMillis _causef ()

                    // rank
                    let _rankf = fun () ->
                                    if input.config.IsEnabledSpectralRanking then
                                        // note that zero scores are OK here
                                        rankByEMD cells input causes
                                    else
                                        let (r,hfo) = totalHistoSums cells ftable scores csstable selcache input.config input.progress input.dag
                                        assert shouldNotHaveZeros r
                                        r, hfo
                    let (ranking,fixes),ranking_time = PerfUtils.runMillis _rankf ()

                    Success(
                        Histogram(
                          {
                            scores = scores;
                            ftable = ftable;
                            csstable = csstable;
                            ranking = ranking;
                            causes = causes;
                            fixes = fixes;
                            score_time = score_time;
                            ftable_time = ftable_time;
                            csstable_time = csstable_time;
                            ranking_time = ranking_time;
                            causes_time = causes_time;
                            sig_threshold_idx = 0;
                            cutoff_idx = 0;
                            weights = new Dictionary<AST.Address,double>();
                          }
                        )
                    )
                with
                | AnalysisCancelled -> Cancellation

            let analyze(app: Microsoft.Office.Interop.Excel.Application)(config: FeatureConf)(dag: Depends.DAG)(alpha: double)(progress: Depends.Progress) =
                let config' = config.validate

                let input : Input = { app = app; config = config'; dag = dag; alpha = alpha; progress = progress; }

                if dag.IsCancelled() then
                    None
                elif input.config.IsCOF then
                    let pipeline = runCOFModel              // produce initial (unsorted) ranking
                                    +> weights              // compute weights
                                    +> reweightRanking      // modify ranking scores
                                    +> canonicalSort        // sort
                                    +> cutoffIndex          // compute initial cutoff index
                                    +> kneeIndexOpt         // optionally compute knee index
                                    +> inferAddressModes    // remove anomaly candidates
                                    +> canonicalSort
                                    +> kneeIndexOpt

                    match pipeline input with
                    | Success(analysis) -> Some (ErrorModel(input, analysis, config'))
                    | Cancellation -> None
                elif input.config.Cluster then
                    let pipeline = runClusterModel

                    match pipeline input with
                    | Success(analysis) -> Some (ErrorModel(input, analysis, config'))
                    | Cancellation -> None
                else
                    let pipeline = runModel                 // produce initial (unsorted) ranking
                                    +> weights              // compute weights
                                    +> reweightRanking      // modify ranking scores
                                    +> canonicalSort        // sort
                                    +> cutoffIndex          // compute initial cutoff index
                                    +> kneeIndexOpt         // optionally compute knee index
                                    +> inferAddressModes    // remove anomaly candidates
                                    +> canonicalSort
                                    +> kneeIndexOpt

                    match pipeline input with
                    | Success(analysis) -> Some (ErrorModel(input, analysis, config'))
                    | Cancellation -> None