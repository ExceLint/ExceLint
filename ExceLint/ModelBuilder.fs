namespace ExceLint
    open System.Collections.Generic
    open System.Collections
    open System
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

                Success({ analysis with weights = weights })
                

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

            let private nonzero(r: Ranking) : Ranking =
                Array.filter (fun (kvp: KeyValuePair<AST.Address,double>) -> kvp.Value > 0.0) r

            // "AngleMin" algorithm
            let private dderiv(y: int[]) : int =
                let mutable anglemin = 1
                let mutable angleminindex = 0
                for index in 0..(y.Length - 3) do
                    let angle = y.[index] + y.[index + 2] - 2 * y.[index + 1]
                    if angle < anglemin then
                        anglemin <- angle
                        angleminindex <- index
                angleminindex

            let private equivalenceClasses(ranking: Ranking)(causes: Causes) : Dict<AST.Address,int> =
                let rankgrps = Array.groupBy (fun (kvp: KeyValuePair<AST.Address,double>) -> kvp.Value) ranking

                let grpids = Array.mapi (fun i (hb,_) -> hb,i) rankgrps |> adict

                let output = Array.map (fun (kvp: KeyValuePair<AST.Address,double>) -> kvp.Key, grpids.[kvp.Value]) ranking |> adict

                output

            let private seekEquivalenceBoundary(ranking: Ranking)(causes: Causes)(cut_idx: int) : int =
                // if we have no anomalies, because either the cut index excludes
                // all cells or because the ranking is zero-length, just return -1
                if cut_idx = -1 || ranking.Length = 0 then
                    -1
                else
                    // a map from AST.Address to equivalence class number
                    let ecs = equivalenceClasses ranking causes

                    // get the equivalence class of the element at the cut index (inclusive)
                    let cutEC = ecs.[ranking.[cut_idx].Key]

                    if cutEC = ecs.[ranking.[cut_idx + 1].Key] then
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
            let private findCutIndex(ranking: Ranking)(thresh: double)(causes: Causes): int =
                // compute total order
                let rank_nums = Array.map (fun (kvp: KeyValuePair<AST.Address,double>) -> int(kvp.Value)) ranking

                // find the index of the "knee"
                let dderiv_idx = dderiv(rank_nums)

                // cut the ranking at the knee index
                let knee_cut = if dderiv_idx = 0 then Array.empty else ranking.[0..dderiv_idx]

                // further cut by scaning through the list to find the
                // index of the last significant score
                let (cut_idx: int, _) = knee_cut
                                        |> Array.mapi (fun i elem -> (i,elem))
                                        |> Array.fold (fun (max_idx: int, mass: double)(i: int, score: KeyValuePair<AST.Address,double>) ->
                                            let cum_mass = mass + score.Value
                                            if cum_mass > thresh then
                                                (max_idx, cum_mass)
                                            else
                                                (i, cum_mass)
                                            ) (-1,0.0)

                // does the cut index straddle an equivalence class?
                seekEquivalenceBoundary ranking causes cut_idx

            let private cutoff(input: Input)(analysis: Analysis) : AnalysisOutcome =
                // compute cutoff
                let c = findCutIndex analysis.ranking analysis.significance_threshold analysis.causes
                Success({ analysis with cutoff = c })

            let private significanceThreshold(input: Input)(analysis: Analysis) : AnalysisOutcome =
                let st : double =
                    // compute total mass of distribution
                    let total_mass = Array.sumBy (fun (kvp: KeyValuePair<AST.Address,double>) -> kvp.Value) analysis.ranking
                    // compute the maximum score allowable
                    total_mass * input.alpha
                Success({ analysis with significance_threshold = st })


            let private canonicalSort(input: Input)(analysis: Analysis) : AnalysisOutcome =
                let r = analysis.ranking
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
                Success({ analysis with ranking = arr })

            let private reweightRanking(input: Input)(analysis: Analysis) : AnalysisOutcome =
                let ranking = 
                    analysis.ranking
                    |> Array.map (fun (kvp: KeyValuePair<AST.Address,double>) ->
                        let addr = kvp.Key
                        let score = kvp.Value
                        new KeyValuePair<AST.Address,double>(addr, analysis.weights.[addr] * score)
                      )
                    |> nonzero

                Pipeline.Success({ analysis with ranking = ranking })

            let private getChangeSetAddresses(cs: ChangeSet) : AST.Address[] =
                Array.map (fun (kvp: KeyValuePair<AST.Address,string>) ->
                    kvp.Key
                ) cs.mutants

            let private mutateDAG(cs: ChangeSet)(dag: Depends.DAG)(app: Microsoft.Office.Interop.Excel.Application)(p: Depends.Progress) : Depends.DAG =
                dag.CopyWithUpdatedFormulas(cs.mutants, app, true, p)

            let private makeFastScoreTable(scores: ScoreTable) : FlatScoreTable =
                let mutable max = 0
                for arr in scores do
                    if arr.Value.Length > max then
                        max <- arr.Value.Length

                let d = new Dict<string*AST.Address,double>(max * scores.Count)
                
                Seq.iter (fun (kvp: KeyValuePair<string,(AST.Address*double)[]>) ->
                    let fname = kvp.Key
                    let arr = kvp.Value
                    Array.iter (fun (addr,score) ->
                        d.Add((fname,addr), score)
                    ) arr
                ) scores

                d

            let private sizeOfConditioningSet(addr: AST.Address)(cells: AST.Address[])(sel: Scope.Selector)(dag: Depends.DAG) : int =
                // get selector ID
                let sID = sel.id addr dag
                let total = Array.sumBy (fun c ->
                                let sID' = sel.id c dag
                                if sID = sID' then
                                    1
                                else
                                    0
                            ) cells
                total
             
            let private conditioningSetWeight(addr: AST.Address)(sel: Scope.Selector)(csstable: ConditioningSetSizeTable)(config: FeatureConf) : double =
                if config.IsEnabled "WeightByConditioningSetSize" then
                    1.0 / double (csstable.[sel].[addr])
                else
                    1.0

            let private getCause(addr: AST.Address)(cells: AST.Address[])(sel: Scope.Selector)(ftable: FreqTable)(scores: FlatScoreTable)(csstable: ConditioningSetSizeTable)(config: FeatureConf)(dag: Depends.DAG) : (HistoBin*int*double)[] =
                Array.map (fun fname -> 
                    // get selector ID
                    let sID = sel.id addr dag
                    // get feature score
                    let fscore = scores.[fname,addr]
                    // get score count
                    let count = ftable.[(fname,sID,fscore)]
                    // get weight coefficient (beta)
                    let beta = conditioningSetWeight addr sel csstable config
                    (fname,sID,fscore),count,beta
                ) (config.EnabledFeatures)

            let private causes(cells: AST.Address[])(ftable: FreqTable)(scores: ScoreTable)(csstable: ConditioningSetSizeTable)(config: FeatureConf)(progress: Depends.Progress)(dag: Depends.DAG) : Causes =
                let fscores = makeFastScoreTable scores

                // get histogram bin heights for every given cell
                // and for every enabled scope and feature
                let carr =
                    Array.map (fun addr ->
                        let causes = Array.map (fun sel ->
                                        if progress.IsCancelled() then
                                            raise AnalysisCancelled

                                        getCause addr cells sel ftable fscores csstable config dag

                                     ) (config.EnabledScopes) |> Array.concat
                        (addr, causes)
                    ) cells

                carr |> adict

            // sum the count of the appropriate feature bin of every feature
            // for the given address
            let private sumFeatureCounts(addr: AST.Address)(sel: Scope.Selector)(ftable: FreqTable)(scores: FlatScoreTable)(config: FeatureConf)(dag: Depends.DAG) : int =
                Array.sumBy (fun fname -> 
                    // get selector ID
                    let sID = sel.id addr dag
                    // get feature score
                    let fscore = scores.[fname,addr]
                    // get score count
                    ftable.[(fname,sID,fscore)]
                ) (config.EnabledFeatures)

            // for every cell and for every requested conditional,
            // find the bin height for the cell, then sum all
            // of these bin heights to produce a total ranking score
            // THIS IS WHERE ALL THE ACTION HAPPENS, FOLKS
            let private totalHistoSums(cells: AST.Address[])(ftable: FreqTable)(scores: ScoreTable)(csstable: ConditioningSetSizeTable)(config: FeatureConf)(progress: Depends.Progress)(dag: Depends.DAG) : Ranking =
                let fscores = makeFastScoreTable scores

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
                                      let count = sumFeatureCounts addr sel ftable fscores config dag

                                      // weigh
                                      beta_sel_i * (double count)
                                  ) (config.EnabledScopes)
                        addr, sum
                    ) cells

                // return KeyValuePairs
                Array.map (fun (addr,sum) -> new KeyValuePair<AST.Address,double>(addr,double sum)) addrSums

            let private countBuckets(ftable: FreqTable) : int =
                // get total number of non-zero buckets in the entire table
                Seq.filter (fun (elem: KeyValuePair<HistoBin,int>) -> elem.Value > 0) ftable
                |> Seq.length

            let private argmin(f: 'a -> int)(xs: 'a[]) : int =
                let fx = Array.map (fun x -> f x) xs

                Array.mapi (fun i res -> (i, res)) fx |>
                Array.fold (fun arg (i, res) ->
                    if arg = -1 || res < fx.[arg] then
                        i
                    else
                        arg
                ) -1 

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

            let private buildFrequencyTable(data: ScoreTable)(progress: Depends.Progress)(dag: Depends.DAG)(config: FeatureConf): FreqTable =
                let d = new Dict<HistoBin,int>()
                Array.iter (fun fname ->
                    Array.iter (fun (sel: Scope.Selector) ->
                        Array.iter (fun (addr: AST.Address, score: double) ->
                            if progress.IsCancelled() then
                                raise AnalysisCancelled

                            let sID = sel.id addr dag
                            if d.ContainsKey (fname,sID,score) then
                                let freq = d.[(fname,sID,score)]
                                d.[(fname,sID,score)] <- freq + 1
                            else
                                d.Add((fname,sID,score), 1)
                            progress.IncrementCounter()
                        ) (data.[fname])
                    ) (config.EnabledScopes)
                ) (config.EnabledFeatures)
                d

            let private buildCSSTable(cells: AST.Address[])(progress: Depends.Progress)(dag: Depends.DAG)(config: FeatureConf): ConditioningSetSizeTable =
                let d = new ConditioningSetSizeTable()
                Array.iter (fun (sel: Scope.Selector) ->
                    Array.iter (fun cell ->
                        if progress.IsCancelled() then
                            raise AnalysisCancelled

                        progress.IncrementCounter()

                        // size of cell's conditioning set when conditioned on sel
                        let n = sizeOfConditioningSet cell cells sel dag

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

                    // get the set of buckets
                    let mutBuckets = runEnabledFeatures (
                                        Array.map (fun (kvp: KeyValuePair<AST.Address,string>) ->
                                            kvp.Key
                                        ) mutants
                                     ) dag' config nop

                    // compute frequency tables
                    let mutFtable = buildFrequencyTable mutBuckets nop dag' config
                    
                    { mutants = mutants; scores = mutBuckets; freqtable = mutFtable }
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
                let def_freq = buildFrequencyTable def_buckets nop dag config
                let def_count = countBuckets def_freq

                // find the variants that minimize the bucket count
                let mutant_counts = Array.map (fun mutant ->
                                        countBuckets mutant.freqtable
                                    ) css

                let mode_idx = argmin (fun mutant ->
                                   // count histogram buckets
                                   countBuckets mutant.freqtable
                               ) css

                if mutant_counts.[mode_idx] < def_count then
                    css.[mode_idx]
                else
                    { mutants = ref_fs; scores = def_buckets; freqtable = def_freq; }

            let private inferAddressModes(input: Input)(analysis: Analysis) : AnalysisOutcome =
                if input.config.IsEnabled "InferAddressModes" then
                    try
                        let cells = analysisBase input.config input.dag

                        // convert ranking into map
                        let rankmap = analysis.ranking
                                      |> Array.map (fun (pair: KeyValuePair<AST.Address,double>) -> (pair.Key,pair.Value))
                                      |> toDict

                        // get all the formulas that ref each cell
                        let refss = Array.map (fun i -> i, input.dag.getFormulasThatRefCell i) cells |> toDict

                        // rank inputs by their impact on the ranking
                        let crank = Array.sortBy (fun input ->
                                        let sum = Array.sumBy (fun formula ->
                                                      rankmap.[formula]
                                                  ) (refss.[input])
                                        -sum
                                    ) cells

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


                        // score
                        let scores = runEnabledFeatures cells dag' input.config input.progress

                        // count freqs
                        let freqs = buildFrequencyTable scores input.progress dag' input.config

                        // compute conditioning set size
                        let csstable = buildCSSTable cells input.progress dag' input.config

                        // rerank
                        let ranking = totalHistoSums cells freqs scores csstable input.config input.progress input.dag

                        // get causes
                        let causes = causes cells freqs scores csstable input.config input.progress input.dag

                        // TODO: we shouldn't just blindly pass along old _time numbers
                        Success({ analysis with scores = scores; ftable = freqs; ranking = ranking; causes = causes; })
                    with
                    | AnalysisCancelled -> Cancellation
                else
                    Success(analysis)

            let private cancellableWait(input: Input)(analysis: Analysis) : AnalysisOutcome =
                let mutable timer = 10
                while not (input.progress.IsCancelled()) && timer > 0 do
                    System.Threading.Thread.Sleep(1000)
                    timer <- timer - 1
                Success(analysis)

            let private runModel(input: Input) : AnalysisOutcome =
                try
                    let cells = (analysisBase input.config input.dag)

                    let _runf = fun () -> runEnabledFeatures cells input.dag input.config input.progress

                    // get scores for each feature: featurename -> (address, score)[]
                    let (scores: ScoreTable,score_time: int64) = PerfUtils.runMillis _runf ()

                    // build frequency table: (featurename, selector, score) -> freq
                    let _freqf = fun () -> buildFrequencyTable scores input.progress input.dag input.config
                    let ftable,ftable_time = PerfUtils.runMillis _freqf ()

                    // build conditioning set size table
                    let _cssf = fun () -> buildCSSTable cells input.progress input.dag input.config
                    let csstable,csstable_time = PerfUtils.runMillis _cssf ()

                    // rank
                    let _rankf = fun () -> totalHistoSums cells ftable scores csstable input.config input.progress input.dag
                    let ranking,ranking_time = PerfUtils.runMillis _rankf ()

                    // save causes
                    let _causef = fun () -> causes cells ftable scores csstable input.config input.progress input.dag
                    let causes,causes_time = PerfUtils.runMillis _causef ()

                    Success(
                        {
                            scores = scores;
                            ftable = ftable;
                            csstable = csstable;
                            ranking = ranking;
                            causes = causes;
                            score_time = score_time;
                            ftable_time = ftable_time;
                            csstable_time = csstable_time;
                            ranking_time = ranking_time;
                            causes_time = causes_time;
                            significance_threshold = 0.05;
                            cutoff = 0;
                            weights = new Dictionary<AST.Address,double>();
                        }
                    )
                with
                | AnalysisCancelled -> Cancellation

            let analyze(app: Microsoft.Office.Interop.Excel.Application)(config: FeatureConf)(dag: Depends.DAG)(alpha: double)(progress: Depends.Progress) =
                let input : Input = { app = app; config = config; dag = dag; alpha = alpha; progress = progress; }

                if dag.IsCancelled() then
                    None
                else
                    let pipeline = runModel
    //                                +> cancellableWait    // for debugging
                                    +> inferAddressModes
                                    +> weights
                                    +> reweightRanking
                                    +> canonicalSort
                                    +> significanceThreshold
                                    +> cutoff

                    match pipeline input with
                    | Success(analysis) -> Some (ErrorModel(input, analysis))
                    | Cancellation -> None

        

