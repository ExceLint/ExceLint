namespace ExceLint
    open System.Collections.Generic
    open System.Collections
    open System
    open ConfUtils

        type ScoreTable = Dict<string,(AST.Address*double)[]>
        type FastScoreTable = Dict<string*AST.Address,double>
        type HistoBin = string*Scope.SelectID*double
        type FreqTable = Dict<HistoBin,int>
        type Weights = IDictionary<AST.Address,double>
        type Ranking = KeyValuePair<AST.Address,double>[]
        type Causes = Dict<AST.Address,(HistoBin*int)[]>
        type ChangeSet = { mutants: KeyValuePair<AST.Address,string>[]; scores: ScoreTable; freqtable: FreqTable }

        module Pipeline =
            exception AnalysisCancelled

            type Input = {
                app: Microsoft.Office.Interop.Excel.Application;
                config: FeatureConf;
                dag: Depends.DAG;
                alpha: double;
                progress: Depends.Progress;
            }

            type Analysis = {
                scores: ScoreTable;
                ftable: FreqTable;
                ranking: Ranking;
                causes: Causes;
                ftable_time: int64;
                ranking_time: int64;
                significance_threshold: double;
                cutoff: int
            }

            type AnalysisOutcome =
            | Success of Analysis
            | Cancellation

            let comb (fn1: Input -> Analysis -> AnalysisOutcome)(fn2: Input -> Analysis -> AnalysisOutcome) : Input -> Analysis -> AnalysisOutcome =
                fun (input: Input)(analysis: Analysis) ->
                    match (fn1 input analysis) with
                    | Success(analysis2) -> fn2 input analysis2
                    | Cancellation -> Cancellation

        type ErrorModel(app: Microsoft.Office.Interop.Excel.Application, config: FeatureConf, dag: Depends.DAG, alpha: double, progress: Depends.Progress) =
            let mutable _was_cancelled = false

            let _input : Pipeline.Input = { app = app; config = config; dag = dag; alpha = alpha; progress = progress; }

            // build model
            let _analysis = ErrorModel.runModel _input

            // find model that minimizes anomalousness
            let _ranking2 = ErrorModel.inferAddressModes _analysis _input

            // compute weights
            let _weights = if config.IsEnabled "WeightByIntrinsicAnomalousness" then
                                ErrorModel.intrinsicAnomalousnessWeights _analysis_base dag
                            else
                                ErrorModel.noWeights _analysis_base dag

            // adjust ranking by intrinsic anomalousness
            let _ranking3 = ErrorModel.nonzero(ErrorModel.reweightRanking _ranking2 _weights)

            // sort ranking
            let _rankingSorted = ErrorModel.canonicalSort  _ranking3

            // compute significance threshold
            let _significanceThreshold : double =
                // compute total mass of distribution
                let total_mass = Array.sumBy (fun (kvp: KeyValuePair<AST.Address,double>) -> kvp.Value) _rankingSorted
                // compute the maximum score allowable
                total_mass * alpha

            // compute cutoff
            let _cutoff = ErrorModel.findCutIndex _rankingSorted _significanceThreshold _causes

            member self.WasCancelled : bool = _was_cancelled

            member self.ScoreTimeInMilliseconds : int64 = _score_time

            member self.FrequencyTableTimeInMilliseconds : int64 = _ftable_time

            member self.RankingTimeInMilliseconds : int64 = _ranking_time

            member self.NumScoreEntries : int = Array.fold (fun acc (pairs: (AST.Address*double)[]) ->
                                                    acc + pairs.Length
                                                ) 0 (_scores.Values |> Seq.toArray)

            member self.NumFreqEntries : int = _ftable.Count

            member self.NumRankedEntries : int = _analysis_base(dag).Length

            member self.causeOf(addr: AST.Address) : KeyValuePair<HistoBin,int>[] =
                Array.map (fun cause ->
                    let (bin,count) = cause
                    new KeyValuePair<HistoBin,int>(bin,count)
                ) _causes.[addr]

            member self.weightOf(addr: AST.Address) : double = _weights.[addr]

            member self.rankByFeatureSum() : Ranking =
                if ErrorModel.rankingIsSane _ranking dag (config.IsEnabled "AnalyzeOnlyFormulas") then
                    _rankingSorted
                else
                    failwith "ERROR: Formula-only analysis returns non-formulas."

            member self.getSignificanceCutoff : int = _cutoff

            member self.inspectSelectorFor(addr: AST.Address, sel: Scope.Selector) : KeyValuePair<AST.Address,(string*double)[]>[] =
                let sID = sel.id addr

                let d = new Dict<AST.Address,(string*double) list>()

                Seq.iter (fun (kvp: KeyValuePair<string,(AST.Address*double)[]>) ->
                    let fname: string = kvp.Key
                    let scores: (AST.Address*double)[] = kvp.Value

                    let valid_scores =
                        Array.choose (fun (addr2,score) ->
                            if sel.id addr2 = sID then
                                Some (addr2,score)
                            else
                                None
                        ) scores

                    Array.iter (fun (addr2,score) ->
                        if d.ContainsKey addr2 then
                            d.[addr2] <- (fname,score) :: d.[addr2]
                        else
                            d.Add(addr2, [(fname,score)])
                    ) valid_scores
                ) _scores

                let debug = Seq.map (fun (kvp: KeyValuePair<AST.Address,(string*double) list>) ->
                                        let addr2: AST.Address = kvp.Key
                                        let scores: (string*double)[] = kvp.Value |> List.toArray

                                        new KeyValuePair<AST.Address,(string*double)[]>(addr2, scores)
                                     ) d

                debug |> Seq.toArray

            // _analysis_base specifies which cells should be ranked:
            // 1. allCells means all cells in the spreadsheet
            // 2. onlyFormulas means only formulas
            static member analysisBase(config: FeatureConf)(d: Depends.DAG) : AST.Address[] =
                if config.IsEnabled("AnalyzeOnlyFormulas") then
                    d.getAllFormulaAddrs()
                else
                    d.allCells()

            // sanity check: are scores monotonically increasing?
            static member monotonicallyIncreasing(r: Ranking) : bool =
                let mutable last = 0.0
                let mutable outcome = true
                for kvp in r do
                    if kvp.Value >= last then
                        last <- kvp.Value
                    else
                        outcome <- false
                outcome

            static member prettyHistoBinDesc(hb: HistoBin) : string =
                let (feature_name,selector,fscore) = hb
                "(" +
                "feature name: " + feature_name + ", " +
                "selector: " + Scope.Selector.ToPretty selector + ", " +
                "feature value: " + fscore.ToString() +
                ")"

            static member private rankingIsSane(r: Ranking)(dag: Depends.DAG)(formulasOnly: bool) : bool =
                if formulasOnly then
                    Array.forall (fun (kvp: KeyValuePair<AST.Address,double>) -> dag.isFormula kvp.Key) r
                else
                    true

            static member private nonzero(r: Ranking) : Ranking =
                Array.filter (fun (kvp: KeyValuePair<AST.Address,double>) -> kvp.Value > 0.0) r

            static member private canonicalSort(r: Ranking) : Ranking =
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
                assert (ErrorModel.monotonicallyIncreasing arr)
                arr

            static member private transitiveInputs(faddr: AST.Address)(dag : Depends.DAG) : AST.Address[] =
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

            static member private refCount(dag: Depends.DAG) : Dict<AST.Address,int> =
                // for each input in the dependence graph, count how many formulas transitively refer to it
                let refcounts = Array.map (fun i -> i,(dag.getFormulasThatRefCell i).Length) (dag.allCells()) |> adict

                // if an input was not available at the time of dependence graph construction,
                // it will not be in dag.allCells() but formulas may refer to it; this
                // adds what refcount information we can discern from the visible parts
                // of the dependence graph
                for f in (dag.getAllFormulaAddrs()) do
                    let inputs = ErrorModel.transitiveInputs f dag
                    for i in inputs do
                        if not (refcounts.ContainsKey i) then
                            refcounts.Add(i, 1)
                        else
                            refcounts.[i] <- refcounts.[i] + 1

                refcounts

            static member private intrinsicAnomalousnessWeights(analysis_base: Depends.DAG -> AST.Address[])(dag: Depends.DAG) : Weights =
                // get the set of cells to be analyzed
                let cells = analysis_base(dag)

                // determine how many formulas refer to each input
                let refcounts = ErrorModel.refCount dag

                // for each cell, compute cumulative reference count. the insight here
                // is that summary rows are counting things that are counted by
                // subcomputations; thus, we should inflate their ranks by how much
                // they over-count.
                // this really only makes sense for formulas, but in case the user
                // asked for a ranking of all cells, we compute refcounts here even
                // for non-formulas
                let weights = Array.map (fun f ->
                                  let inputs = ErrorModel.transitiveInputs f dag
                                  let weight = double (Array.sumBy (fun i -> refcounts.[i]) inputs)
                                  f,weight
                              ) cells

                weights |> dict

            static member private noWeights(analysis_base: Depends.DAG -> AST.Address[])(dag: Depends.DAG) : Weights =
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

            static member private reweightRanking(r: Ranking)(w: Weights) : Ranking =
                r |> Array.map (fun (kvp: KeyValuePair<AST.Address,double>) ->
                    let addr = kvp.Key
                    let score = kvp.Value
                    new KeyValuePair<AST.Address,double>(addr, w.[addr] * score)
                  )

            static member private getChangeSetAddresses(cs: ChangeSet) : AST.Address[] =
                Array.map (fun (kvp: KeyValuePair<AST.Address,string>) ->
                    kvp.Key
                ) cs.mutants

            static member private mutateDAG(cs: ChangeSet)(dag: Depends.DAG)(app: Microsoft.Office.Interop.Excel.Application)(p: Depends.Progress) : Depends.DAG =
                dag.CopyWithUpdatedFormulas(cs.mutants, app, true, p)

            static member private equivalenceClasses(ranking: Ranking)(causes: Causes) : Dict<AST.Address,int> =
                let rankgrps = Array.groupBy (fun (kvp: KeyValuePair<AST.Address,double>) -> kvp.Value) ranking

                let grpids = Array.mapi (fun i (hb,_) -> hb,i) rankgrps |> adict

                let output = Array.map (fun (kvp: KeyValuePair<AST.Address,double>) -> kvp.Key, grpids.[kvp.Value]) ranking |> adict

                output

            static member private seekEquivalenceBoundary(ranking: Ranking)(causes: Causes)(cut_idx: int) : int =
                if ranking.Length = 0 then
                    -1
                else
                    let ecs = ErrorModel.equivalenceClasses ranking causes

                    if ecs.[ranking.[cut_idx].Key] = ecs.[ranking.[cut_idx + 1].Key] then
                        // find the first index that is different by scanning backward
                        if cut_idx <= 0 then
                            -1
                        else
                            let mutable seek_idx = ecs.[ranking.[cut_idx - 1].Key]
                            while seek_idx = ecs.[ranking.[cut_idx].Key] && seek_idx >= 0 do
                                seek_idx <- seek_idx - 1
                            seek_idx
                    else
                        cut_idx

            // returns the index of the last element to KEEP
            // returns -1 if you should keep nothing
            static member private findCutIndex(ranking: Ranking)(thresh: double)(causes: Causes): int =
                // compute total order
                let rank_nums = Array.map (fun (kvp: KeyValuePair<AST.Address,double>) -> int(kvp.Value)) ranking

                // find the index of the "knee"
                let dderiv_idx = ErrorModel.dderiv(rank_nums)

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
                ErrorModel.seekEquivalenceBoundary ranking causes cut_idx

            static member private toDict(arr: ('a*'b)[]) : Dict<'a,'b> =
                // assumes that 'a is unique
                let d = new Dict<'a,'b>(arr.Length)
                Array.iter (fun (a,b) ->
                    d.Add(a,b)
                ) arr
                d

            static member private inferAddressModes(input: Pipeline.Input)(analysis: Pipeline.Analysis) : Pipeline.AnalysisOutcome =
                if input.config.IsEnabled "InferAddressModes" then
                    let cells = ErrorModel.analysisBase input.config input.dag

                    // convert ranking into map
                    let rankmap = analysis.ranking
                                  |> Array.map (fun (pair: KeyValuePair<AST.Address,double>) -> (pair.Key,pair.Value))
                                  |> ErrorModel.toDict

                    // get all the formulas that ref each cell
                    let refss = Array.map (fun i -> i, input.dag.getFormulasThatRefCell i) cells |> ErrorModel.toDict

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
                                       let cs = ErrorModel.chooseLikelyAddressMode i refs accdag input.config input.progress input.app

                                       // update DAG
                                       ErrorModel.mutateDAG cs accdag input.app input.progress
                                   else
                                       accdag
                               ) input.dag crank


                    // score
                    let scores = ErrorModel.runEnabledFeatures cells dag' input.config input.progress

                    // count freqs
                    let freqs = ErrorModel.buildFrequencyTable scores input.progress dag' input.config
                    
                    // rerank
                    let ranking = ErrorModel.rank cells freqs scores input.config;

                    // get causes
                    let causes = ErrorModel.causes (ErrorModel.analysisBase input.config input.dag) freqs scores input.config;

                    // TODO: we shouldn't just blindly pass along old _time numbers
                    Pipeline.Success(
                        {
                            scores = scores;
                            ftable = freqs;
                            ranking = ranking;
                            causes = causes;
                            ftable_time = analysis.ftable_time;
                            ranking_time = analysis.ranking_time;
                            significance_threshold = 0.05;
                            cutoff = 0;
                        }
                    )
                else
                    analysis.ranking

            static member private runModel(input: Pipeline.Input) : Pipeline.AnalysisOutcome =
                try
                    let _runf = fun () -> ErrorModel.runEnabledFeatures (ErrorModel.analysisBase input.config input.dag) input.dag input.config input.progress

                    // get scores for each feature: featurename -> (address, score)[]
                    let (scores: ScoreTable,score_time: int64) = PerfUtils.runMillis _runf ()

                    // build frequency table: (featurename, selector, score) -> freq
                    let _freqf = fun () -> ErrorModel.buildFrequencyTable scores input.progress input.dag input.config
                    let ftable,ftable_time = PerfUtils.runMillis _freqf ()

                    // rank
                    let _rankf = fun () -> ErrorModel.rank (ErrorModel.analysisBase input.config input.dag) ftable scores input.config
                    let ranking,ranking_time = PerfUtils.runMillis _rankf ()

                    // save causes
                    let _causef = fun () -> ErrorModel.causes (ErrorModel.analysisBase input.config input.dag) ftable scores input.config
                    let causes,causes_time = PerfUtils.runMillis _causef ()

                    Pipeline.Success(
                        {
                            scores = scores;
                            ftable = ftable;
                            ranking = ranking;
                            causes = causes;
                            ftable_time = ftable_time;
                            ranking_time = ranking_time + causes_time;
                            significance_threshold = 0.05;
                            cutoff = 0;
                        }
                    )
                with
                | AnalysisCancelled -> Pipeline.Cancellation

            static member private causes(cells: AST.Address[])(ftable: FreqTable)(scores: ScoreTable)(config: FeatureConf) : Causes =
                let fscores = ErrorModel.makeFastScoreTable scores

                // get histogram bin heights for every given cell
                // and for every enabled scope and feature
                let carr =
                    Array.map (fun addr ->
                        let causes = Array.map (fun sel ->
                                         ErrorModel.getFeatureCounts addr sel ftable fscores config
                                     ) (config.EnabledScopes) |> Array.concat
                        (addr, causes)
                    ) cells

                carr |> adict

            static member private rank(cells: AST.Address[])(ftable: FreqTable)(scores: ScoreTable)(config: FeatureConf) : Ranking =
                let fscores = ErrorModel.makeFastScoreTable scores

                // get sums for every given cell
                // and for every enabled scope
                let addrSums: (AST.Address*int)[] =
                    Array.map (fun addr ->
                        let sum = Array.sumBy (fun sel ->
                                      ErrorModel.sumFeatureCounts addr sel ftable fscores config
                                  ) (config.EnabledScopes)
                        addr, sum
                    ) cells

                // return KeyValuePairs
                Array.map (fun (addr,sum) -> new KeyValuePair<AST.Address,double>(addr,double sum)) addrSums

            static member private countBuckets(ftable: FreqTable) : int =
                // get total number of non-zero buckets in the entire table
                Seq.filter (fun (elem: KeyValuePair<HistoBin,int>) -> elem.Value > 0) ftable
                |> Seq.length

            static member private argmin(f: 'a -> int)(xs: 'a[]) : int =
                let fx = Array.map (fun x -> f x) xs

                Array.mapi (fun i res -> (i, res)) fx |>
                Array.fold (fun arg (i, res) ->
                    if arg = -1 || res < fx.[arg] then
                        i
                    else
                        arg
                ) -1 

            static member private transpose(mat: 'a[][]) : 'a[][] =
                // assumes that all subarrays are the same length
                Array.map (fun i ->
                    Array.map (fun j ->
                        mat.[j].[i]
                    ) [| 0..mat.Length - 1 |]
                ) [| 0..(mat.[0]).Length - 1 |]

            static member private genChanges(cell: AST.Address)(refs: AST.Address[])(dag: Depends.DAG)(config: FeatureConf)(progress: Depends.Progress)(app: Microsoft.Office.Interop.Excel.Application) : ChangeSet[] =
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
                let fsT = ErrorModel.transpose fs'

                Array.map (fun (addrs_exprs: (AST.Address*AST.Expression)[]) ->
                    // generate formulas for each expr AST
                    let mutants = Array.map (fun (addr, expr: AST.Expression) ->
                                    new KeyValuePair<AST.Address,string>(addr, expr.WellFormedFormula)
                                    ) addrs_exprs

                    // get new DAG
                    let dag' = dag.CopyWithUpdatedFormulas(mutants, app, true, progress)

                    // get the set of buckets
                    let mutBuckets = ErrorModel.runEnabledFeatures (
                                        Array.map (fun (kvp: KeyValuePair<AST.Address,string>) ->
                                            kvp.Key
                                        ) mutants
                                     ) dag' config ErrorModel.nop

                    // compute frequency tables
                    let mutFtable = ErrorModel.buildFrequencyTable mutBuckets ErrorModel.nop dag' config
                    
                    { mutants = mutants; scores = mutBuckets; freqtable = mutFtable }
                ) fsT


            static member private chooseLikelyAddressMode(input: AST.Address)(refs: AST.Address[])(dag: Depends.DAG)(config: FeatureConf)(progress: Depends.Progress)(app: Microsoft.Office.Interop.Excel.Application) : ChangeSet =
                // generate all variants for the formulas that refer to this cell
                let css = ErrorModel.genChanges input refs dag config progress app

                // count the buckets for the default
                let ref_fs = Array.map (fun (ref: AST.Address) ->
                                new KeyValuePair<AST.Address,string>(ref,dag.getFormulaAtAddress(ref))
                             ) refs
                let def_buckets = ErrorModel.runEnabledFeatures (
                                        Array.map (fun (kvp: KeyValuePair<AST.Address,string>) ->
                                            kvp.Key
                                        ) ref_fs
                                     ) dag config ErrorModel.nop
                let def_freq = ErrorModel.buildFrequencyTable def_buckets ErrorModel.nop dag config
                let def_count = ErrorModel.countBuckets def_freq

                // find the variants that minimize the bucket count
                let mutant_counts = Array.map (fun mutant ->
                                        ErrorModel.countBuckets mutant.freqtable
                                    ) css

                let mode_idx = ErrorModel.argmin (fun mutant ->
                                   // count histogram buckets
                                   ErrorModel.countBuckets mutant.freqtable
                               ) css

                if mutant_counts.[mode_idx] < def_count then
                    css.[mode_idx]
                else
                    { mutants = ref_fs; scores = def_buckets; freqtable = def_freq; }

            static member private runEnabledFeatures(cells: AST.Address[])(dag: Depends.DAG)(config: FeatureConf)(progress: Depends.Progress) =
                config.EnabledFeatures |>
                Array.map (fun fname ->
                    // get feature lambda
                    let feature = config.FeatureByName fname

                    let fvals =
                        Array.map (fun cell ->
                            progress.IncrementCounter()
                            cell, feature cell dag
                        ) cells
                    
                    fname, fvals
                ) |> adict

            static member private buildFrequencyTable(data: ScoreTable)(progress: Depends.Progress)(dag: Depends.DAG)(config: FeatureConf): FreqTable =
                let d = new Dict<HistoBin,int>()
                Array.iter (fun fname ->
                    Array.iter (fun (sel: Scope.Selector) ->
                        Array.iter (fun (addr: AST.Address, score: double) ->
                            let sID = sel.id addr
                            if d.ContainsKey (fname,sID,score) then
                                let freq = d.[(fname,sID,score)]
                                d.[(fname,sID,score)] <- freq + 1
                            else
                                d.Add((fname,sID,score), 1)
                            progress.IncrementCounter()
                        ) (data.[fname])
                    ) (Scope.Selector.Kinds)
                ) (config.EnabledFeatures)
                d

            static member private makeFastScoreTable(scores: ScoreTable) : FastScoreTable =
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

            // sum the count of the appropriate feature bin of every feature
            // for the given address
            static member private sumFeatureCounts(addr: AST.Address)(sel: Scope.Selector)(ftable: FreqTable)(scores: FastScoreTable)(config: FeatureConf) : int =
                Array.sumBy (fun fname -> 
                    // get selector ID
                    let sID = sel.id addr
                    // get feature score
                    let fscore = scores.[fname,addr]
                    // get score count
                    ftable.[(fname,sID,fscore)]
                ) (config.EnabledFeatures)

            static member private getFeatureCounts(addr: AST.Address)(sel: Scope.Selector)(ftable: FreqTable)(scores: FastScoreTable)(config: FeatureConf) : (HistoBin*int)[] =
                Array.map (fun fname -> 
                    // get selector ID
                    let sID = sel.id addr
                    // get feature score
                    let fscore = scores.[fname,addr]
                    // get score count
                    let count = ftable.[(fname,sID,fscore)]
                    (fname,sID,fscore),count
                ) (config.EnabledFeatures)

            // "AngleMin" algorithm
            static member private dderiv(y: int[]) : int =
                let mutable anglemin = 1
                let mutable angleminindex = 0
                for index in 0..(y.Length - 3) do
                    let angle = y.[index] + y.[index + 2] - 2 * y.[index + 1]
                    if angle < anglemin then
                        anglemin <- angle
                        angleminindex <- index
                angleminindex

            static member private nop = Depends.Progress.NOPProgress()

