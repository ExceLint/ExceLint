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

        type ErrorModel(app: Microsoft.Office.Interop.Excel.Application, config: FeatureConf, dag: Depends.DAG, alpha: double, progress: Depends.Progress) =
            // _analysis_base specifies which cells should be ranked:
            // 1. allCells means all cells in the spreadsheet
            // 2. onlyFormulas means only formulas
            let _analysis_base(d: Depends.DAG) : AST.Address[] = if config.IsEnabled("AnalyzeOnlyFormulas") then d.getAllFormulaAddrs() else d.allCells()

            let _significanceThreshold : int =
                // round to integer
                int (
                    // get total number of counts (maximum possible histogram height)
                    double (_analysis_base(dag).Length * config.EnabledFeatures.Length * config.EnabledScopes.Length)
                    // times signficance
                    * alpha
                )

            // build model
            let (_scores: ScoreTable,
                 _ftable: FreqTable,
                 _ranking: Ranking,
                 _causes: Causes,
                 _score_time: int64,
                 _ftable_time: int64,
                 _ranking_time: int64) = ErrorModel.runModel _analysis_base dag config progress

            // find model that minimizes anomalousness
            let _ranking2 = if config.IsEnabled "InferAddressModes" then
                                ErrorModel.inferAddressModes _analysis_base _ranking dag config (ErrorModel.nop) app
                            else
                                _ranking

            // compute weights
            let _weights = if config.IsEnabled "WeightByIntrinsicAnomalousness" then
                               ErrorModel.intrinsicAnomalousnessWeights _analysis_base dag
                           else
                               ErrorModel.noWeights _analysis_base dag

            // adjust ranking by intrinsic anomalousness
            let _ranking3 = ErrorModel.nonzero(ErrorModel.reweightRanking _ranking2 _weights)

            // sort ranking
            let _rankingSorted = ErrorModel.canonicalSort  _ranking3

            // compute cutoff
            let _cutoff = ErrorModel.findCutIndex _rankingSorted _significanceThreshold _causes

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
                Array.sortWith (fun a b ->
                                       let a_addr: AST.Address = a.Key
                                       let a_score: double = a.Value
                                       let b_addr: AST.Address = b.Key
                                       let b_score: double = b.Value
                                       if a_score < b_score then
                                           -1
                                       else if a_score = b_score then
                                           if (a_addr.Col) < (b_addr.Col) then
                                               -1
                                           else if a_addr.Col = b_addr.Col && a_addr.Row <= b_addr.Row then
                                               -1
                                           else
                                               1
                                       else
                                           1
                                   ) r

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

            static member private intrinsicAnomalousnessWeights(analysis_base: Depends.DAG -> AST.Address[])(dag: Depends.DAG) : Weights =
                // get the set of cells to be analyzed
                let cells = analysis_base(dag)

                // for each input in the dependence graph, count how many formulas transitively refer to it
                let refcounts = Array.map (fun i -> i,(dag.getFormulasThatRefCell i).Length) (dag.allCells()) |> dict

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
                let ecs = ErrorModel.equivalenceClasses ranking causes
                if ecs.[ranking.[cut_idx].Key] = ecs.[ranking.[cut_idx + 1].Key] then
                    // find the first index that is different by scanning backward
                    let mutable seek = ecs.[ranking.[cut_idx - 1].Key]
                    while seek = ecs.[ranking.[cut_idx].Key] && seek > 0 do
                        seek <- seek - 1
                    seek
                else
                    cut_idx

            static member private findCutIndex(ranking: Ranking)(thresh: int)(causes: Causes): int =
                let sigThresh = thresh

                // compute total order
                let rank_nums = Array.map (fun (kvp: KeyValuePair<AST.Address,double>) -> int(kvp.Value)) ranking

                // find the index of the "knee"
                let dderiv_idx = ErrorModel.dderiv(rank_nums)

                // cut the ranking at the knee index
                let knee_cut = ranking.[0..dderiv_idx]

                // further cut by scaning through the list to find the
                // index of the last significant score
                let cut_idx: int = knee_cut
                                    |> Array.mapi (fun i elem -> (i,elem))
                                    |> Array.fold (fun (acc: int)(i: int, score: KeyValuePair<AST.Address,double>) ->
                                        if score.Value > double sigThresh then
                                            acc
                                        else
                                            i
                                        ) (knee_cut.Length - 1)

                // does the cut index straddle an equivalence class?
                ErrorModel.seekEquivalenceBoundary ranking causes cut_idx

            static member private toDict(arr: ('a*'b)[]) : Dict<'a,'b> =
                // assumes that 'a is unique
                let d = new Dict<'a,'b>(arr.Length)
                Array.iter (fun (a,b) ->
                    d.Add(a,b)
                ) arr
                d

            static member private inferAddressModes(analysis_base: Depends.DAG -> AST.Address[])(r: Ranking)(dag: Depends.DAG)(config: FeatureConf)(progress: Depends.Progress)(app: Microsoft.Office.Interop.Excel.Application) : Ranking =
                let cells = analysis_base(dag)

                // convert ranking into map
                let rankmap = r |> Array.map (fun (pair: KeyValuePair<AST.Address,double>) -> (pair.Key,pair.Value))
                                |> ErrorModel.toDict

                // get all the formulas that ref each cell
                let refss = Array.map (fun input -> input, dag.getFormulasThatRefCell input) cells |> ErrorModel.toDict

                // rank inputs by their impact on the ranking
                let crank = Array.sortBy (fun input ->
                                let sum = Array.sumBy (fun formula ->
                                              rankmap.[formula]
                                          ) (refss.[input])
                                -sum
                            ) cells

                // for each input cell, try changing all refs to either abs or rel;
                // if anomalousness drops, keep new interpretation
                let dag' = Array.fold (fun accdag input ->
                               // get referring formulas
                               let refs = refss.[input]

                               if refs.Length <> 0 then
                                   // run inference
                                   let cs = ErrorModel.chooseLikelyAddressMode input refs accdag config progress app

                                   // update DAG
                                   ErrorModel.mutateDAG cs accdag app progress
                               else
                                   accdag
                           ) dag crank


                // score
                let scores = ErrorModel.runEnabledFeatures cells dag' config progress

                // count freqs
                let freqs = ErrorModel.buildFrequencyTable scores progress dag' config

                // rerank
                ErrorModel.rank cells freqs scores config

            static member private runModel(analysisbase: Depends.DAG -> AST.Address[])(dag: Depends.DAG)(config: FeatureConf)(progress: Depends.Progress) =
                let _runf = fun () -> ErrorModel.runEnabledFeatures (analysisbase dag) dag config progress

                // get scores for each feature: featurename -> (address, score)[]
                let (scores: ScoreTable,score_time: int64) = PerfUtils.runMillis _runf ()

                // build frequency table: (featurename, selector, score) -> freq
                let _freqf = fun () -> ErrorModel.buildFrequencyTable scores progress dag config
                let ftable,ftable_time = PerfUtils.runMillis _freqf ()

                // rank
                let _rankf = fun () -> ErrorModel.rank (analysisbase dag) ftable scores config
                let ranking,ranking_time = PerfUtils.runMillis _rankf ()

                // save causes
                let _causef = fun () -> ErrorModel.causes (analysisbase dag) ftable scores config
                let causes,causes_time = PerfUtils.runMillis _causef ()

                (scores, ftable, ranking, causes, score_time, ftable_time, ranking_time + causes_time)

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

