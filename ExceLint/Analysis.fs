namespace ExceLint
    open System.Collections.Generic
    open System.Collections
    open System
    open ConfUtils

        type ScoreTable = Dict<string,(AST.Address*double)[]>
        type FastScoreTable = Dict<string*AST.Address,double>
        type FreqTable = Dict<string*Scope.SelectID*double,int>
        type Ranking = KeyValuePair<AST.Address,double>[]
        type Mutant = { mutants: KeyValuePair<AST.Address,string>[]; scores: ScoreTable; freqtable: FreqTable }

        type ErrorModel(config: FeatureConf, dag: Depends.DAG, alpha: double, progress: Depends.Progress) =
            let _significanceThreshold : int =
                // round to integer
                int (
                    // get total number of counts
                    double (dag.allCells().Length * config.EnabledFeatures.Length * config.EnabledScopes.Length)
                    // times signficance
                    * alpha
                )

            // build model
            let (_scores: ScoreTable,
                 _ftable: FreqTable,
                 _ranking: Ranking,
                 _score_time: int64,
                 _ftable_time: int64,
                 _ranking_time: int64) = ErrorModel.runModel dag config progress

            // find model that minimizes anomalousness
            let _ranking' = ErrorModel.inferAddressModes _ranking dag config (ErrorModel.nop)

            // compute cutoff
            let _cutoff = ErrorModel.findCutIndex _ranking _significanceThreshold

            member self.ScoreTimeInMilliseconds : int64 = _score_time

            member self.FrequencyTableTimeInMilliseconds : int64 = _ftable_time

            member self.RankingTimeInMilliseconds : int64 = _ranking_time

            member self.NumScoreEntries : int = Array.fold (fun acc (pairs: (AST.Address*double)[]) ->
                                                    acc + pairs.Length
                                                ) 0 (_scores.Values |> Seq.toArray)

            member self.NumFreqEntries : int = _ftable.Count

            member self.NumRankedEntries : int = dag.allCells().Length

            member self.rankByFeatureSum() : Ranking = _ranking

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

            static member private mutantDAG(mutants: Mutant[])(dag: Depends.DAG) : Depends.DAG =
                let fs = Array.fold (fun (acc: KeyValuePair<AST.Address,string> list)(mutant: Mutant) ->
                             (mutant.mutants |> Array.toList) @ acc
                         ) (List.empty) mutants |> Array.ofList
                dag.CopyWithUpdatedFormulas fs

            static member private findCutIndex(ranking: Ranking)(thresh: int): int =
                let sigThresh = thresh

                // compute total order
                let rank_nums = Array.map (fun (kvp: KeyValuePair<AST.Address,double>) -> int(kvp.Value)) ranking

                // find the index of the "knee"
                let dderiv_idx = ErrorModel.dderiv(rank_nums)

                // cut the ranking at the knee index
                let knee_cut = ranking.[0..dderiv_idx]

                // the ranking may include scores above the significance threshold, so
                // scan through the list to find the index of the last significant score
                let cut_idx: int = knee_cut
                                    |> Array.mapi (fun i elem -> (i,elem))
                                    |> Array.fold (fun (acc: int)(i: int, score: KeyValuePair<AST.Address,double>) ->
                                        if score.Value > double sigThresh then
                                            acc
                                        else
                                            i
                                        ) (knee_cut.Length - 1)

                cut_idx

            static member private toDict(arr: ('a*'b)[]) : Dict<'a,'b> =
                // assumes that 'a is unique
                let d = new Dict<'a,'b>(arr.Length)
                Array.iter (fun (a,b) ->
                    d.Add(a,b)
                ) arr
                d

            static member private inferAddressModes(r: Ranking)(dag: Depends.DAG)(config: FeatureConf)(progress: unit -> unit) : Ranking =
                let cells = dag.allCells()

                // convert ranking into map
                let rankmap = r |> Array.map (fun (pair: KeyValuePair<AST.Address,double>) -> (pair.Key,pair.Value))
                                |> ErrorModel.toDict

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
                let mutants = Array.map (fun input ->
                                if refss.[input].Length <> 0 then
                                    Some(ErrorModel.chooseLikelyAddressMode input refss.[input] dag config progress)
                                else
                                    None
                              ) crank |> Array.choose id

                // create new score and frequency tables
                let dag' = ErrorModel.mutantDAG mutants dag

                // score
                let scores = ErrorModel.runEnabledFeatures cells dag' config progress

                // count freqs
                let freqs = ErrorModel.buildFrequencyTable scores progress dag' config

                // rerank
                ErrorModel.rank cells freqs scores config

            static member private runModel(dag: Depends.DAG)(config: FeatureConf)(progress: Depends.Progress) =
                let _progf = fun () -> progress.IncrementCounter()
                let _runf = fun () -> ErrorModel.runEnabledFeatures (dag.allCells()) dag config _progf

                // get scores for each feature: featurename -> (address, score)[]
                let (scores: ScoreTable,score_time: int64) = PerfUtils.runMillis _runf ()

                // build frequency table: (featurename, selector, score) -> freq
                let _freqf = fun () -> ErrorModel.buildFrequencyTable scores _progf dag config
                let ftable,ftable_time = PerfUtils.runMillis _freqf ()

                // rank
                let _rankf = fun () -> ErrorModel.rank (dag.allCells()) ftable scores config
                let ranking,ranking_time = PerfUtils.runMillis _rankf ()

                (scores, ftable, ranking, score_time, ftable_time, ranking_time)

            static member private rank(cells: AST.Address[])(ftable: FreqTable)(scores: ScoreTable)(config: FeatureConf) : Ranking =
                let fscores = ErrorModel.makeFastScoreTable scores

                // get sums for every given cell
                // and for every enabled scope
                let addrSums: (AST.Address*int)[] =
                    Array.map (fun addr ->
                        let sum = Array.sumBy (fun sel ->
                                      ErrorModel.sumFeatureCounts addr (Scope.Selector.AllCells) ftable fscores config
                                  ) (config.EnabledScopes)
                        addr, sum
                    ) cells

                // rank by sum (smallest first)
                let rankedAddrs: (AST.Address*int)[] = Array.sortBy (fun (addr,sum) -> sum) addrSums

                // return KeyValuePairs
                Array.map (fun (addr,sum) -> new KeyValuePair<AST.Address,double>(addr,double sum)) rankedAddrs

            static member private countBuckets(ftable: FreqTable) : int =
                // get total number of non-zero buckets in the entire table
                Seq.filter (fun (elem: KeyValuePair<string*Scope.SelectID*double,int>) -> elem.Value > 0) ftable
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

            static member private genMutants(cell: AST.Address)(refs: AST.Address[])(dag: Depends.DAG)(config: FeatureConf) : Mutant[] =
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
                    // generate formulas for each AST
                    let mutants = Array.map (fun (addr,ast: AST.Expression) ->
                                    new KeyValuePair<AST.Address,string>(addr,ast.ToFormula)
                                    ) addrs_exprs

                    // get new DAG
                    let dag' = dag.CopyWithUpdatedFormulas mutants

                    // get the set of buckets
                    let mutBuckets = ErrorModel.runEnabledFeatures (
                                        Array.map (fun (kvp: KeyValuePair<AST.Address,string>) ->
                                            kvp.Key
                                        ) mutants
                                     ) dag' config ErrorModel.nop

                    // compute frequency tables
                    let mutFtable = ErrorModel.buildFrequencyTable mutBuckets ErrorModel.nop dag config
                    
                    { mutants = mutants; scores = mutBuckets; freqtable = mutFtable }
                ) fsT


            static member private chooseLikelyAddressMode(input: AST.Address)(refs: AST.Address[])(dag: Depends.DAG)(config: FeatureConf)(progress: unit -> unit) : Mutant =
                // generate all variants for the formulas that refer to this cell
                let mutants = ErrorModel.genMutants input refs dag config

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
                                        failwith "these counts must be wrong"
                                        ErrorModel.countBuckets mutant.freqtable
                                    ) mutants
                let mode_idx = ErrorModel.argmin (fun mutant ->
                                   // count histogram buckets
                                   ErrorModel.countBuckets mutant.freqtable
                               ) mutants

                if mutant_counts.[mode_idx] < def_count then
                    mutants.[mode_idx]
                else
                    { mutants = ref_fs; scores = def_buckets; freqtable = def_freq; }

            static member private runEnabledFeatures(cells: AST.Address[])(dag: Depends.DAG)(config: FeatureConf)(progress: unit -> unit) =
                config.EnabledFeatures |>
                Array.map (fun fname ->
                    // get feature lambda
                    let feature = config.FeatureByName fname

                    let fvals =
                        Array.map (fun cell ->
                            progress()
                            cell, feature cell dag
                        ) cells
                    
                    fname, fvals
                ) |> adict

            static member private buildFrequencyTable(data: ScoreTable)(incrProgress: unit -> unit)(dag: Depends.DAG)(config: FeatureConf): FreqTable =
                let d = new Dict<string*Scope.SelectID*double,int>()
                Array.iter (fun fname ->
                    Array.iter (fun (sel: Scope.Selector) ->
                        Array.iter (fun (addr: AST.Address, score: double) ->
                            let sID = sel.id addr
                            if d.ContainsKey (fname,sID,score) then
                                let freq = d.[(fname,sID,score)]
                                d.[(fname,sID,score)] <- freq + 1
                            else
                                d.Add((fname,sID,score), 1)
                            incrProgress()
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
                        if d.ContainsKey(fname,addr) then
                            let bin = d.[fname,addr]
                            let me = score
                            let arr' = arr
                            failwith "huh?"
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

            static member private nop = fun () -> ()

