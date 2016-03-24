namespace ExceLint
    open System.Collections.Generic
    open System.Collections
    open System

    module Analysis =
        // we're using C# Dictionary instead of F# map
        // for debugging (it's inspectable) and speed purposes
        type Dict<'a,'b> = Dictionary<'a,'b>
        let adict(a: seq<('a*'b)>) = new Dict<'a,'b>(a |> dict)

        type RunnerMap = Map<string, AST.Address -> Depends.DAG -> double>

        // a C#-friendly configuration object that is also pure/fluent
        type FeatureConf private (userConf: Map<string,Feature.Capability>) =
            let _base = Feature.BaseFeature.run 

            let _capabilities : Map<string,Feature.Capability> =
                [   Degree.InDegree.capability;
                    Degree.OutDegree.capability;
                    Degree.CombinedDegree.capability;
                    Vector.DeepInputVectorRelativeL2NormSum.capability;
                    Vector.DeepOutputVectorRelativeL2NormSum.capability;
                    Vector.DeepInputVectorAbsoluteL2NormSum.capability;
                    Vector.DeepOutputVectorAbsoluteL2NormSum.capability;
                    Vector.DeepInputVectorMixedL2NormSum.capability;
                    Vector.DeepOutputVectorMixedL2NormSum.capability
                    Vector.ShallowInputVectorRelativeL2NormSum.capability;
                    Vector.ShallowOutputVectorRelativeL2NormSum.capability;
                    Vector.ShallowInputVectorAbsoluteL2NormSum.capability;
                    Vector.ShallowOutputVectorAbsoluteL2NormSum.capability;
                    Vector.ShallowInputVectorMixedL2NormSum.capability;
                    Vector.ShallowOutputVectorMixedL2NormSum.capability;
                    Proximity.Above.capability;
                    Proximity.Below.capability;
                    Proximity.Left.capability;
                    Proximity.Right.capability
                ] |> Map.ofList

            let nop(cell: AST.Address)(dag: Depends.DAG) : double = 0.0

            let _config = Map.fold (fun (acc: Map<string,Feature.Capability>)(fname: string)(cap: Feature.Capability) ->
                            let cap' : Feature.Capability =
                                {   enabled = cap.enabled;
                                    kind = cap.kind;
                                    runner = if cap.enabled then cap.runner else nop
                                }
                            Map.add fname cap acc
                          ) _capabilities userConf

            let _features : RunnerMap = Map.map (fun (fname: string)(cap: Feature.Capability) -> cap.runner) _config

            new() = FeatureConf(Map.empty)

            // fluent constructors
            member self.enableInDegree() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Degree.InDegree.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.enableOutDegree() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Degree.OutDegree.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.enableCombinedDegree() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Degree.CombinedDegree.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.enableDeepInputVectorRelativeL2NormSum() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Vector.DeepInputVectorRelativeL2NormSum.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.enableDeepOutputVectorRelativeL2NormSum() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Vector.DeepOutputVectorRelativeL2NormSum.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.enableDeepInputVectorAbsoluteL2NormSum() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Vector.DeepInputVectorAbsoluteL2NormSum.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.enableDeepOutputVectorAbsoluteL2NormSum() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Vector.DeepOutputVectorAbsoluteL2NormSum.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.enableDeepInputVectorMixedL2NormSum() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Vector.DeepInputVectorMixedL2NormSum.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.enableDeepOutputVectorMixedL2NormSum() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Vector.DeepOutputVectorMixedL2NormSum.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.enableShallowInputVectorRelativeL2NormSum() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Vector.ShallowInputVectorRelativeL2NormSum.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.enableShallowOutputVectorRelativeL2NormSum() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Vector.ShallowOutputVectorRelativeL2NormSum.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.enableShallowInputVectorAbsoluteL2NormSum() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Vector.ShallowInputVectorAbsoluteL2NormSum.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.enableShallowOutputVectorAbsoluteL2NormSum() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Vector.ShallowOutputVectorAbsoluteL2NormSum.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.enableShallowInputVectorMixedL2NormSum() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Vector.ShallowInputVectorMixedL2NormSum.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.enableShallowOutputVectorMixedL2NormSum() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Vector.ShallowOutputVectorMixedL2NormSum.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.enableProximityAbove() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Proximity.Above.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.enableProximityBelow() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Proximity.Below.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.enableProximityLeft() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Proximity.Left.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.enableProximityRight() : FeatureConf =
                FeatureConf(
                    let (name,cap) = Proximity.Right.capability
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            member self.analyzeRelativeToAllCells() : FeatureConf =
                FeatureConf(
                    let name = "ScopeAllCells"
                    let cap : Feature.Capability = { enabled = true; kind = Feature.ConfigKind.Scope; runner = nop}
                    _config.Add(name, cap)
                )
            member self.analyzeRelativeToColumns() : FeatureConf =
                FeatureConf(
                    let name = "ScopeColumns"
                    let cap : Feature.Capability = { enabled = true; kind = Feature.ConfigKind.Scope; runner = nop}
                    _config.Add(name, cap)
                )
            member self.analyzeRelativeToRows() : FeatureConf =
                FeatureConf(
                    let name = "ScopeRows"
                    let cap : Feature.Capability = { enabled = true; kind = Feature.ConfigKind.Scope; runner = nop}
                    _config.Add(name, cap)
                )

            // getters
            member self.FeatureByName
                with get(name) = _features.[name]
            member self.EnabledFeatures
                with get() : string[] = 
                    _config |>
                        Map.toArray |>
                        Array.choose (fun (fname,cap) ->
                                        if cap.enabled && cap.kind = Feature.ConfigKind.Feature then
                                            Some fname
                                        else None)
            member self.EnabledScopes
                with get() : Scope.Selector[] =
                    _config |>
                        Map.toArray |>
                        Array.choose (fun (confname,cap) ->
                                        if cap.enabled && cap.kind = Feature.ConfigKind.Scope then
                                            match confname with
                                            | "ScopeAllCells" -> Some Scope.AllCells
                                            | "ScopeColumns" -> Some Scope.SameColumn
                                            | "ScopeRows" -> Some Scope.SameRow
                                            | _ -> failwith "Unknown scope selector."
                                        else None)

        // a C#-friendly error model constructor
        type ErrorModel(config: FeatureConf, dag: Depends.DAG, alpha: double, progress: Depends.Progress) =

            // train model on construction
            // a score table: featurename -> (address, score)
            let _data: Dict<string,(AST.Address*double)[]> =
                config.EnabledFeatures |>
                Array.map (fun fname ->
                    // get feature lambda
                    let feature = config.FeatureByName fname

                    let fvals =
                        Array.map (fun cell ->
                            progress.IncrementCounter()
                            cell, feature cell dag
                        ) (dag.allCells())
                    
                    fname, fvals
                ) |> adict

            // a frequency table: (featurename, selector, score) -> freq
            let ftable: Dict<string*Scope.SelectID*double,int> =
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
                            progress.IncrementCounter()
                        ) (_data.[fname])
                    ) (Scope.Selector.Kinds)
                ) (config.EnabledFeatures)
                d

            /// <summary>Analyzes the given cell using all of the configured features and produces a score.</summary>
            /// <param name="cell">the address of a formula cell</param>
            /// <returns>a score</returns>
            member self.score(cell: AST.Address)(fname: string) : double =
                // get feature by name
                let f = config.FeatureByName fname

                // get feature value for this cell
                f cell dag

            member private self.rankByFeatureAndSelector(fname: string)(dag: Depends.DAG)(sel: Scope.Selector) : (AST.Address*int)[] =
                // sort by least common
                let sorted = Array.map (fun addr ->
                                 let sID = sel.id addr
                                 let score = self.score addr fname
                                 let freq = ftable.[(fname,sID,score)]
                                 addr, freq
                             ) (dag.allCells()) |>
                             Array.sortBy (fun (addr,freq) -> freq)
                // return the address and its rank
                sorted |> Array.mapi (fun i (addr,freq) -> addr,i )

            member private self.rankAllFeatures(dag: Depends.DAG)(sel: Scope.Selector) : Dict<AST.Address,int[]> =
                // get per-feature rank for every cell
                let d = new Dict<AST.Address,int[]>()

                Array.iteri (fun i fname ->
                    let ranks = self.rankByFeatureAndSelector fname dag sel
                    Array.iter (fun (addr,rank) ->
                        if d.ContainsKey addr then
                            let arr = d.[addr]
                            arr.[i] <- rank
                            d.[addr] <- arr
                        else
                            let arr = Array.zeroCreate(config.EnabledFeatures.Length)
                            arr.[i] <- rank
                            d.Add(addr, arr)
                    ) ranks
                ) (config.EnabledFeatures)
                d

            member private self.sumFeatureRanksAndSort(ranks: Dict<AST.Address,int[]>) : KeyValuePair<AST.Address,double>[] =
                Seq.map (fun (kvp: KeyValuePair<AST.Address,int[]>) ->
                    new KeyValuePair<AST.Address,double>(kvp.Key, double(Array.sum (kvp.Value)))
                ) ranks
                |> Seq.toArray
                |> Array.sortBy (fun kvp -> kvp.Value)

            // sum the count of the appropriate feature bin of every feature
            // for the given address
            member private self.sumFeatureCounts(addr: AST.Address)(sel: Scope.Selector) : int =
                Array.sumBy (fun fname -> 
                    // get feature
                    let feature = config.FeatureByName fname
                    // get selector ID
                    let sID = sel.id addr
                    // get feature score
                    let fscore = feature addr dag
                    // get score count
                    ftable.[(fname,sID,fscore)]
                ) (config.EnabledFeatures)

            member self.rankByFeatureSum() : KeyValuePair<AST.Address,double>[] =
                // get sums for every address
                // and for every enabled scope
                let addrSums: (AST.Address*int)[] =
                    Array.map (fun addr ->
                        let sum = Array.sumBy (fun sel ->
                                        self.sumFeatureCounts addr Scope.Selector.AllCells
                                  ) (config.EnabledScopes)
                        addr, sum
                    ) (dag.allCells())

                // rank by sum (smallest first)
                let rankedAddrs: (AST.Address*int)[] = Array.sortBy (fun (addr,sum) -> sum) addrSums

                // return KeyValuePairs
                let totalOrder = Array.map (fun (addr,sum) -> new KeyValuePair<AST.Address,double>(addr,double sum)) rankedAddrs

                let winners = self.cutRankBySignificance totalOrder

                winners

            member private self.mergeRanks(multiRanks: Dict<AST.Address,int[]>[])(dag: Depends.DAG) : Dict<AST.Address,int[]> =
                let d = new Dict<AST.Address, int[]>()

                Array.iter (fun addr ->
                    let allRanks = 
                        Array.fold (fun (acc: int list)(ranking: Dict<AST.Address,int[]>) ->
                            (ranking.[addr] |> List.ofArray) @ acc
                        ) ([]) multiRanks
                    d.Add(addr, allRanks |> List.toArray)
                ) (dag.allCells())

                d

            /// <summary>
            /// Finds the maxmimum significance height
            /// </summary>
            /// <returns>maxmimum significance height count (an int)</returns>
            member self.getSignificanceCutoff : int =
                // round to integer
                int (
                    // get total number of counts
                    double (dag.allCells().Length * config.EnabledFeatures.Length * config.EnabledScopes.Length)
                    // times signficance
                    * alpha
                )

            member private self.cutRankBySignificance(ranking: KeyValuePair<AST.Address,double>[]): KeyValuePair<AST.Address,double>[] =
                let cutoff = self.getSignificanceCutoff

                let cutrank = Array.fold (fun (acc: KeyValuePair<AST.Address,double> list)(score: KeyValuePair<AST.Address,double>) ->
                                    if score.Value > double cutoff then
                                        acc
                                    else
                                        score :: acc
                                ) (List.empty) ranking |> List.rev |> List.toArray

                let rank_nums = Array.map (fun (kvp: KeyValuePair<AST.Address,double>) -> int(kvp.Value)) cutrank

                let dderiv_idx = self.dderiv(rank_nums)

                cutrank.[0..dderiv_idx]

            /// <summary>Ranks all the cells in the workbook by their anomalousness.</summary>
            /// <returns>an KeyValuePair<AST.Address,int>[] of (address,score) ranked from most to least anomalous</returns>
            member self.rankWithScore() : KeyValuePair<AST.Address,double>[] =
                // get the number of features
                let fsize = double(config.EnabledFeatures.Length)

                // find per-feature ranks for every cell in the DAG and compute total rank
                let theRankings = Array.map (fun scope ->
                                      self.rankAllFeatures dag scope
                                  ) (config.EnabledScopes)

                let mergedRankings = self.mergeRanks theRankings dag

                self.sumFeatureRanksAndSort mergedRankings

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
                ) _data

                let debug = Seq.map (fun (kvp: KeyValuePair<AST.Address,(string*double) list>) ->
                                        let addr2: AST.Address = kvp.Key
                                        let scores: (string*double)[] = kvp.Value |> List.toArray

                                        new KeyValuePair<AST.Address,(string*double)[]>(addr2, scores)
                                     ) d

                debug |> Seq.toArray

            (*
            import fileinput

            def dderiv(y):
               anglemin = 1
               angleminindex = 0
               ymin = 0
               for index in range(0,len(y)-2):
                   angle = y[index] + y[index+2] - 2 * y[index+1]
                   # print (y[index],y[index+1],y[index+2])
                   # assume evenly spaced
                   # print angle
                   if (angle < anglemin):
                       anglemin = angle
                       angleminindex = index
                       ymin = y[index]
               return (angleminindex+1, ymin)

            list = []

            for line in fileinput.input():
               line = line.rstrip('\r\n')
               val = float(line)
               list.append(val)
               pass

            print dderiv(list)
            *)

            member self.dderiv(y: int[]) : int =
                let mutable anglemin = 1
                let mutable angleminindex = 0
                for index in 0..(y.Length - 3) do
                    let angle = y.[index] + y.[index + 2] - 2 * y.[index + 1]
                    if angle < anglemin then
                        anglemin <- angle
                        angleminindex <- index
                angleminindex