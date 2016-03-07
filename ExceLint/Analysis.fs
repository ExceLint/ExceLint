namespace ExceLint
    open System.Collections.Generic
    open System.Collections

    module Analysis =
        // we're using C# Dictionary instead of F# map
        // for debugging (it's inspectable) and speed purposes
        type Dict<'a,'b> = Dictionary<'a,'b>
        let adict(a: seq<('a*'b)>) = new Dict<'a,'b>(a |> dict)

        type ConfigKind =
            | Feature
            | Scope

        // a C#-friendly configuration object that is also pure/fluent
        type FeatureConf private (userConf: Map<string,(bool*ConfigKind)>) =
            let _base = Feature.BaseFeature.run 
            let _defaults = Map.ofSeq [
                ("indegree", (false, ConfigKind.Feature));
                ("combineddegree", (false, ConfigKind.Feature));
                ("outdegree", (false, ConfigKind.Feature));
                ("vRelL2normsum", (false, ConfigKind.Feature));
                ("dRelL2normsum", (false, ConfigKind.Feature));
                ("vAbsL2normsum", (false, ConfigKind.Feature));
                ("dAbsL2normsum", (false, ConfigKind.Feature));
                ("allCells", (false, ConfigKind.Scope));
                ("columns", (false, ConfigKind.Scope));
                ("rows", (false, ConfigKind.Scope));
            ]
            let _config = Map.fold (fun acc key value -> Map.add key value acc) _defaults userConf

            let _features = Map.ofSeq [
                ("indegree", fun (cell)(dag) -> if fst _config.["indegree"] then Degree.InDegree.run cell dag else _base cell dag);
                ("combineddegree", fun (cell)(dag) -> if fst _config.["combineddegree"] then (Degree.InDegree.run cell dag + Degree.OutDegree.run cell dag) else _base cell dag);
                ("outdegree", fun (cell)(dag) -> if fst _config.["outdegree"] then Degree.OutDegree.run cell dag else _base cell dag);
                ("vRelL2normsum", fun (cell)(dag) -> if fst _config.["vRelL2normsum"] then Vector.FormulaRelativeL2NormSum.run cell dag else _base cell dag);
                ("dRelL2normsum", fun (cell)(dag) -> if fst _config.["dRelL2normsum"] then Vector.DataRelativeL2NormSum.run cell dag else _base cell dag);
                ("vAbsL2normsum", fun (cell)(dag) -> if fst _config.["vAbsL2normsum"] then Vector.FormulaAbsoluteL2NormSum.run cell dag else _base cell dag);
                ("dAbsL2normsum", fun (cell)(dag) -> if fst _config.["dAbsL2normsum"] then Vector.DataAbsoluteL2NormSum.run cell dag else _base cell dag);
            ]

            new() = FeatureConf(Map.empty)

            // fluent constructors
            member self.enableInDegree() : FeatureConf =
                FeatureConf(_config.Add("indegree", (true, ConfigKind.Feature)))
            member self.enableOutDegree() : FeatureConf =
                FeatureConf(_config.Add("outdegree", (true, ConfigKind.Feature)))
            member self.enableCombinedDegree() : FeatureConf =
                FeatureConf(_config.Add("combineddegree", (true, ConfigKind.Feature)))
            member self.enableFormulaRelativeL2NormSum() : FeatureConf =
                FeatureConf(_config.Add("vRelL2normsum", (true, ConfigKind.Feature)))
            member self.enableDataRelativeL2NormSum() : FeatureConf =
                FeatureConf(_config.Add("dRelL2normsum", (true, ConfigKind.Feature)))
            member self.enableFormulaAbsoluteL2NormSum() : FeatureConf =
                FeatureConf(_config.Add("vAbsL2normsum", (true, ConfigKind.Feature)))
            member self.enableDataAbsoluteL2NormSum() : FeatureConf =
                FeatureConf(_config.Add("dAbsL2normsum", (true, ConfigKind.Feature)))
            member self.analyzeRelativeToAllCells() : FeatureConf =
                FeatureConf(_config.Add("allCells", (true, ConfigKind.Scope)))
            member self.analyzeRelativeToColumns() : FeatureConf =
                FeatureConf(_config.Add("columns", (true, ConfigKind.Scope)))
            member self.analyzeRelativeToRows() : FeatureConf =
                FeatureConf(_config.Add("rows", (true, ConfigKind.Scope)))

            // getters
            member self.EnabledFeatures
                with get(name) = _features.[name]
            member self.Features
                with get() : string[] = 
                    _config |>
                        Map.toArray |>
                        Array.choose (fun (confname,(enabled,ckind)) ->
                                        if enabled && ckind = ConfigKind.Feature then
                                            Some confname
                                        else None)
            member self.EnabledScopes
                with get() : Scope.Selector[] =
                    _config |>
                        Map.toArray |>
                        Array.choose (fun (confname,(enabled,ckind)) ->
                                        if enabled && ckind = ConfigKind.Scope then
                                            match confname with
                                            | "allCells" -> Some Scope.AllCells
                                            | "columns" -> Some Scope.SameColumn
                                            | "rows" -> Some Scope.SameRow
                                            | _ -> failwith "Unknown scope selector."
                                        else None)

        // a C#-friendly error model constructor
        type ErrorModel(config: FeatureConf, dag: Depends.DAG, alpha: double) =

            // train model on construction
            // a score table: featurename -> (address, score)
            let _data: Dict<string,(AST.Address*double)[]> =
                config.Features |>
                Array.map (fun fname ->
                    // get feature lambda
                    let feature = config.EnabledFeatures fname

                    let fvals =
                        Array.map (fun cell ->
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
                        ) (_data.[fname])
                    ) (Scope.Selector.Kinds)
                ) (config.Features)
                d

            /// <summary>Analyzes the given cell using all of the configured features and produces a score.</summary>
            /// <param name="cell">the address of a formula cell</param>
            /// <returns>a score</returns>
            member self.score(cell: AST.Address)(fname: string) : double =
                // get feature by name
                let f = config.EnabledFeatures fname

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
                            let arr = Array.zeroCreate(config.Features.Length)
                            arr.[i] <- rank
                            d.Add(addr, arr)
                    ) ranks
                ) (config.Features)
                d

            member private self.sumFeatureRanks(ranks: Dict<AST.Address,int[]>) : KeyValuePair<AST.Address,double>[] =
                Seq.map (fun (kvp: KeyValuePair<AST.Address,int[]>) ->
                    new KeyValuePair<AST.Address,double>(kvp.Key, double(Array.sum (kvp.Value)))
                ) ranks
                |> Seq.toArray
                |> Array.sortBy (fun kvp -> kvp.Value)

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

            /// <summary>Ranks all the cells in the workbook by their anomalousness.</summary>
            /// <returns>an KeyValuePair<AST.Address,int>[] of (address,score) ranked from most to least anomalous</returns>
            member self.rankWithScore() : KeyValuePair<AST.Address,double>[] =
                // get the number of features
                let fsize = double(config.Features.Length)

                // 
                // find per-feature ranks for every cell in the DAG and compute total rank
                let theRankings = Array.map (fun scope ->
                                      self.rankAllFeatures dag scope
                                  ) (config.EnabledScopes)

                let mergedRankings = self.mergeRanks theRankings dag

                self.sumFeatureRanks mergedRankings
