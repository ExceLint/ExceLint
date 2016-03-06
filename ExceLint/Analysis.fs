namespace ExceLint
    open System.Collections.Generic
    open System.Collections

    module Analysis =

        // we're using C# Dictionary instead of F# map
        // for debugging (it's inspectable) and speed purposes
        type Dict<'a,'b> = Dictionary<'a,'b>
        let adict(a: seq<('a*'b)>) = new Dict<'a,'b>(a |> dict)

        // a C#-friendly configuration object that is also pure/fluent
        type FeatureConf private (userConf: Map<string,bool>) =
            let _base = Feature.BaseFeature.run 
            let _defaults = Map.ofSeq [
                ("indegree", false);
                ("combineddegree", false);
                ("outdegree", false);
                ("vRelL2normsum", false);
                ("dRelL2normsum", false);
                ("vAbsL2normsum", false);
                ("dAbsL2normsum", false);
            ]
            let _config = Map.fold (fun acc key value -> Map.add key value acc) _defaults userConf

            let _features = Map.ofSeq [
                ("indegree", fun (cell)(dag) -> if _config.["indegree"] then Degree.InDegree.run cell dag else _base cell dag);
                ("combineddegree", fun (cell)(dag) -> if _config.["combineddegree"] then (Degree.InDegree.run cell dag + Degree.OutDegree.run cell dag) else _base cell dag);
                ("outdegree", fun (cell)(dag) -> if _config.["outdegree"] then Degree.OutDegree.run cell dag else _base cell dag);
                ("vRelL2normsum", fun (cell)(dag) -> if _config.["vRelL2normsum"] then Vector.FormulaRelativeL2NormSum.run cell dag else _base cell dag);
                ("dRelL2normsum", fun (cell)(dag) -> if _config.["dRelL2normsum"] then Vector.DataRelativeL2NormSum.run cell dag else _base cell dag);
                ("vAbsL2normsum", fun (cell)(dag) -> if _config.["vAbsL2normsum"] then Vector.FormulaAbsoluteL2NormSum.run cell dag else _base cell dag);
                ("dAbsL2normsum", fun (cell)(dag) -> if _config.["dAbsL2normsum"] then Vector.DataAbsoluteL2NormSum.run cell dag else _base cell dag);
            ]

            new() = FeatureConf(Map.empty)

            // fluent constructors
            member self.enableInDegree() : FeatureConf =
                FeatureConf(_config.Add("indegree", true))
            member self.enableOutDegree() : FeatureConf =
                FeatureConf(_config.Add("outdegree", true))
            member self.enableCombinedDegree() : FeatureConf =
                FeatureConf(_config.Add("combineddegree", true))
            member self.enableFormulaRelativeL2NormSum() : FeatureConf =
                FeatureConf(_config.Add("vRelL2normsum", true))
            member self.enableDataRelativeL2NormSum() : FeatureConf =
                FeatureConf(_config.Add("dRelL2normsum", true))
            member self.enableFormulaAbsoluteL2NormSum() : FeatureConf =
                FeatureConf(_config.Add("vAbsL2normsum", true))
            member self.enableDataAbsoluteL2NormSum() : FeatureConf =
                FeatureConf(_config.Add("dAbsL2normsum", true))

            // getters
            member self.Feature
                with get(name) = _features.[name]
            member self.Features
                with get() : string[] = 
                    _config |>
                        Map.toArray |>
                        Array.choose (fun (feature,enabled) ->
                                        if enabled then
                                            Some feature
                                        else None)

        // a C#-friendly error model constructor
        type ErrorModel(config: FeatureConf, dag: Depends.DAG, alpha: double) =

            // train model on construction
            let _data: Dict<string,double[]> =
                config.Features |>
                Array.map (fun fname ->
                    // get feature lambda
                    let feature = config.Feature fname

                    // run feature on every cell
                    let fvals = dag.allCells() |>
                                Array.map (fun cell -> feature cell dag) 
                    fname, fvals
                ) |> adict

            // a 3D frequency table: (featurename, score) -> freq
            let ftable: Dict<string*double,int> =
                let d = new Dict<string*double,int>()
                Array.iter (fun fname ->
                    Array.iter (fun score ->
                        if d.ContainsKey (fname,score) then
                            let freq = d.[(fname,score)]
                            d.[(fname,score)] <- freq + 1
                        else
                            d.Add((fname,score), 1)
                    ) (_data.[fname])
                ) (config.Features)
                d

            /// <summary>Analyzes the given cell using all of the configured features and produces a score.</summary>
            /// <param name="cell">the address of a formula cell</param>
            /// <returns>a score</returns>
            member self.score(cell: AST.Address)(fname: string) : double =
                // get feature by name
                let f = config.Feature fname

                // get feature value for this cell
                f cell dag

            member private self.rankByFeature(fname: string)(dag: Depends.DAG) : (AST.Address*int)[] =
                // sort by least common
                let sorted = Array.map (fun addr ->
                                 let score = self.score addr fname
                                 let freq = ftable.[(fname,score)]
                                 addr, freq
                             ) (dag.allCells()) |>
                             Array.sortBy (fun (addr,freq) -> freq)
                // return the address and its rank
                sorted |> Array.mapi (fun i (addr,freq) -> addr,i )

            member private self.rankAllFeatures(dag: Depends.DAG) : Dict<AST.Address,int[]> =
                // get per-feature rank for every cell
                let d = new Dict<AST.Address,int[]>()

                Array.iteri (fun i fname ->
                    let ranks = self.rankByFeature fname dag
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

            /// <summary>Ranks all the cells in the workbook by their anomalousness.</summary>
            /// <returns>an KeyValuePair<AST.Address,int>[] of (address,score) ranked from most to least anomalous</returns>
            member self.rankWithScore() : KeyValuePair<AST.Address,double>[] =
                // get the number of features
                let fsize = double(config.Features.Length)

                // find per-feature ranks for every cell in the DAG and compute total rank
                self.sumFeatureRanks (self.rankAllFeatures dag)
