namespace ExceLint
    open System.Collections.Generic

    module Analysis =

        // a C#-friendly configuration object that is also pure/fluent
        type FeatureConf private (userConf: Map<string,bool>) =
            let _base = Feature.BaseFeature.run 
            let _defaults = Map.ofSeq [
                ("indegree", false);
                ("combineddegree", false);
                ("outdegree", false);
                ("relL2normsum", false);
            ]
            let _config = Map.fold (fun acc key value -> Map.add key value acc) _defaults userConf

            let _features = Map.ofSeq [
                ("indegree", fun (cell)(dag) -> if _config.["indegree"] then Degree.InDegree.run cell dag else _base cell dag);
                ("combineddegree", fun (cell)(dag) -> if _config.["combineddegree"] then (Degree.InDegree.run cell dag + Degree.OutDegree.run cell dag) else _base cell dag);
                ("outdegree", fun (cell)(dag) -> if _config.["outdegree"] then Degree.OutDegree.run cell dag else _base cell dag);
                ("relL2normsum", fun (cell)(dag) -> if _config.["relL2normsum"] then Vector.FormulaRelativeL2NormSum.run cell dag else _base cell dag);
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
                FeatureConf(_config.Add("relL2normsum", true))

            // getters
            member self.Feature
                with get(name) = _features.[name]
            member self.Features
                with get() = _features |> Map.toArray |> Array.map fst

        // a C#-friendly error model constructor
        type ErrorModel(cdebug: FeatureConf, dag: Depends.DAG, alpha: double) =
            let config = (new FeatureConf()).enableFormulaRelativeL2NormSum()

            // train model on construction
            let _data =
                let allFormulaCells = dag.getAllFormulaAddrs()

                config.Features |>
                Array.map (fun fname ->
                    // get feature lambda
                    let feature = config.Feature(fname)

                    // run feature on every cell
                    let fvals = Array.map (fun cell -> feature cell dag) allFormulaCells
                    fname, fvals
                ) |>
                Map.ofArray

            // TODO: DEBUGGING: HARDCODED
            let dataarr = _data.["relL2normsum"] |> Array.sort

            let ftable = Array.fold (fun (acc: Map<double,int>)(elem: double) ->
                            if (acc.ContainsKey(elem)) then
                                acc.Add(elem, acc.[elem] + 1)
                            else
                                acc.Add(elem, 1)
                            ) (Map.empty) dataarr 
//                            |> Map.toArray |> Array.sortBy (fun (key,value) -> value)



            /// <summary>Analyzes the given cell using all of the configured features and produces a score.</summary>
            /// <param name="cell">the address of a formula cell</param>
            /// <returns>a score</returns>
            member self.score(cell: AST.Address) : double =
                

                // get feature scores
                let fs = Array.map (fun fname ->
                            // get feature lambda
                            let f = config.Feature fname

                            // get feature value for this cell
                            let t = f cell dag

                            t

                            // determine probability
//                            let p = BasicStats.cdf t _data.[fname]

                            // do two-tailed test
//                            if p > (1.0 - alpha) then 1.0 else 0.0
                         ) (config.Features)

                // combine scores
                Array.sum fs

            /// <summary>Ranks all the cells in the workbook by their anomalousness.</summary>
            /// <returns>an AST.Address[] ranked from most to least anomalous</returns>
            member self.rank() : AST.Address[] =
                // rank by analysis score (rev to sort from high to low)
                let ranks = Array.sortBy (fun addr ->
                                let score = self.score(addr)
                                let freq = ftable.[score]
                                freq
                            ) (dag.getAllFormulaAddrs())

                ranks

            /// <summary>Ranks all the cells in the workbook by their anomalousness.</summary>
            /// <returns>an KeyValuePair<AST.Address,int>[] of (address,score) ranked from most to least anomalous</returns>
            member self.rankWithScore() : KeyValuePair<AST.Address,int>[] =
                // get all cells
                let output = dag.getAllFormulaAddrs() |>

                    // get scores
                    Array.map (fun c -> new KeyValuePair<AST.Address, int>(c, ftable.[self.score c])) |>

                    // rank
                    Array.sortBy (fun (pair: KeyValuePair<AST.Address, int>) -> pair.Value)

                output