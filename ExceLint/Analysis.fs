namespace ExceLint
    module Analysis =

        // a C#-friendly configuration object that is also pure/fluent
        type FeatureConf private (userConf: Map<string,bool>) =
            let _base = Feature.BaseFeature.run 
            let _defaults = Map.ofSeq [
                ("indegree", false);
                ("combineddegree", false);
                ("outdegree", false);
            ]
            let _config = Map.fold (fun acc key value -> Map.add key value acc) _defaults userConf

            let _features = Map.ofSeq [
                ("indegree", fun (cell)(dag) -> if _config.["indegree"] then Degree.InDegree.run cell dag else _base cell dag);
                ("combineddegree", fun (cell)(dag) -> if _config.["combineddegree"] then (Degree.InDegree.run cell dag + Degree.OutDegree.run cell dag) else _base cell dag);
                ("outdegree", fun (cell)(dag) -> if _config.["outdegree"] then Degree.OutDegree.run cell dag else _base cell dag);
            ]

            new() = FeatureConf(Map.empty)

            // fluent constructors
            member self.enableInDegree() : FeatureConf =
                FeatureConf(_config.Add("indegree", true))
            member self.enableOutDegree() : FeatureConf =
                FeatureConf(_config.Add("outdegree", true))
            member self.enableCombinedDegree() : FeatureConf =
                FeatureConf(_config.Add("combineddegree", true))

            // getters
            member self.Feature
                with get(name) = _features.[name]
            member self.Features
                with get() = _features |> Map.toArray |> Array.map fst

        // a C#-friendly error model constructor
        type ErrorModel(config: FeatureConf, dag: Depends.DAG) =
            let _feature_data =
                let allCells = dag.allCells()

                // train model
                config.Features |>
                Array.map (fun fname ->
                    // get feature lambda
                    let feature = config.Feature(fname)

                    // run feature on every cell
                    let fvals = Array.map (fun cell -> feature cell dag) allCells
                    fname, fvals
                )

            let _total = dag.allCells().Length

            // this method must be C#-friendly (no currying)
            member self.analyze(cell: AST.Address) =
                failwith "not implemented"