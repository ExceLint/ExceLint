module Analysis

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
        let _feature_counts =
            let allCells = dag.allCells()

            // train model
            config.Features |>
            Array.map (fun fname ->
                // get feature lambda
                let feature = config.Feature(fname)

                // count
                let map = Array.fold (fun (m: Map<int,int>) cell ->
                                let indeg = feature cell dag
                                if not (m.ContainsKey(indeg)) then
                                    m.Add(indeg, 1)
                                else
                                    m.Add(indeg, m.Item(indeg) + 1)
                            ) (new Map<int,int>([])) allCells
                fname, map
            )

        let _total = dag.allCells().Length

        // this method must be C#-friendly (no currying)
        member self.analyze() =
            let allCells = dag.allCells()

            // run analysis for each cell
            Array.map (fun cell ->
                Array.fold (fun score fname ->
                    // get feature lambda
                    let feature = config.Feature(fname)

                    // get 
//                        _feature_counts.[fname]

                    failwith "not done"
                ) 0 (config.Features)
            ) allCells

    let degreeAnalysis dag =
        let cellDegrees = Array.map (fun cell -> cell, (Degree.InDegree.run cell dag), (Degree.OutDegree.run cell dag)) (dag.allCells())
        let cellDegreesGTZero = Array.filter (fun (_, indeg, outdeg) -> (indeg + outdeg) > 0) cellDegrees
        let totalCells = System.Convert.ToDouble(cellDegreesGTZero.Length)

        // histogram for indegree
        let hist_indeg = Array.fold (fun (m: Map<int,int>) (_, indeg, _) ->
                            if not (m.ContainsKey(indeg)) then
                                m.Add(indeg, 1)
                            else
                                m.Add(indeg, m.Item(indeg) + 1)
                            ) (new Map<int,int>([])) cellDegreesGTZero
        // histogram for outdegree
        let hist_outdeg = Array.fold (fun (m: Map<int,int>) (_, _, outdeg) ->
                            if not (m.ContainsKey(outdeg)) then
                                m.Add(outdeg, 1)
                            else
                                m.Add(outdeg, m.Item(outdeg) + 1)
                            ) (new Map<int,int>([])) cellDegreesGTZero

        // histogram for combined
        let hist_combo = Array.fold (fun (m: Map<int,int>) (_, indeg, outdeg) ->
                            let combined = indeg + outdeg
                            if not (m.ContainsKey(combined)) then
                                m.Add(combined, 1)
                            else
                                m.Add(combined, m.Item(combined) + 1)
                            ) (new Map<int,int>([])) cellDegreesGTZero

        let sorted = Array.sortBy (fun (addr, indeg, outdeg) ->
                            let combo = indeg + outdeg
                            let prIndeg = hist_indeg.[indeg]
                            let prOutdeg = hist_outdeg.[outdeg]
                            let prCombo = hist_combo.[combo]
                            prIndeg + prOutdeg + prCombo
                        ) cellDegreesGTZero

        let output = sorted
                        |> Array.map (fun (addr, _, _) -> addr)
                        |> Array.mapi (fun i addr -> new KeyValuePair<AST.Address,int>(addr,i))

        output