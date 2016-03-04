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
                ("vRelL2normsum", false);
                ("dRelL2normsum", false);
            ]
            let _config = Map.fold (fun acc key value -> Map.add key value acc) _defaults userConf

            let _features = Map.ofSeq [
                ("indegree", fun (cell)(dag) -> if _config.["indegree"] then Degree.InDegree.run cell dag else _base cell dag);
                ("combineddegree", fun (cell)(dag) -> if _config.["combineddegree"] then (Degree.InDegree.run cell dag + Degree.OutDegree.run cell dag) else _base cell dag);
                ("outdegree", fun (cell)(dag) -> if _config.["outdegree"] then Degree.OutDegree.run cell dag else _base cell dag);
                ("vRelL2normsum", fun (cell)(dag) -> if _config.["vRelL2normsum"] then Vector.FormulaRelativeL2NormSum.run cell dag else _base cell dag);
                ("dRelL2normsum", fun (cell)(dag) -> if _config.["dRelL2normsum"] then Vector.DataRelativeL2NormSum.run cell dag else _base cell dag);
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
        type ErrorModel(cdebug: FeatureConf, dag: Depends.DAG, alpha: double) =
            let config = (new FeatureConf()).enableFormulaRelativeL2NormSum().enableDataRelativeL2NormSum()

            // train model on construction
            let _data: Map<string,double[]> =
                let allCells = dag.allCells()

                config.Features |>
                Array.map (fun fname ->
                    // get feature lambda
                    let feature = config.Feature(fname)

                    // run feature on every cell
                    let fvals = Array.map (fun cell -> feature cell dag) allCells
                    fname, fvals
                ) |>
                Map.ofArray

            // a 3D frequency table: (featurename, score, count)
            let ftables: Map<string,Map<double,int>> =
                Array.fold (fun (oacc: Map<string,Map<double,int>>)(fname: string) ->
                    let imap = Array.fold (fun (iacc: Map<double,int>)(elem: double) ->
                                    if (iacc.ContainsKey(elem)) then
                                        iacc.Add(elem, iacc.[elem] + 1)
                                    else
                                        iacc.Add(elem, 1)
                                ) (Map.empty) (_data.[fname]) 
                    oacc.Add(fname, imap)
                ) (Map.empty) (config.Features)

            /// <summary>Analyzes the given cell using all of the configured features and produces a score.</summary>
            /// <param name="cell">the address of a formula cell</param>
            /// <returns>a score</returns>
            member self.score(cell: AST.Address)(fname: string) : double =
                // get feature by name
                let f = config.Feature fname

                // get feature value for this cell
                f cell dag

            member private self.rankByFeature(fname: string)(dag: Depends.DAG) : Map<AST.Address,int> =
                Array.map (fun c ->
                    let score = self.score c fname
                    let ftable = ftables.[fname]
                    c, ftable.[score]
                ) (dag.allCells()) |>
                Array.sortBy (fun (c,freq) -> freq) |>
                Array.mapi (fun i (c,s) -> c,i )
                |> Map.ofArray

            /// <summary>Ranks all the cells in the workbook by their anomalousness.</summary>
            /// <returns>an KeyValuePair<AST.Address,int>[] of (address,score) ranked from most to least anomalous</returns>
            member self.rankWithScore() : KeyValuePair<AST.Address,double>[] =
                // get the number of features
                let fsize = double(config.Features.Length)

                // get per-feature rank for every cell
                let ranks = Array.map (fun fname -> self.rankByFeature fname dag) (config.Features)

                let debug_ranks = Array.map (fun cell ->
                                    let rs = Array.map (fun (ranking: Map<AST.Address,int>) ->
                                                if (ranking.ContainsKey cell) then
                                                    ranking.[cell]
                                                else
                                                    0
                                             ) ranks
                                    cell, rs
                                  ) (dag.allCells())
                                  |> Array.sortBy (fun (cell, rs) -> double(Array.sum rs) / fsize)

                // compute and return average rank for each cell
                Array.map (fun (c: AST.Address) ->
                    let crank: int = Array.sumBy (
                                        fun (ranking: Map<AST.Address,int>) ->
                                            // TODO: FIX: this is a dirty hack
                                            // because one of my cells is missing (WTF?!)
                                            if (ranking.ContainsKey c) then
                                                ranking.[c]
                                            else
                                                0
                                     ) ranks
                    let avgrank = double(crank) / fsize
                    new KeyValuePair<AST.Address, double>(c, avgrank)
                ) (dag.allCells())