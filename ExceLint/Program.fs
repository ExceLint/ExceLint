namespace ExceLint
    open COMWrapper
    open System.IO
    open System.Collections.Generic

    module Analysis =
        // a C#-friendly configuration object that is also pure/fluent
        type FeatureConf private (userConf: Map<string,bool>) =
            let _defaults = Map.ofSeq [
                ("indegree", false);
                ("combineddegree", false);
                ("outdegree", false);
            ]
            let _config = Map.fold (fun acc key value -> Map.add key value acc) _defaults userConf

            let _features = Map.ofSeq [
                ("indegree", fun (cell)(dag) -> if _config.["indegree"] then (Degree.getIndegreeForCell(cell)(dag)) else 0);
                ("combineddegree", fun (cell)(dag) -> if _config.["combineddegree"] then (Degree.getIndegreeForCell(cell)(dag) + Degree.getOutdegreeForCell(cell)(dag)) else 0);
                ("outdegree", fun (cell)(dag) -> if _config.["outdegree"] then (Degree.getOutdegreeForCell(cell)(dag)) else 0);
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
            let cellDegrees = Array.map (fun cell -> cell, (Degree.getIndegreeForCell cell dag), (Degree.getOutdegreeForCell cell dag)) (dag.allCells())
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

        [<EntryPoint>]
        let main argv = 
            let INPUTFILE = @"..\..\..\ExceLintTests\TestData\Public_debt-ratios_advanced-SummarySheetOnly.xlsx"
            let OUTPUTFILE = "output.csv"
            let ALPHA = 0.05

            let outFile = new StreamWriter(OUTPUTFILE)

            printfn "Starting Excel..."
            let app = new COMWrapper.Application()

            printfn "Opening workbook..."
            let workbook = app.OpenWorkbook(INPUTFILE)

            printfn "Building dependence graph..."
            let dag = workbook.buildDependenceGraph()

            printfn "Computing indegree and outdegree..."
            let cellDegrees = Array.map (fun cell -> cell, (Degree.getIndegreeForCell cell dag), (Degree.getOutdegreeForCell cell dag)) (dag.allCells())
            let cellDegreesGTZero = Array.filter (fun (_, indeg, outdeg) -> (indeg + outdeg) > 0) cellDegrees
            let totalCells = System.Convert.ToDouble(cellDegreesGTZero.Length)

            printfn "Computing probabilities..."
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
                                - (prIndeg * prOutdeg * prCombo)
                            ) cellDegreesGTZero
                            |> Array.map (fun (addr, _, _) -> addr)
                            |> Array.mapi (fun i addr -> new KeyValuePair<AST.Address,int>(addr,i))

            printfn "Writing output..."
            let headers = [| "\"cell\",\"indegree\",\"outdegree\",\"combined\",\"pr_indeg\",\"pr_outdeg\",\"pr_combined\",\"indeg_anom\",\"outdeg_anom\",\"combo_anom\",\"tot_anom\"" |]
            let output = Array.sortBy (fun (_, indeg, outdeg) -> - (indeg + outdeg)) cellDegreesGTZero |>
                         Array.map (fun (addr: AST.Address, indeg: int, outdeg: int) ->
                            let a1_addr = addr.A1Local()
                            let combined = indeg + outdeg
                            let indeg_prob = System.Convert.ToDouble(hist_indeg.Item(indeg)) / totalCells
                            let outdeg_prob = System.Convert.ToDouble(hist_outdeg.Item(outdeg)) / totalCells
                            let combo_prob = System.Convert.ToDouble(hist_combo.Item(combined)) / totalCells
                            let indeg_anom = indeg_prob < ALPHA
                            let outdeg_anom = outdeg_prob < ALPHA
                            let combo_anom = combo_prob < ALPHA
                            let total_anom = (if indeg_anom then 1 else 0) + (if outdeg_anom then 1 else 0) + (if combo_anom then 1 else 0)
                            sprintf "\"%s\",%i,%i,%i,%f,%f,%f,%b,%b,%b,%i" a1_addr indeg outdeg combined indeg_prob outdeg_prob combo_prob indeg_anom outdeg_anom combo_anom total_anom
                         )
            Array.map (fun (row: string) -> outFile.WriteLine(row)) (Array.append headers output) |> ignore

            outFile.Close()

            printfn "Done."
            0
