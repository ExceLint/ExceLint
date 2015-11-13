namespace ExceLint
    open COMWrapper
    open System.IO

    module Analysis =

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
            let allCells = dag.allCells()

            let cellDegrees = Array.map (fun cell -> cell, (Degree.getIndegreeForCell cell dag), (Degree.getOutdegreeForCell cell dag)) allCells
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

            printfn "Writing output..."
            let headers = [| "\"cell\",\"indegree\",\"outdegree\",\"combined\",\"pr_indeg\",\"pr_outdeg\",\"pr_combined\",\"indeg_anom\",\"outdeg_anom\",\"combo_anom\"" |]
            let output = Array.sortBy (fun (_, indeg, outdeg) -> - (indeg + outdeg)) cellDegrees |>
                         Array.map (fun (addr: AST.Address, indeg: int, outdeg: int) ->
                            let a1_addr = addr.A1Local()
                            let combined = indeg + outdeg
                            let indeg_prob = System.Convert.ToDouble(hist_indeg.Item(indeg)) / totalCells
                            let outdeg_prob = System.Convert.ToDouble(hist_outdeg.Item(outdeg)) / totalCells
                            let combo_prob = System.Convert.ToDouble(hist_combo.Item(combined)) / totalCells
                            let indeg_anom = indeg_prob < ALPHA
                            let outdeg_anom = outdeg_prob < ALPHA
                            let combo_anom = combo_prob < ALPHA
                            sprintf "\"%s\",%i,%i,%i,%f,%f,%f,%b,%b,%b" a1_addr indeg outdeg combined indeg_prob outdeg_prob combo_prob indeg_anom outdeg_anom combo_anom
                         )
            Array.map (fun (row: string) -> outFile.WriteLine(row)) (Array.append headers output) |> ignore

            outFile.Close()

            printfn "Done."
            0
