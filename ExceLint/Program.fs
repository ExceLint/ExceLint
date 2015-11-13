namespace ExceLint
    open COMWrapper
    open System.IO

    module Analysis =

        [<EntryPoint>]
        let main argv = 
            let filename = @"..\..\..\ExceLintTests\TestData\Public_debt-ratios_advanced-SummarySheetOnly.xlsx"
            let output = "output.csv"

            let outFile = new StreamWriter(output)

            printfn "Starting Excel..."
            let app = new COMWrapper.Application()

            printfn "Opening workbook..."
            let workbook = app.OpenWorkbook(filename)

            printfn "Building dependence graph..."
            let dag = workbook.buildDependenceGraph()

            printfn "Computing indegree and outdegree..."
            let allCells = dag.allCells()

            let cellDegrees = Array.map (fun cell -> cell, (Degree.getIndegreeForCell cell dag), (Degree.getOutdegreeForCell cell dag)) allCells

            printfn "Writing output..."
            Array.sortBy (fun (_, indeg, outdeg) -> indeg + outdeg) cellDegrees |>
            Array.map (fun (addr: AST.Address, indeg: int, outdeg: int) ->
                let a1_addr = addr.A1Local()
                let row = sprintf "\"%s\",%i,%i" a1_addr indeg outdeg
                outFile.WriteLine(row)
            )
            |> ignore

            outFile.Close()

            printfn "Done."
            0
