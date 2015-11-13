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
            Array.map (fun (addr: AST.Address, indeg: int, outdeg: int) ->
                let row = sprintf "%s,%i,%i" (addr.ToString()) indeg outdeg
                outFile.WriteLine(row)
            ) cellDegrees |> ignore

            outFile.Close()

            printfn "Done."
            0
