open COMWrapper
open ExceLint
open ExceLint.Utils
open ExceLintFileFormats
open System.Collections.Generic

[<EntryPoint>]
let main argv = 

    let config =
            try
                Args.processArgs argv
            with
            | e ->
                printfn "%A" e.Message
                System.Environment.Exit 1
                failwith "never gets called but keeps F# happy"

    let workbooks = Array.map (fun fname ->
                                  let wbname = System.IO.Path.GetFileName fname
                                  let path = System.IO.Path.GetDirectoryName fname
                                  System.IO.Path.Combine(path, wbname)
                              ) (config.files)

    using(new Application()) (fun app ->
        using(new CorpusStats(config.output_file)) (fun csv ->
            for workbook in workbooks do
                printfn "Opening: %A" workbook

                let wb = app.OpenWorkbook(workbook)
            
                printfn "Building dependence graph: %A" workbook
                let graph = wb.buildDependenceGraph()

                // get all formula addresses
                let fs = graph.getAllFormulaAddrs()

                // get all formula ASTs
                let fs_asts = fs |>
                              Array.map (fun addr -> addr, graph.getFormulaAtAddress addr) |>
                              Array.map (fun (addr,astr) -> Parcel.parseFormulaAtAddress addr astr)

                // get operator counts from ASTs
                let ops = fs_asts |>
                          Array.fold (fun (acc: Dictionary<string,int>)(ast: AST.Expression) ->
                               let strs = Parcel.operatorNamesFromExpr ast
                               for str in strs do
                                   if not (acc.ContainsKey str) then
                                       acc.Add(str, 0)
                                   acc.[str] <- acc.[str] + 1
                               acc
                          ) (new Dictionary<string,int>())
            
                // write rows to CSV, one per operator
                for pair in ops do
                    let row = new CorpusStatsRow()
                    row.Workbook <- System.IO.Path.GetFileName workbook
                    row.Operator <- pair.Key
                    row.Count <- pair.Value
                    csv.WriteRow row
        )
    )

    0 // exit normally
