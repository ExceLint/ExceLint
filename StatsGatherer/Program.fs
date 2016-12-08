open COMWrapper
open ExceLint
open ExceLint.Utils
open ExceLintFileFormats
open System.Collections.Generic
open System.Diagnostics
open System.Threading

type FDict = Dictionary<AST.Address,string>
type FCount = Dictionary<string,int>

exception TimeoutExceeded

let TIMEOUT_S = 30L

let asts(fsd: FDict)(err: ParserErrors)(ucount: int byref)(sw: Stopwatch) : AST.Expression[] =
    let mutable i: int = 0

    let fsda = Seq.toArray fsd
    
    let do_work(pair: KeyValuePair<AST.Address,string>) = 
        // check timeout
        if sw.ElapsedMilliseconds / 1000L > TIMEOUT_S then
            raise TimeoutExceeded

        let addr = pair.Key
        let astr = pair.Value
        try
            Some(Parcel.parseFormulaAtAddress addr astr)
        with
        | ex ->
            // log error as a side-effect
            let erow = new ParserErrorsRow()
            erow.Path <- addr.Path
            erow.Workbook <- addr.WorkbookName
            erow.Worksheet <- addr.WorksheetName
            erow.Address <- addr.A1Local()
            erow.Formula <- astr
            err.WriteRow erow

            // thread-safe bump count
            Interlocked.Increment(ref i) |> ignore

            // we failed; return nothing
            None

    let output =
        if fsda.Length > 100 then
            fsda
                |> Array.Parallel.map do_work
                |> Seq.choose id
                |> Seq.toArray
        else
            fsda
                |> Array.map do_work
                |> Seq.choose id
                |> Seq.toArray

    ucount <- i
    output

let ast_count(fs_asts: AST.Expression[]) : FCount =
    fs_asts |>
        Array.fold (fun (acc: Dictionary<string,int>)(ast: AST.Expression) ->
            let strs = Parcel.operatorNamesFromExpr ast
            for str in strs do
                if not (acc.ContainsKey str) then
                    acc.Add(str, 0)
                acc.[str] <- acc.[str] + 1
            acc
        ) (new Dictionary<string,int>())

let copy_workbook(workbook: string)(tmpdir: string) : string =
    match Application.MagicBytes(workbook) with
    | Application.CWFileType.XLS ->
        let newpath = System.IO.Path.Combine(
                        tmpdir,
                        System.IO.Path.GetFileName(workbook) + ".xls"
                        )
        System.IO.File.Copy(workbook, newpath)
        newpath
    | Application.CWFileType.XLSX ->
        let newpath = System.IO.Path.Combine(
                        tmpdir,
                        System.IO.Path.GetFileName(workbook) + ".xlsx"
                        )
        System.IO.File.Copy(workbook, newpath, overwrite = true)
        newpath
    | _ ->
        printfn "Not an Excel file: %A" workbook
        failwith (sprintf "Not an Excel file: %A" workbook)

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

    let workbooks = Seq.map (fun fname ->
                                  let wbname = System.IO.Path.GetFileName fname
                                  let path = System.IO.Path.GetDirectoryName fname
                                  System.IO.Path.Combine(path, wbname)
                             ) (config.files)

    let tmpdir = System.IO.Path.GetTempPath()

    using(new Application()) (fun app ->
        using(new CorpusStats(config.output_file)) (fun csv ->
            using(new ParserErrors(config.error_file)) (fun err ->
                using(new ExceptionLog(config.exception_file)) (fun exlog ->
                    for workbook in workbooks do
                    printfn "Opening: %A" workbook

                    try
                        // determine file type and copy to tmp with appropriate name
                        let workbook' = copy_workbook workbook tmpdir

                        try
                            using(app.OpenWorkbook(workbook')) (fun wb ->
                                // START TIME-CONSUMING
                                let sw = new Stopwatch()
                                sw.Start()

                                printfn "Reading workbook formulas: %A" workbook'
                                let fsd = wb.Formulas;

                                // get all formula ASTs
                                let mutable ucount = 0
                                let fs_asts = asts fsd err &ucount sw

                                // END TIME-CONSUMING
                                sw.Stop()

                                // get operator counts from ASTs
                                let ops = ast_count fs_asts
                    
                                // add unparseable formula count
                                ops.Add("unparseable",ucount)
            
                                // write rows to CSV, one per operator
                                for pair in ops do
                                    let row = new CorpusStatsRow()
                                    row.Workbook <- System.IO.Path.GetFileName workbook'
                                    row.Operator <- pair.Key
                                    row.Count <- pair.Value
                                    csv.WriteRow row
                            )
                        finally
                            System.IO.File.Delete workbook'
                    with
                    | TimeoutExceeded ->
                        let exrow = new ExceptionLogRow()
                        exrow.Workbook <- workbook
                        exrow.Error <- "Timeout"
                        exlog.WriteRow exrow

                        printfn "Analysis timeout: %A" workbook
                    | ex ->
                        let exrow = new ExceptionLogRow()
                        exrow.Workbook <- workbook
                        exrow.Error <- ex.Message
                        exlog.WriteRow exrow

                        printfn "Cannot open workbook: %A" workbook
                )
            )
        )
    )

    printfn "Press any key to continue..."
    System.Console.ReadKey() |> ignore

    0 // exit normally
