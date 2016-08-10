module CUSTODES
    open System
    open System.Diagnostics
    open System.Collections.Generic
    open CUSTODESGrammar
    open System.Runtime.InteropServices
    open System.Text
    open System.Threading

    type Tool =
    | GroundTruth
    | CUSTODES
    | AmCheck
    | UCheck
    | Dimension
    | Excel
        member self.Accessor(row: CSV.CUSTODESGroundTruth.Row) =
            match self with
            | GroundTruth -> row.GroundTruth
            | CUSTODES -> row.Custodes
            | AmCheck -> row.AmCheck
            | UCheck -> row.UCheck
            | Dimension -> row.Dimension
            | Excel -> row.Excel
        static member All = [| GroundTruth; CUSTODES; AmCheck; UCheck; Dimension; Excel |]

    [<DllImport("kernel32.dll", CharSet = CharSet.Auto)>]
    extern int GetShortPathName(
        [<MarshalAs(UnmanagedType.LPTStr)>]
        string path,
        [<MarshalAs(UnmanagedType.LPTStr)>]
        StringBuilder shortPath,
        int shortPathLength
    )

    let shortPath(p: string) : string =
        let sb = new StringBuilder(255)
        GetShortPathName(p, sb, sb.Capacity) |> ignore
        sb.ToString()

    type ShellResult =
    | STDOUT of string
    | STDERR of string

    let private runCommand(cpath: string)(args: string[]) : ShellResult =
        using(new Process()) (fun (p) ->
            p.StartInfo.FileName <- @"c:\windows\system32\cmd.exe"
            p.StartInfo.Arguments <- "/c \"" + cpath + " " + String.Join(" ", args) + "\" 2>&1"
            p.StartInfo.UseShellExecute <- false
            p.StartInfo.RedirectStandardOutput <- true
            p.StartInfo.RedirectStandardError <- true

            let output = new StringBuilder()

            using(new AutoResetEvent(false)) (fun oWH ->

                    // attach event handler
                    p.OutputDataReceived.Add
                        (fun e ->
                            if (e.Data = null) then
                                oWH.Set() |> ignore
                            else
                                output.AppendLine(e.Data) |> ignore
                        )

                    // start process
                    p.Start() |> ignore

                    // begin listening
                    p.BeginOutputReadLine()
                    p.BeginErrorReadLine()

                    // wait indefinitely for process to terminate
                    p.WaitForExit()

                    // wait on handle
                    oWH.WaitOne() |> ignore

                    if p.ExitCode = 0 then
                        STDOUT (output.ToString())
                    else
                        STDERR (output.ToString())
            )
        )

    let runCUSTODES(spreadsheet: string)(custodesPath: string)(javaPath: string) : CUSTODESSmells =
        let outputPath = IO.Path.GetTempPath()

        let invocation = fun () -> runCommand (shortPath javaPath) [| "-jar"; shortPath custodesPath; shortPath spreadsheet; shortPath outputPath; |]
        match invocation() with
        | STDOUT output ->
            // debug
            printfn "DEBUG:\n\n%A" output
            parse output
        | STDERR error -> failwith error

    type Output(spreadsheet: string, custodesPath: string, javaPath: string) =
        let absSpreadsheetPath = IO.Path.GetFullPath(spreadsheet)
        let absCustodesPath = IO.Path.GetFullPath(custodesPath)
        let absJavaPath = IO.Path.GetFullPath(javaPath)

        let workbookname = IO.Path.GetFileNameWithoutExtension(absSpreadsheetPath)
        let path = IO.Path.GetDirectoryName(absSpreadsheetPath)

        // run custodes
        let cOutput = runCUSTODES absSpreadsheetPath absCustodesPath absJavaPath

        // convert to parcel addresses and flatten
        let canonicalOutput = Seq.map (fun (pair: KeyValuePair<Worksheet,Address[]>) ->
                                let worksheetname = pair.Key
                                Array.map (fun (a: Address) ->
                                    AST.Address.FromA1String(a, worksheetname, workbookname, path)
                                ) pair.Value
                              ) cOutput
                              |> Seq.concat |> Seq.toArray

        // this will remove duplicates, if there are any
        let canonicalOutputHS = new HashSet<AST.Address>(canonicalOutput)

        member self.NumSmells = canonicalOutputHS.Count
        member self.Smells = canonicalOutputHS

    let addresses(tool: Tool)(row: CSV.CUSTODESGroundTruth.Row)(path: string) : AST.Address[] =
        // get cell address array
        let cells = tool.Accessor(row).Replace(" ", "").Split(',')

        // convert to real address references
        Array.map (fun straddr ->
            AST.Address.FromA1String(
                straddr,
                row.Worksheet,
                row.Spreadsheet,
                path
            )
        ) cells

    type GroundTruth(folderPath: string) =
        let raw = CSV.CUSTODESGroundTruth.Load(CSV.CUSTODESGroundTruthPath)

        let d = new Dictionary<Tool,HashSet<AST.Address>>()

        do
            Seq.iter (fun (row: CSV.CUSTODESGroundTruth.Row) ->
                Array.iter (fun tool ->
                    if not (d.ContainsKey(tool)) then
                        d.Add(CUSTODES, new HashSet<AST.Address>())

                    let cells = addresses tool row folderPath

                    Array.iter (fun addr ->
                        d.[CUSTODES].Add(addr) |> ignore
                    ) cells
                ) Tool.All
            ) raw.Rows

        member self.Table = d
