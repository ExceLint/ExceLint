module CUSTODES
    open System
    open System.Diagnostics
    open System.Collections.Generic
    open CUSTODESGrammar
    open System.Runtime.InteropServices
    open System.Text
    open System.Threading
    open ExceLint
    open ExceLintFileFormats

    type Tool =
    | GroundTruth
    | CUSTODES
    | AmCheck
    | UCheck
    | Dimension
    | Excel
        member self.Accessor(row: CUSTODESGroundTruthRow) =
            match self with
            | GroundTruth -> row.GroundTruth
            | CUSTODES -> row.CUSTODES
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

    type Output(canonicalOutput: AST.Address[]) =
        member self.NumSmells = canonicalOutput.Length
        member self.Smells = canonicalOutput

    type OutputResult =
    | OKOutput of Output*int64
    | BadOutput of string*int64

    let runCUSTODES(spreadsheet: string)(custodesPath: string)(javaPath: string) : CUSTODESParse*int64 =
        let outputPath = IO.Path.GetTempPath()

        let runf = fun () -> runCommand (shortPath javaPath) [| "-jar"; "-d64"; "-Xms2g"; "-Xmx4g"; shortPath custodesPath; shortPath spreadsheet; shortPath outputPath; |]

        // time invocation
        let sw = new System.Diagnostics.Stopwatch()
        sw.Start()
        let output = runf()
        sw.Stop()

        let parsed = match output with
                     | STDOUT output ->
                        parse output
                     | STDERR error ->
                         match parseException error with
                         | Some (ex) -> CFailure(ex)
                         | None -> CFailure(error)
        parsed, sw.ElapsedMilliseconds

    let CUSTODESToAddress(addrstr: Address)(worksheetname: string)(workbookname: string)(path: string) : AST.Address =
        // we force the mode to absolute because
        // that's how Depends reads them
        AST.Address.FromA1StringForceMode(
            addrstr.ToUpper(),
            AST.AddressMode.Absolute,
            AST.AddressMode.Absolute,
            worksheetname,
            (if workbookname.EndsWith(".xls") then workbookname else workbookname + ".xls"),
            IO.Path.GetFullPath(path)   // ensure absolute path
        )

    let getOutput(spreadsheet: string, custodesPath: string, javaPath: string) : OutputResult =
        let absSpreadsheetPath = IO.Path.GetFullPath(spreadsheet)
        let absCustodesPath = IO.Path.GetFullPath(custodesPath)
        let absJavaPath = IO.Path.GetFullPath(javaPath)

        let workbookname = IO.Path.GetFileNameWithoutExtension(absSpreadsheetPath)
        let path = IO.Path.GetDirectoryName(absSpreadsheetPath)

        // run custodes
        let (cOutput,cTime) = runCUSTODES absSpreadsheetPath absCustodesPath absJavaPath

        match cOutput with
        | CFailure(err) -> BadOutput(err,cTime)
        | CSuccess(o) ->
            // convert to parcel addresses and flatten
            let canonicalOutput = Seq.map (fun (pair: KeyValuePair<Worksheet,Address[]>) ->
                                    let worksheetname = pair.Key
                                    Array.map (fun (addrstr: Address) ->
                                        CUSTODESToAddress addrstr worksheetname workbookname path
                                    ) pair.Value
                                  ) o
                                  |> Seq.concat |> Seq.toArray

            OKOutput (Output(canonicalOutput),cTime)

    let addresses(tool: Tool)(row: CUSTODESGroundTruthRow)(workbook_paths: Dictionary<string, string>) : AST.Address[] =
        let cells_str = tool.Accessor(row).Replace(" ", "")

        if String.IsNullOrEmpty(cells_str) then
            [||]
        else
            // get cell address array
            let cells = cells_str.Split(',')

            // convert to real address references
            Array.map (fun addrstr ->
                if String.IsNullOrEmpty(addrstr) then
                    None
                else
                    if not (workbook_paths.ContainsKey row.Spreadsheet) then
                        None
                    else
                        // get workbook absolute path
                        let path = workbook_paths.[row.Spreadsheet]

                        // get address
                        Some(CUSTODESToAddress addrstr row.Worksheet row.Spreadsheet path)
            ) cells
            |> Array.choose id

    type GroundTruth(workbook_paths: Dictionary<string,string>, gtpath: string) =
        let raw = CUSTODESGroundTruth.Load(gtpath)

        let d = new Dictionary<Tool,HashSet<AST.Address>>()

        do
            Seq.iter (fun (row: CUSTODESGroundTruthRow) ->
                Array.iter (fun tool ->
                    if not (d.ContainsKey(tool)) then
                        d.Add(tool, new HashSet<AST.Address>())

                    let cells = addresses tool row workbook_paths

                    Array.iter (fun addr ->
                        d.[tool].Add(addr) |> ignore
                    ) cells
                ) Tool.All
            ) raw.Rows

        member self.Table = d
        member self.isTrueSmell(addr: AST.Address) : bool = d.[Tool.GroundTruth].Contains(addr)
        member self.isFlaggedByExcel(addr: AST.Address) : bool = d.[Tool.Excel].Contains(addr)
        member self.differs(addr: AST.Address)(custodesFlagged: bool) : bool =
            custodesFlagged = d.[Tool.CUSTODES].Contains(addr) 
        member self.TrueSmellsbyWorkbook(workbookname: string) = new HashSet<AST.Address>(Seq.filter (fun (addr: AST.Address) -> addr.WorkbookName = workbookname) (self.Table.[Tool.GroundTruth]))
        member self.AmCheckbyWorkbook(workbookname: string) = new HashSet<AST.Address>(Seq.filter (fun (addr: AST.Address) -> addr.WorkbookName = workbookname) (self.Table.[Tool.AmCheck]))
        member self.UCheckbyWorkbook(workbookname: string) = new HashSet<AST.Address>(Seq.filter (fun (addr: AST.Address) -> addr.WorkbookName = workbookname) (self.Table.[Tool.UCheck]))
        member self.DimensionbyWorkbook(workbookname: string) = new HashSet<AST.Address>(Seq.filter (fun (addr: AST.Address) -> addr.WorkbookName = workbookname) (self.Table.[Tool.Dimension]))
        member self.ExcelbyWorkbook(workbookname: string) = new HashSet<AST.Address>(Seq.filter (fun (addr: AST.Address) -> addr.WorkbookName = workbookname) (self.Table.[Tool.Excel]))
