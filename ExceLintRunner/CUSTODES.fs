module CUSTODES
    open System
    open System.Diagnostics
    open System.Collections.Generic
    open CUSTODESGrammar
    open System.Runtime.InteropServices
    open System.Text

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
        // setup temp file
//        let tmpdir = IO.Path.GetTempPath()
//        let tempOutput = IO.Path.Combine(tmpdir, "CUSTODES_STDOUT_STDERR.txt")

//        let args' = Array.append args [| ">"; tempOutput; |]

        // setup command
        let p = new Process();
        p.StartInfo.FileName <- @"c:\windows\system32\cmd.exe";
        p.StartInfo.Arguments <- "/k \"" + cpath + " " + String.Join(" ", args) + "\""
        p.StartInfo.UseShellExecute <- false
        p.StartInfo.RedirectStandardOutput <- true

        // DEBUG
        let DEBUGRUNNING = p.StartInfo.FileName + " " +  p.StartInfo.Arguments
        printfn "INVOCATION: %s" DEBUGRUNNING

        // start command
        let successful = p.Start()

        let result =
            if successful then
                p.BeginOutputReadLine()
                let output = p.StandardError.ReadToEnd()

//                let sb = new StringBuilder()
//                while p.StandardOutput.Peek() > -1 do
//                    sb.Append(p.StandardOutput.ReadLine()) |> ignore
//
//                let output = sb.ToString()
                STDOUT output
            else
                STDERR "no"

        // wait until process terminates
        p.WaitForExit()

        result

    let runCUSTODES(spreadsheet: string)(custodesPath: string)(javaPath: string) : CUSTODESSmells =
        let outputPath = IO.Path.GetTempPath()
        let invocation = fun () -> runCommand javaPath [| "-jar"; shortPath custodesPath; shortPath spreadsheet; shortPath outputPath; |]
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
                                    AST.Address.FromString(a, worksheetname, workbookname, path)
                                ) pair.Value
                              ) cOutput
                              |> Seq.concat |> Seq.toArray

        member self.NumSmells = cOutput.Count