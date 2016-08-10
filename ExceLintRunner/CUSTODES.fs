module CUSTODES
    open System
    open System.Diagnostics
    open System.Collections.Generic
    open CUSTODESGrammar
    open System.Runtime.InteropServices
    open System.Text
    open System.Threading

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
        let timeout = 5000

        using(new Process()) (fun (p) ->
            p.StartInfo.FileName <- @"c:\windows\system32\cmd.exe"
            p.StartInfo.Arguments <- "/c \"" + cpath + " " + String.Join(" ", args) + "\" 2>&1"
            p.StartInfo.UseShellExecute <- false
            p.StartInfo.RedirectStandardOutput <- true
            p.StartInfo.RedirectStandardError <- true

            printfn "Invoking: %s %s" (p.StartInfo.FileName) (p.StartInfo.Arguments)

            let output = new StringBuilder()
            let error = new StringBuilder()

            using(new AutoResetEvent(false)) (fun oWH ->
                using(new AutoResetEvent(false)) (fun eWH ->

                    // attach event handlers
                    p.OutputDataReceived.Add
                        (fun e ->
                            if (e.Data = null) then
                                oWH.Set() |> ignore
                            else
                                output.AppendLine(e.Data) |> ignore
                        )
                    p.ErrorDataReceived.Add
                        (fun e ->
                            if (e.Data = null) then
                                eWH.Set() |> ignore
                            else
                                error.AppendLine(e.Data) |> ignore
                        )

                    // start process
                    p.Start() |> ignore

                    // begin listening
                    p.BeginOutputReadLine()
                    p.BeginErrorReadLine()

                    // wait indefinitely for process to terminate
                    p.WaitForExit()

                    // wait on handles
                    if (oWH.WaitOne() && eWH.WaitOne()) then
                        if p.ExitCode = 0 then
                            let o = output.ToString()
                            printfn "%s" o
                            STDOUT o
                        else
                            STDERR (output.ToString())
                    else
                        STDERR (output.ToString())
                )
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

        member self.NumSmells = cOutput.Count