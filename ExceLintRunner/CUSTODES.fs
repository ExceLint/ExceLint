module CUSTODES
    open System
    open System.Diagnostics
    open System.Collections.Generic

    type ShellResult =
    | STDOUT of string
    | STDERR of string

    let private runCommand(cpath: string)(args: string[]) : ShellResult =
        // setup command
        let p = new Process()
        p.StartInfo.FileName <- cpath
        p.StartInfo.Arguments <- String.Join(" ", args)
        p.StartInfo.UseShellExecute <- false
        p.StartInfo.RedirectStandardOutput <- true
        p.StartInfo.RedirectStandardError <- true

        // start command
        p.Start() |> ignore

        // capture output
        let output = p.StandardOutput.ReadToEnd()
        let err = p.StandardError.ReadToEnd()

        // wait until process terminates
        p.WaitForExit()

        // return STDOUT if everything is cool, STDERR otherwise
        if not (String.IsNullOrEmpty(err)) then
            STDOUT output
        else
            STDERR err

    let runCUSTODES(spreadsheet: string)(custodesPath: string)(javaPath: string) : CUSTODESGrammar.CUSTODESSmells =
        let outputPath = IO.Path.GetTempPath()
        let invocation = fun () -> runCommand javaPath [| "-jar"; custodesPath; spreadsheet; outputPath; |]
        match invocation() with
        | STDOUT output -> CUSTODESGrammar.parse output
        | STDERR error -> failwith error

    type Run(spreadsheet: string, custodesPath: string, javaPath: string) =
        let cOutput = runCUSTODES spreadsheet custodesPath javaPath