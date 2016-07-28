open COMWrapper
open System

    [<EntryPoint>]
    let main argv = 
        let config = Args.processArgs argv

        Console.CancelKeyPress.Add(fun _ -> printfn "Do something")

        let app = new Application()

        for file in config.files do
            let shortf = (System.IO.Path.GetFileName file)

            printfn "Analyzing: %A" shortf
            let wb = app.OpenWorkbook(file)
            let graph = wb.buildDependenceGraph()
            printfn "DAG built: %A" shortf

        printfn "Analysis complete.  Press Enter to continue."
        System.Console.ReadLine() |> ignore

        0
