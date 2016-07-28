open COMWrapper
open System

    [<EntryPoint>]
    let main argv = 
        let config = Args.processArgs argv

        Console.CancelKeyPress.Add(fun _ -> printfn "Do something")

        let app = new Application()

        for file in config.files do
//            let wb = app.OpenWorkbook(@"..\..\TestData\AddressModes.xlsx")
//            let graph = wb.buildDependenceGraph()

            printfn "Analyzing: %A" (System.IO.Path.GetFileName file)

        printfn "Analysis complete.  Press Enter to continue."
        System.Console.ReadLine() |> ignore

        0
