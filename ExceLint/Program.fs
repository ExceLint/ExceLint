namespace ExceLint
    open COMWrapper

    module Analysis =

        [<EntryPoint>]
        let main argv = 
            let filename = argv.[0]

            let app = new COMWrapper.Application()

            let workbook = app.OpenWorkbook(filename)

            let dag = workbook.buildDependenceGraph()

            printfn "%s" (dag.ToDOT())

            let baz = System.Console.ReadLine()

            0
