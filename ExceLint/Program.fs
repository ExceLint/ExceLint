namespace ExceLint
    open COMWrapper

    module Analysis =

        [<EntryPoint>]
        let main argv = 
            let filename = @"..\..\..\ExceLintTests\TestData\7_gradebook_xls.xlsx"

            let app = new COMWrapper.Application()

            let workbook = app.OpenWorkbook(filename)

            let dag = workbook.buildDependenceGraph()

            printfn "%s" (dag.ToDOT())

            0
