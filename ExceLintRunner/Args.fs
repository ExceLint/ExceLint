module Args

open System.IO

    type Config(dpath: string, csv: string, fc: ExceLint.FeatureConf) =
        member self.files: string[] =
            Directory.EnumerateFiles(dpath, "*.xls?", SearchOption.AllDirectories) |> Seq.toArray
        member self.csv: string = csv
        member self.FeatureConf = fc

    let usage() : unit =
        printfn "ExceLintRunner.exe [input directory] [output csv] [opt1 ... opt7]"
        printfn 
            "Recursively finds all Excel (*.xls and *.xlsx) files in [input directory], \
            opens them, runs ExceLint, and prints output statistics to [output csv]. \
            Press Ctrl-C to cancel an analysis."
        printfn "\nwhere\n"
        printfn "opt1 = <true> or <false>: condition by all cells"
        printfn "opt2 = <true> or <false>: condition by columns"
        printfn "opt3 = <true> or <false>: condition by rows"
        printfn "opt4 = <true> or <false>: condition by levels"
        printfn "opt5 = <true> or <false>: infer address modes"
        printfn "opt6 = <true> or <false>: weigh by intrinsic anomalousness"
        printfn "opt7 = <true> or <false>: weigh by conditioning set size"
        printfn "\nExample:"
        printfn "ExceLintRunner \"C:\\data\" output.csv true true true true false false true"

        System.Environment.Exit(1)

    let processArgs(argv: string[]) : Config =
        if argv.Length <> 9 then
            usage()
        let dpath = argv.[0]
        let csv = argv.[1]
        let optCondAllCells = System.Boolean.Parse(argv.[2])
        let optCondCols = System.Boolean.Parse(argv.[3])
        let optCondRows = System.Boolean.Parse(argv.[4])
        let optCondLevels = System.Boolean.Parse(argv.[5])
        let optAddrMode = System.Boolean.Parse(argv.[6])
        let optIntrinsicAnom = System.Boolean.Parse(argv.[7])
        let optCondSetSz = System.Boolean.Parse(argv.[8])

        let mutable featureConf = new ExceLint.FeatureConf()

        featureConf <- featureConf.enableShallowInputVectorMixedL2NormSum()
        featureConf <- if optCondAllCells then featureConf.analyzeRelativeToAllCells() else featureConf
        featureConf <- if optCondCols then featureConf.analyzeRelativeToColumns() else featureConf
        featureConf <- if optCondRows then featureConf.analyzeRelativeToRows() else featureConf
        featureConf <- if optCondLevels then featureConf.analyzeRelativeToLevels() else featureConf
        featureConf <- if optAddrMode then featureConf.inferAddressModes() else featureConf
        featureConf <- if optIntrinsicAnom then featureConf.weightByIntrinsicAnomalousness() else featureConf
        featureConf <- if optCondSetSz then featureConf.weightByConditioningSetSize() else featureConf

        Config(dpath, csv, featureConf)

