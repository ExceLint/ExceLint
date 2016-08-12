module Args

open System.IO
open System.Text.RegularExpressions

    type Config(dpath: string, opath: string, jpath: string, cpath: string, verbose: bool, csv: string, fc: ExceLint.FeatureConf) =
        member self.files: string[] =
            Directory.EnumerateFiles(dpath, "*.xls?", SearchOption.AllDirectories) |> Seq.toArray
        member self.csv: string = csv
        member self.isVerbose : bool = verbose
        member self.verbose_csv(wbname: string) = Path.Combine(opath, Regex.Replace(wbname,"[^A-Za-z0-9_-]","") + ".csv")
        member self.FeatureConf = fc
        member self.CustodesPath = cpath
        member self.JavaPath = jpath
        member self.InputDirectory = dpath
        member self.OutputPath = opath

    let usage() : unit =
        printfn "ExceLintRunner.exe [input directory] [output path] [java path] [opt_verbose] [opt1 ... opt7]"
        printfn 
            "Recursively finds all Excel (*.xls and *.xlsx) files in [input directory], \n\
            opens them, runs ExceLint, and prints output statistics to a file called \n\
            exceline_output.csv in [output path].\n\n\
            [java path] and [CUSTODES path] are needed in order to conduct a comparison\n\
            against the CUSTODES tool.\n\n\
            Press Ctrl-C to cancel an analysis."
        printfn "\nwhere\n"
        printfn "opt_verbose = <true> or <false>: log per-spreadsheet flagged cells as separate csvs"
        printfn "opt1 = <true> or <false>: condition by all cells"
        printfn "opt2 = <true> or <false>: condition by columns"
        printfn "opt3 = <true> or <false>: condition by rows"
        printfn "opt4 = <true> or <false>: condition by levels"
        printfn "opt5 = <true> or <false>: infer address modes"
        printfn "opt6 = <true> or <false>: weigh by intrinsic anomalousness"
        printfn "opt7 = <true> or <false>: weigh by conditioning set size"
        printfn "\nExample:"
        printfn "ExceLintRunner \"C:\\data\" output.csv true true true true true false false true"

        System.Environment.Exit(1)

    let processArgs(argv: string[]) : Config =
        if argv.Length <> 12 then
            usage()
        let dpath = argv.[0]    // input directory
        let opath = argv.[1]    // output directory
        let jpath = argv.[2]    // java path
        let cpath = argv.[3]    // CUSTODES path

        let optVerbose = System.Boolean.Parse(argv.[4])
        let optCondAllCells = System.Boolean.Parse(argv.[5])
        let optCondCols = System.Boolean.Parse(argv.[6])
        let optCondRows = System.Boolean.Parse(argv.[7])
        let optCondLevels = System.Boolean.Parse(argv.[8])
        let optAddrMode = System.Boolean.Parse(argv.[9])
        let optIntrinsicAnom = System.Boolean.Parse(argv.[10])
        let optCondSetSz = System.Boolean.Parse(argv.[11])

        let csv = Path.Combine(opath, "excelint_output.csv")

        let mutable featureConf = new ExceLint.FeatureConf()

        featureConf <- featureConf.enableShallowInputVectorMixedL2NormSum()
        featureConf <- if optCondAllCells then featureConf.analyzeRelativeToAllCells() else featureConf
        featureConf <- if optCondCols then featureConf.analyzeRelativeToColumns() else featureConf
        featureConf <- if optCondRows then featureConf.analyzeRelativeToRows() else featureConf
        featureConf <- if optCondLevels then featureConf.analyzeRelativeToLevels() else featureConf
        featureConf <- if optAddrMode then featureConf.inferAddressModes() else featureConf
        featureConf <- if optIntrinsicAnom then featureConf.weightByIntrinsicAnomalousness() else featureConf
        featureConf <- if optCondSetSz then featureConf.weightByConditioningSetSize() else featureConf

        Config(dpath, opath, jpath, cpath, optVerbose, csv, featureConf)

