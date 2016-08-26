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
        member self.DebugPath = Path.Combine(opath, "debug.csv")

    let usage() : unit =
        printfn "ExceLintRunner.exe <input directory> <output path> <java path> [flags]"
        printfn 
            "Recursively finds all Excel (*.xls and *.xlsx) files in <input directory>, \n\
            opens them, runs ExceLint, and prints output statistics to a file called \n\
            exceline_output.csv in <output path>.\n\n\
            <java path> and <CUSTODES path> are needed in order to conduct a comparison\n\
            against the CUSTODES tool.\n\n\
            Press Ctrl-C to cancel an analysis."
        printfn "\nwhere\n"
        printfn "[flags] consists of any of the following options, which are TRUE when"
        printfn "present and FALSE when omitted:"
        printfn "\n"
        printfn "-verbose    log per-spreadsheet flagged cells as separate csvs"
        printfn "-spectral   use spectral outliers, otherwise use summation outliers;"
        printfn "            forces the use of -sheets below and disables -allcells,"
        printfn "            -columns, -rows, and -levels"
        printfn "-allcells   condition by all cells"
        printfn "-columns    condition by columns"
        printfn "-rows       condition by rows"
        printfn "-levels     condition by levels"
        printfn "-sheets     condition by sheets"
        printfn "-addrmode   infer address modes"
        printfn "-intrinsic  weigh by intrinsic anomalousness"
        printfn "-css        weigh by conditioning set size"
        printfn "\nExample:\n"
        printfn "ExceLintRunner \"C:\\data\" output.csv true true true true true false false true"
        printfn "ExceLintRunner.exe \"C:\\data\" \"C:\\output\" \"C:\\ProgramData\\Oracle\\Java\\javapath\\java.exe\" \"C:\\CUSTODES2\\cc2.jar\" -verbose -allcells -rows -columns -levels -css"

        System.Environment.Exit(1)

    let processArgs(argv: string[]) : Config =
        if argv.Length < 4 || argv.Length > 14 then
            usage()
        let dpath = argv.[0]    // input directory
        let opath = argv.[1]    // output directory
        let jpath = argv.[2]    // java path
        let cpath = argv.[3]    // CUSTODES path

        let csv = Path.Combine(opath, "excelint_output.csv")

        let flags = argv.[4 .. argv.Length - 1]

        let (isVerb,fConf) = Array.fold (fun (isVerb: bool,conf: ExceLint.FeatureConf) flag ->
                                 match flag with
                                 | "-verbose" -> true, conf
                                 | "-spectral" -> isVerb, conf.spectralRanking()
                                 | "-allcells" -> isVerb, if conf.IsEnabledSpectralRanking then conf else conf.analyzeRelativeToAllCells()
                                 | "-columns" -> isVerb, if conf.IsEnabledSpectralRanking then conf else conf.analyzeRelativeToColumns()
                                 | "-rows" -> isVerb, if conf.IsEnabledSpectralRanking then conf else conf.analyzeRelativeToRows()
                                 | "-levels" -> isVerb, if conf.IsEnabledSpectralRanking then conf else conf.analyzeRelativeToLevels()
                                 | "-sheets" -> isVerb, if conf.IsEnabledSpectralRanking then conf.analyzeRelativeToSheet() else conf
                                 | "-addrmode" -> isVerb, conf.inferAddressModes()
                                 | "-intrinsic" -> isVerb, conf.weightByIntrinsicAnomalousness()
                                 | "-css" -> isVerb, conf.weightByConditioningSetSize()
                                 | s -> failwith ("Unrecognized option: " + s)
                             ) (false,new ExceLint.FeatureConf()) flags

        Config(dpath, opath, jpath, cpath, isVerb, csv, fConf)

