module Args

open System.IO
open System.Text.RegularExpressions

    type Config(dpath: string, opath: string, jpath: string, cpath: string, gpath: string, verbose: bool, noexit: bool, csv: string, fc: ExceLint.FeatureConf) =
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
        member self.DontExitWithoutKeystroke = noexit
        member self.CustodesGroundTruthCSV = gpath

    let usage() : unit =
        printfn "ExceLintRunner.exe <input directory> <output directory> <ground truth CSV> <java path> <CUSTODES JAR> [flags]"
        printfn 
            "Recursively finds all Excel (*.xls and *.xlsx) files in <input directory>, \n\
            opens them, runs ExceLint, and prints output statistics to a file called \n\
            exceline_output.csv in <output directory>. The file <ground truth CSV> is\n\
            used to compute false positive rates with respect to CUSTODES ground truth.\n\
            You can obtain this file from the ExceLint repository at:\n\
            \"ExceLint\\data\\analyses\\CUSTODES\"\n\n\
            <java path> and <CUSTODES JAR> are needed in order to conduct a comparison\n\
            against the CUSTODES tool. The CUSTODES JAR is available in the ExceLint\n\
            repository at:\n\
            \"ExceLint\\data\\analyses\\CUSTODES2\cc2.jar\"\n\n\
            Press Ctrl-C to cancel an analysis."
        printfn "\nwhere\n"
        printfn "[flags] consists of any of the following options, which are TRUE when"
        printfn "present and FALSE when omitted:"
        printfn "\n"
        printfn "-verbose    log per-spreadsheet flagged cells as separate CSVs"
        printfn "-noexit     prompt user to press a key before exiting"
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
        printfn "-inputstoo  analyze inputs as well; by default ExceLint only"
        printfn "            analyzes formulas"
        printfn "\nExample:\n"
        printfn "ExceLintRunner.exe \"C:\\data\" \"C:\\output\" \"C:\\CUSTODES\\smell_detection_result.csv\" \"C:\\ProgramData\\Oracle\\Java\\javapath\\java.exe\" \"C:\\CUSTODES\\cc2.jar\" -verbose -allcells -rows -columns -levels -css"
        printfn "\nHelp:\n"
        printfn "ExceLintRunner.exe -help"

        System.Environment.Exit(1)

    let processArgs(argv: string[]) : Config =
        if argv.Length < 5 || argv.Length > 15 || (Array.contains "-help" argv) || (Array.contains "--help" argv) then
            usage()
        let dpath = argv.[0]    // input directory
        let opath = argv.[1]    // output directory
        let gpath = argv.[2]    // path to ground truth CSV
        let jpath = argv.[3]    // java path
        let cpath = argv.[4]    // CUSTODES path

        let csv = Path.Combine(opath, "excelint_output.csv")

        let flags = argv.[5 .. argv.Length - 1]

        let (isVerb,noExit,fConf) = Array.fold (fun (isVerb: bool, noExit: bool, conf: ExceLint.FeatureConf) flag ->
                                        match flag with
                                        | "-verbose" -> true, noExit, conf
                                        | "-noexit" -> isVerb, true, conf
                                        | "-spectral" -> isVerb, noExit, conf.spectralRanking(true)
                                        | "-allcells" -> isVerb, noExit, conf.analyzeRelativeToAllCells(true)
                                        | "-columns" -> isVerb, noExit, conf.analyzeRelativeToColumns(true)
                                        | "-rows" -> isVerb, noExit, conf.analyzeRelativeToRows(true)
                                        | "-levels" -> isVerb, noExit, conf.analyzeRelativeToLevels(true)
                                        | "-sheets" -> isVerb, noExit, conf.analyzeRelativeToSheet(true)
                                        | "-addrmode" -> isVerb, noExit, conf.inferAddressModes(true)
                                        | "-intrinsic" -> isVerb, noExit, conf.weightByIntrinsicAnomalousness(true)
                                        | "-css" -> isVerb, noExit, conf.weightByConditioningSetSize(true)
                                        | "-inputstoo" -> isVerb, noExit, conf.analyzeOnlyFormulas(false)
                                        | s -> failwith ("Unrecognized option: " + s)
                                    ) (false,false,new ExceLint.FeatureConf()) flags

        let fConf' = fConf.validate

        if fConf <> fConf' then
            let (changed,removed,added) = fConf.diff fConf'

            if changed.Count > 0 || removed.Count > 0 || added.Count > 0 then
                printfn "Some options changed to conform with internal constraints:"

            if changed.Count > 0 then
                Set.iter (fun k ->
                    printfn "'%s' changed from %b to %b" k (fConf.IsEnabled k) (fConf'.IsEnabled k)
                ) changed

            if removed.Count > 0 then
                Set.iter (fun k ->
                    printfn "'%s' disabled." k
                ) changed

            if added.Count > 0 then
                Set.iter (fun k ->
                    printfn "'%s' enabled." k
                ) changed

        Config(dpath, opath, jpath, cpath, gpath, isVerb, noExit, csv, fConf')

