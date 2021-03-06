﻿module Args

open System.IO
open System.Text.RegularExpressions
    type Knobs = { verbose: bool; dont_exit: bool; alpha: double; oldNNjaccard: bool; kmedioidjaccard: bool; nocustodes: bool; tronly: bool; shuffle: bool; }

    type Config(dpath: string, opath: string, jpath: string, cpath: string, egpath: string, gpath: string, knobs: Knobs, csv: string, fc: ExceLint.FeatureConf) =
        do
            printfn "\n------------------------------------"
            printfn "Running with the following options: "
            printfn "------------------------------------"
            printfn "Input directory: %s" dpath
            printfn "Output directory: %s" opath
            printfn "Output stats CSV: %s" csv
            printfn "ExceLint ground truth CSV: %s" egpath
            if knobs.tronly then
                printfn "ONLY RUNNING BENCHMARKS WITH TRUE REF ANNOTATIONS"
            if knobs.nocustodes then
                printfn "NOT RUNNING CUSTODES."
            else
                printfn "CUSTODES ground truth CSV: %s" gpath
                printfn "Java path: %s" jpath
                printfn "CUSTODES JAR path: %s" cpath
            printfn "Verbose mode: %b" knobs.verbose
            printfn "No-exit mode: %b" knobs.dont_exit
            printfn "Shuffle input files: %b" knobs.shuffle
            if knobs.oldNNjaccard then
                printfn "Comparing against old nearest-neighbor algorithm."
            if knobs.kmedioidjaccard then
                printfn "Comparing against k-medioid algorithm."
            printfn "Threshold: %f" knobs.alpha
            Array.iter (fun (opt,enabled) -> printfn "%s: %b" opt enabled) (ExceLint.FeatureConf.simpleConf(fc.rawConf) |> Map.toArray)
            printfn "------------------------------------\n"

        member self.files: string[] =
            Directory.EnumerateFiles(dpath, "*.xls?", SearchOption.AllDirectories) |> Seq.toArray
        member self.csv: string = csv
        member self.isVerbose : bool = knobs.verbose
        member self.verbose_csv(wbname: string) = Path.Combine(opath, Regex.Replace(wbname,"[^A-Za-z0-9_-]","") + ".csv")
        member self.clustering_csv(wbname: string)(clustername: string) = Path.Combine(opath, Regex.Replace(wbname,"[^A-Za-z0-9_-]","") + "_" + clustername + ".csv")
        member self.FeatureConf = fc
        member self.CustodesPath = cpath
        member self.JavaPath = jpath
        member self.InputDirectory = dpath
        member self.OutputDirectory = opath
        member self.DebugPath = Path.Combine(opath, "debug.csv")
        member self.DontExitWithoutKeystroke = knobs.dont_exit
        member self.CustodesGroundTruthCSV = gpath
        member self.ExceLintGroundTruthCSV = egpath
        member self.alpha = knobs.alpha
        member self.CompareAgainstOldNN = knobs.oldNNjaccard
        member self.CompareAgainstKMedioid = knobs.kmedioidjaccard
        member self.DontRunCUSTODES = knobs.nocustodes
        member self.TrueRefOnly = knobs.tronly
        member self.Shuffle = knobs.shuffle

    let usage() : unit =
        printfn "ExceLintRunner.exe <input directory> <output directory> <ExceLint ground truth CSV> <CUSTODES ground truth CSV> <java path> <CUSTODES JAR> [flags]"
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
        printfn "-tronly     only run those benchmarks that have true ref annotations"
        printfn "-noexit     prompt user to press a key before exiting"
        printfn "-noshuffle  don't shuffle input spreadsheets"
        printfn "-nocustodes don't do a comparison against CUSTODES"
        printfn "-resultant  bin by resultant vector, otherwise bin by L2 norm sum"
        printfn "-spectral   find outliers by earth mover's distance, otherwise use raw frequency;"
        printfn "            forces the use of -sheets below and disables -allcells,"
        printfn "            -columns, -rows, and -levels"
        printfn "-cluster    find outliers by entropy-minimizing clustering; forces the"
        printfn "            use of -sheets and disables -allcells, -columns, -rows,"
        printfn "            and -levels"
        printfn "-oldcluster compares LSH-NN clustering against slow NN."
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
        printfn "-thresh <n> sets max %% to inspect at n%%; default 5%%"
        printfn "-dist <d>   selects a distance metric for the clustering configuration;"
        printfn "            one of:"
        printfn "            NN     nearest neighbor euclidean distance (single linkage)"
        printfn "            EMD    earth mover's distance (complete linkage)"
        printfn "            MC     mean centroid (complete linkage)"
        printfn "\nExample:\n"
        printfn "ExceLintRunner.exe \"C:\\data\" \"C:\\output\" \"C:\\EXCELINT\\ground_truth.csv\" \"C:\\CUSTODES\\smell_detection_result.csv\" \"C:\\ProgramData\\Oracle\\Java\\javapath\\java.exe\" \"C:\\CUSTODES\\cc2.jar\" -verbose -allcells -rows -columns -levels -css"
        printfn "\nHelp:\n"
        printfn "ExceLintRunner.exe -help"

        System.Environment.Exit(1)

    let processArgs(argv: string[]) : Config =
        if argv.Length < 6 || argv.Length > 16 || (Array.contains "-help" argv) || (Array.contains "--help" argv) then
            usage()
        let dpath  = System.IO.Path.GetFullPath argv.[0]   // input directory
        let opath  = System.IO.Path.GetFullPath argv.[1]   // output directory
        let egpath = System.IO.Path.GetFullPath argv.[2]   // path to ExceLint ground truth CSV
        let gpath  = System.IO.Path.GetFullPath argv.[3]   // path to CUSTODES ground truth CSV
        let jpath  = System.IO.Path.GetFullPath argv.[4]   // java path
        let cpath  = System.IO.Path.GetFullPath argv.[5]   // CUSTODES path

        let csv = Path.Combine(opath, "excelint_output.csv")

        let flags = argv.[6 .. argv.Length - 1] |> Array.toList

        let rec optParse = (fun (args: string list)(knobs: Knobs)(conf: ExceLint.FeatureConf) ->
                               match args with
                               | [] -> knobs, conf
                               // KNOBS
                               | "-verbose"    :: rest -> optParse rest { knobs with verbose = true } conf
                               | "-noexit"     :: rest -> optParse rest { knobs with dont_exit = true } conf
                               | "-oldcluster" :: rest -> optParse rest { knobs with oldNNjaccard = true } conf
                               | "-kmedioid"   :: rest -> optParse rest { knobs with kmedioidjaccard = true } conf
                               | "-nocustodes" :: rest -> optParse rest { knobs with nocustodes = true } conf
                               | "-tronly"     :: rest -> optParse rest { knobs with tronly = true } conf
                               | "-noshuffle"  :: rest -> optParse rest { knobs with shuffle = false } conf
                               // FEATURECONF
                               | "-cluster"    :: rest -> optParse rest knobs (conf.enableShallowInputVectorMixedFullCVectorResultantOSI true)
                               | "-allcells"   :: rest -> optParse rest knobs (conf.analyzeRelativeToAllCells true)
                               | "-columns"    :: rest -> optParse rest knobs (conf.analyzeRelativeToColumns true)
                               | "-rows"       :: rest -> optParse rest knobs (conf.analyzeRelativeToRows true)
                               | "-levels"     :: rest -> optParse rest knobs (conf.analyzeRelativeToLevels true)
                               | "-sheets"     :: rest -> optParse rest knobs (conf.analyzeRelativeToSheet true)
                               | "-addrmode"   :: rest -> optParse rest knobs (conf.inferAddressModes true)
                               | "-intrinsic"  :: rest -> optParse rest knobs (conf.weightByIntrinsicAnomalousness true)
                               | "-css"        :: rest -> optParse rest knobs (conf.weightByConditioningSetSize true)
                               | "-inputstoo"  :: rest -> optParse rest knobs (conf.analyzeOnlyFormulas false)
                               | "-thresh"     :: d :: rest ->
                                   let alpha = System.Convert.ToDouble d / 100.0
                                   if alpha < 0.0 || alpha > 1.0 then
                                       failwith "Threshold must be between 0 and 100."
                                   optParse rest { knobs with alpha = alpha } conf
                               | "-dist"      :: s :: rest ->
                                   let conf' = match s with
                                               | "NN"  -> conf.enableDistanceNearestNeighbor true
                                               | "EMD" -> conf.enableDistanceEarthMover true
                                               | "MC"  -> conf.enableDistanceMeanCentroid true
                                   optParse rest knobs conf'
                               | s :: rest -> failwith ("Unrecognized option: " + s)
                           )

        // init with defaults
        let (knobs,fConf) = optParse flags { verbose = false; dont_exit = false; alpha = 0.05; oldNNjaccard = false; kmedioidjaccard = false; nocustodes = false; tronly = false; shuffle = true;} (new ExceLint.FeatureConf())

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

        Config(dpath, opath, jpath, cpath, egpath, gpath, knobs, csv, fConf')

