open COMWrapper
open System
open System.IO

    [<EntryPoint>]
    let main argv = 
        let config = Args.processArgs argv

        Console.CancelKeyPress.Add(fun _ -> printfn "Do something")

        let app = new Application()

        let thresh = 0.05

        let csv = new CSV.ExceLintRunnerLogger([])
        let mutable csv_file = csv.Append([])

        using (new StreamWriter(config.csv)) (fun sw ->
            for file in config.files do
            let shortf = (System.IO.Path.GetFileName file)

            printfn "Analyzing: %A" shortf
            let wb = app.OpenWorkbook(file)
            let graph = wb.buildDependenceGraph()
            printfn "DAG built: %A" shortf
            let analysis = ExceLint.ModelBuilder.analyze (app.XLApplication()) config.FeatureConf graph thresh (Depends.Progress.NOPProgress())

            csv_file <- match analysis with
                            | Some(a) ->

                                let row = CSV.ExceLintRunnerLogger.Row(
                                            benchmarkName = shortf,
                                            numCells = graph.allCells().Length,
                                            numFormulas = graph.getAllFormulaAddrs().Length,
                                            sigThresh = thresh,
                                            depTimeMs = graph.AnalysisMilliseconds,
                                            scoreTimeMs = a.ScoreTimeInMilliseconds,
                                            freqTimeMs = a.FrequencyTableTimeInMilliseconds,
                                            rankingTimeMs = a.RankingTimeInMilliseconds,
                                            numAnom = a.getSignificanceCutoff,
                                            optCondAllCells = config.FeatureConf.IsEnabledOptCondAllCells,
                                            optCondRows = config.FeatureConf.IsEnabledOptCondRows,
                                            optCondCols = config.FeatureConf.IsEnabledOptCondCols,
                                            optCondLevels = config.FeatureConf.IsEnabledOptCondLevels,
                                            optAddrmodeInference = config.FeatureConf.IsEnabledOptAddrmodeInference,
                                            optWeightIntrinsicAnom = config.FeatureConf.IsEnabledOptWeightIntrinsicAnomalousness,
                                            optWeightConditionSetSz = config.FeatureConf.IsEnabledOptWeightConditioningSetSize
                                          )

                                csv_file.Append([row])
                                
                            | None ->
                                printfn "Analysis failed: %A" shortf
                                csv_file

            csv_file.Save sw
        )

        printfn "Analysis complete.  Press Enter to continue."
        System.Console.ReadLine() |> ignore

        0
