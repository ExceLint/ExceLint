open COMWrapper
open System
open System.IO
open System.Collections.Generic
open ExceLint

    [<EntryPoint>]
    let main argv = 
        let config = Args.processArgs argv

        Console.CancelKeyPress.Add(fun _ -> printfn "Do something")

        let app = new Application()

        let thresh = 0.05

        let csv = new CSV.ExceLintStats([])
        let mutable csv_file = csv.Append([])

        using (new StreamWriter(config.csv)) (fun sw ->
            for file in config.files do
            let shortf = (System.IO.Path.GetFileName file)

            printfn "Analyzing: %A" shortf
            let wb = app.OpenWorkbook(file)
            let graph = wb.buildDependenceGraph()
            printfn "DAG built: %A" shortf
            let model_opt = ExceLint.ModelBuilder.analyze (app.XLApplication()) config.FeatureConf graph thresh (Depends.Progress.NOPProgress())

            csv_file <- match model_opt with
                            | Some(model) ->

                                // global stats
                                let row = CSV.ExceLintStats.Row(
                                            benchmarkName = shortf,
                                            numCells = graph.allCells().Length,
                                            numFormulas = graph.getAllFormulaAddrs().Length,
                                            sigThresh = thresh,
                                            depTimeMs = graph.AnalysisMilliseconds,
                                            scoreTimeMs = model.ScoreTimeInMilliseconds,
                                            freqTimeMs = model.FrequencyTableTimeInMilliseconds,
                                            rankingTimeMs = model.RankingTimeInMilliseconds,
                                            numAnom = model.getSignificanceCutoff,
                                            optCondAllCells = config.FeatureConf.IsEnabledOptCondAllCells,
                                            optCondRows = config.FeatureConf.IsEnabledOptCondRows,
                                            optCondCols = config.FeatureConf.IsEnabledOptCondCols,
                                            optCondLevels = config.FeatureConf.IsEnabledOptCondLevels,
                                            optAddrmodeInference = config.FeatureConf.IsEnabledOptAddrmodeInference,
                                            optWeightIntrinsicAnom = config.FeatureConf.IsEnabledOptWeightIntrinsicAnomalousness,
                                            optWeightConditionSetSz = config.FeatureConf.IsEnabledOptWeightConditioningSetSize
                                          )

                                let output = csv_file.Append([row])

                                // per-workbook stats
                                if (config.isVerbose) then
                                    let per_csv = CSV.WorkbookStats([])

                                    using (new StreamWriter(config.verbose_csv shortf)) (fun per_sw ->
                                        let mutable per_csv_file = per_csv.Append([])

                                        let ranking = model.rankByFeatureSum()

                                        Array.mapi (fun i (kvp: KeyValuePair<AST.Address,double>) ->
                                            let per_row = CSV.WorkbookStats.Row(
                                                              flaggedCellAddr = kvp.Key.A1FullyQualified(),
                                                              rank = i,
                                                              score = kvp.Value
                                                          )
                                            per_csv_file <- per_csv_file.Append([per_row])
                                        ) ranking |> ignore
                                    )

                                output
                                
                            | None ->
                                printfn "Analysis failed: %A" shortf
                                csv_file

            csv_file.Save sw
        )

        printfn "Analysis complete.  Press Enter to continue."
        System.Console.ReadLine() |> ignore

        0
