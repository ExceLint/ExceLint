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

        using (new StreamWriter(config.csv)) (fun sw ->
            
            // write headers
            sw.Write(CSV.ExceLintStatsHeaders)
            sw.Flush()

            for file in config.files do
                let shortf = (System.IO.Path.GetFileName file)

                printfn "Opening: %A" shortf
                let wb = app.OpenWorkbook(file)
            
                printfn "Building dependence graph: %A" shortf
                let graph = wb.buildDependenceGraph()

                printfn "Running ExceLint analysis: %A" shortf
                let model_opt = ExceLint.ModelBuilder.analyze (app.XLApplication()) config.FeatureConf graph thresh (Depends.Progress.NOPProgress())

                printfn "Running CUSTODES analysis: %A" shortf
                let custodes = CUSTODES.Output(file, config.CustodesPath, config.JavaPath)

                match model_opt with
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
                                causesTimeMs = model.CausesTimeInMilliseconds,
                                conditioningSetSzTimeMs = model.ConditioningSetSizeTimeInMilliseconds,
                                numAnom = model.getSignificanceCutoff,
                                numCustodesSmells = custodes.NumSmells,
                                optCondAllCells = config.FeatureConf.IsEnabledOptCondAllCells,
                                optCondRows = config.FeatureConf.IsEnabledOptCondRows,
                                optCondCols = config.FeatureConf.IsEnabledOptCondCols,
                                optCondLevels = config.FeatureConf.IsEnabledOptCondLevels,
                                optAddrmodeInference = config.FeatureConf.IsEnabledOptAddrmodeInference,
                                optWeightIntrinsicAnom = config.FeatureConf.IsEnabledOptWeightIntrinsicAnomalousness,
                                optWeightConditionSetSz = config.FeatureConf.IsEnabledOptWeightConditioningSetSize
                              )

                    // append to streamwriter & flush stream
                    sw.Write (csv.Append([row]).SaveToString())
                    sw.Flush()

                    // per-workbook stats
                    if (config.isVerbose) then
                        let per_csv = new CSV.WorkbookStats([])

                        using (new StreamWriter(config.verbose_csv shortf)) (fun per_sw ->
                            let ranking = model.rankByFeatureSum()

                            Array.mapi (fun i (kvp: KeyValuePair<AST.Address,double>) ->
                                let per_row = CSV.WorkbookStats.Row(
                                                    flaggedCellAddr = kvp.Key.A1FullyQualified(),
                                                    rank = i,
                                                    score = kvp.Value
                                                )

                                // append to streamwriter
                                per_sw.Write (per_csv.Append([per_row]).SaveToString())
                            ) ranking |> ignore
                        )

                    printfn "Analysis complete: %A" shortf
                                
                | None ->
                    printfn "Analysis failed: %A" shortf
        )

        0
