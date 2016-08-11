open COMWrapper
open System
open System.IO
open System.Collections.Generic
open ExceLint

    [<EntryPoint>]
    let main argv = 
        let config = Args.processArgs argv

//        Console.CancelKeyPress.Add(fun _ -> printfn "Do something")

        using(new Application()) (fun app ->

            let thresh = 0.05

            let csv = new CSV.ExceLintStats([])

            let truth = new CUSTODES.GroundTruth(config.InputDirectory)

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

                                // convert to set
                                let excelint_flags = new HashSet<AST.Address>(Array.map (fun (kvp: KeyValuePair<AST.Address,double>) -> kvp.Key ) ranking)

                                // append all ExceLint flagged cells
                                Array.mapi (fun i (kvp: KeyValuePair<AST.Address,double>) ->
                                    let addr = kvp.Key
                                    let per_row = CSV.WorkbookStats.Row(
                                                        flaggedCellAddr = addr.A1FullyQualified(),
                                                        flaggedByExcelint = true,
                                                        flaggedByCustodes = custodes.Smells.Contains addr,
                                                        cliSameAsV1 = truth.differs addr (custodes.Smells.Contains addr),
                                                        rank = i,
                                                        score = kvp.Value,
                                                        custodesTrueSmell = truth.isTrueSmell addr
                                                    )

                                    // append to streamwriter
                                    per_sw.Write (per_csv.Append([per_row]).SaveToString())
                                ) ranking |> ignore

                                // HashSet difference is, annoyingly, side-effecting
                                let except_excelint = new HashSet<AST.Address>(custodes.Smells)
                                except_excelint.ExceptWith(excelint_flags)
                                let except_excelint_arr = except_excelint |> Seq.toArray

                                // append all remaining CUSTODES cells
                                Array.map (fun (addr: AST.Address) ->
                                    let per_row = CSV.WorkbookStats.Row(
                                                        flaggedCellAddr = addr.A1FullyQualified(),
                                                        flaggedByExcelint = false,
                                                        flaggedByCustodes = true,
                                                        cliSameAsV1 = truth.differs addr (custodes.Smells.Contains addr),
                                                        rank = 999999999,
                                                        score = 0.0,
                                                        custodesTrueSmell = truth.isTrueSmell addr
                                                    )

                                    // append to streamwriter
                                    per_sw.Write (per_csv.Append([per_row]).SaveToString())
                                ) except_excelint_arr |> ignore
                            )

                        printfn "Analysis complete: %A" shortf
                                
                    | None ->
                        printfn "Analysis failed: %A" shortf
            )
        )

        0
