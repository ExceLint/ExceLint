open COMWrapper
open System
open System.IO
open System.Collections.Generic
open ExceLint

    type Stats = {
        shortname: string;
        threshold: double;
        custodes_flagged: HashSet<AST.Address>;
        excelint_not_custodes: HashSet<AST.Address>;
        custodes_not_excelint: HashSet<AST.Address>;
        true_smells_this_wb : HashSet<AST.Address>;
        true_smells_not_found_by_excelint: HashSet<AST.Address>;
        true_smells_not_found_by_custodes: HashSet<AST.Address>;
        true_smells_not_found: HashSet<AST.Address>;
        excel_this_wb: HashSet<AST.Address>;
        excelint_true_smells: HashSet<AST.Address>;
        custodes_true_smells: HashSet<AST.Address>;
        excelint_excel_intersect: HashSet<AST.Address>;
        custodes_excel_intersect: HashSet<AST.Address>;
    }

    let hs_difference<'a>(hs1: HashSet<'a>)(hs2: HashSet<'a>) : HashSet<'a> =
        let hs3 = new HashSet<'a>(hs1)
        hs3.ExceptWith(hs2)
        hs3

    let hs_intersection<'a>(hs1: HashSet<'a>)(hs2: HashSet<'a>) : HashSet<'a> =
        let hs3 = new HashSet<'a>(hs1)
        hs3.IntersectWith(hs2)
        hs3

    let hs_union<'a>(hs1: HashSet<'a>)(hs2: HashSet<'a>) : HashSet<'a> =
        let hs3 = new HashSet<'a>(hs1)
        hs3.UnionWith(hs2)
        hs3

    let rankToSet(ranking: Pipeline.Ranking)(model: ErrorModel) : HashSet<AST.Address> =
        Array.mapi (fun i (kvp: KeyValuePair<AST.Address,double>) -> (i, kvp.Key)) ranking
        |> Array.filter (fun (i,e) -> i <= model.Cutoff)
        |> Array.map (fun (i,e) -> e)
        |> (fun arr -> new HashSet<AST.Address>(arr))


    let per_append_excelint(sw: StreamWriter)(csv: CSV.WorkbookStats)(truth: CUSTODES.GroundTruth)(custodes: CUSTODES.OutputResult)(model: ErrorModel)(ranking: Pipeline.Ranking) : unit =
        let smells = match custodes with
                     | CUSTODES.OKOutput c -> c.Smells
                     | _ -> new HashSet<AST.Address>()

        // append all ExceLint flagged cells
        Array.mapi (fun i (kvp: KeyValuePair<AST.Address,double>) ->
            let addr = kvp.Key
            let per_row = CSV.WorkbookStats.Row(
                                path = addr.A1Path(),
                                workbook = addr.WorkbookName,
                                worksheet = addr.WorksheetName,
                                addr = addr.A1Local(),
                                flaggedByExcelint = (i <= model.Cutoff),
                                flaggedByCustodes = smells.Contains addr,
                                flaggedByExcel = truth.isFlaggedByExcel(addr),
                                cliSameAsV1 = truth.differs addr (smells.Contains addr),
                                rank = i,
                                score = kvp.Value,
                                custodesTrueSmell = truth.isTrueSmell addr
                            )

            // append to streamwriter
            sw.Write (csv.Append([per_row]).SaveToString())
        ) ranking |> ignore

    let per_append_custodes(sw: StreamWriter)(csv: CSV.WorkbookStats)(truth: CUSTODES.GroundTruth)(custodes: CUSTODES.OutputResult)(model: ErrorModel)(ranking: Pipeline.Ranking)(custodes_not_excelint: HashSet<AST.Address>) : unit =
        let smells = match custodes with
                     | CUSTODES.OKOutput c -> c.Smells
                     | _ -> new HashSet<AST.Address>()

        // append all remaining CUSTODES cells
        Array.map (fun (addr: AST.Address) ->
            let per_row = CSV.WorkbookStats.Row(
                                path = addr.A1Path(),
                                workbook = addr.WorkbookName,
                                worksheet = addr.WorksheetName,
                                addr = addr.A1Local(),
                                flaggedByExcelint = false,
                                flaggedByCustodes = true,
                                flaggedByExcel = truth.isFlaggedByExcel(addr),
                                cliSameAsV1 = truth.differs addr (smells.Contains addr),
                                rank = 999999999,
                                score = 0.0,
                                custodesTrueSmell = truth.isTrueSmell addr
                            )

            // append to streamwriter
            sw.Write (csv.Append([per_row]).SaveToString())
        ) (custodes_not_excelint |> Seq.toArray) |> ignore

    let per_append_true_smells(sw: StreamWriter)(csv: CSV.WorkbookStats)(truth: CUSTODES.GroundTruth)(custodes: CUSTODES.OutputResult)(model: ErrorModel)(ranking: Pipeline.Ranking)(true_smells_not_found: HashSet<AST.Address>) : unit =
        let smells = match custodes with
                     | CUSTODES.OKOutput c -> c.Smells
                     | _ -> new HashSet<AST.Address>()

        // append all true smells found by neither tool
        Array.map (fun (addr: AST.Address) ->
            let per_row = CSV.WorkbookStats.Row(
                                path = addr.A1Path(),
                                workbook = addr.WorkbookName,
                                worksheet = addr.WorksheetName,
                                addr = addr.A1Local(),
                                flaggedByExcelint = false,
                                flaggedByCustodes = false,
                                flaggedByExcel = truth.isFlaggedByExcel(addr),
                                cliSameAsV1 = truth.differs addr (smells.Contains addr),
                                rank = 999999999,
                                score = 0.0,
                                custodesTrueSmell = true
                            )

            // append to streamwriter
            sw.Write (csv.Append([per_row]).SaveToString())
        ) (true_smells_not_found |> Seq.toArray) |> ignore

    let per_append_debug(sw: StreamWriter)(csv: CSV.DebugInfo)(model: ErrorModel)(custodes_smells: HashSet<AST.Address>)(config: Args.Config)(ranking: Pipeline.Ranking) : unit =
        // warn user if CUSTODES analysis contains cells not analyzed by ExceLint
        let rset = Array.map (fun (kvp: KeyValuePair<AST.Address, double>) -> kvp.Key) ranking
                   |> (fun arr -> new HashSet<AST.Address>(arr))
        let not_excelint_at_all = hs_difference custodes_smells rset

        if not_excelint_at_all.Count <> 0 then
            // are these related to formulas?  if so, this is a sign that something went wrong
            let all_comp = new HashSet<AST.Address>(model.DependenceGraph.allComputationCells())
            let missed_formula_related = hs_intersection not_excelint_at_all all_comp

            printfn "WARNING: CUSTODES analysis contains %d cells not analyzed by ExceLint," (not_excelint_at_all.Count)
            printfn "         %d of which are formula-related (either inputs or formulas)." missed_formula_related.Count
            if missed_formula_related.Count > 0 then
                printfn "         Writing to %s" config.DebugPath

                // append all true smells found by neither tool
                Array.map (fun (addr: AST.Address) ->
                    let per_row = CSV.DebugInfo.Row(
                                        path = addr.A1Path(),
                                        workbook = addr.WorkbookName,
                                        worksheet = addr.WorksheetName,
                                        addr = addr.A1Local()
                                    )

                    // append to streamwriter
                    sw.Write (csv.Append([per_row]).SaveToString())
                ) (missed_formula_related |> Seq.toArray) |> ignore

    let append_stats(stats: Stats)(sw: StreamWriter)(csv: CSV.ExceLintStats)(model: ErrorModel)(custodes: CUSTODES.OutputResult)(config: Args.Config) : unit =
        // write stats
        let row = CSV.ExceLintStats.Row(
                    benchmarkName = stats.shortname,
                    numCells = model.DependenceGraph.allCells().Length,
                    numFormulas = model.DependenceGraph.getAllFormulaAddrs().Length,
                    sigThresh = stats.threshold,
                    depTimeMs = model.DependenceGraph.AnalysisMilliseconds,
                    scoreTimeMs = model.ScoreTimeInMilliseconds,
                    freqTimeMs = model.FrequencyTableTimeInMilliseconds,
                    rankingTimeMs = model.RankingTimeInMilliseconds,
                    causesTimeMs = model.CausesTimeInMilliseconds,
                    conditioningSetSzTimeMs = model.ConditioningSetSizeTimeInMilliseconds,
                    numAnom = model.Cutoff,
                    custodesFail = (match custodes with | CUSTODES.BadOutput _ -> true | _ -> false),
                    custodesFailMsg = (match custodes with | CUSTODES.BadOutput msg -> msg | _ -> ""),
                    numCustodesSmells = stats.custodes_flagged.Count,
                    trueSmells = stats.true_smells_this_wb.Count,
                    excelintTrueSmellsFound = stats.excelint_true_smells.Count,
                    custodesTrueSmellsFound = stats.custodes_true_smells.Count,
                    excelintCustodesTrueSmellIntersect = (hs_intersection stats.excelint_true_smells stats.custodes_true_smells).Count,
                    trueSmellsMissedByBoth = (hs_difference stats.true_smells_this_wb (hs_union stats.excelint_true_smells stats.custodes_true_smells)).Count,
                    excelFlags = stats.excel_this_wb.Count,
                    excelintExcelIntersect = stats.excelint_excel_intersect.Count,
                    custodesExcelIntersect = stats.custodes_excel_intersect.Count,
                    excelFlagsMissedByBoth = (hs_difference stats.excel_this_wb (hs_union stats.excelint_excel_intersect stats.custodes_excel_intersect)).Count,
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

    let analyze (file: String)(thresh: double)(app: Application)(config: Args.Config)(truth: CUSTODES.GroundTruth)(csv: CSV.ExceLintStats)(debug_csv: CSV.DebugInfo)(sw: StreamWriter)(debug_sw: StreamWriter) =
        let shortf = (System.IO.Path.GetFileName file)

        printfn "Opening: %A" shortf

        let wb = app.OpenWorkbook(file)
            
        printfn "Building dependence graph: %A" shortf
        let graph = wb.buildDependenceGraph()

        printfn "Running ExceLint analysis: %A" shortf
        let model_opt = ExceLint.ModelBuilder.analyze (app.XLApplication()) config.FeatureConf graph thresh (Depends.Progress.NOPProgress())

        printfn "Running CUSTODES analysis: %A" shortf
        let custodes = CUSTODES.getOutput(file, config.CustodesPath, config.JavaPath)

        match model_opt with
        | Some(model) ->
            // per-workbook stats
            let per_csv = new CSV.WorkbookStats([])

            using (new StreamWriter(config.verbose_csv shortf)) (fun per_sw ->
                // write header
                // write headers
                per_sw.Write(CSV.WorkbookStatsHeaders)
                per_sw.Flush()

                let ranking = model.rankByFeatureSum()

                // get the set of cells flagged by ExceLint
                let excelint_flags = rankToSet ranking model

                // get workbook selector
                let this_wb = ranking.[0].Key.WorkbookName

                // get the set of cells flagged by CUSTODES
                let custodes_flags = match custodes with
                                        | CUSTODES.OKOutput c -> c.Smells
                                        | CUSTODES.BadOutput _ -> new HashSet<AST.Address>()

                // find true smells found by neither tool
                let true_smells_this_wb = truth.TrueSmellsbyWorkbook this_wb
                let true_smells_not_found_by_excelint = hs_difference true_smells_this_wb excelint_flags
                let true_smells_not_found_by_custodes = hs_difference true_smells_this_wb custodes_flags
                let true_smells_not_found = hs_intersection true_smells_not_found_by_excelint true_smells_not_found_by_custodes

                // overall stats
                let excel_this_wb = truth.ExcelbyWorkbook this_wb
                let excelint_true_smells = hs_intersection excelint_flags true_smells_this_wb
                assert (excelint_true_smells.IsSubsetOf(excelint_flags))
                let custodes_true_smells = hs_intersection custodes_flags true_smells_this_wb
                assert (custodes_true_smells.IsSubsetOf(custodes_flags))
                let excelint_excel_intersect = hs_intersection excelint_flags excel_this_wb
                let custodes_excel_intersect = hs_intersection custodes_flags excel_this_wb

                let stats = {
                    shortname = shortf;
                    threshold = thresh;
                    custodes_flagged = custodes_flags;
                    excelint_not_custodes = hs_difference excelint_flags custodes_flags;
                    custodes_not_excelint = hs_difference custodes_flags excelint_flags;
                    true_smells_this_wb = true_smells_this_wb;
                    true_smells_not_found_by_excelint = true_smells_not_found_by_excelint;
                    true_smells_not_found_by_custodes = true_smells_not_found_by_custodes;
                    true_smells_not_found = true_smells_not_found;
                    excel_this_wb =  excel_this_wb;
                    excelint_true_smells = excelint_true_smells;
                    custodes_true_smells = custodes_true_smells;
                    excelint_excel_intersect = excelint_excel_intersect;
                    custodes_excel_intersect = custodes_excel_intersect;
                }

                // write to per-workbook CSV
                per_append_excelint per_sw per_csv truth custodes model ranking
                per_append_custodes per_sw per_csv truth custodes model ranking stats.excelint_not_custodes
                per_append_true_smells per_sw per_csv truth custodes model ranking true_smells_not_found

                // write overall stats to CSV
                append_stats stats sw csv model custodes config

                // sanity check
                per_append_debug debug_sw debug_csv model stats.custodes_flagged config ranking
            )

            printfn "Analysis complete: %A" shortf
        | None ->
            printfn "Analysis failed: %A" shortf

    [<EntryPoint>]
    let main argv = 
        let config = Args.processArgs argv

        Console.CancelKeyPress.Add(
            (fun _ ->
                printfn "Ctrl-C received.  Cancelling..."
                System.Environment.Exit(1)
            )
        )

        using(new Application()) (fun app ->

            let thresh = 0.05

            let csv = new CSV.ExceLintStats([])
            let debug_csv = new CSV.DebugInfo([])

            let truth = new CUSTODES.GroundTruth(config.InputDirectory)

            using (new StreamWriter(config.csv)) (fun sw ->
                using (new StreamWriter(config.DebugPath)) (fun debug_sw ->
            
                    // write headers
                    sw.Write(CSV.ExceLintStatsHeaders)
                    sw.Flush()

                    debug_sw.Write(CSV.DebugInfoHeaders)
                    debug_sw.Flush()

                    for file in config.files do
                        
                        let shortf = (System.IO.Path.GetFileName file)

                        try
                            analyze file thresh app config truth csv debug_csv sw debug_sw
                        with
                        | e -> printfn "Cannot open workbook %A because:\n%A" shortf e.Message
                )
            )
        )

        printfn "Press Enter to continue."
        Console.ReadLine() |> ignore

        0
