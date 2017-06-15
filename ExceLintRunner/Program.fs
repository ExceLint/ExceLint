open COMWrapper
open System
open System.IO
open System.Collections.Generic
open ExceLint
open ExceLint.Utils
open ExceLintFileFormats

    type Stats = {
        shortname: string;
        threshold: double;
        custodes_flagged: HashSet<AST.Address>;
        excelint_not_custodes: HashSet<AST.Address>;
        custodes_not_excelint: HashSet<AST.Address>;
        true_ref_bugs_this_wb: HashSet<AST.Address>;
        excelint_true_ref_bugs: HashSet<AST.Address>;
        custodes_true_ref_bugs: HashSet<AST.Address>;
        true_smells_this_wb : HashSet<AST.Address>;
        true_smells_not_found_by_excelint: HashSet<AST.Address>;
        true_smells_not_found_by_custodes: HashSet<AST.Address>;
        true_smells_not_found: HashSet<AST.Address>;
        excel_this_wb: HashSet<AST.Address>;
        excelint_true_smells: HashSet<AST.Address>;
        custodes_true_smells: HashSet<AST.Address>;
        excelint_excel_intersect: HashSet<AST.Address>;
        custodes_excel_intersect: HashSet<AST.Address>;
        custodes_time: int64;
        excelint_jaccard: double;
        excelint_delta_k: int;
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

    let rankToSet(ranking: CommonTypes.Ranking)(cutoff: int) : HashSet<AST.Address> =
        Array.mapi (fun i (kvp: KeyValuePair<AST.Address,double>) -> (i, kvp.Key)) ranking
        |> Array.filter (fun (i,e) -> i <= cutoff)
        |> Array.map (fun (i,e) -> e)
        |> (fun arr -> new HashSet<AST.Address>(arr))

    let per_append_excelint(csv: WorkbookStats)(etruth: ExceLintGroundTruth)(ctruth: CUSTODES.GroundTruth)(custodes_o: CUSTODES.OutputResult option)(model: ErrorModel)(ranking: CommonTypes.Ranking)(dag: Depends.DAG) : unit =
        let output =
            match custodes_o with
            | Some custodes ->
                match custodes with
                     | CUSTODES.OKOutput(c,_) -> c.Smells
                     | _ -> [||]
            | None -> [||]

        let coutputd = new Dictionary<AST.Address,int>()
        for i in [0..output.Length - 1] do
            coutputd.Add(output.[i], i)

        let smells = new HashSet<AST.Address>(output)

        // append all ExceLint flagged cells
        Array.mapi (fun i (kvp: KeyValuePair<AST.Address,double>) ->
            let addr = kvp.Key
            let per_row = WorkbookStatsRow()
            per_row.Path <- addr.A1Path()
            per_row.Workbook <- addr.WorkbookName
            per_row.Worksheet <- addr.WorksheetName
            per_row.Address <- addr.A1Local()
            per_row.IsFormula <- dag.isFormula addr
            per_row.IsFlaggedByExceLint <- (i <= model.Cutoff)
            per_row.IsFlaggedByCUSTODES <- smells.Contains addr
            per_row.IsFlaggedByExcel <- ctruth.isFlaggedByExcel(addr)
            per_row.CLISameAsV1 <- ctruth.differs addr (smells.Contains addr)
            per_row.ExceLintRank <- i
            per_row.CUSTODESRank <- if coutputd.ContainsKey(addr) then coutputd.[addr] else 999999999
            per_row.Score <- kvp.Value
            per_row.IsExceLintTrueBug <- etruth.IsABug addr
            per_row.IsCUSTODESTrueSmell <- ctruth.isTrueSmell addr

            csv.WriteRow per_row
        ) ranking |> ignore

    let per_append_custodes(csv: WorkbookStats)(etruth: ExceLintGroundTruth)(ctruth: CUSTODES.GroundTruth)(custodes_o: CUSTODES.OutputResult option)(model: ErrorModel)(ranking: CommonTypes.Ranking)(custodes_not_excelint: HashSet<AST.Address>)(dag: Depends.DAG) : unit =
        let output =
            match custodes_o with
            | Some custodes ->
                match custodes with
                     | CUSTODES.OKOutput(c,_) -> c.Smells
                     | _ -> [||]
            | None -> [||]

        let smells = new HashSet<AST.Address>(output)

        // append all remaining CUSTODES cells
        Array.iter (fun (addr: AST.Address) ->
            let per_row = WorkbookStatsRow()
            per_row.Path <- addr.A1Path()
            per_row.Workbook <- addr.WorkbookName
            per_row.Worksheet <- addr.WorksheetName
            per_row.Address <- addr.A1Local()
            per_row.IsFormula <- dag.isFormula addr
            per_row.IsFlaggedByExceLint <- false
            per_row.IsFlaggedByCUSTODES <- true
            per_row.IsFlaggedByExcel <- ctruth.isFlaggedByExcel(addr)
            per_row.CLISameAsV1 <- ctruth.differs addr (smells.Contains addr)
            per_row.ExceLintRank <- 999999999
            per_row.Score <- 0.0
            per_row.IsExceLintTrueBug <- etruth.IsABug addr
            per_row.IsCUSTODESTrueSmell <- ctruth.isTrueSmell addr

            csv.WriteRow per_row
        ) (custodes_not_excelint |> Seq.toArray)

    let per_append_true_smells(csv: WorkbookStats)(etruth: ExceLintGroundTruth)(ctruth: CUSTODES.GroundTruth)(custodes_o: CUSTODES.OutputResult option)(model: ErrorModel)(ranking: CommonTypes.Ranking)(true_smells_not_found: HashSet<AST.Address>)(dag: Depends.DAG) : unit =
        let output =
            match custodes_o with
            | Some custodes ->
                match custodes with
                     | CUSTODES.OKOutput(c,_) -> c.Smells
                     | _ -> [||]
            | None -> [||]

        let smells = new HashSet<AST.Address>(output)

        // append all true smells found by neither tool
        Array.iter (fun (addr: AST.Address) ->
            let per_row = WorkbookStatsRow()
            per_row.Path <- addr.A1Path()
            per_row.Workbook <- addr.WorkbookName
            per_row.Worksheet <- addr.WorksheetName
            per_row.Address <- addr.A1Local()
            per_row.IsFormula <- dag.isFormula addr
            per_row.IsFlaggedByExceLint <- false
            per_row.IsFlaggedByCUSTODES <- false
            per_row.IsFlaggedByExcel <- ctruth.isFlaggedByExcel(addr)
            per_row.CLISameAsV1 <- ctruth.differs addr (smells.Contains addr)
            per_row.ExceLintRank <- 999999999
            per_row.CUSTODESRank <- 999999999
            per_row.Score <- 0.0
            per_row.IsExceLintTrueBug <- etruth.IsABug addr
            per_row.IsCUSTODESTrueSmell <- true

            csv.WriteRow per_row
        ) (true_smells_not_found |> Seq.toArray)

    let per_append_debug(csv: DebugInfo)(model: ErrorModel)(custodes_smells: HashSet<AST.Address>)(config: Args.Config)(ranking: CommonTypes.Ranking) : unit =
        // warn user if CUSTODES analysis contains cells not analyzed by ExceLint
        let rset = Array.map (fun (kvp: KeyValuePair<AST.Address, double>) -> kvp.Key) ranking
                   |> (fun arr -> new HashSet<AST.Address>(arr))
        let not_excelint_at_all = hs_difference custodes_smells rset

        if not_excelint_at_all.Count <> 0 then
            // are these related to formulas?  if so, this is a sign that something went wrong
            let all_formulas = new HashSet<AST.Address>(model.DependenceGraph.getAllFormulaAddrs())
            let missed_formulas = hs_intersection not_excelint_at_all all_formulas

            printfn "WARNING: CUSTODES analysis contains %d cells not analyzed by ExceLint," (not_excelint_at_all.Count)
            printfn "         %d of which are formulas." missed_formulas.Count
            if missed_formulas.Count > 0 then
                printfn "         Writing to %s" config.DebugPath

                // append all true smells found by neither tool
                Array.iter (fun (addr: AST.Address) ->
                    let per_row = DebugInfoRow()
                    per_row.Path <- addr.A1Path()
                    per_row.Workbook <- addr.WorkbookName
                    per_row.Worksheet <- addr.WorksheetName
                    per_row.Address <- addr.A1Local()

                    csv.WriteRow per_row
                ) (missed_formulas |> Seq.toArray)

    let precision(tp: int)(fp: int) : double =
        if tp = 0 && fp = 0 then
            1.0
        else
            let tp' = double tp
            let fp' = double fp
            tp' / (tp' + fp')

    let recall(tp: int)(fn: int) : double =
        if tp = 0 && fn = 0 then
            1.0
        else
            let tp' = double tp
            let fn' = double fn
            tp' / (tp' + fn')

    let append_stats(stats: Stats)(csv: ExceLintStats)(model: ErrorModel)(custodes_o: CUSTODES.OutputResult option)(config: Args.Config) : unit =
        
        let min_excelint_score =
            if model.ranking().Length = 0 then
                0.0
            else
                Array.map (fun (kvp: KeyValuePair<AST.Address,double>) -> kvp.Value) (model.ranking()) |> Array.min

        // write stats
        let row = ExceLintStatsRow()
        row.BenchmarkName <- stats.shortname
        row.NumCells <- model.DependenceGraph.allCells().Length
        row.NumFormulas <- model.DependenceGraph.getAllFormulaAddrs().Length
        row.SigThresh <- stats.threshold
        row.DepTimeMs <- model.DependenceGraph.AnalysisMilliseconds
        row.ScoreTimeMs <- model.ScoreTimeInMilliseconds
        row.FreqTimeMs <- model.FrequencyTableTimeInMilliseconds
        row.RankingTimeMs <- model.RankingTimeInMilliseconds
        row.CausesTimeMs <- model.CausesTimeInMilliseconds
        row.ConditioningSetSzTimeMs <- model.ConditioningSetSizeTimeInMilliseconds
        row.ExceLintFlags <- model.Cutoff + 1
        row.ExceLintPrecisionVsCustodesGT <- precision (stats.excelint_true_smells.Count) (row.ExceLintFlags - stats.excelint_true_smells.Count)
        row.ExceLintRecallVsCustodesGT <- recall (stats.excelint_true_smells.Count) (stats.true_smells_this_wb.Count - stats.excelint_true_smells.Count)
        row.CUSTODESPrecisionVsCustodesGT <- precision (stats.custodes_true_smells.Count) (stats.custodes_flagged.Count - stats.custodes_true_smells.Count)
        row.CUSTODESRecallVsCustodesGT <- recall (stats.custodes_true_smells.Count) (stats.true_smells_this_wb.Count - stats.custodes_true_smells.Count)
        row.MinAnomScore <- min_excelint_score
        row.CUSTODESTimeMs <- stats.custodes_time
        row.CUSTODESFailed <- match custodes_o with | Some custodes -> (match custodes with | CUSTODES.BadOutput _ -> true | _ -> false) | None -> true
        row.CUSTODESFailureMsg <- match custodes_o with | Some custodes -> (match custodes with | CUSTODES.BadOutput(msg,_) -> msg | _ -> "") | None -> "did not run CUSTODES"
        row.NumExceLintTrueRefBugsFound <- stats.excelint_true_ref_bugs.Count
        row.NumCUSTODESTrueRefBugsFound <- stats.custodes_true_ref_bugs.Count
        row.NumCUSTODESSmells <- stats.custodes_flagged.Count
        row.NumTrueSmells <- stats.true_smells_this_wb.Count
        row.NumExceLintTrueSmellsFound <- stats.excelint_true_smells.Count
        row.NumCUSTODESTrueSmellsFound <- stats.custodes_true_smells.Count
        row.NumExceLintCUSTODESTrueSmellsIntersect <- (hs_intersection stats.excelint_true_smells stats.custodes_true_smells).Count
        row.NumTrueSmellsMissedByBoth <- (hs_difference stats.true_smells_this_wb (hs_union stats.excelint_true_smells stats.custodes_true_smells)).Count
        row.NumExcelFlags <- stats.excel_this_wb.Count
        row.NumExceLintExcelIntersect <- stats.excelint_excel_intersect.Count
        row.NumCUSTODESExcelIntersect <- stats.custodes_excel_intersect.Count
        row.NumExcelMissedByBoth <- (hs_difference stats.excel_this_wb (hs_union stats.excelint_excel_intersect stats.custodes_excel_intersect)).Count
        row.OptSpectral <- config.FeatureConf.IsEnabledSpectralRanking
        row.OptCondAllCells <- config.FeatureConf.IsEnabledOptCondAllCells
        row.OptCondRows <- config.FeatureConf.IsEnabledOptCondRows
        row.OptCondCols <- config.FeatureConf.IsEnabledOptCondCols
        row.OptCondLevels <- config.FeatureConf.IsEnabledOptCondLevels
        row.OptCondSheets <- config.FeatureConf.IsEnabledOptCondSheets
        row.OptAddrModeInference <- config.FeatureConf.IsEnabledOptAddrmodeInference
        row.OptWeightIntrinsicAnom <- config.FeatureConf.IsEnabledOptWeightIntrinsicAnomalousness
        row.OptWeightConditionSetSz <- config.FeatureConf.IsEnabledOptWeightConditioningSetSize
        row.ExceLintJaccardDistance <- stats.excelint_jaccard
        row.ExceLintDeltaK <- stats.excelint_delta_k

        csv.WriteRow row

    let kmedioidsJaccardIndex(shortf: string)(model: ErrorModel)(config: Args.Config)(graph: Depends.DAG)(app: Application) : double =
        try
            printfn "Running ExceLint k-medioids analysis: %A" shortf
            let k = model.Clustering.Count
            let ex_clusters = model.Clustering

            let input = CommonTypes.SimpleInput (app.XLApplication()) config.FeatureConf graph
            let km_clusters = KMedioidsClusterModelBuilder.getClustering input k

            // assign IDs to clusters
            let correspondence = CommonFunctions.JaccardCorrespondence km_clusters ex_clusters
            let ex_ids: CommonTypes.ClusterIDs = CommonFunctions.numberClusters ex_clusters
            let mutable maxId = ex_ids.Values |> Seq.max
            let km_ids: CommonTypes.ClusterIDs =
                km_clusters
                |> Seq.map (fun cl ->
                    let ex_cluster_opt = correspondence.[Some cl]
                    match ex_cluster_opt with
                    | Some ex_cluster ->
                        cl, ex_ids.[ex_cluster]
                    | None ->
                        maxId <- maxId + 1
                        cl, maxId
                    ) |> adict

            // write clustering logs
            Clustering.writeClustering(ex_clusters, ex_ids, config.clustering_csv shortf "clustering_excelint")
            Clustering.writeClustering(km_clusters, km_ids, config.clustering_csv shortf "clustering_kmedioids")

            CommonFunctions.ClusteringJaccardIndex km_clusters ex_clusters correspondence
        with
        | _ -> 0.0

    let oldClusterAlgoJaccardIndex(shortf: string)(model: ErrorModel)(config: Args.Config)(graph: Depends.DAG)(app: Application) : double*int =
        try
            printfn "Running old ExceLint cluster analysis: %A" shortf
            let fc' = config.FeatureConf.enableOldClusteringAlgorithm true

            let model_opt' = ExceLint.ModelBuilder.analyze (app.XLApplication()) fc' graph (config.alpha) (Depends.Progress.NOPProgress())
            match model_opt' with
            | Some model' ->
                let ex_k = model.Clustering.Count
                let ex_clusters = model.Clustering
                let old_k = model.Clustering.Count
                let oldex_clusters = model'.Clustering

                // how many more clusters old model has than new one
                let delta_k = old_k - ex_k

                // assign IDs to clusters
                let correspondence = CommonFunctions.JaccardCorrespondence oldex_clusters ex_clusters
                let ex_ids: CommonTypes.ClusterIDs = CommonFunctions.numberClusters ex_clusters
                let mutable maxId = ex_ids.Values |> Seq.max
                let old_ids: CommonTypes.ClusterIDs =
                    oldex_clusters
                    |> Seq.map (fun cl ->
                        let ex_cluster_opt = correspondence.[Some cl]
                        match ex_cluster_opt with
                        | Some ex_cluster ->
                            cl, ex_ids.[ex_cluster]
                        | None ->
                            maxId <- maxId + 1
                            cl, maxId
                       ) |> adict

                // write clustering logs
                Clustering.writeClustering(ex_clusters, ex_ids, config.clustering_csv shortf "clustering_excelint")
                Clustering.writeClustering(oldex_clusters, old_ids, config.clustering_csv shortf "clustering_OLDexcelint")
                    
                CommonFunctions.ClusteringJaccardIndex oldex_clusters ex_clusters correspondence, delta_k
            | None -> 0.0,0
        with
        | _ -> 0.0,0

    let analyze (file: String)(app: Application)(config: Args.Config)(etruth: ExceLintGroundTruth)(ctruth: CUSTODES.GroundTruth)(csv: ExceLintStats)(debug_csv: DebugInfo) =
        let shortf = (System.IO.Path.GetFileName file)

        printfn "Opening: %A" shortf
        let wb = app.OpenWorkbook(file)
            
        printfn "Building dependence graph: %A" shortf
        let graph = wb.buildDependenceGraph()

        printfn "Running ExceLint analysis: %A" shortf
        let model_opt = ExceLint.ModelBuilder.analyze (app.XLApplication()) config.FeatureConf graph (config.alpha) (Depends.Progress.NOPProgress())

        let (jdist,delta_k) =
            match model_opt with
            | Some model ->
                if config.CompareAgainstOldNN then
                    oldClusterAlgoJaccardIndex shortf model config graph app
                elif config.CompareAgainstKMedioid then
                    kmedioidsJaccardIndex shortf model config graph app, 0
                else
                    0.0, 0
            | None -> 0.0, 0

        let custodes_o =
            if not config.DontRunCUSTODES then
                printfn "Running CUSTODES analysis: %A" shortf
                Some (CUSTODES.getOutput(file, config.CustodesPath, config.JavaPath))
            else
                None

        match model_opt with
        | Some(model) ->
            using (new WorkbookStats(config.verbose_csv shortf)) (fun wbstats ->

                let ranking = model.ranking()

                // get the set of cells flagged by ExceLint
                let excelint_flags = rankToSet ranking model.Cutoff
                // get the set of cells in ExceLint's ranking
                let excelint_analyzed = rankToSet ranking (ranking.Length - 1)

                // get workbook name
                let this_wb = wb.WorkbookName

                // get the set of cells flagged by CUSTODES
                let (custodes_total_order,custodes_time) =
                    match custodes_o with
                    | Some custodes ->
                        match custodes with
                        | CUSTODES.OKOutput(c,t) -> c.Smells, t
                        | CUSTODES.BadOutput(_,t) -> [||], t
                    | None -> [||], 0L

                let custodes_flags = new HashSet<AST.Address>(custodes_total_order)

                // find true ref bugs
                let true_ref_bugs_this_wb = etruth.TrueRefBugsByWorkbook this_wb
                let excelint_true_ref_bugs = hs_intersection excelint_flags true_ref_bugs_this_wb
                let custodes_true_ref_bugs = hs_intersection custodes_flags true_ref_bugs_this_wb

                // find true smells found by neither tool
                let true_smells_this_wb = ctruth.TrueSmellsbyWorkbook this_wb
                let true_smells_not_found_by_excelint = hs_difference true_smells_this_wb excelint_flags
                let true_smells_not_found_by_custodes = hs_difference true_smells_this_wb custodes_flags
                let true_smells_not_found = hs_intersection true_smells_not_found_by_excelint true_smells_not_found_by_custodes

                // overall stats
                let excel_this_wb = ctruth.ExcelbyWorkbook this_wb
                let excelint_true_smells = hs_intersection excelint_flags true_smells_this_wb
                assert (excelint_true_smells.IsSubsetOf(excelint_flags))
                let custodes_true_smells = hs_intersection custodes_flags true_smells_this_wb
                assert (custodes_true_smells.IsSubsetOf(custodes_flags))
                let excelint_excel_intersect = hs_intersection excelint_flags excel_this_wb
                let custodes_excel_intersect = hs_intersection custodes_flags excel_this_wb

                let stats = {
                    shortname = shortf;
                    threshold = config.alpha;
                    custodes_flagged = custodes_flags;
                    excelint_not_custodes = hs_difference excelint_flags custodes_flags;
                    custodes_not_excelint = hs_difference custodes_flags excelint_flags;
                    true_ref_bugs_this_wb = true_ref_bugs_this_wb;
                    excelint_true_ref_bugs = excelint_true_ref_bugs;
                    custodes_true_ref_bugs = custodes_true_ref_bugs;
                    true_smells_this_wb = true_smells_this_wb;
                    true_smells_not_found_by_excelint = true_smells_not_found_by_excelint;
                    true_smells_not_found_by_custodes = true_smells_not_found_by_custodes;
                    true_smells_not_found = true_smells_not_found;
                    excel_this_wb =  excel_this_wb;
                    excelint_true_smells = excelint_true_smells;
                    custodes_true_smells = custodes_true_smells;
                    excelint_excel_intersect = excelint_excel_intersect;
                    custodes_excel_intersect = custodes_excel_intersect;
                    custodes_time = custodes_time;
                    excelint_jaccard = jdist;
                    excelint_delta_k = delta_k;
                }

                // write to per-workbook CSV
                per_append_excelint wbstats etruth ctruth custodes_o model ranking graph
                let custodes_not_in_ranking = hs_difference (stats.excelint_not_custodes) excelint_analyzed
                per_append_custodes wbstats etruth ctruth custodes_o model ranking custodes_not_in_ranking graph
                let true_smells_not_in_ranking = hs_difference (hs_difference true_smells_not_found excelint_analyzed) custodes_not_in_ranking
                per_append_true_smells wbstats etruth ctruth custodes_o model ranking true_smells_not_in_ranking graph

                // write overall stats to CSV
                append_stats stats csv model custodes_o config

                // sanity checks
                assert ((hs_intersection excelint_analyzed custodes_not_in_ranking).Count = 0)
                assert ((hs_intersection excelint_analyzed true_smells_not_in_ranking).Count = 0)
                assert ((hs_intersection custodes_not_in_ranking true_smells_not_in_ranking).Count = 0)
                per_append_debug debug_csv model stats.custodes_flagged config ranking
            )

            printfn "Analysis complete: %A" shortf
        | None ->
            printfn "Analysis failed: %A" shortf

    [<EntryPoint>]
    let main argv = 
        let config =
            try
                Args.processArgs argv
            with
            | e ->
                printfn "%A" e.Message
                System.Environment.Exit 1
                failwith "never gets called but keeps F# happy"

        Console.CancelKeyPress.Add(
            (fun _ ->
                printfn "Ctrl-C received.  Cancelling..."
                System.Environment.Exit 1
            )
        )

        using(new Application()) (fun app ->

            let csv = new ExceLintStats(config.csv)
            let debug_csv = new DebugInfo(config.DebugPath)

            let workbook_paths = Array.map (fun fname ->
                                     let wbname = System.IO.Path.GetFileName fname
                                     let path = System.IO.Path.GetDirectoryName fname
                                     wbname, path
                                 ) (config.files) |> adict

            let custodes_gt = new CUSTODES.GroundTruth(workbook_paths, config.CustodesGroundTruthCSV)
            let excelint_gt = ExceLintGroundTruth.Load(config.ExceLintGroundTruthCSV)

            for file in config.files do
                        
                let shortf = (System.IO.Path.GetFileName file)

                try
                    analyze file app config excelint_gt custodes_gt csv debug_csv
                with
                | e ->
                    printfn "Cannot analyze workbook %A because:\n%A" shortf e.Message
                    printfn "Stacktrace:\n%A" e.StackTrace
        )

        if config.DontExitWithoutKeystroke then
            printfn "Press Enter to continue."
            Console.ReadLine() |> ignore

        0
