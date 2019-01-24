open COMWrapper
open System
open System.IO
open System.Collections.Generic
open ExceLint
open ExceLint.Utils
open ExceLintFileFormats
open System.Threading
open MathNet.Numerics.Distributions
open System.Diagnostics
open FastDependenceAnalysis

    type BugClass = HashSet<AST.Address>
    type Stats = {
        shortname: string;
        threshold: double;
        excelint_flagged: HashSet<AST.Address>;
        custodes_flagged: HashSet<AST.Address>;
        excelint_not_custodes: HashSet<AST.Address>;
        custodes_not_excelint: HashSet<AST.Address>;
        num_inconsistent_bugs_this_wb: int;
        excelint_inconsistent_TP: int;
        excelint_inconsistent_FP: int;
        num_missing_formulas_this_wb: int;
        excelint_missing_formula_TP: int;
        excelint_missing_formula_FP: int;
        num_whitespace_ops_this_wb: int;
        excelint_op_on_ws_TP: int;
        excelint_op_on_ws_FP: int;
        custodes_inconsistent_TP: int;
        custodes_inconsistent_FP: int;
        custodes_missing_formula_TP: int;
        custodes_missing_formula_FP: int;
        custodes_op_on_ws_TP: int;
        custodes_op_on_ws_FP: int;
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
        cells: int;
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

    /// <summary>
    /// Computes the expected number of true positives obtained
    /// by flagging cells at random.
    /// </summary>
    /// <param name="m">population size</param>
    /// <param name="r">number of true bugs in population</param>
    /// <param name="n">sample size</param>
    let expectedNumRandomCorrectFlags(m: int)(r: int)(n: int) : double =
        // n * (r/m)
        // where n: sample size
        //       r: # of true bugs in population
        //       m: population size
        (double n) * ((double r) / (double m))

    let PValue(numcells: int)(numbugs: int)(numtp: int)(numflags: int) : double =
        Hypergeometric.PMF(numcells, numbugs, numflags, numtp)

    let per_append_excelint(csv: WorkbookStats)(etruth: ExceLintGroundTruth)(ctruth: CUSTODES.GroundTruth)(custodes_o: CUSTODES.OutputResult option)(model: ErrorModel)(ranking: CommonTypes.Ranking)(dag: Graph) : unit =
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

    let per_append_custodes(csv: WorkbookStats)(etruth: ExceLintGroundTruth)(ctruth: CUSTODES.GroundTruth)(custodes_o: CUSTODES.OutputResult option)(model: ErrorModel)(ranking: CommonTypes.Ranking)(custodes_not_excelint: HashSet<AST.Address>)(dag: Graph) : unit =
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

    let per_append_true_smells(csv: WorkbookStats)(etruth: ExceLintGroundTruth)(ctruth: CUSTODES.GroundTruth)(custodes_o: CUSTODES.OutputResult option)(model: ErrorModel)(ranking: CommonTypes.Ranking)(true_smells_not_found: HashSet<AST.Address>)(dag: Graph) : unit =
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

    let append_stats(stats: Stats)(csv: ExceLintStats)(custodes_o: CUSTODES.OutputResult option)(config: Args.Config)(graphs: Graphs)(models: ErrorModel[]) : unit =
        // write stats
        let row = ExceLintStatsRow()
        row.BenchmarkName <- stats.shortname
        row.NumCells <- graphs.NumCells
        row.NumFormulas <- graphs.NumFormulas
        row.SigThresh <- stats.threshold
        row.TimeMarshalingMs <- graphs.TimeMarshalingMilliseconds
        row.TimeParsingMs <- graphs.TimeParsingMilliseconds
        row.TimeGraphConstruct <- graphs.TimeDependenceAnalysisMilliseconds
        row.ScoreTimeMs <- models |> Array.sumBy (fun model -> model.ScoreTimeInMilliseconds)
        row.FreqTimeMs <- models |> Array.sumBy (fun model -> model.FrequencyTableTimeInMilliseconds)
        row.RankingTimeMs <-  models |> Array.sumBy (fun model -> model.RankingTimeInMilliseconds)
        row.CausesTimeMs <-  models |> Array.sumBy (fun model -> model.CausesTimeInMilliseconds)
        row.ConditioningSetSzTimeMs <-  models |> Array.sumBy (fun model -> model.ConditioningSetSizeTimeInMilliseconds)
        row.ExceLintFlags <- stats.excelint_flagged.Count;
        row.ExceLintPrecisionVsCustodesGT <- precision (stats.excelint_true_smells.Count) (row.ExceLintFlags - stats.excelint_true_smells.Count)
        row.ExceLintRecallVsCustodesGT <- recall (stats.excelint_true_smells.Count) (stats.true_smells_this_wb.Count - stats.excelint_true_smells.Count)
        row.CUSTODESPrecisionVsCustodesGT <- precision (stats.custodes_true_smells.Count) (stats.custodes_flagged.Count - stats.custodes_true_smells.Count)
        row.CUSTODESRecallVsCustodesGT <- recall (stats.custodes_true_smells.Count) (stats.true_smells_this_wb.Count - stats.custodes_true_smells.Count)
        row.CUSTODESTimeMs <- stats.custodes_time
        row.CUSTODESFailed <- match custodes_o with | Some custodes -> (match custodes with | CUSTODES.BadOutput _ -> true | _ -> false) | None -> true
        row.CUSTODESFailureMsg <- match custodes_o with | Some custodes -> (match custodes with | CUSTODES.BadOutput(msg,_) -> msg | _ -> "") | None -> "did not run CUSTODES"
        row.NumTrueRefBugs <- stats.num_inconsistent_bugs_this_wb
        row.ExceLintTrueRefTruePositives <- stats.excelint_inconsistent_TP
        row.ExceLintTrueRefFalsePositives <- stats.excelint_inconsistent_FP
        row.NumMissingFormulaBugs <- stats.num_missing_formulas_this_wb
        row.ExceLintMissingFormulaTruePositives <- stats.excelint_missing_formula_TP
        row.ExceLintMissingFormulaFalsePositives <- stats.excelint_missing_formula_FP
        row.NumWhitespaceOpBugs <- stats.num_whitespace_ops_this_wb
        row.ExceLintWhitespaceOpTruePositives <- stats.excelint_op_on_ws_TP
        row.ExceLintWhitespaceOpFalsePositives <- stats.excelint_op_on_ws_TP
        row.CUSTODESTrueRefTruePositives <- stats.custodes_inconsistent_TP
        row.CUSTODESTrueRefFalsePositives <- stats.custodes_inconsistent_FP
        row.CUSTODESMissingFormulaTruePositives <- stats.custodes_missing_formula_TP
        row.CUSTODESMissingFormulaFalsePositives <- stats.custodes_missing_formula_FP
        row.CUSTODESWhitespaceOpTruePositives <- stats.custodes_op_on_ws_TP
        row.CUSTODESWhitespaceOpFalsePositives <- stats.custodes_op_on_ws_FP
        row.ExceLintPrecisionVsTrueRefBugs <- precision stats.excelint_inconsistent_TP stats.excelint_inconsistent_FP
        row.ExceLintRecallVsTrueRefBugs <- recall stats.excelint_inconsistent_TP (stats.num_inconsistent_bugs_this_wb - stats.excelint_inconsistent_TP)
        row.ExceLintMissingFormulaPrecision <- precision stats.excelint_missing_formula_TP stats.excelint_missing_formula_FP
        row.ExceLintMissingFormulaRecall <- recall stats.excelint_missing_formula_TP (stats.num_missing_formulas_this_wb - stats.excelint_missing_formula_TP)
        row.ExceLintWhitespaceOpPrecision <- precision stats.excelint_op_on_ws_TP stats.excelint_op_on_ws_FP
        row.ExceLintWhitespaceOpRecall <- recall stats.excelint_op_on_ws_TP (stats.num_whitespace_ops_this_wb - stats.excelint_op_on_ws_TP)
        row.CUSTODESPrecisionVsTrueRefBugs <- precision stats.custodes_inconsistent_TP stats.custodes_inconsistent_FP
        row.CUSTODESRecallVsTrueRefBugs <- recall stats.custodes_inconsistent_TP (stats.num_inconsistent_bugs_this_wb - stats.custodes_inconsistent_TP)
        row.CUSTODEStMissingFormulaPrecision <- precision stats.custodes_missing_formula_TP stats.custodes_missing_formula_FP
        row.CUSTODEStMissingFormulaRecall <- recall stats.custodes_missing_formula_TP (stats.num_missing_formulas_this_wb - stats.custodes_missing_formula_TP)
        row.CUSTODEStWhitespaceOpPrecision <- precision stats.custodes_op_on_ws_TP stats.custodes_op_on_ws_FP
        row.CUSTODEStWhitespaceOpRecall <- recall stats.custodes_op_on_ws_TP (stats.num_whitespace_ops_this_wb - stats.custodes_op_on_ws_TP)
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

        csv.WriteRow row

    let oldClusterAlgoJaccardIndex(shortf: string)(model: ErrorModel)(config: Args.Config)(graph: Graph)(app: Application) : double*int =
        try
            printfn "Running old ExceLint cluster analysis: %A" shortf
            let fc' = config.FeatureConf.enableOldClusteringAlgorithm true

            let model_opt' = ExceLint.ModelBuilder.analyze (app.XLApplication()) fc' graph (config.alpha) (Progress.NOPProgress())
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

    let write_flags(cells: HashSet<HashSet<AST.Address>>)(config: Args.Config)(name: string) : unit =
        let path = System.IO.Path.Combine(config.OutputDirectory, name)
        Clustering.writeClustering(cells, path)

    type BugKind =
    | ReferenceError
    | MissingFormula
    | WhitespaceOp
        member self.ToNum =
            match self with
            | ReferenceError -> 0
            | MissingFormula -> 1
            | WhitespaceOp   -> 2
        member self.Discriminator(etruth: ExceLintFileFormats.ExceLintGroundTruth) =
            match self with
            | ReferenceError -> etruth.IsAnInconsistentFormulaBug
            | MissingFormula -> etruth.IsAMissingFormulaBug
            | WhitespaceOp   -> etruth.IsAWhitespaceBug
        member self.ErrorClass =
            match self with
            | ReferenceError -> ErrorClass.INCONSISTENT
            | MissingFormula -> ErrorClass.MISSINGFORMULA
            | WhitespaceOp -> ErrorClass.WHITESPACE

    let dual_for_addr(bc: Dictionary<BugClass,BugClass>)(addr: AST.Address) : KeyValuePair<BugClass,BugClass> option =
        // search bc
        let mutable found_dual = None
        for kvp in bc do
            let left = kvp.Key
            let right = kvp.Value
            if (left.Contains addr) || (right.Contains addr) then
                found_dual <- Some kvp
        found_dual

    let count_TP_for_kind(etruth: ExceLintFileFormats.ExceLintGroundTruth)(wbname: string)(flags: HashSet<AST.Address>)(kind: BugKind) : int =
        let counts = new Dict<BugClass*BugClass,int>()

        let duals = etruth.KindBugsForWorkbook(wbname, kind.ErrorClass)

        for addr in flags do
            let dual = dual_for_addr duals addr
            match dual with
            | Some d ->
                // we know that these are all bugs; get dual
                let tup = d.Key, d.Value
                // count
                if not (counts.ContainsKey tup) then
                    counts.Add(tup, 1)
                else
                    // add one if we have not exceeded our max count for this dual
                    let total_bugs = (etruth.NumBugsForBugClass (d.Key))
                    if counts.[tup] < total_bugs then
                        counts.[tup] <- counts.[tup] + 1
            | None ->
                // is it actually a bug?
                if kind.Discriminator etruth addr then
                    // yes it is
                    // make a singleton bugclass and count it
                    let bugclass = new HashSet<AST.Address>([addr])
                    let duals = (bugclass,bugclass)
                    counts.Add(duals, 1)

        Seq.sum counts.Values

    let count_FP_for_kind(etruth: ExceLintFileFormats.ExceLintGroundTruth)(flags: HashSet<AST.Address>)(kind: BugKind) : int =
        let discriminator =
            match kind with
            | ReferenceError -> etruth.IsAnInconsistentFormulaBug
            | MissingFormula -> etruth.IsAMissingFormulaBug
            | WhitespaceOp   -> etruth.IsAWhitespaceBug

        let mutable i = 0
        for addr in flags do
            if not (discriminator addr) then
                i <- i + 1
        i

    type SoundnessCount = { ncells: int; nnomatch: int; }

    let soundness_count(model_opt: ErrorModel option)(dag: Graph) : SoundnessCount =
        // get analysis base
        let cells = match model_opt with | Some m -> m.AllCells | None -> failwith "does not apply"

        // save set of cells that hashes to the same fingerprint
        let fd = new Dict<Countable,HashSet<AST.Address>>()

        // save all vectors for cells at given address
        let addrv = new Dict<AST.Address,Countable[]>()

        // for each cell, get vectors and fingerprint
        cells |>
        Seq.iter (fun cell ->
            let vs = Vector.ShallowInputVectorMixedFullCVectorResultantOSI.getPaperVectors cell dag |> Array.map (fun v -> v)
            let fingerprint = (Vector.ShallowInputVectorMixedFullCVectorResultantOSI.run cell dag).LocationFree

            // save vectors
            addrv.Add(cell, vs)

            // init hashset
            if not (fd.ContainsKey(fingerprint)) then
                fd.Add(fingerprint, new HashSet<AST.Address>())

            // get set
            let hs = fd.[fingerprint]

            // add to set
            hs.Add cell |> ignore
        )

        // for each fingerprint, count
        // how many of those cells' vector sets do not match
        let mutable nomatch = 0
        fd |>
        Seq.iter (fun (kvp: KeyValuePair<Countable,HashSet<AST.Address>>) -> 
            let addrs = kvp.Value |> Seq.toArray
            if addrs.Length > 1 then
                // get the first set of vectors
                let vs0 = addrv.[addrs.[0]] |> Set.ofArray
                for addr in addrs do
                    // get the second set of vectors
                    let vsi = addrv.[addr] |> Set.ofArray
                    if vs0 <> vsi then
                        nomatch <- nomatch + 1
        )

        { ncells = Seq.length cells; nnomatch = nomatch; }


    let analyze (file: String)(app: Application)(config: Args.Config)(etruth: ExceLintGroundTruth)(ctruth: CUSTODES.GroundTruth)(csv: ExceLintStats)(debug_csv: DebugInfo) =
        let shortf = (System.IO.Path.GetFileName file)

        printfn "Opening: %A" shortf
        let wb = app.OpenWorkbook(file)
            
        printfn "Building dependence graph: %A" shortf
        let graphs = wb.buildDependenceGraph()

        printfn "Running ExceLint analysis: %A" shortf
        
        let model_opts =
            graphs.Worksheets |>
            Array.map (fun g -> g, ExceLint.ModelBuilder.analyze (app.XLApplication()) config.FeatureConf g (config.alpha) (Progress.NOPProgress()))

        let scounts = model_opts |> Array.map (fun (g, model_opt) -> soundness_count model_opt g)
        let scount = scounts |> Array.reduce (fun acc sc -> { ncells = acc.ncells + sc.ncells; nnomatch = acc.nnomatch + sc.nnomatch; })

        let custodes_o =
            if not config.DontRunCUSTODES then
                printfn "Running CUSTODES analysis: %A" shortf
                Some (CUSTODES.getOutput(file, config.CustodesPath, config.JavaPath))
            else
                None

        // get the set of cells flagged by ExceLint
        let excelint_flags =
            model_opts
            |> Array.map (fun (g, model_opt) ->
                   match model_opt with
                   | Some(model) ->
                       let ranking = model.ranking()
                       rankToSet ranking model.Cutoff
                   | None ->
                       printfn "Analysis failed for worksheet: %A" g.Worksheet
                       new HashSet<AST.Address>()
               )
            |> Array.reduce (fun acc hs -> hs_union acc hs)

        using (new WorkbookStats(config.verbose_csv shortf)) (fun wbstats ->
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
            let num_inconsistent_bugs_this_wb = etruth.NumBugKindBugsForWorkbook(this_wb,ErrorClass.INCONSISTENT)
            let excelint_inconsistent_TP = count_TP_for_kind etruth this_wb excelint_flags ReferenceError
            let excelint_inconsistent_FP = count_FP_for_kind etruth excelint_flags ReferenceError
            assert (num_inconsistent_bugs_this_wb >= excelint_inconsistent_TP)

            let custodes_inconsistent_TP = count_TP_for_kind etruth this_wb custodes_flags ReferenceError
            let custodes_inconsistent_FP = count_FP_for_kind etruth custodes_flags ReferenceError
            assert (num_inconsistent_bugs_this_wb >= custodes_inconsistent_TP)

            // find missing formula bugs
            let num_missing_formula_bugs_this_wb = etruth.NumBugKindBugsForWorkbook(this_wb,ErrorClass.MISSINGFORMULA)
            let excelint_missing_formula_TP = count_TP_for_kind etruth this_wb excelint_flags MissingFormula
            let excelint_missing_formula_FP = count_FP_for_kind etruth excelint_flags MissingFormula
            let custodes_missing_formula_TP = count_TP_for_kind etruth this_wb custodes_flags MissingFormula
            let custodes_missing_formula_FP = count_FP_for_kind etruth custodes_flags MissingFormula

            // find whitespace bugs
            let num_whitespace_bugs_this_wb = etruth.NumBugKindBugsForWorkbook(this_wb,ErrorClass.WHITESPACE)
            let excelint_whitespace_TP = count_TP_for_kind etruth this_wb excelint_flags WhitespaceOp
            let excelint_whitespace_FP = count_FP_for_kind etruth excelint_flags WhitespaceOp
            let custodes_whitespace_TP = count_TP_for_kind etruth this_wb custodes_flags WhitespaceOp
            let custodes_whitespace_FP = count_FP_for_kind etruth custodes_flags WhitespaceOp
                
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
                excelint_flagged = excelint_flags;
                custodes_flagged = custodes_flags;
                excelint_not_custodes = hs_difference excelint_flags custodes_flags;
                custodes_not_excelint = hs_difference custodes_flags excelint_flags;
                num_inconsistent_bugs_this_wb = num_inconsistent_bugs_this_wb;
                excelint_inconsistent_TP = excelint_inconsistent_TP;
                excelint_inconsistent_FP = excelint_inconsistent_FP;
                num_missing_formulas_this_wb = num_missing_formula_bugs_this_wb;
                excelint_missing_formula_TP = excelint_missing_formula_TP;
                excelint_missing_formula_FP = excelint_missing_formula_FP;
                num_whitespace_ops_this_wb = num_whitespace_bugs_this_wb;
                excelint_op_on_ws_TP = excelint_whitespace_TP;
                excelint_op_on_ws_FP = excelint_whitespace_FP;
                custodes_inconsistent_TP = custodes_inconsistent_TP;
                custodes_inconsistent_FP = custodes_inconsistent_FP;
                custodes_missing_formula_TP = custodes_missing_formula_TP;
                custodes_missing_formula_FP = custodes_missing_formula_FP;
                custodes_op_on_ws_TP = custodes_whitespace_TP;
                custodes_op_on_ws_FP = custodes_whitespace_FP;
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
                cells = scount.ncells;
            }

            let models = model_opts |> Array.map (fun (_, mo) -> mo) |> Array.choose id

            // write overall stats to CSV
            append_stats stats csv custodes_o config graphs models

            // write set of flagged excelint cells, custodes cells, and true smells to external CSV dump
            let eflags = new HashSet<HashSet<AST.Address>>([excelint_flags])
            let cflags = new HashSet<HashSet<AST.Address>>([custodes_flags])
            let smells = new HashSet<HashSet<AST.Address>>([true_smells_this_wb])
            write_flags eflags config ("excelint_flags-"+this_wb+".csv")
            write_flags cflags config ("custodes_flags-"+this_wb+".csv")
            write_flags smells config ("true_smells-"+this_wb+".csv")
        )

        printfn "Analysis complete: %A" shortf

    let fyshuffle(a: 'a[]) : 'a[] =
        let n = a.Length
        let r = new System.Random()
        for i = 0 to n - 1 do
            let j = r.Next(i,n)
            let tmp = a.[i]
            a.[i] <- a.[j]
            a.[j] <- tmp
        a

    [<EntryPoint>]
    let main argv = 
        // global stopwatch
        let sw = new Stopwatch()
        sw.Start |> ignore

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

            let files = if config.Shuffle then fyshuffle(config.files) else config.files
            for file in files do
                        
                let shortf = (System.IO.Path.GetFileName file)

                if not config.TrueRefOnly || (config.TrueRefOnly && excelint_gt.HasTrueRefAnnotations shortf) then
                    try
                        analyze file app config excelint_gt custodes_gt csv debug_csv
                    with
                    | e ->
                        printfn "Cannot analyze workbook %A because:\n%A" shortf e.Message
                        printfn "Stacktrace:\n%A" e.StackTrace
        )

        sw.Stop |> ignore
        printfn "Total elapsed time: %A seconds" ((float sw.ElapsedMilliseconds) / 1000.0)

        if config.DontExitWithoutKeystroke then
            printfn "Press Enter to continue."
            Console.ReadLine() |> ignore

        0
