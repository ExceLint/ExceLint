namespace ExceLint
    open System
    open System.Collections.Generic
    open Utils
    open CommonTypes
    open CommonFunctions

    module COFModelBuilder =

        let private getCOFFixes(scores: ScoreTable)(dag: Depends.DAG)(normalizeRefSpace: bool)(normalizeSSSpace: bool)(BDD: Dictionary<string,Dictionary<AST.Address,Vector.SquareVector>>)(DDs: Dictionary<Vector.WorksheetName,Vector.DistDict>) : Dictionary<AST.Address,HashSet<AST.Address>> =
            let d = new Dictionary<AST.Address,HashSet<AST.Address>>()

            for kvp in scores do
                for (addr,score) in kvp.Value do
                    let k = Vector.COFk addr dag normalizeRefSpace normalizeSSSpace
                    let DD = DDs.[addr.WorksheetName]
                    let nb = Vector.Nk_cells addr k dag normalizeRefSpace normalizeSSSpace BDD DD
                    if not (d.ContainsKey addr) then
                        d.Add(addr, nb)
                    else
                        d.[addr] <- HsUnion d.[addr] nb

            d

        let private rankCOFScores(scores: ScoreTable) : Ranking =
            let d = new Dictionary<AST.Address,double>()
                
            for kvp in scores do
                for (addr,score) in kvp.Value do
                    let dscore: double = match score with
                                            | Num n -> n
                                            | _ -> failwith "Wrong Countable type for COF."
                    if not (d.ContainsKey addr) then
                        d.Add(addr, dscore)
                    else
                        d.[addr] <- d.[addr] + dscore

            Seq.sortBy (fun (kvp: KeyValuePair<AST.Address,double>) -> kvp.Value) d
            |> Seq.toArray

        let runCOFModel(input: Input) : AnalysisOutcome =
            try
                let cells = (analysisBase input.config input.dag)

                let _runf = fun () -> runEnabledFeatures cells input.dag input.config input.progress

                // get COF scores for each feature: featurename -> (address, score)[]
                let (scores: ScoreTable,score_time: int64) = PerfUtils.runMillis _runf ()

                // rank
                let _rankf = fun () -> rankCOFScores scores
                let (ranking,ranking_time) = PerfUtils.runMillis _rankf ()

                // get fixes
                let _fixf = fun () -> getCOFFixes scores input.dag input.config.NormalizeRefs input.config.NormalizeSS (input.config.BDD input.dag) (input.config.DD input.dag)
                let (fixes,fixes_time) = PerfUtils.runMillis _fixf ()

                Success(COF
                    {
                        scores = scores;
                        ranking = ranking;
                        fixes = fixes;
                        fixes_time = fixes_time;
                        score_time = score_time;
                        ranking_time = ranking_time;
                        sig_threshold_idx = 0;
                        cutoff_idx = 0;
                        weights = new Dictionary<AST.Address,double>();
                    }
                )
            with
            | AnalysisCancelled -> Cancellation