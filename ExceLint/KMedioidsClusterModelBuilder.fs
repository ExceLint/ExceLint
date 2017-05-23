namespace ExceLint
    open System
    open System.Collections.Generic
    open Utils
    open CommonTypes
    open CommonFunctions
    open Distances

    type MedoidClustering = Dict<AST.Address,HashSet<AST.Address>>

    module KMedioidsClusterModelBuilder =
        let private dist(d: DistanceF)(addr1: AST.Address)(addr2: AST.Address) : double =
            let addr1hs = new HashSet<AST.Address>([addr1])
            let addr2hs = new HashSet<AST.Address>([addr2])
            d addr1hs addr2hs

        let private findClosestMedioids(medioids: AST.Address[])(points: AST.Address[])(d: DistanceF) : MedoidClustering =
            let clustering =
                points |>
                Array.fold (fun (acc: MedoidClustering)(point: AST.Address) ->
                    let closest = medioids |> argmin (fun m -> dist d point m)
                    if acc.ContainsKey closest then
                        acc.[closest].Add point |> ignore
                    else
                        let hs = new HashSet<AST.Address>()
                        hs.Add(point) |> ignore
                        acc.Add(closest, hs) |> ignore
                    acc
                ) (new Dict<AST.Address,HashSet<AST.Address>>())
            clustering

        let private kmedioids(k: int)(cells: AST.Address[])(s: ScoreTable)(hb_inv: InvertedHistogram)(d: DistanceF)(r: Random) : Clustering =
            // choose k random indices using rejection sampling
            let seeds = Array.fold (fun xs i ->
                            let mutable ri = r.Next(cells.Length)
                            while Set.contains ri xs do
                                ri <- r.Next(cells.Length)
                            Set.add ri xs
                        ) Set.empty [| 0 .. k - 1 |]
                        |> Set.toArray

            // get the medioid point representations
            let mutable medioids: AST.Address[] = seeds |> Array.map (fun i -> cells.[i])

            // associate every point with its closest medioid
            let mutable clustering = findClosestMedioids medioids cells d

            let mutable cost_decreased = true

            // while the cost decreases, find new medoids
            while cost_decreased do
                cost_decreased <- false

                // for each medioid
                for m in medioids do
                    // swap medioid and any other point
                    for c in cells do
                        
                        failwith "hang on!"


            failwith "gettin there"


        let runClusterModel(input: Input)(k: int) : AnalysisOutcome =
            try
                if (analysisBase input.config input.dag).Length <> 0 then
                    // initialize selector cache
                    let selcache = Scope.SelectorCache()

                    // determine the set of cells to be analyzed
                    let cells = analysisBase input.config input.dag

                    // get all NLFRs for every formula cell
                    let _runf = fun () -> runEnabledFeatures cells input.dag input.config input.progress
                    let (ns: ScoreTable,score_time: int64) = PerfUtils.runMillis _runf ()

                    // scale
                    let nlfrs: ScoreTable = ScaleBySheet ns

                    // make HistoBin lookup by address
                    let hb_inv = invertedHistogram nlfrs selcache input.dag input.config

                    // DEFINE DISTANCE
                    let DISTANCE =
                        match input.config.DistanceMetric with
                        | DistanceMetric.NearestNeighbor -> min_dist hb_inv
                        | DistanceMetric.EarthMover -> earth_movers_dist hb_inv
                        | DistanceMetric.MeanCentroid -> cent_dist hb_inv

                    // init RNG
                    let r = new Random()

                    // get clustering
                    let c = kmedioids k cells nlfrs hb_inv DISTANCE r

                    Success(Cluster
                        {
                            scores = m.Scores;
                            ranking = m.Ranking;
                            score_time = m.ScoreTimeMs;
                            ranking_time = m.RankingTimeMs;
                            sig_threshold_idx = 0;
                            cutoff_idx = m.Cutoff;
                            weights = new Dictionary<AST.Address,double>();
                        }
                    )
                else
                    CantRun "Cannot perform analysis. This worksheet contains no formulas."
            with
            | AnalysisCancelled -> Cancellation