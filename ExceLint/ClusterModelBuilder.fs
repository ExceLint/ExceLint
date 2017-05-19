namespace ExceLint
    open System
    open System.Collections.Generic
    open Utils
    open CommonTypes
    open CommonFunctions

    module ClusterModelBuilder =
        // define distance
        let min_dist(hb_inv: InvertedHistogram) =
            (fun (source: HashSet<AST.Address>)(target: HashSet<AST.Address>) ->
                let pairs = cartesianProduct source target

                let compute = (fun (a,b) ->
                                let (_,_,ac) = hb_inv.[a]
                                let (_,_,bc) = hb_inv.[b]
                                ac.EuclideanDistance bc
                                )
                let f = (fun (a,b) ->
                            compute (a,b)
                        )
                let minpair = argmin f pairs
                let dist = f minpair
                dist * (double source.Count)
            )

        // this is kinda-sorta EMD; it has no notion of flows because I have
        // no idea what that means in terms of spreadsheet formula fixes
        let earth_movers_dist(hb_inv: InvertedHistogram) =
            (fun (source: HashSet<AST.Address>)(target: HashSet<AST.Address>) ->
                let compute = (fun (a,b) ->
                                let (_,_,ac) = hb_inv.[a]
                                let (_,_,bc) = hb_inv.[b]
                                ac.EuclideanDistance bc
                                )

                let f = (fun (a,b) ->
                            compute (a,b)
                        )

                // for every point in source, find the closest point in target
                let ds = source
                            |> Seq.map (fun addr ->
                                let pairs = target |> Seq.map (fun t -> addr,t)
                                let min_dist: double = pairs |> Seq.map f |> Seq.min
                                min_dist
                            )

                Seq.sum ds
            )

        // define distance (min distance between clusters)
        let cent_dist(hb_inv: InvertedHistogram) =
            (fun (source: HashSet<AST.Address>)(target: HashSet<AST.Address>) ->
                // Euclidean distance with a small twist:
                // The distance between any two cells on different
                // sheets is defined as infinity.
                // This ensures that we always agglomerate intra-sheet
                // before agglomerating inter-sheet.
                let s_centroid = centroid source hb_inv
                let t_centroid = centroid target hb_inv
                let dist = if s_centroid.SameSheet t_centroid then
                                s_centroid.EuclideanDistance t_centroid
                            else
                                Double.PositiveInfinity
                dist * (double source.Count)
            )

        // find s vector guaranteed to be longer than any of the given vectors
        let diagonalScaleFactor(ss: ScoreTable) : double =
            let points = ss
                            |> Seq.map (fun kvp -> kvp.Value |> Seq.map (fun (_,c: Countable) -> c.Location))
                            |> Seq.concat
                            |> Seq.toArray
            let min_init = points.[0]
            let max_init = points.[0]
            let minf = (fun (a: double)(b: double) -> if a < b then a else b )
            let maxf = (fun (a: double)(b: double) -> if a > b then a else b )
            let min = points |> Seq.fold (fun (minc:Countable)(c:Countable) -> minc.ElementwiseOp c minf) min_init
            let max = points |> Seq.fold (fun (maxc:Countable)(c:Countable) -> maxc.ElementwiseOp c maxf) max_init
            min.EuclideanDistance max

        /// <summary>
        /// Scales the countables in the ScoreTable by a per-
        /// sheet diagonal scale factor.
        /// </summary>
        /// <param name="s"></param>
        let private ScaleBySheet(s: ScoreTable) : ScoreTable =
            // scale each sheet separately
            s
            |> Seq.map (fun kvp ->
                    let feature = kvp.Key
                    let dist = kvp.Value
                    let scaledDist =
                        dist
                        |> Seq.groupBy (fun (addr,co) ->
                                addr.WorksheetName
                            )
                        |> Seq.map (fun (wsname,cells) ->
                                // this is conditioned on sheet name, so
                                // turn into a ScoreTable again
                                let acells = cells |> Seq.toArray
                                let stForSheet: ScoreTable = [(feature,acells)] |> adict
                                // compute scale factor
                                let factor = diagonalScaleFactor stForSheet
                                // scale each vector unless it is off-sheet
                                cells |>
                                Seq.map (fun (addr,co) ->
                                    if co.IsOffSheet then
                                        // don't scale
                                        addr,co
                                    else
                                        // scale
                                        addr,co.UpdateResultant (co.ToCVectorResultant.ScalarMultiply factor)
                                )
                                
                            )
                        |> Seq.concat
                        |> Seq.toArray
                    feature, scaledDist
                )
            |> adict

        // the total sum of squares
        let TSS(C: Clustering)(ih: InvertedHistogram) : double =
            let all_observations = C |> Seq.concat |> Seq.toArray |> Array.map (fun addr -> ToCountable addr ih)
            let n = all_observations.Length
            let mean = Countable.Mean all_observations 

            [|0..n-1|]
            |> Array.sumBy (fun i ->
                    let obs = all_observations.[i]
                    let error = obs.Sub(mean)
                    error.VectorMultiply(error)
                )

        // the within-cluster sum of squares
        let WCSS(C: Clustering)(ih: InvertedHistogram) : double =
            let k = C.Count
            let clusters = C |> Seq.toArray |> Array.map (fun c -> c |> Seq.toArray |> Array.map (fun addr -> ToCountable addr ih))
            let means = clusters |> Array.map (fun c -> Countable.Mean(c))
            let ns = clusters |> Array.map (fun cs -> cs.Length)

            // for every cluster
            [|0..k-1|]
            |> Array.sumBy (fun i ->
                let ni = ns.[i]

                // for every observation
                [|0..ni-1|]
                |> Array.sumBy (fun j ->
                        // compute squared error
                        let obs = clusters.[i].[j]
                        let error = obs.Sub(means.[i])
                        error.VectorMultiply(error)
                    )
                // and double sum
            )

        // the between-cluster sum of squares
        let BCSS(C: Clustering)(ih: InvertedHistogram) : double =
            let k = C.Count
            let clusters = C |> Seq.toArray |> Array.map (fun c -> c |> Seq.toArray |> Array.map (fun addr -> ToCountable addr ih))
            let means = clusters |> Array.map (fun c -> Countable.Mean(c))
            let mean = clusters |> Array.concat
                                |> (fun cs -> Countable.Mean(cs))
            let ns = clusters |> Array.map (fun cs -> cs.Length)
            let n = Array.sum ns

            // for every cluster
            [|0..k-1|]
            |> Array.sumBy (fun i ->
                let ni = ns.[i]
                let error = means.[i].Sub(mean)
                (double ni) * error.VectorMultiply(error)
            )

        let F(C: Clustering)(ih: InvertedHistogram) : double =
            let k = double C.Count
            let n = double (C |> Seq.sumBy (fun cl -> cl.Count))

            // variance is sum of squared error divided by sample size
            let bc_var = (BCSS C ih) / (k - 1.0)
            let wc_var = (WCSS C ih) / (n - k)

            // F is the ratio of the between-cluster variance to the within-cluster variance
            bc_var / wc_var

        type ClusterModel(input: Input) =
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

            let refvect_same(source: HashSet<AST.Address>)(target: HashSet<AST.Address>) : bool =
                // check that all of the location-free vectors in the source and target are the same
                let slf = source
                            |> Seq.map (fun addr ->
                                            let (_,_,ac) = hb_inv.[addr]
                                            ac.LocationFree
                                        ) |> Seq.toArray
                let tlf = target
                            |> Seq.map (fun addr ->
                                            let (_,_,ac) = hb_inv.[addr]
                                            ac.LocationFree
                                        ) |> Seq.toArray
                (Array.forall (fun s -> s.Equals(slf.[0])) slf) &&
                (Array.forall (fun t -> t.Equals(slf.[0])) tlf)

            let pp(c: Set<AST.Address>) : string =
                c
                |> Seq.map (fun a -> a.A1Local())
                |> (fun addrs -> "[" + String.Join(",", addrs) + "] with centroid " + (centroid c hb_inv).ToString())

            let mutable log: ClusterStep list = []
            let mutable per_log = []
            let mutable steps_ms: int64 list = []

            // DEFINE DISTANCE
            let DISTANCE =
                match input.config.DistanceMetric with
                | DistanceMetric.NearestNeighbor -> min_dist hb_inv
                | DistanceMetric.EarthMover -> earth_movers_dist hb_inv
                | DistanceMetric.MeanCentroid -> cent_dist hb_inv

            // compute initial NN table
            let keymaker = (fun (addr: AST.Address) ->
                                let (_,_,co) = hb_inv.[addr]
                                LSHCalc.h7 co
                            )
            let keyexists = (fun addr1 addr2 ->
                                failwith "Duplicate keys should not happen."
                            )
            let hs = HashSpace<AST.Address>(cells, keymaker, keyexists, LSHCalc.h7unmasker, DISTANCE)

            let mutable probable_knee = false

            member self.CanStep : bool =
                hs.NearestNeighborTable.Length > 1
            member private self.IsKnee(s: HashSet<AST.Address>)(t: HashSet<AST.Address>) : bool =
                // the first time we merge two clusters that have
                // different resultants, we've probably hit the knee
                not (refvect_same s t)

            // determine whether the next step will be the knee
            // without actually doing anohter agglomeration step
            member self.NextStepIsKnee : bool =
                let nn_next = hs.NextNearestNeighbor
                let source = nn_next.FromCluster
                let target = nn_next.ToCluster
                self.IsKnee source target

            member self.Step() : bool =
                let sw = new System.Diagnostics.Stopwatch()
                sw.Start()

                // get the two clusters that minimize distance
                let nn_next = hs.NextNearestNeighbor
                let source = nn_next.FromCluster
                let target = nn_next.ToCluster

                if self.IsKnee source target then
                    probable_knee <- true

                // record merge in log
                let clusters = hs.Clusters
                log <- {
                            beyond_knee = probable_knee;
                            source = Set.ofSeq source;
                            target = Set.ofSeq target;
                            distance = DISTANCE source target;
                            f = F clusters hb_inv;
                            within_cluster_sum_squares = WCSS clusters hb_inv;
                            between_cluster_sum_squares = BCSS clusters hb_inv;
                            total_sum_squares = TSS clusters hb_inv;
                            num_clusters = clusters.Count;
                        } :: log

                // dump clusters to log
                let mutable clusterlog = []
                hs.Clusters
                |> Seq.iter (fun cl ->
                    cl
                    |> Seq.iter (fun addr ->
                        let v = ToCountable addr hb_inv
                        match v with
                        | FullCVectorResultant(x,y,z,dx,dy,dz,dc) ->
                            let row = new ExceLintFileFormats.VectorDumpRow()
                            row.clusterID <- hs.ClusterID cl
                            row.x <- x
                            row.y <- y
                            row.z <- z
                            row.dx <- dx
                            row.dy <- dy
                            row.dz <- dz
                            row.dc <- dc
                            clusterlog <- row :: clusterlog
                        | _ -> () // we don't care about other vector types
                    )
                )
                per_log <- (List.rev clusterlog) :: per_log

                // merge them
                hs.Merge source target

                sw.Stop()
                steps_ms <- sw.ElapsedMilliseconds :: steps_ms

                // tell the user whether more steps remain
                self.CanStep

            member self.WritePerLogs() =
                (List.rev per_log)
                |> List.iteri (fun i per_log ->
                    // open file
                    let veccsvw = new ExceLintFileFormats.VectorDump("C:\\Users\\dbarowy\\Desktop\\clusterdump\\vectorstep" + i.ToString() + ".csv")

                    // write rows
                    List.iter (fun row ->
                        veccsvw.WriteRow row
                    ) per_log

                    // close file
                    veccsvw.Dispose()
                )
                        
            member self.WriteLog() =
                // init CSV writer
                let csvw = new ExceLintFileFormats.ClusterSteps("C:\\Users\\dbarowy\\Desktop\\clusterdump\\clustersteps.csv")

                // write rows
                List.rev log
                |> List.iter (fun step ->
                        let row = new ExceLintFileFormats.ClusterStepsRow()
                        row.Show <- step.beyond_knee
                        row.Merge <- (pp step.source) + " with " + (pp step.target)
                        row.Distance <- step.distance
                        row.FScore <- step.f
                        row.WCSS <- step.within_cluster_sum_squares
                        row.BCSS <- step.between_cluster_sum_squares
                        row.TSS <- step.total_sum_squares
                        row.k <- step.num_clusters

                        csvw.WriteRow row      
                    )

                // close file
                csvw.Dispose()

            member self.Clustering = hs.Clusters

            member self.Ranking =
                let numfrm = cells.Length

                // keep a record of reported cells
                let rptd = new HashSet<AST.Address>()

                // for each step in the log,
                // add each source address and distance (score) to the ranking
                let rnk = 
                    log
                    |> List.rev
                    |> List.map (fun (step : ClusterStep) ->
                            if step.beyond_knee then
                            Some(
                                Seq.map (fun addr -> 
                                    if not (rptd.Contains addr) then
                                        let retval = Some(new KeyValuePair<AST.Address,double>(addr, step.distance))
                                        rptd.Add(addr) |> ignore
                                        retval
                                    else
                                        None
                                ) step.source
                                |> Seq.choose id
                                |> Seq.toList
                            )
                            else
                            None
                        )
                    |> List.choose id
                    |> List.concat
                    |> List.toArray

                // the ranking must contain all the formulas and nothing more;
                // some formulas may never be reported depending on location
                // of knee
                assert (rnk.Length <= numfrm)

                rnk

            member self.RankingTimeMs = List.sum steps_ms
            member self.ScoreTimeMs = score_time
            member self.Scores = nlfrs
            member self.Cutoff = self.Ranking.Length - 1
            member self.LSHTree = hs.HashTree

        type LSHViz(input: Input) =
            let c = ClusterModel input
            member self.ToGraphViz = c.LSHTree.ToGraphViz

        let runClusterModel(input: Input) : AnalysisOutcome =
            try
                if (analysisBase input.config input.dag).Length <> 0 then
                    let m = ClusterModel input

                    let mutable notdone = true
                    while notdone do
                        notdone <- m.Step()

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