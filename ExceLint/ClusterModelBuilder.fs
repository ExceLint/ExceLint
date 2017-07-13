namespace ExceLint
    open System
    open System.Collections.Generic
    open Utils
    open CommonTypes
    open CommonFunctions
    open Distances

    module ClusterModelBuilder =
        /// <summary>
        /// Compute Anselin's local Moran for the given point and neighbors (Anselin 1995).
        /// </summary>
        /// <param name="point">An address.</param>
        /// <param name="neighbors">A set of addresses.</param>
        /// <param name="z">Value function for point at given address.</param>
        /// <param name="w">Neighborhood weight for points at given addresses.</param>
        let LISA(point: AST.Address)(neighbors: HashSet<AST.Address>)(z: AST.Address -> double)(w: AST.Address -> AST.Address -> double) : double =
            let nsWithoutPt = HashSetUtils.differenceElem neighbors point
            let zmean = BasicStats.mean (nsWithoutPt |> Seq.map z |> Seq.toArray)
            (z point - zmean) * (neighbors |> Seq.sumBy (fun neigh -> (w point neigh) * (z neigh - zmean)))

        // define "neighbor" relation
        let W(a1: AST.Address)(a2: AST.Address) : double =
            if a1.SameAs a2 then
                0.0
            else if isAdjacent a1 a2 then
                1.0
            else
                0.0

        // define value function
        let X(a: AST.Address)(dag: Depends.DAG)(conf: FeatureConf)(s: FlatScoreTable) : double =
            // enabled features (in principle, ExceLint can have many enabled simultaneously)
            assert (conf.EnabledFeatures.Length = 1)

            // the one enabled feature
            let feat = conf.EnabledFeatures.[0]

            if dag.isFormula a then
                // value is L2 norm of non-location resultant if
                // address is a formula
                s.[feat,a].ToCVectorResultant.L2Norm
//                s.[feat,a].L2Norm
            else
                // is it a number?

                // read value
                let value = dag.readCOMValueAtAddress a
                let mutable num = 0.0

                if Double.TryParse(value, &num) then
                    // numbers
                    num
                else if String.IsNullOrWhiteSpace value then
                    // blanks
                    0.0
                else
                    // text strings
                    0.0

        /// <summary>
        /// Compute Moran's I for the given set of points.
        /// </summary>
        /// <param name="points">A set of addresses.</param>
        /// <param name="x">Value function for point at given address.</param>
        /// <param name="w">Weight function for two points.</param>
        let Moran(points: HashSet<AST.Address>)(x: AST.Address -> double)(w: AST.Address -> AST.Address -> double) : double =
            let xmean = BasicStats.mean (points |> Seq.map x |> Seq.toArray)

            let N = double (Seq.length points)
            let pairs = CommonFunctions.cartesianProduct points points
            let W = pairs
                    |> Seq.map (fun (i,j) -> w i j)
                    |> Seq.sum
            let scale = N / W

            let debug_xs = Seq.map x points
            let debug_ws = Seq.map (fun (i,j) -> w i j) pairs

            let numerator = pairs
                            |> Seq.sumBy (fun (i, j) -> (w i j) * (x i - xmean) * (x j - xmean))
            let denominator = points
                              |> Seq.sumBy (fun i -> (x i - xmean) * (x i - xmean))

            if N = 1.0 then
                // singleton clusters might as well be random
                0.0
            else if denominator = 0.0 then
                // any time the denominator is 0, we either have a
                // singleton cluster (handled above) or all of
                // the values are exactly the same-- a perfect cluster
                1.0
            else
                scale * numerator / denominator

        /// <summary>
        /// Compute Moran's I for the given set of points. C#-friendly.
        /// </summary>
        /// <param name="points">A set of addresses.</param>
        /// <param name="x">Value function for point at given address.</param>
        /// <param name="w">Weight function for two points.</param>
        let MoranCS(points: HashSet<AST.Address>, x: Func<AST.Address,double>, w: Func<AST.Address, AST.Address, double>) : double =
            let x' = fun a -> x.Invoke a
            let w' = fun a1 a2 -> w.Invoke(a1, a2)
            Moran points x' w'

        let refvect_same(source: HashSet<AST.Address>)(target: HashSet<AST.Address>)(hb_inv: InvertedHistogram) : bool =
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

        let pp(c: Set<AST.Address>)(hb_inv: InvertedHistogram) : string =
            c
            |> Seq.map (fun a -> a.A1Local())
            |> (fun addrs -> "[" + String.Join(",", addrs) + "] with centroid " + (centroid c hb_inv).ToString())

        type ClusterModel(input: Input) =
            let mutable clusteringAtKnee = None
            let mutable inCriticalRegion = false

            // track merge targets so that we
            // can tell when they become merge sources
            let merge_targets = new HashSet<HashSet<AST.Address>>()

            // determine the set of cells to be analyzed
            let cells = analysisBase input.config input.dag

            // get all NLFRs for every formula cell
            let _runf = fun () -> runEnabledFeatures cells input.dag input.config input.progress
            let (ns: ScoreTable,feat_time: int64) = PerfUtils.runMillis _runf ()

            // scale
            let _runscale = fun () -> ScaleBySheet ns
            let (nlfrs: ScoreTable,scale_time: int64) = PerfUtils.runMillis _runscale () 

            // flatten
            let _runflatten = fun () -> makeFlatScoreTable nlfrs
            let (fnlfrs: FlatScoreTable,flatten_time: int64) = PerfUtils.runMillis _runflatten ()

            // make HistoBin lookup by address
            let _runhisto = fun () -> invertedHistogram nlfrs input.dag input.config
            let (hb_inv: InvertedHistogram,invert_time: int64) = PerfUtils.runMillis _runhisto ()

            let score_time = feat_time + scale_time + invert_time + flatten_time

            let mutable log: ClusterStep list = []
            let mutable per_log = []
            let mutable steps_ms: int64 list = []

            // DEFINE DISTANCE
            let DISTANCE =
                match input.config.DistanceMetric with
                | DistanceMetric.NearestNeighbor -> min_dist hb_inv
                | DistanceMetric.EarthMover -> earth_movers_dist hb_inv
                | DistanceMetric.MeanCentroid -> cent_dist hb_inv

            // do region inference
            let rTree = BinaryMinEntropyTree.Infer cells hb_inv
            let regions = BinaryMinEntropyTree.RectangularClustering rTree hb_inv

            // compute NN table
            let keymaker = (fun (addr: AST.Address) ->
                                let (_,_,co) = hb_inv.[addr]
                                LSHCalc.h7 co
                            )
            let keyexists = (fun addr1 addr2 ->
                                failwith "Duplicate keys should not happen."
                            )
            let hs = HashSpace<AST.Address>(regions, keymaker, keyexists, LSHCalc.h7unmasker, DISTANCE)

            let mutable probable_knee = false

            member self.InitialClustering : Clustering =
                let regions' = CopyImmutableToMutableClustering regions
                let isRect = BinaryMinEntropyTree.ClusteringContainsOnlyRectangles regions'
                assert isRect
                regions'

            member self.NumCells : int = cells.Length
            member self.CanStep : bool =
                Seq.length (hs.NearestNeighborTable) > 1

            member private self.IsKnee(s: HashSet<AST.Address>)(t: HashSet<AST.Address>) : bool =
                // the first time we merge two clusters that have
                // different resultants, we've probably hit the knee
                not (refvect_same s t hb_inv)

            member self.NearestNeighborForCluster(c: HashSet<AST.Address>) : HashSet<AST.Address> =
                hs.NearestNeighbor c

            member private self.IsFoot(s: HashSet<AST.Address>) : bool =
                // the first time a former merge target becomes
                // a merge source, we're probably done
                merge_targets.Contains s

            // determine whether the next step will be the knee
            // without actually doing another agglomeration step
            member self.NextStepIsKnee : bool =
                if not (Seq.isEmpty hs.NearestNeighborTable) then
                    let nn_next = hs.NextNearestNeighbor
                    let source = nn_next.FromCluster
                    let target = nn_next.ToCluster
                    self.IsKnee source target
                else
                    // if there are no nearest neighbors, it's becase
                    // all the formulas are exactly the same, and thus
                    // there is no knee
                    true

            member self.Step() : bool =
                // update progress
                input.progress.IncrementCounter()

                let sw = new System.Diagnostics.Stopwatch()
                sw.Start()

                // get the two clusters that minimize distance
                let nn_next = hs.NextNearestNeighbor
                let source = nn_next.FromCluster
                let target = nn_next.ToCluster

                if self.IsKnee source target then
                    // only update ONCE
                    clusteringAtKnee <-
                        match clusteringAtKnee with
                        | Some cak -> Some cak
                        | None -> Some (CopyClustering hs.Clusters)

                    inCriticalRegion <- true

                if self.IsFoot source then
                    inCriticalRegion <- false

                // record merge in log
                let clusters = hs.Clusters
                log <- {
                            source = Set.ofSeq source;
                            target = Set.ofSeq target;
                            distance = DISTANCE source target;
                            // NOTE: Stats commented out because they are
                            //       expensive to compute and not obviously
                            //       useful.
//                            f = F clusters hb_inv;
//                            within_cluster_sum_squares = WCSS clusters hb_inv;
//                            between_cluster_sum_squares = BCSS clusters hb_inv;
//                            total_sum_squares = TSS clusters hb_inv;
                            f = 0.0;
                            within_cluster_sum_squares = 0.0;
                            between_cluster_sum_squares = 0.0;
                            total_sum_squares = 0.0;
                            num_clusters = clusters.Count;
                            in_critical_region = inCriticalRegion;
                        } :: log

                // dump clusters to log if debugging
                if input.config.DebugMode then
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
                merge_targets.Add target |> ignore

                // tell the user whether more steps remain
                let canstep = self.CanStep

                sw.Stop()
                steps_ms <- sw.ElapsedMilliseconds :: steps_ms

                canstep

            member self.WritePerLogs(path: string) =
                (List.rev per_log)
                |> List.iteri (fun i per_log ->
                    // open file
                    let veccsvw = new ExceLintFileFormats.VectorDump(path + "\\vectorstep" + i.ToString() + ".csv")

                    // write rows
                    List.iter (fun row ->
                        veccsvw.WriteRow row
                    ) per_log

                    // close file
                    veccsvw.Dispose()
                )
                        
            member self.WriteLog(path: string) =
                // init CSV writer
                let csvw = new ExceLintFileFormats.ClusterSteps(path + "\\clustersteps.csv")

                // write rows
                List.rev log
                |> List.iter (fun step ->
                        let row = new ExceLintFileFormats.ClusterStepsRow()
                        row.Show <- step.in_critical_region
                        row.Merge <- (pp step.source hb_inv) + " with " + (pp step.target hb_inv)
                        row.Distance <- step.distance
                        row.FScore <- step.f
                        row.WCSS <- step.within_cluster_sum_squares
                        row.BCSS <- step.between_cluster_sum_squares
                        row.TSS <- step.total_sum_squares
                        row.WCSS <- 0.0
                        row.BCSS <- 0.0
                        row.TSS <- 0.0
                        row.k <- step.num_clusters

                        csvw.WriteRow row      
                    )

                // close file
                csvw.Dispose()

            member self.CurrentClustering = hs.Clusters

            member self.ClusteringAtKnee =
                match clusteringAtKnee with
                | Some clustering -> clustering
                    // if there was no knee, then
                    // the knee and the foot are the same
                    // and it's the latest clustering
                | None -> self.CurrentClustering

            member self.DebugClusterSteps = log

            member self.Ranking : Ranking =
                let numfrm = cells.Length

                // keep a record of reported cells
                let rptd = new HashSet<AST.Address>()

                // for each step in the log,
                // add each source address and distance (score) to the ranking
                let rnk = 
                    log
                    |> List.rev
                    |> List.map (fun (step : ClusterStep) ->
                        if step.in_critical_region then
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

            member self.TotalEntropy : double =
                let addrs = self.CurrentClustering |> Seq.concat |> Seq.distinct |> Seq.toArray
                let rmap = BinaryMinEntropyTree.MakeCells addrs hb_inv

                self.CurrentClustering
                |> Seq.map (fun c -> BinaryMinEntropyTree.AddressSetEntropy (c |> Seq.toArray) rmap)
                |> Seq.sum

            member self.ScoreForCell(addr: AST.Address) : Countable =
                let (_,_,c) = hb_inv.[addr]
                c

            member self.RankingTimeMs = List.sum steps_ms
            member self.ScoreTimeMs = score_time
            member self.Scores = nlfrs
            member self.Cutoff = self.Ranking.Length - 1
            member self.LSHTree = hs.HashTree
            member self.AnalysisBase = cells
            member self.Analysis =
                Success(
                    Cluster {
                        scores = self.Scores;
                        ranking = self.Ranking;
                        score_time = self.ScoreTimeMs;
                        ranking_time = self.RankingTimeMs;
                        sig_threshold_idx = 0;
                        cutoff_idx = self.Cutoff;
                        weights = self.AnalysisBase |> Array.map (fun a -> a, 0.0) |> adict;
                        clustering = self.ClusteringAtKnee;
                    }
                )

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

                    // write debug logs when in debug mode
                    if input.config.DebugMode then
                        // create output dir
                        let path = IO.Path.Combine(
                                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                                    "ExceLintDebugOutput"
                                   )
                        IO.Directory.CreateDirectory(path) |> ignore
                        // write logs
                        m.WriteLog path
                        m.WritePerLogs path

                    m.Analysis
                else
                    CantRun "Cannot perform analysis. This worksheet contains no formulas."
            with
            | AnalysisCancelled -> Cancellation