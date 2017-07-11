namespace ExceLint
    open System
    open System.Collections.Generic
    open Utils
    open CommonTypes
    open CommonFunctions
    open Distances

    module OldClusterModel =
        // true iff on two clusters are on the same sheet;
        // does not check entire cluster since, by induction,
        // clusters on other sheets will never be merged
        let private zfilter = (fun (C1: HashSet<AST.Address>)(C2: HashSet<AST.Address>) ->
                                let c1: AST.Address = Seq.head C1
                                let c2: AST.Address = Seq.head C2
                                c1.A1Worksheet() = c2.A1Worksheet()
                              )

        let private initialClustering(scoretable: ScoreTable)(dag: Depends.DAG)(config: FeatureConf) : Clustering =
            let t: Clustering = new Clustering()

            Array.iter (fun fname ->
                Array.iter (fun (addr: AST.Address, score: Countable) ->
                    t.Add(new HashSet<AST.Address>([addr])) |> ignore
                ) (scoretable.[fname])
            ) (config.EnabledFeatures)
            t

        let private induceCompleteGraphExcluding(xs: seq<'a>)(pred: 'a -> 'a -> bool) : seq<'a*'a> =
            cartesianProduct xs xs
            |> Seq.filter (fun (source,target) ->
                                source <> target &&     // no self edges
                                (pred source target)    // filter disallowed targets
                            )

        let private pairwiseClusterDistances(C: Clustering)(d: DistanceF): SortedSet<Edge> =
            // get all pairs of clusters and add to set
            let G: Edge[] = induceCompleteGraphExcluding (C |> Seq.toArray) zfilter |> Seq.map (fun (a,b) -> Edge(a,b)) |> Seq.toArray

            let edges = new SortedSet<Edge>(new MinDistComparer(d))
            G |> Array.iter (fun e -> edges.Add(e) |> ignore)

            edges

        // all updates here are as side-effects
        let updatePairwiseClusterDistancesAndCluster(C: Clustering)(ih: InvertedHistogram)(d: DistanceF)(edges: SortedSet<Edge>)(source: HashSet<AST.Address>)(target: HashSet<AST.Address>) : unit =
            let edgecount = edges.Count
                
            // find vertices neither source nor target
            // and make sure that they're on the same sheet
            let Ccount = C.Count
            let C' = C |> Seq.filter (fun v -> v <> source && v <> target && zfilter v source) |> Seq.toArray
            let Cpcount = C'.Length

            if Ccount <= Cpcount then
                failwith "not possible"

            // remove source from graph
            C.Remove(source) |> ignore
                
            // remove edges incident on source and target
            let removals = edges |> Seq.filter (fun edge ->
                                        let (edge_to,edge_from) = edge.tupled
                                        edge_to = source ||
                                        edge_from = source ||
                                        edge_to = target ||
                                        edge_from = target
                                    )
                                    |> Seq.toArray // laziness is bad
            if removals.Length = 0 then
                failwith "no!"
            removals |> Array.iter (fun edge -> edges.Remove(edge) |> ignore)

            // add source to target
            source |> Seq.iter (fun addr -> target.Add(addr) |> ignore)

            // add edges incident on merged source-target
            C'
            |> Seq.iter (fun v ->
                    edges.Add(Edge(v,target)) |> ignore
                    edges.Add(Edge(target,v)) |> ignore
                )

            let edgecount' = edges.Count
            if edgecount' > edgecount then
                failwith "marcia marcia marcia!!!"

        type OldClusterModel(input: Input) =
            // determine the set of cells to be analyzed
            let cells = analysisBase input.config input.dag

            // get all NLFRs for every formula cell
            let _runf = fun () -> runEnabledFeatures cells input.dag input.config input.progress
            let (ns: ScoreTable,score_time: int64) = PerfUtils.runMillis _runf ()

            // filter and scale
            let factor = diagonalScaleFactor ns
            let nlfrs: ScoreTable =
                ns
                |> Seq.map (fun kvp ->
                                kvp.Key,
                                kvp.Value
                                // remove all off-sheet refs
                                |> Array.map (fun (addr,c) ->
                                    if c.IsOffSheet then
                                        None
                                    else 
                                        Some (addr,c)
                                    )
                                |> Array.choose id
                                |> Array.map (fun (addr,c) ->
                                    addr,
                                    // only scale the resultant, not the location
                                    c.UpdateResultant (c.ToCVectorResultant.ScalarMultiply factor)
                                    )
                            )
                |> Seq.toArray
                |> toDict

            // make HistoBin lookup by address
            let hb_inv = invertedHistogram nlfrs input.dag input.config

            // initially assign every cell to its own cluster
            let clusters = initialClustering nlfrs input.dag input.config
            let mutable clusteringAtKnee = None

            // create cluster ID map
            let (_,ids: Dict<HashSet<AST.Address>,int>) =
                Seq.fold (fun (idx,m) cl ->
                    m.Add(cl, idx)
                    (idx + 1, m)
                ) (0,new Dict<HashSet<AST.Address>,int>()) clusters

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

            // get initial pairwise distances
            let edges = pairwiseClusterDistances clusters DISTANCE

            // compute initial NN table
            let keymaker = (fun (cluster: HashSet<AST.Address>) ->
                                // any of the addresses in the cluster
                                // are capable of producing a
                                // 'representative' hash because they
                                // all share a common prefix
                                let addr = cluster |> Seq.toArray |> (fun c -> c.[0])
                                let (_,_,co) = hb_inv.[addr]
                                LSHCalc.h7 co
                            )
            let keyexists = (fun addr1 addr2 ->
                                failwith "Duplicate keys should not happen."
                            )
            let initialClustering = HashSpace.DegenerateClustering cells
            let hs = HashSpace<AST.Address>(initialClustering, keymaker, keyexists, LSHCalc.h7unmasker, DISTANCE)

            let mutable probable_knee = false

            member self.CanStep : bool = clusters.Count > 1
            member private self.IsKnee(s: HashSet<AST.Address>)(t: HashSet<AST.Address>) : bool =
                // the first time we merge two clusters that have
                // different resultants, we've probably hit the knee
                not (refvect_same s t)

            // determine whether the next step will be the knee
            // without actually doing anohter agglomeration step
            member self.NextStepIsKnee : bool =
                let e = edges.Min
                let (source,target) = e.tupled
                self.IsKnee source target

            member self.Step() : bool =
                let sw = new System.Diagnostics.Stopwatch()
                sw.Start()

                // get the two clusters that minimize distance
                let e = edges.Min

                let (source,target) = e.tupled

                if self.IsKnee source target then
                    // only update ONCE
                    clusteringAtKnee <-
                        match clusteringAtKnee with
                        | Some cak -> Some cak
                        | None -> Some (CopyClustering hs.Clusters)

                    probable_knee <- true

                // record merge in log
                log <- {
                            // NOTE: Stats commented out because they are
                            //       expensive to compute and not obviously
                            //       useful.
                            in_critical_region = probable_knee;
                            source = Set.ofSeq source;
                            target = Set.ofSeq target;
                            distance = DISTANCE source target;
                            f = F clusters hb_inv;
//                            within_cluster_sum_squares = WCSS clusters hb_inv;
                            within_cluster_sum_squares = 0.0;
//                            between_cluster_sum_squares = BCSS clusters hb_inv;
                            between_cluster_sum_squares = 0.0;
//                            total_sum_squares = TSS clusters hb_inv;
                            total_sum_squares = 0.0;
                            num_clusters = clusters.Count;
                        } :: log

                // dump clusters to log
                let mutable clusterlog = []
                clusters
                |> Seq.iter (fun cl ->
                        cl
                        |> Seq.iter (fun addr ->
                            let v = ToCountable addr hb_inv
                            match v with
                            | FullCVectorResultant(x,y,z,dx,dy,dz,dc) ->
                                let row = new ExceLintFileFormats.VectorDumpRow()
                                row.clusterID <- ids.[cl]
                                row.x <- x
                                row.y <- y
                                row.z <- z
                                row.dx <- dx
                                row.dy <- dy
                                row.dz <- dz
                                row.dc <- dc
                                clusterlog <- row :: clusterlog
                            | _ -> ()
                        )
                    )
                per_log <- (List.rev clusterlog) :: per_log

                // merge them
                updatePairwiseClusterDistancesAndCluster clusters hb_inv DISTANCE edges source target

                sw.Stop()
                steps_ms <- sw.ElapsedMilliseconds :: steps_ms

                // tell the user whether more steps remain
                clusters.Count > 1 && edges.Count > 0

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
                        row.Show <- step.in_critical_region
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

            member self.Clustering = clusters

            member self.ClusteringAtKnee =
                match clusteringAtKnee with
                | Some c -> c
                | None -> failwith "No clustering available. Did you actually run the model?"

            member self.Ranking =
                let numfrm = input.dag.getAllFormulaAddrs().Length

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

            member self.RankingTimeMs = List.sum steps_ms
            member self.ScoreTimeMs = score_time
            member self.Scores = nlfrs
            member self.Cutoff = self.Ranking.Length - 1

        let runClusterModel(input: Input) : AnalysisOutcome =
            try
                if (analysisBase input.config input.dag).Length <> 0 then
                    let m = OldClusterModel input

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
                            clustering = m.ClusteringAtKnee;
                        }
                    )
                else
                    CantRun "Cannot perform analysis. This worksheet contains no formulas."
            with
            | AnalysisCancelled -> Cancellation