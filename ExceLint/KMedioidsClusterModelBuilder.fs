namespace ExceLint
    open System
    open System.Collections.Generic
    open Utils
    open CommonTypes
    open CommonFunctions
    open Distances

    type MedioidClustering = Dict<AST.Address,HashSet<AST.Address>>

    module KMedioidsClusterModelBuilder =
        let NRUNS = 100

        let private dist(d: DistanceF)(addr1: AST.Address)(addr2: AST.Address) : double =
            let addr1hs = new HashSet<AST.Address>([addr1])
            let addr2hs = new HashSet<AST.Address>([addr2])
            d addr1hs addr2hs

        let private findClosestMedioids(medioids: AST.Address[])(points: AST.Address[])(d: DistanceF) : MedioidClustering =
            let clustering =
                points |>
                Array.fold (fun (acc: MedioidClustering)(point: AST.Address) ->
                    let distances = Array.map (fun m -> dist d point m) medioids
                    let smallest_idx = [| 0 .. medioids.Length - 1|] |> argmin (fun i -> distances.[i])
                    let closest = medioids.[smallest_idx]
                    if acc.ContainsKey closest then
                        acc.[closest].Add point |> ignore
                    else
                        let hs = new HashSet<AST.Address>()
                        hs.Add(point) |> ignore
                        acc.Add(closest, hs) |> ignore
                    acc
                ) (new Dict<AST.Address,HashSet<AST.Address>>())
            clustering

        let replace(arr: 'a[])(e: 'a)(idx: int) : 'a[] =
            let arr' = Array.copy arr
            arr'.[idx] <- e
            arr'

        let cost(c: MedioidClustering)(d: DistanceF) : double =
            Seq.fold (fun (bigsum: double)(kvp: KeyValuePair<AST.Address,HashSet<AST.Address>>) ->
                let medioid = kvp.Key
                let points = kvp.Value
                bigsum +
                Seq.fold (fun (littlesum: double)(point: AST.Address) ->
                    dist d point medioid
                ) 0.0 points
            ) 0.0 c

        let medioidClustering2Clustering(c: MedioidClustering) : Clustering =
            new HashSet<HashSet<AST.Address>>(c.Values)

        let private debugClusterings(clusterings: MedioidClustering[]) : unit =
            // DEBUG: export all of the clusterings
            // convert to correct format
            let cls = Array.map (fun cluster -> medioidClustering2Clustering cluster) clusterings
            // init idmap array
            let maps = Array.init clusterings.Length (fun i -> new Dict<HashSet<AST.Address>, int>())
            Array.iteri (fun i clustering ->
                if i = 0 then
                    // first map
                    maps.[0] <- numberClusters clustering
                else
                    // use the first clustering's map to find the cluster numbers for the second clustering
                    let correspondence = JaccardCorrespondence clustering cls.[0]
                    let m = clustering |>
                        Seq.map (fun cluster ->
                            let cl_orig = correspondence.[cluster]
                            let num = maps.[0].[cl_orig]
                            cluster, num
                        ) |> adict
                    maps.[i] <- m
            ) cls

            for i in 0 .. cls.Length - 1 do
                let cl = cls.[i]
                ExceLintFileFormats.Clustering.writeClustering(cl, maps.[i], sprintf "C:\Users\dbarowy\Desktop\debug\%A.csv" i)

        /// <summary>
        /// This is the partitioning around medioids (PAM) algorithm.
        /// </summary>
        /// <param name="k"></param>
        /// <param name="cells"></param>
        /// <param name="s"></param>
        /// <param name="hb_inv"></param>
        /// <param name="d"></param>
        /// <param name="r"></param>
        let private kmedioids(k: int)(cells: AST.Address[])(s: ScoreTable)(hb_inv: InvertedHistogram)(d: DistanceF)(r: Random) : MedioidClustering =
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
                for i in [| 0 .. medioids.Length - 1 |] do

                    // swap medioid with every other point
                    for c in cells do
                        // new set of medioids
                        let medioids' = replace medioids c i
                        // find closest medioids
                        let clustering' = findClosestMedioids medioids' cells d
                        // if the cost decreases, keep the new clustering
                        let oldcost = cost clustering d
                        let newcost = cost clustering' d
                        if newcost < oldcost then
                            clustering <- clustering'
                            cost_decreased <- true

            clustering

        let getClustering(input: Input)(k: int) : Clustering =
            assert ((analysisBase input.config input.dag).Length <> 0)

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

            // run k-medioids
            let clusterings = Array.Parallel.map (fun _ -> kmedioids k cells nlfrs hb_inv DISTANCE r) [| 0 .. NRUNS - 1 |]
            
            // costs
            let costs = Array.map (fun clustering -> cost clustering DISTANCE) clusterings

            // debug
            debugClusterings clusterings

            // return the lowest sum-of-distances cost
            let lowest_cost_clustering = argmin (fun i -> costs.[i]) [| 0 .. NRUNS - 1 |]

            // convert into standard format
            medioidClustering2Clustering clusterings.[lowest_cost_clustering]