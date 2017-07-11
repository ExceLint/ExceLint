namespace ExceLint
    open System.Collections.Generic
    open CommonTypes
    open Utils

    module HashSetUtils =
        let difference<'a>(hs1: HashSet<'a>)(hs2: HashSet<'a>) : HashSet<'a> =
            let hs3 = new HashSet<'a>(hs1)
            hs3.ExceptWith(hs2)
            hs3

        let differenceElem<'a>(hs: HashSet<'a>)(elem: 'a) : HashSet<'a> =
            let hs2 = new HashSet<'a>(hs)
            hs2.Remove(elem) |> ignore
            hs2

        let intersection<'a>(hs1: HashSet<'a>)(hs2: HashSet<'a>) : HashSet<'a> =
            let hs3 = new HashSet<'a>(hs1)
            hs3.IntersectWith(hs2)
            hs3

        let union<'a>(hs1: HashSet<'a>)(hs2: HashSet<'a>) : HashSet<'a> =
            let hs3 = new HashSet<'a>(hs1)
            hs3.UnionWith(hs2)
            hs3

        let unionElem<'a>(hs: HashSet<'a>)(elem: 'a) : HashSet<'a> =
            let hs2 = new HashSet<'a>(hs)
            hs2.Add elem |> ignore
            hs2

        let inPlaceUnion<'a>(source: HashSet<'a>)(target: HashSet<'a>) : unit =
            target.UnionWith(source)

        let equals<'a>(hs1: HashSet<'a>)(hs2: HashSet<'a>) : bool =
            let hsu = union hs1 hs2
            hs1.Count = hsu.Count

    type DistanceF<'p> = HashSet<'p> -> HashSet<'p> -> double

    type NN<'p> =
        struct
            val c_from: HashSet<'p>
            val c_to: HashSet<'p>
            val com: UInt128
            val mask: UInt128
            val d: float

            new(cluster_from: HashSet<'p>, cluster_to: HashSet<'p>, common_prefix: UInt128, common_mask: UInt128, distance: float) =
                { c_from = cluster_from; c_to = cluster_to; com = common_prefix; mask = common_mask; d = distance }
            member self.FromCluster = self.c_from
            member self.ToCluster = self.c_to
            member self.CommonPrefix = self.com
            member self.CommonMask = self.mask
            member self.Distance = self.d
        end

    type HashSpace<'p when 'p : equality>(clustering: GenericClustering<'p>, keymaker: HashSet<'p> -> UInt128, keyexists: 'p -> 'p -> 'p, unmasker: UInt128 -> UInt128, d: DistanceF<'p>) =
        // extract points
        let points = clustering |> Seq.concat |> Seq.distinct |> Seq.toArray

        // initialize tree
        // use degenerate cluster (one point per cluster)
        // to satisfy keymaker
        let t = Seq.fold (fun (t': CRTNode<'p>)(point: 'p) ->
                    let cluster = new HashSet<'p>([| point |])
                    let key = keymaker cluster
                    t'.InsertOr key point keyexists
                ) (CRTRoot<'p>() :> CRTNode<'p>) points

        let tviz = t.ToGraphViz

        // initial mask
        let imsk = UInt128.Zero.Sub(UInt128.One)

        // dict of clusters
//        let pt2Cluster =
//            points
//            |> Seq.map (fun p ->
//                 p,new HashSet<'p>([p])
//               )
//            |> adict

        let pt2Cluster =
            clustering
            |> Seq.map (fun c ->
                c
                |> Seq.map (fun p ->
                    p, c
                )
            )
            |> Seq.concat
            |> adict

        // create cluster ID map
        let (_,ids: Dict<HashSet<'p>,int>) =
            Seq.fold (fun (idx,m) cl ->
                m.Add(cl, idx)
                (idx + 1, m)
            ) (0,new Dict<HashSet<'p>,int>()) (pt2Cluster.Values)

        // initialize NN table
        let nn =
            clustering
            |> Seq.map (fun (cluster: HashSet<'p>) ->
                 // get key
                 let key = keymaker cluster

                 // get common mask
//                 failwith "huh"

                 // index NN entry by cluster
                 cluster, HashSpace.NearestCluster t cluster key imsk unmasker pt2Cluster d
            )
            |> adict

        member self.NearestNeighborTable : seq<NN<'p>> = nn.Values |> Seq.sortBy (fun nn -> nn.Distance)
        member self.HashTree: CRTNode<'p> = t
        member self.ClusterID(c: HashSet<'p>) = ids.[c]
        member self.Merge(source: HashSet<'p>)(target: HashSet<'p>) : unit =
            // add all points in source cluster to the target cluster
            HashSetUtils.inPlaceUnion source target

            // update all points that map to the source so that
            // they now map to the target
            source
            |> Seq.iter (fun p -> pt2Cluster.[p] <- target)

            // remove the source cluster from the NN table
            nn.Remove source |> ignore

            // stop if we're down to the last entry
            if nn.Count > 1 then
                // update all entries whose nearest neighbor was
                // either the source or the target
                let nn_old = new Dict<HashSet<'p>,NN<'p>>(nn)
                nn_old |> Seq.iter (fun kvp ->
                              let nent = kvp.Value
                              let cl_source = kvp.Key
                              let cl_target = nent.ToCluster

                              if cl_target = source || cl_target = target then
                                  // find set of new nearest neighbors & update entry
                                  let key = nent.CommonPrefix
                                  let initial_mask = nent.CommonMask
                                  nn.[cl_source] <- HashSpace.NearestCluster t cl_source key initial_mask unmasker pt2Cluster d
                          )

        member self.NextNearestNeighbor : NN<'p> = self.NearestNeighborTable |> Seq.head
        member self.Clusters : HashSet<HashSet<'p>> =
            let hss =  pt2Cluster.Values
            let h = new HashSet<HashSet<'p>>()
            hss |> Seq.iter (fun hs -> h.Add hs |> ignore)
            h

        /// <summary>
        /// Finds the set of closest points to a given cluster key,
        /// excluding those points already in the cluster, using
        /// the LSH prefix tree.
        /// </summary>
        /// <param name="cluster">the cluster in question</param>
        /// <param name="root">the root of the prefix tree</param>
        /// <param name="key">the key representing the cluster</param>
        /// <param name="initial_mask">the last-searched mask</param>
        /// <param name="unmasker">a function that gives the next mask given the current mask</param>
        static member private NearestPoints(cluster: HashSet<'p>)(root: CRTNode<'p>)(key: UInt128)(initial_mask: UInt128)(unmasker: UInt128 -> UInt128) : seq<'p>*UInt128 =
            let mutable neighbors = Seq.empty<'p>
            let mutable mask = initial_mask
            let mutable no_more_neighbors = false
            while (Seq.isEmpty neighbors) && not no_more_neighbors do
                let st_opt = root.LookupSubtree key mask
                match st_opt with
                // don't include neighbors in the traversal that are
                // already in the cluster
                | Some(st) -> neighbors <- (st.LRTraversal |> Seq.filter (fun (p: 'p) -> not (cluster.Contains p)))
                | None -> ()
                if mask = UInt128.Zero then
                    no_more_neighbors <- true
                // only adjust mask if neighbors is empty
                else if Seq.isEmpty neighbors then
                    mask <- unmasker mask

            assert not (Seq.isEmpty neighbors)
            neighbors, mask

        /// <summary>
        /// Finds the closest cluster to a given cluster identified
        /// by an LSH key and mask.
        /// </summary>
        /// <param name="root">the root of the prefix tree</param>
        /// <param name="source">the cluster in question</param>
        /// <param name="key">a key belonging to any one of the points in the cluster</param>
        /// <param name="mask">the cluster's common bitmask</param>
        /// <param name="unmasker">a function that gives the next mask given the current mask</param>
        /// <param name="pt2Cluster">a mapping from points to clusters</param>
        /// <param name="d">a cluster-to-cluster distance function</param>
        static member private NearestCluster(root: CRTNode<'p>)(source: HashSet<'p>)(key: UInt128)(mask: UInt128)(unmasker: UInt128 -> UInt128)(pt2Cluster: Dict<'p,HashSet<'p>>)(d: DistanceF<'p>) =
            // get nearest neighbors
            let (ns,new_mask) = HashSpace.NearestPoints source root key mask unmasker

            // find the cluster corresponding to each point nearest to cl_source;
            // exclude the cl_source cluster itself
            let nhs = Seq.map (fun p -> pt2Cluster.[p]) ns |> Seq.distinct |> Seq.filter (fun cl -> cl <> source)

            // find the closest cluster
            let closest = Utils.argmin (fun c -> d source c) nhs

            // compute distance
            let dst = d source closest

            // return updated entry
            NN(source, closest, key, new_mask, dst)

         static member DegenerateClustering<'p>(cells: 'p[]) : GenericClustering<'p> =
             cells
             |> Array.map (fun c -> new HashSet<'p>([|c|]))
             |> (fun arr -> new HashSet<HashSet<'p>>(arr))
