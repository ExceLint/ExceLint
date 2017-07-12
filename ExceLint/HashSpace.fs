namespace ExceLint
    open System.Collections.Generic
    open CommonTypes
    open Utils

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

    type HashSpace<'p when 'p : equality>(clustering: ImmutableGenericClustering<'p>, keymaker: 'p -> UInt128, keyexists: 'p -> 'p -> 'p, unmasker: UInt128 -> UInt128, d: DistanceF<'p>) =
        // make a mutable copy of immutable clustering
        // because the Merge procedure is side-effecting
        let clustering' = CommonFunctions.CopyImmutableToMutableClustering clustering
        
        // extract points
        let points = clustering' |> Seq.concat |> Seq.toArray

        // initialize tree
        let t = Seq.fold (fun (t': CRTNode<'p>)(point: 'p) ->
                    let key = keymaker point
                    t'.InsertOr key point keyexists
                ) (CRTRoot<'p>() :> CRTNode<'p>) points

        // lookup cluster by point
        let pt2Cluster =
            clustering'
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
            ) (0,new Dict<HashSet<'p>,int>()) clustering'

        // initialize NN table
        let nn =
            clustering'
            |> Seq.map (fun (cluster: HashSet<'p>) ->
                 let keys = cluster |> Seq.map (fun p -> keymaker p) |> Seq.toArray

                 // get key, any key
                 let key = keys.[0]

                 // find common mask
                 let cmsk = UInt128.calcMask 0 (UInt128.LongestCommonPrefix keys)

                 // index NN entry by cluster
                 cluster, HashSpace.NearestCluster t cluster key cmsk unmasker pt2Cluster d
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
