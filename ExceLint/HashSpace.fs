namespace ExceLint
    open System.Collections.Generic
    open Utils

    module HashSetUtils =
        let difference<'a>(hs1: HashSet<'a>)(hs2: HashSet<'a>) : HashSet<'a> =
            let hs3 = new HashSet<'a>(hs1)
            hs3.ExceptWith(hs2)
            hs3

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
            val unm: bool
            val d: float

            new(cluster_from: HashSet<'p>, cluster_to: HashSet<'p>, common_prefix: UInt128, unmask_after_merge: bool, distance: float) =
                { c_from = cluster_from; c_to = cluster_to; com = common_prefix; unm = unmask_after_merge; d = distance }
            member self.FromCluster = self.c_from
            member self.ToCluster = self.c_to
            member self.CommonPrefix = self.com
            member self.UnmaskAfterMerge = self.unm
            member self.Distance = self.d
        end

    type HashSpace<'p when 'p : equality>(points: seq<'p>, keymaker: 'p -> UInt128, keyexists: 'p -> 'p -> 'p, unmasker: UInt128 -> UInt128, d: DistanceF<'p>) =
        // initialize tree
        let t = Seq.fold (fun (t': CRTNode<'p>)(point: 'p) ->
                    let key = keymaker point
                    t'.InsertOr key point keyexists
                ) (CRTRoot<'p>() :> CRTNode<'p>) points

        // initial mask
        let imsk = UInt128.Zero.Sub(UInt128.One)

        // dict of clusters
        let pt2Cluster =
            points
            |> Seq.map (fun p ->
                 p,new HashSet<'p>([p])
               )
            |> adict

        // initialize NN table
        let nn =
            points
            |> Seq.map (fun (p: 'p) ->
                 // get key
                 let key = keymaker p

                 // get the initial 'cluster'
                 let c1 = new HashSet<'p>([p])

                 // get nearest neighbors
                 let (ns,mask) = HashSpace.NearestNeighbors c1 t key imsk unmasker

                 // convert each neighboring point into a cluster (i.e., a HashSet)
                 let nhs = Seq.map (fun n -> pt2Cluster.[n]) ns

                 // choose the closest neighbor
                 let c2 = Utils.argmin (fun c -> d c1 c) nhs

                 // the new cluster
                 let c_after = HashSetUtils.union c1 c2

                 // adjust mask after merge?
                 // yes iff c2 = c1 union ns
                 // in other words, the entire subtree is cluster c2
                 let unm = HashSetUtils.equals c_after (HashSetUtils.union c1 (new HashSet<'p>(ns)))

                 // compute distance
                 let dst = d c1 c2

                 NN(c1, c2, mask, unm, dst)
            )
            |> Seq.sortBy (fun nn -> nn.Distance)
            |> Seq.toArray

        member self.Key(point: 'p) : UInt128 = keymaker point
        member self.NearestNeighborTable : NN<'p>[] = nn
        member self.HashTree: CRTNode<'p> = t
        member self.Merge(source: HashSet<'p>)(target: HashSet<'p>) : unit =
            // add all points in source cluster to the target cluster
            HashSetUtils.inPlaceUnion source target

            // update all points that map to the source so that
            // they now map to the target
            source
            |> Seq.iter (fun p -> pt2Cluster.[p] <- target)
        member self.NextNearestNeighbor : NN<'p> =
            // TODO HERE
            let (ns,mask) = HashSpace.NearestNeighbors c1 t key imsk unmasker
            failwith "hey"
        member self.Clusters : HashSet<HashSet<'p>> =
            let hss =  pt2Cluster.Values
            let h = new HashSet<HashSet<'p>>()
            hss |> Seq.iter (fun hs -> h.Add hs |> ignore)
            h


        /// <summary>
        /// Finds the set of closest points to a given cluster key,
        /// excluding those points already in the cluster.
        /// </summary>
        /// <param name="cluster">the cluster in question</param>
        /// <param name="root">the root of the prefix tree</param>
        /// <param name="key">the key representing the cluster</param>
        /// <param name="initial_mask">the last-searched mask</param>
        /// <param name="unmasker">a function that gives the next mask given the current mask</param>
        static member private NearestNeighbors(cluster: HashSet<'p>)(root: CRTNode<'p>)(key: UInt128)(initial_mask: UInt128)(unmasker: UInt128 -> UInt128) : seq<'p>*UInt128 =
            let mutable neighbors = Seq.empty<'p>
            let mutable mask = initial_mask 
            while (Seq.isEmpty neighbors) && mask <> UInt128.Zero do
                let st_opt = root.LookupSubtree key mask
                match st_opt with
                // don't include neighbors in the traversal that are
                // already in the cluster
                | Some(st) -> neighbors <- (st.LRTraversal |> Seq.filter (fun (p: 'p) -> not (cluster.Contains p)))
                | None -> ()
                // only adjust mask if neighbors is empty
                if Seq.isEmpty neighbors then
                    mask <- unmasker mask

            assert not (Seq.isEmpty neighbors)
            neighbors, mask