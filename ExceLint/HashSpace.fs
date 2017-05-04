namespace ExceLint
    open System.Collections.Generic

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

                    let tstr = t'.ToGraphViz

                    // debug
                    let keys = Set.ofSeq t'.LRKeyTraversal
                    if Set.contains key keys then
                        failwith "whoa nelly"

                    let t'' = t'.InsertOr key point keyexists
                    
                    t''
                ) (CRTRoot<'p>() :> CRTNode<'p>) points

        // initial mask
        let imsk = UInt128.Zero.Sub(UInt128.One)

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
                 let nhs = Seq.map (fun n -> new HashSet<'p>([n])) ns

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
            |> Seq.toArray

        member self.Key(point: 'p) : UInt128 = keymaker point
        member self.NearestNeighborTable : NN<'p>[] = nn
        member self.HashTree: CRTNode<'p> = t

        static member private NearestNeighbors(points: HashSet<'p>)(t: CRTNode<'p>)(key: UInt128)(initial_mask: UInt128)(unmasker: UInt128 -> UInt128) : seq<'p>*UInt128 =
            let mutable neighbors = Seq.empty<'p>
            let mutable mask = initial_mask 
            while (Seq.isEmpty neighbors) && mask <> UInt128.Zero do
                let st_opt = t.LookupSubtree key mask
                match st_opt with
                // don't include neighbors in the traversal that are
                // already in the cluster
                | Some(st) -> neighbors <- (st.LRTraversal |> Seq.filter (fun (p: 'p) -> not (points.Contains p)))
                | None -> ()
                // only adjust mask if neighbors is empty
                if Seq.isEmpty neighbors then
                    mask <- unmasker mask

            assert not (Seq.isEmpty neighbors)
            neighbors, mask