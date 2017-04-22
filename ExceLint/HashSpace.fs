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
        end

    type HashSpace(points: seq<'p>, keymaker: 'p -> UInt128, keyexists: 'p -> 'p -> 'p, unmasker: UInt128 -> UInt128, d: DistanceF<'p>) =
        // initialize tree
        let t = Seq.fold (fun (t': CRTNode<'p>)(point: 'p) ->
                    let key = keymaker point
                    t'.InsertOr key point keyexists
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

                 // choose a neighbor arbitrarily
                 let n = Seq.head ns

                 // get ns as a set
                 let ns' = new HashSet<'p>(ns)

                 // the new cluster
                 let c2 = HashSetUtils.unionElem c1 n

                 // adjust mask after merge?
                 // yes iff c2 = c1 union ns'
                 // in other words, the entire subtree is cluster c2
                 let unm = c2 = HashSetUtils.union c1 ns'

                 // compute distance
                 let dst = d c1 ns'

                 NN(c1, c2, mask, unm, dst)
            )
            |> Seq.toArray

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