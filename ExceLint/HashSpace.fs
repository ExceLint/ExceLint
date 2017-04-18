namespace ExceLint
    type HashSpace(points: seq<'p>, keymaker: 'p -> UInt128, keyexists: 'p -> 'p -> 'p, unmasker: UInt128 -> UInt128) =
        // initialize tree
        let t = Seq.fold (fun (t': CRTNode<'p>)(point: 'p) ->
                    let key = keymaker point
                    t'.InsertOr key point keyexists
                ) (CRTRoot<'p>() :> CRTNode<'p>) points

        member self.NearestNeighbors(point: 'p) : seq<'p> =
            let key = keymaker point
            let mutable points = Seq.empty<'p>
            let mutable mask = unmasker (UInt128.Zero.Sub(UInt128.One))
            while (Seq.isEmpty points) && mask <> UInt128.Zero do
                let st_opt = t.LookupSubtree key mask
                match st_opt with
                | Some(st) -> points <- st.LRTraversal
                | None -> ()
                mask <- unmasker mask
            points