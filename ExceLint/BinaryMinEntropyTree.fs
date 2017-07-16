namespace ExceLint

    open System.Collections.Generic
    open System.Collections.Immutable
    open CommonTypes
    open CommonFunctions
    open Utils

    type Cells = Dict<AST.Address,Countable>
    type Coord = AST.Address * AST.Address
    type LeftTop = Coord
    type RightBottom = Coord
    type Region = LeftTop * RightBottom

    [<AbstractClass>]
    type BinaryMinEntropyTree(lefttop: AST.Address, rightbottom: AST.Address, subtree_kind: SubtreeKind) =
        abstract member Region : string
        default self.Region : string = lefttop.A1Local() + ":" + rightbottom.A1Local()
        abstract member ToGraphViz : int -> int*string
        abstract member Subtree : SubtreeKind
        default self.Subtree = subtree_kind

        static member AddressSetEntropy(addrs: AST.Address[])(rmap: Dict<AST.Address,Countable>) : double =
            if addrs.Length = 0 then
                // the decomposition where one side is the empty set
                // is the worst decomposition, unless it is the only
                // decomposition
                System.Double.PositiveInfinity
            else
                // get values
                let vs = addrs |> Array.map (fun a -> rmap.[a].ToCVectorResultant)

                // count
                let cs = BasicStats.counts vs

                // compute probability vector
                let ps = BasicStats.empiricalProbabilities cs

                // compute entropy
                let entropy = BasicStats.entropy ps

                entropy

        static member GraphViz(t: BinaryMinEntropyTree) : string =
            let (_,graph) = t.ToGraphViz 0
            "graph {" + graph + "}"

        /// <summary>
        /// Measure the entropy of a clustering, where the number of cells
        /// inside clusters is used to determine frequency.
        /// </summary>
        /// <param name="t">A BinaryMinEntropyTree  </param>
        static member TreeEntropy(t: BinaryMinEntropyTree) : double =
            // get regions from tree
            let tRegions = BinaryMinEntropyTree.Regions t |> Array.map (fun leaf -> leaf.Cells)

            // count
            let cs = tRegions |> Array.map (fun reg -> reg.Count)

            // debug
            let debug_cs = cs |> Array.sort |> (fun arr -> "c(" + System.String.Join(", ", arr) + ")")

            // compute probability vector
            let ps = BasicStats.empiricalProbabilities cs

            // compute entropy
            let entropy = BasicStats.entropy ps

            entropy

        /// <summary>
        /// The difference in tree entropy between t2 and t1. A negative number
        /// denotes a decrease in entropy from t1 to t2 whereas a positive number
        /// denotes an increase in entropy from t1 to t2.
        /// </summary>
        /// <param name="t1"></param>
        /// <param name="t2"></param>
        static member TreeEntropyDiff(t1: BinaryMinEntropyTree)(t2: BinaryMinEntropyTree) : double =
            let t1e = BinaryMinEntropyTree.TreeEntropy t1
            let t2e = BinaryMinEntropyTree.TreeEntropy t2
            t2e - t1e

        static member private MinEntropyPartition(rmap: Cells)(indivisibles: HashSet<HashSet<AST.Address>>)(vert: bool) : AST.Address[]*AST.Address[] =
            // which axis we use depends on whether we are
            // decomposing verticalls or horizontally
            let indexer = (fun (a: AST.Address) -> if vert then a.X else a.Y)

            // extract addresses
            let addrs = Array.ofSeq rmap.Keys

            // find bounding box
            let (lt,rb) = Utils.BoundingRegion (rmap.Keys) 0

            let parts = Array.map (fun i ->
                               // partition addresses by "less than index of a",
                               // e.g., if vert then "less than x"
                               let (left,right) = addrs |> Array.partition (fun a -> indexer a < i)

                               // if left and right divide any indivisible clusters,
                               // ignore this partitioning
                               let left' = new HashSet<AST.Address>(left)
                               let right' = new HashSet<AST.Address>(right)
                               let ok = indivisibles |>
                                        Seq.fold (fun acc indiv ->
                                            // either the left side contains the entire indivisible set
                                            // or the right side contains the entire indivisible set
                                            // or neither side contains any of the indivisible set
                                            acc && (
                                                let l_int = HashSetUtils.intersection left' indiv
                                                let r_int = HashSetUtils.intersection right' indiv
                                                (l_int.Count = indiv.Count && r_int.Count = 0) ||
                                                (l_int.Count = 0 && r_int.Count = indiv.Count) ||
                                                (l_int.Count = 0 && r_int.Count = 0)
                                            )
                                        ) true
                               if ok then
                                   Some(left, right)
                               else
                                   let i' = indivisibles
                                   let l' = left
                                   let r' = right
                                   None
                           ) [| indexer lt .. indexer rb + 1 |]
                        |> Array.choose id

            assert (parts.Length <> 0)

            let debug_es =
                parts
                |> Array.map (fun (l,r) ->
                    // compute entropy
                    let entropy_left = BinaryMinEntropyTree.AddressSetEntropy l rmap
                    let entropy_right = BinaryMinEntropyTree.AddressSetEntropy r rmap

                    // total for left and right
                    (entropy_left + entropy_right),entropy_left,entropy_right
                )

            parts
            |> Utils.argmin (fun (l,r) ->
                // compute entropy
                let entropy_left = BinaryMinEntropyTree.AddressSetEntropy l rmap
                let entropy_right = BinaryMinEntropyTree.AddressSetEntropy r rmap

                // total for left and right
                entropy_left + entropy_right
            )

        static member MakeCells(hb_inv: InvertedHistogram) : Cells =
            let addrs = hb_inv.Keys
            let d = new Dict<AST.Address, Countable>()
            for addr in addrs do
                let (_,_,c) = hb_inv.[addr]
                d.Add(addr, c)
            d

        static member Infer(hb_inv: InvertedHistogram)(indivisibles: HashSet<HashSet<AST.Address>>) : BinaryMinEntropyTree =
            let rmap = BinaryMinEntropyTree.MakeCells hb_inv
            BinaryMinEntropyTree.Decompose rmap indivisibles

//        static member private Decompose(rmap: Cells)(indivisibles: HashSet<HashSet<AST.Address>>)(parent_opt: Inner option) : BinaryMinEntropyTree =
//            // get bounding region
//            let (lefttop,rightbottom) = Utils.BoundingRegion rmap.Keys 0
//
//            // base case 1: there's only 1 cell
//            if lefttop = rightbottom then
//                Leaf(lefttop, rightbottom, parent_opt, rmap) :> BinaryMinEntropyTree
//            else
//                // find the minimum entropy decompositions
//                let (left,right) = BinaryMinEntropyTree.MinEntropyPartition rmap indivisibles true
//                let (top,bottom) = BinaryMinEntropyTree.MinEntropyPartition rmap indivisibles false
//
//                // compute entropies again
//                let e_vert = BinaryMinEntropyTree.AddressSetEntropy left rmap +
//                                BinaryMinEntropyTree.AddressSetEntropy right rmap
//                let e_horz = BinaryMinEntropyTree.AddressSetEntropy top rmap +
//                                BinaryMinEntropyTree.AddressSetEntropy bottom rmap
//
//                // split vertically or horizontally (favor vert for ties)
//                let (entropy,p1,p2) =
//                    if e_vert <= e_horz then
//                        e_vert, left, right
//                    else
//                        e_horz, top, bottom
//
//                // base case 2: (perfect decomposition & right values same as left values)
//                if entropy = 0.0 && rmap.[p1.[0]].ToCVectorResultant = rmap.[p2.[0]].ToCVectorResultant then
//                    Leaf(lefttop, rightbottom, parent_opt, rmap) :> BinaryMinEntropyTree
//                // recursive case
//                else
//                    let p1_rmap = p1 |> Array.map (fun a -> a,rmap.[a]) |> adict
//                    let p2_rmap = p2 |> Array.map (fun a -> a,rmap.[a]) |> adict
//
//                    let node = Inner(lefttop, rightbottom)
//                    let p1node = BinaryMinEntropyTree.Decompose p1_rmap indivisibles (Some node)
//                    let p2node = BinaryMinEntropyTree.Decompose p2_rmap indivisibles (Some node)
//                    node.AddLeft p1node
//                    node.AddRight p2node
//                    node :> BinaryMinEntropyTree

        static member private Decompose (initial_rmap: Cells)(indivisibles: HashSet<HashSet<AST.Address>>) : BinaryMinEntropyTree =
            let mutable todos = [ (Root, initial_rmap) ]
            let mutable linkUp = []
            let mutable root_opt = None

            // process work list
            while not todos.IsEmpty do
                // grab next item
                let (subtree_kind, rmap) = todos.Head
                todos <- todos.Tail

                // get bounding region
                let (lefttop,rightbottom) = Utils.BoundingRegion rmap.Keys 0

                // base case 1: there's only 1 cell
                if lefttop = rightbottom then
                    let leaf = Leaf(lefttop, rightbottom, subtree_kind, rmap) :> BinaryMinEntropyTree
                    // add leaf to to link-up list
                    linkUp <- leaf :: linkUp

                    // is this leaf the root?
                    match subtree_kind with
                    | Root -> root_opt <- Some leaf
                    | _ -> ()
                else
                    // find the minimum entropy decompositions
                    let (left,right) = BinaryMinEntropyTree.MinEntropyPartition rmap indivisibles true
                    let (top,bottom) = BinaryMinEntropyTree.MinEntropyPartition rmap indivisibles false

                    // compute entropies again
                    let e_vert = BinaryMinEntropyTree.AddressSetEntropy left rmap +
                                    BinaryMinEntropyTree.AddressSetEntropy right rmap
                    let e_horz = BinaryMinEntropyTree.AddressSetEntropy top rmap +
                                    BinaryMinEntropyTree.AddressSetEntropy bottom rmap

                    // split vertically or horizontally (favor vert for ties)
                    let (entropy,p1,p2) =
                        if e_vert <= e_horz then
                            e_vert, left, right
                        else
                            e_horz, top, bottom

                    // base case 2: perfect decomposition & right values same as left values
                    if entropy = 0.0 && rmap.[p1.[0]].ToCVectorResultant = rmap.[p2.[0]].ToCVectorResultant ||
                    // base case 3: cluster is indivisible, so min entropy is "infinite"
                       System.Double.IsPositiveInfinity entropy
                    then
                        let leaf = Leaf(lefttop, rightbottom, subtree_kind, rmap) :> BinaryMinEntropyTree

                        // is this leaf the root?
                        match subtree_kind with
                        | Root -> root_opt <- Some leaf
                        | _ -> ()

                        // add leaf to to link-up list
                        linkUp <- leaf :: linkUp
                    else
                        let p1_rmap = p1 |> Array.map (fun a -> a,rmap.[a]) |> adict
                        let p2_rmap = p2 |> Array.map (fun a -> a,rmap.[a]) |> adict

                        let node = Inner(lefttop, rightbottom, subtree_kind)

                        // is this node the root?
                        match subtree_kind with
                        | Root -> root_opt <- Some (node :> BinaryMinEntropyTree)
                        | _ -> ()

                        // add next nodes to work list
                        todos <- (LeftOf node, p1_rmap) :: (RightOf node, p2_rmap) :: todos
                        
            // process "link-up" list
            while not linkUp.IsEmpty do
                // grab next item
                let node = linkUp.Head
                linkUp <- linkUp.Tail

                match node.Subtree with
                | LeftOf parent ->
                    // add parent to linkup list
                    linkUp <- (parent :> BinaryMinEntropyTree) :: linkUp

                    // make link
                    parent.AddLeft node
                | RightOf parent ->
                    // add parent to linkup list
                    linkUp <- (parent :> BinaryMinEntropyTree) :: linkUp

                    // make link
                    parent.AddRight node
                | Root -> ()    // do nothing

            match root_opt with
            | Some root -> root
            | None -> failwith "this should never happen"

        /// <summary>return the leaves of the tree, in order of smallest to largest region</summary>
        static member Regions(tree: BinaryMinEntropyTree) : Leaf[] =
            match tree with
            | :? Inner as i -> Array.append (BinaryMinEntropyTree.Regions (i.Left)) (BinaryMinEntropyTree.Regions (i.Right))
            | :? Leaf as l -> [| l |]
            | _ -> failwith "Unknown tree node type."

        static member Clustering(tree: BinaryMinEntropyTree) : ImmutableClustering =
            let regions = BinaryMinEntropyTree.Regions tree
            let cs = regions |> Array.map (fun leaf -> leaf.Cells)
            makeImmutableGenericClustering cs

        static member MutableClustering(tree: BinaryMinEntropyTree) : Clustering =
            let regions = BinaryMinEntropyTree.Regions tree
            let cs = regions |> Array.map (fun leaf -> new HashSet<AST.Address>(leaf.Cells))
            new Clustering(cs)

        static member ClusterIsRectangular(c: HashSet<AST.Address>) : bool =
            let boundingbox = Utils.BoundingBoxHS c 0
            let diff = HashSetUtils.difference boundingbox c
            let isRect = diff.Count = 0
            isRect

        static member ClusteringContainsOnlyRectangles(cs: Clustering) : bool =
            cs |> Seq.fold (fun a c -> a && BinaryMinEntropyTree.ClusterIsRectangular c) true

        static member MergeIsRectangular(source: HashSet<AST.Address>)(target: HashSet<AST.Address>) : bool =
            let merged = HashSetUtils.union source target
            BinaryMinEntropyTree.ClusterIsRectangular merged

        static member CoaleseAdjacentClusters(coal_vert: bool)(clusters: ImmutableClustering)(hb_inv: ROInvertedHistogram) : ImmutableClustering =
            // sort cells array depending on coalesce direction:
            // 1. coalesce vertically means sort horizontally (small to large x values)
            // 2. coalesce horizontally means sort vertically (small to large y values)
            let cells =
                clusters
                |> Seq.map (fun cluster -> cluster |> Seq.toArray)
                |> Array.concat 
                |> Array.sortBy (fun addr -> if coal_vert then (addr.X, addr.Y) else (addr.Y, addr.X))

            // algorithm mutates clusters'
            let clusters' = CopyImmutableToMutableClustering clusters

            let revLookup = ReverseClusterLookup clusters'

            let adjacent = if coal_vert then
                                (fun (c1: AST.Address)(c2: AST.Address) -> c1.X = c2.X && c1.Y < c2.Y)
                            else
                                (fun (c1: AST.Address)(c2: AST.Address) -> c1.Y = c2.Y && c1.X < c2.X)

            for i in 0 .. cells.Length - 2 do
                // get cell
                let cell = cells.[i]
                // get cluster of the cell
                let source = revLookup.[cell]
                // get possibly adjacent cell
                let maybeAdj = cells.[i+1]

                // if maybeAdj is already in the same cluster, move on
                if not (source.Contains maybeAdj) then
                    // cell's countable
                    let (_,_,co_cell) = hb_inv.[cell]
                    // maybeAdj's countable
                    let (_,_,co_maybe) = hb_inv.[maybeAdj]
                    // is maybeAdj adjacent to cell and has the same cvector?
                    let isAdj = adjacent cell maybeAdj
                    let sameCV = co_cell.ToCVectorResultant = co_maybe.ToCVectorResultant
                    if isAdj && sameCV then
                        // get cluster for maybeAdj
                        let target = revLookup.[maybeAdj]
                        // if I merge source and target, does the merge remain rectangular?
                        let isRect = BinaryMinEntropyTree.MergeIsRectangular source target
                        if isRect then
                            // add every cell from source to target hashset
                            HashSetUtils.inPlaceUnion source target
                            // update all reverse lookups for source cells
                            for c in source do
                                revLookup.[c] <- target
                            // remove source from clusters
                            clusters'.Remove source |> ignore

            ToImmutableClustering clusters'

        static member RectangularClustering(tree: BinaryMinEntropyTree)(hb_inv: ROInvertedHistogram) : ImmutableClustering =
            // coalesce all cells that have the same cvector,
            // ensuring that all merged clusters remain rectangular
            let regs = BinaryMinEntropyTree.Regions tree
            let clusters = regs |> Array.map (fun leaf -> leaf.Cells) |> makeImmutableGenericClustering

            // coalesce vertical ordering horizontally
            let clusters' = BinaryMinEntropyTree.CoaleseAdjacentClusters false clusters hb_inv

            // coalesce horizontal ordering vertically
            let clusters'' = BinaryMinEntropyTree.CoaleseAdjacentClusters true clusters' hb_inv

            // return clustering
            clusters''

    and Inner(lefttop: AST.Address, rightbottom: AST.Address, subtree: SubtreeKind) =
        inherit BinaryMinEntropyTree(lefttop, rightbottom, subtree)
        let mutable left = None
        let mutable right = None
        member self.AddLeft(n: BinaryMinEntropyTree) : unit =
            left <- Some n
        member self.AddRight(n: BinaryMinEntropyTree) : unit =
            right <- Some n
        member self.Left =
            match left with
            | Some l -> l
            | None -> failwith "Cannot traverse tree until it is constructed!"
        member self.Right =
            match right with
            | Some r -> r
            | None -> failwith "Cannot traverse tree until it is constructed!"
        override self.ToGraphViz(i: int) =
            let start = "\"" + i.ToString() + "\""
            let start_node = start + " [label=\"" + self.Region + "\"]\n"
            let (j,ledge) = match left with
                            | Some l ->
                                let (i',graph) = l.ToGraphViz (i + 1)
                                i', start + " -- " + "\"" + (i + 1).ToString() + "\"" + "\n" + graph
                            | None -> i,""
            let (k,redge) = match right with
                            | Some r ->
                                let (j',graph) = r.ToGraphViz (j + 1)
                                j', start + " -- " + "\"" + (j + 1).ToString() + "\"" + "\n" + graph
                            | None -> j,""
            k,start_node + ledge + redge

    and Leaf(lefttop: AST.Address, rightbottom: AST.Address, subtree: SubtreeKind, cells: Cells) =
        inherit BinaryMinEntropyTree(lefttop, rightbottom, subtree)
        member self.Cells : ImmutableHashSet<AST.Address> = (new HashSet<AST.Address>(cells.Keys)).ToImmutableHashSet()
        override self.ToGraphViz(i: int)=
            let start = "\"" + i.ToString() + "\""
            let node = start + " [label=\"" + self.Region + "\"]\n"
            i,node

    and SubtreeKind =
    | LeftOf of Inner
    | RightOf of Inner
    | Root