namespace ExceLint

    open System.Collections.Generic
    open CommonTypes
    open Utils

    type Cells = Dict<AST.Address,Countable>
    type Coord = AST.Address * AST.Address
    type LeftTop = Coord
    type RightBottom = Coord
    type Region = LeftTop * RightBottom

    [<AbstractClass>]
    type BinaryMinEntropyTree(lefttop: AST.Address, rightbottom: AST.Address) =
        abstract member Region: string
        default self.Region : string = lefttop.A1Local() + ":" + rightbottom.A1Local()

        static member private AddressEntropy(addrs: AST.Address[])(rmap: Dict<AST.Address,Countable>) : double =
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
                let ps = BasicStats.empiricalProbabilities cs vs.Length

                // debug string
                let debug = "x = c(" + System.String.Join(", ", cs) + ")"

                // compute entropy
                let entropy = BasicStats.entropy ps

                entropy

        static member private MinEntropyPartition(rmap: Cells)(vert: bool) : AST.Address[]*AST.Address[] =
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
                               left, right
                           ) [| indexer lt .. indexer rb + 1 |]

            let debug_es =
                parts
                |> Array.map (fun (l,r) ->
                    // compute entropy
                    let entropy_left = BinaryMinEntropyTree.AddressEntropy l rmap
                    let entropy_right = BinaryMinEntropyTree.AddressEntropy r rmap

                    // total for left and right
                    (entropy_left + entropy_right),entropy_left,entropy_right
                )

            parts
            |> Utils.argmin (fun (l,r) ->
                // compute entropy
                let entropy_left = BinaryMinEntropyTree.AddressEntropy l rmap
                let entropy_right = BinaryMinEntropyTree.AddressEntropy r rmap

                // total for left and right
                entropy_left + entropy_right
            )


        static member Infer(addrs: AST.Address[])(hb_inv: InvertedHistogram) : BinaryMinEntropyTree =
            let d = new Dict<AST.Address, Countable>()
            for addr in addrs do
                let (_,_,c) = hb_inv.[addr]
                d.Add(addr, c)
            BinaryMinEntropyTree.Decompose d None

        static member private Decompose(rmap: Cells)(parent_opt: Inner option) : BinaryMinEntropyTree =
            // get bounding region
            let (lefttop,rightbottom) = Utils.BoundingRegion rmap.Keys 0

            // base case 1: there's only 1 cell
            if lefttop = rightbottom then
                Leaf(lefttop, rightbottom, parent_opt, rmap) :> BinaryMinEntropyTree
            else
                // find the minimum entropy decompositions
                let (left,right) = BinaryMinEntropyTree.MinEntropyPartition rmap true
                let (top,bottom) = BinaryMinEntropyTree.MinEntropyPartition rmap false

                // compute entropies again
                let e_vert_l = BinaryMinEntropyTree.AddressEntropy left rmap
                let e_vert_r = BinaryMinEntropyTree.AddressEntropy right rmap
                let e_horz_t = BinaryMinEntropyTree.AddressEntropy top rmap
                let e_horz_b = BinaryMinEntropyTree.AddressEntropy bottom rmap

                let e_vert = e_vert_l + e_vert_r
                let e_horz = e_horz_t + e_horz_b

                // split vertically or horizontally (favor vert for ties)
                if e_vert <= e_horz then
                    // base case 2: (perfect decomposition & right values same as left values)
                    if e_vert = 0.0 && rmap.[left.[0]].ToCVectorResultant = rmap.[right.[0]].ToCVectorResultant then
                        Leaf(lefttop, rightbottom, parent_opt, rmap) :> BinaryMinEntropyTree
                    else
                        let l_rmap = left   |> Array.map (fun a -> a,rmap.[a]) |> adict
                        let r_rmap = right  |> Array.map (fun a -> a,rmap.[a]) |> adict

                        let node = Inner(lefttop, rightbottom)
                        let lnode = BinaryMinEntropyTree.Decompose l_rmap (Some node)
                        let rnode = BinaryMinEntropyTree.Decompose r_rmap (Some node)
                        node.AddLeft lnode
                        node.AddRight rnode
                        node :> BinaryMinEntropyTree
                else
                    // base case 2: (perfect decomposition & top values same as bottom values)
                    if e_horz = 0.0 && rmap.[top.[0]].ToCVectorResultant = rmap.[bottom.[0]].ToCVectorResultant then
                        Leaf(lefttop, rightbottom, parent_opt, rmap) :> BinaryMinEntropyTree
                    else
                        let t_rmap = top    |> Array.map (fun a -> a,rmap.[a]) |> adict
                        let b_rmap = bottom |> Array.map (fun a -> a,rmap.[a]) |> adict

                        let node = Inner(lefttop, rightbottom)
                        let tnode = BinaryMinEntropyTree.Decompose t_rmap (Some node)
                        let bnode = BinaryMinEntropyTree.Decompose b_rmap (Some node)
                        node.AddLeft tnode
                        node.AddRight bnode
                        node :> BinaryMinEntropyTree

        /// <summary>return the leaves of the tree, in order of smallest to largest region</summary>
        static member Regions(tree: BinaryMinEntropyTree) : Leaf[] =
            match tree with
            | :? Inner as i -> Array.append (BinaryMinEntropyTree.Regions (i.Left)) (BinaryMinEntropyTree.Regions (i.Right))
            | :? Leaf as l -> [| l |]
            | _ -> failwith "Unknown tree node type."

        static member private MergeIsRectangular(source: HashSet<AST.Address>)(target: HashSet<AST.Address>) : bool =
            let merged = HashSetUtils.union source target
            let bb_merged = Utils.BoundingBoxHS merged 0
            (HashSetUtils.difference bb_merged merged).Count = 0

        static member RectangularClustering(tree: BinaryMinEntropyTree)(hb_inv: InvertedHistogram) : Clustering =
            // coalesce all cells that have the same cvector,
            // ensuring that all merged clusters remain rectangular
            let regs = BinaryMinEntropyTree.Regions tree
            let clusters = regs |> Array.map (fun leaf -> leaf.Cells)
            let revLookup = new Dict<AST.Address,HashSet<AST.Address>>()

            // reverse lookup
            for c in clusters do
                for a in c do
                    revLookup.Add(a,c)

            // vertically sorted array of all cells
            let cells_sorted_vert =
                clusters
                |> Array.map (fun cluster -> cluster |> Seq.toArray)
                |> Array.concat 
                |> Array.sortBy (fun addr -> (addr.X, addr.Y))

            // horizontally sorted array of all cells
            let cells_sorted_horiz =
                cells_sorted_vert
                |> Array.sortBy (fun addr -> (addr.Y, addr.X))

            let coalesce : AST.Address[] -> bool -> unit = (fun (cells: AST.Address[])(vert: bool) ->
                let adjacent = if vert then
                                   (fun (c1: AST.Address)(c2: AST.Address) -> c1.X = c2.X && c1.Y < c2.Y)
                               else
                                   (fun (c1: AST.Address)(c2: AST.Address) -> c1.Y = c2.Y && c1.X < c2.X)

                for i in 0 .. cells.Length - 2 do
                    // get cell
                    let cell = cells.[i]
                    // get cluster of the cell
                    let source = revLookup.[cell]
                    // get adjacent cell, if there is one
                    let maybeAdj = cells.[i+1]
                    // cell's countable
                    let (_,_,co_cell) = hb_inv.[cell]
                    // maybeAdj's countable
                    let (_,_,co_maybe) = hb_inv.[maybeAdj]
                    // is maybeAdj adjacent to cell and has the same cvector?
                    if (adjacent cell maybeAdj) && co_cell.ToCVectorResultant = co_maybe.ToCVectorResultant then
                        // get cluster for maybeAdj
                        let target = revLookup.[maybeAdj]
                        // if I merge source and target, does the merge remain rectangular?
                        if BinaryMinEntropyTree.MergeIsRectangular source target then
                            // add every cell from source to target hashset
                            HashSetUtils.inPlaceUnion source target
                            // update all reverse lookups for source cells
                            for c in source do
                                revLookup.[c] <- target
            )

            // coalesce vertically
            coalesce cells_sorted_vert true

            // coalesce horizontally
            coalesce cells_sorted_horiz false

            // flatten reverse lookup and return clustering
            new Clustering(revLookup.Values |> Seq.distinct)

    and Inner(lefttop: AST.Address, rightbottom: AST.Address) =
        inherit BinaryMinEntropyTree(lefttop, rightbottom)
        let mutable left = None
        let mutable right = None
        let mutable parent = None
        member self.AddParent(p: Inner) : unit =
            parent <- Some p
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

    and Leaf(lefttop: AST.Address, rightbottom: AST.Address, parent: Inner option, cells: Cells) =
        inherit BinaryMinEntropyTree(lefttop, rightbottom)
        member self.Cells : HashSet<AST.Address> = new HashSet<AST.Address>(cells.Keys)
