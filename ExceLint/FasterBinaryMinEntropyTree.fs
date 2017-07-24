namespace ExceLint

    open System.Collections.Generic
    open System.Collections.Immutable
    open CommonTypes
    open CommonFunctions
    open Utils

    type Coord = int*int
    type LeftTop = Coord
    type RightBottom = Coord
    type Region = LeftTop * RightBottom

    module Reg =
        let XLo(r: Region) =
            let ((x_lo,_),_) = r
            x_lo

        let XHi(r: Region) =
            let (_,(x_hi,_)) = r
            x_hi

        let YLo(r: Region) =
            let ((_,y_lo),_) = r
            y_lo

        let YHi(r: Region) =
            let (_,(_,y_hi)) = r
            y_hi

    [<AbstractClass>]
    type FasterBinaryMinEntropyTree(lefttop: int*int, rightbottom: int*int, subtree_kind: ParentRelation) =
        abstract member Region : string
        default self.Region : string = lefttop.ToString() + ":" + rightbottom.ToString()
        abstract member ToGraphViz : int -> int*string
        abstract member Subtree : ParentRelation
        default self.Subtree = subtree_kind

        static member GraphViz(t: FasterBinaryMinEntropyTree) : string =
            let (_,graph) = t.ToGraphViz 0
            "graph {" + graph + "}"

        static member Condition(cs: ImmutableClustering)(attribute: AST.Address -> int)(value: int)(noEmptyClusters) : ImmutableClustering =
            cs
            |> Seq.map (fun cluster ->
                cluster
                |> Seq.filter (fun addr ->
                    (attribute addr) = value
                )
                |> (fun cluster2 ->
                    if noEmptyClusters && Seq.length cluster2 = 0 then
                        None
                    else
                        Some((new HashSet<AST.Address>(cluster2)).ToImmutableHashSet())
                   )
            )
            |> Seq.choose id
            |> (fun cs' -> cs'.ToImmutableHashSet())

        /// <summary>
        /// Measure the normalized entropy of a clustering, where the number of cells
        /// inside clusters is used to determine frequency.
        /// </summary>
        /// <param name="c">A Clustering</param>
        static member NormalizedClusteringEntropy(c: ImmutableClustering) : double =
            // count
            let cs = c |> Seq.map (fun reg -> reg.Count) |> Seq.toArray

            // n
            let n = Array.sum cs

            // compute probability vector
            let ps = BasicStats.empiricalProbabilities cs

            // compute entropy
            let entropy = BasicStats.normalizedEntropy ps n

            entropy

        /// <summary>
        /// Measure the entropy of a clustering, where the number of cells
        /// inside clusters is used to determine frequency.
        /// </summary>
        /// <param name="c">A Clustering</param>
        static member ClusteringEntropy(c: Clustering) : double =
            // count
            let cs = c |> Seq.map (fun reg -> reg.Count) |> Seq.toArray

            // compute probability vector
            let ps = BasicStats.empiricalProbabilities cs

            // compute entropy
            let entropy = BasicStats.entropy ps

            entropy

        /// <summary>
        /// The difference in clustering entropy between cTo and cFrom. A negative number
        /// denotes a decrease in entropy from cFrom to cTo whereas a positive number
        /// denotes an increase in entropy from cFrom to cTo.
        /// </summary>
        /// <param name="cFrom">original clustering</param>
        /// <param name="cTo">new clustering</param>
        static member ClusteringEntropyDiff(cFrom: ImmutableClustering)(cTo: ImmutableClustering) : double =
            let c1e = FasterBinaryMinEntropyTree.NormalizedClusteringEntropy cFrom
            let c2e = FasterBinaryMinEntropyTree.NormalizedClusteringEntropy cTo
            c2e - c1e

        static member private MinEntropyPartition(lefttop: int*int)(rightbottom: int*int)(z: int)(fsc: FastSheetCounter)(vert: bool) : Region*Region=
            let (minX,minY) = lefttop
            let (maxX,maxY) = rightbottom

            // which axis we use depends on whether we are
            // decomposing vertically or horizontally
            if vert then
                let part = Utils.argmin (fun i ->
                               let entropy_one = fsc.EntropyFor z minX (i - 1) minY maxY
                               let entropy_two = fsc.EntropyFor z i maxX minY maxY
                               entropy_one + entropy_two
                           ) [| minX .. maxX |]

                let region_one = (minX,minY),(part - 1,maxY)
                let region_two = (part,minY),(maxX,maxY)
            
                region_one,region_two
            else
                let part = Utils.argmin (fun i ->
                               let entropy_one = fsc.EntropyFor z minX maxX minY (i - 1)
                               let entropy_two = fsc.EntropyFor z minX maxX i maxY
                               entropy_one + entropy_two
                           ) [| minY .. maxY |]

                let region_one = (minX,minY),(maxX,part - 1)
                let region_two = (minX,part),(maxX,maxY)
            
                region_one,region_two

        static member Infer(fsc: FastSheetCounter)(z: int)(hb_inv: ROInvertedHistogram) : FasterBinaryMinEntropyTree =
            FasterBinaryMinEntropyTree.Decompose fsc z hb_inv

        static member private IsHomogeneous(c: AST.Address[])(ih: ROInvertedHistogram) : bool =
            let (_,_,rep) = ih.[c.[0]]
            c |>
            Array.forall (fun a ->
                let (_,_,co) = ih.[a]
                co = rep
            )

        static member private Decompose (fsc: FastSheetCounter)(z: int)(ih: ROInvertedHistogram) : FasterBinaryMinEntropyTree =
            let initial_lt = (fsc.MinXForWorksheet z, fsc.MinYForWorksheet z)
            let initial_rb = (fsc.MaxXForWorksheet z, fsc.MaxYForWorksheet z)
            let mutable todos = [ (Root, (initial_lt, initial_rb)) ]
            let mutable linkUp = []
            let mutable root_opt = None

            // process work list
            while not todos.IsEmpty do
                // grab next item
                let (parentRelation, (lefttop, rightbottom)) = todos.Head
                todos <- todos.Tail

                // base case 1: there's only 1 cell
                if lefttop = rightbottom then
                    let leaf = Leaf(lefttop, rightbottom, parentRelation) :> FasterBinaryMinEntropyTree
                    // add leaf to link-up list
                    linkUp <- leaf :: linkUp

                    // is this leaf the root?
                    match parentRelation with
                    | Root -> root_opt <- Some leaf
                    | _ -> ()
                else
                    // find the minimum entropy decomposition
                    let (left,right) = FasterBinaryMinEntropyTree.MinEntropyPartition lefttop rightbottom z fsc true
                    let (top,bottom) = FasterBinaryMinEntropyTree.MinEntropyPartition lefttop rightbottom z fsc false

                    // compute entropies again
                    let e_vert = fsc.EntropyFor z (Reg.XLo left) (Reg.XHi left) (Reg.YLo left) (Reg.YHi left) +
                                 fsc.EntropyFor z (Reg.XLo right) (Reg.XHi right) (Reg.YLo right) (Reg.YHi right)
                    let e_horz = fsc.EntropyFor z (Reg.XLo top) (Reg.XHi top) (Reg.YLo top) (Reg.YHi top) +
                                 fsc.EntropyFor z (Reg.XLo bottom) (Reg.XHi bottom) (Reg.YLo bottom) (Reg.YHi bottom)

                    // split vertically or horizontally (favor vert for ties)
                    let (entropy,p1,p2) =
                        if e_vert <= e_horz then
                            e_vert, left, right
                        else
                            e_horz, top, bottom

                    // base case 2: perfect decomposition & p1 values same as p2 values
                    let rep_p1 = fsc.ValueFor (Reg.XLo p1) (Reg.YLo p1) z
                    let rep_p2 = fsc.ValueFor (Reg.XLo p2) (Reg.YLo p2) z
                    if entropy = 0.0 && rep_p1 = rep_p2
                    then
                        let leaf = Leaf(lefttop, rightbottom, parentRelation) :> FasterBinaryMinEntropyTree

                        // is this leaf the root?
                        match parentRelation with
                        | Root -> root_opt <- Some leaf
                        | _ -> ()

                        // add leaf to to link-up list
                        linkUp <- leaf :: linkUp
                    else
                        // "recursive" case
                        let node = Inner(lefttop, rightbottom, parentRelation)

                        // is this node the root?
                        match parentRelation with
                        | Root -> root_opt <- Some (node :> FasterBinaryMinEntropyTree)
                        | _ -> ()

                        // add next nodes to work list
                        todos <- (LeftOf node, p1) :: (RightOf node, p2) :: todos

                // process "link-up" list
            while not linkUp.IsEmpty do
                // grab next item
                let node = linkUp.Head
                linkUp <- linkUp.Tail

                match node.Subtree with
                | LeftOf parent ->
                    // add parent to linkup list
                    linkUp <- (parent :> FasterBinaryMinEntropyTree) :: linkUp

                    // make link
                    parent.AddLeft node
                | RightOf parent ->
                    // add parent to linkup list
                    linkUp <- (parent :> FasterBinaryMinEntropyTree) :: linkUp

                    // make link
                    parent.AddRight node
                | Root -> ()    // do nothing

            match root_opt with
            | Some root -> root
            | None -> failwith "this should never happen"

        /// <summary>return the leaves of the tree, in order of smallest to largest region</summary>
        static member Regions(tree: FasterBinaryMinEntropyTree) : Leaf[] =
            match tree with
            | :? Inner as i -> Array.append (FasterBinaryMinEntropyTree.Regions (i.Left)) (FasterBinaryMinEntropyTree.Regions (i.Right))
            | :? Leaf as l -> [| l |]
            | _ -> failwith "Unknown tree node type."

        static member MergeIndivisibles(ic: ImmutableClustering)(indivisibles: ImmutableClustering) : ImmutableClustering =
            let cs = ToMutableClustering ic

            // get rmap
            let reverseLookup = ReverseClusterLookupMutable cs

            // coalesce indivisibles
            for i in indivisibles do
                // get the biggest cluster
                let biggest = i |> Seq.map (fun a -> reverseLookup.[a]) |> Seq.maxBy (fun c -> c.Count)
                for a in i do
                    if not (biggest.Contains a) then
                        // get the cluster that currently belongs to
                        let c = reverseLookup.[a]

                        // merge all of the cells from c into biggest
                        if not (biggest = c) then
                            HashSetUtils.inPlaceUnion c biggest

                            // remove c from clustering
                            cs.Remove c |> ignore

            // restore immutability
            ToImmutableClustering cs

        static member Clustering(tree: FasterBinaryMinEntropyTree)(ih: ROInvertedHistogram)(indivisibles: ImmutableClustering) : ImmutableClustering =
            // coalesce rectangular regions
            let cs = FasterBinaryMinEntropyTree.RectangularClustering tree ih

            // merge indivisible clusters
            let cs' = FasterBinaryMinEntropyTree.MergeIndivisibles cs indivisibles

            cs'

        static member ClusterIsRectangular(c: HashSet<AST.Address>) : bool =
            let boundingbox = Utils.BoundingBoxHS c 0
            let diff = HashSetUtils.difference boundingbox c
            let isRect = diff.Count = 0
            isRect

        static member ClusteringContainsOnlyRectangles(cs: Clustering) : bool =
            cs |> Seq.fold (fun a c -> a && FasterBinaryMinEntropyTree.ClusterIsRectangular c) true

        static member CellMergeIsRectangular(source: AST.Address)(target: ImmutableHashSet<AST.Address>) : bool =
            let merged = new HashSet<AST.Address>(target.Add source)
            FasterBinaryMinEntropyTree.ClusterIsRectangular merged

        static member ImmMergeIsRectangular(source: ImmutableHashSet<AST.Address>)(target: ImmutableHashSet<AST.Address>) : bool =
            let merged = source.Union target
            FasterBinaryMinEntropyTree.ClusterIsRectangular (new HashSet<AST.Address>(merged))

        static member MergeIsRectangular(source: HashSet<AST.Address>)(target: HashSet<AST.Address>) : bool =
            let merged = HashSetUtils.union source target
            FasterBinaryMinEntropyTree.ClusterIsRectangular merged

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
            let clusters' = ToMutableClustering clusters

            let revLookup = ReverseClusterLookupMutable clusters'

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
                        let isRect = FasterBinaryMinEntropyTree.MergeIsRectangular source target
                        if isRect then
                            // add every cell from source to target hashset
                            HashSetUtils.inPlaceUnion source target
                            // update all reverse lookups for source cells
                            for c in source do
                                revLookup.[c] <- target
                            // remove source from clusters
                            clusters'.Remove source |> ignore

            ToImmutableClustering clusters'

        static member RectangularClustering(tree: FasterBinaryMinEntropyTree)(hb_inv: ROInvertedHistogram) : ImmutableClustering =
            // coalesce all cells that have the same cvector,
            // ensuring that all merged clusters remain rectangular
            let regs = FasterBinaryMinEntropyTree.Regions tree
            let clusters = regs |> Array.map (fun leaf -> leaf.Cells hb_inv) |> makeImmutableGenericClustering

            FasterBinaryMinEntropyTree.RectangularCoalesce clusters hb_inv

        static member RectangularCoalesce(cs: ImmutableClustering)(hb_inv: ROInvertedHistogram) : ImmutableClustering =
            let mutable clusters = cs
            let mutable changed = true

            let mutable timesAround = 1

            while changed do
                // coalesce vertical ordering horizontally
                let clusters' = FasterBinaryMinEntropyTree.CoaleseAdjacentClusters false clusters hb_inv

                // coalesce horizontal ordering vertically
                let clusters'' = FasterBinaryMinEntropyTree.CoaleseAdjacentClusters true clusters' hb_inv

                if CommonFunctions.SameClustering clusters clusters'' then
                    changed <- false
                else
                    if timesAround > 10000 then
                        failwith "Coalesce convergence error."

                    changed <- true
                    clusters <- clusters''

                timesAround <- timesAround + 1

            // return clustering
            clusters

    and Inner(lefttop: int*int, rightbottom: int*int, parentRel: ParentRelation) =
        inherit FasterBinaryMinEntropyTree(lefttop, rightbottom, parentRel)
        let mutable left = None
        let mutable right = None
        member self.AddLeft(n: FasterBinaryMinEntropyTree) : unit =
            left <- Some n
        member self.AddRight(n: FasterBinaryMinEntropyTree) : unit =
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

    and Leaf(lefttop: int*int, rightbottom: int*int, parentRel: ParentRelation) =
        inherit FasterBinaryMinEntropyTree(lefttop, rightbottom, parentRel)
        member self.Cells(ih: ROInvertedHistogram) : ImmutableHashSet<AST.Address> =
            let (ltX,ltY) = lefttop
            let (rbX,rbY) = rightbottom
            ih
            |> Seq.filter (fun kvp ->
                   let addr = kvp.Key
                   addr.X >= ltX && addr.X <= rbX &&
                   addr.Y >= ltY && addr.Y <= rbY
               )
            |> Seq.map (fun kvp -> kvp.Key)
            |> (fun xs -> xs.ToImmutableHashSet<AST.Address>())
        override self.ToGraphViz(i: int)=
            let start = "\"" + i.ToString() + "\""
            let node = start + " [label=\"" + self.Region + "\"]\n"
            i,node

    and ParentRelation =
    | LeftOf of Inner
    | RightOf of Inner
    | Root