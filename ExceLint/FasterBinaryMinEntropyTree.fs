namespace ExceLint

    open System.Collections.Generic
    open System.Collections.Immutable
    open CommonTypes
    open CommonFunctions
    open Utils

    type Crd = int*int
    type LT = Crd
    type RB = Crd
    type Rgn = LT * RB

    module RUtil =
        let XLo(r: Rgn) =
            let ((x_lo,_),_) = r
            x_lo

        let XHi(r: Rgn) =
            let (_,(x_hi,_)) = r
            x_hi

        let YLo(r: Rgn) =
            let ((_,y_lo),_) = r
            y_lo

        let YHi(r: Rgn) =
            let (_,(_,y_hi)) = r
            y_hi

    [<AbstractClass>]
    type FasterBinaryMinEntropyTree(lefttop: int*int, rightbottom: int*int, z: int, parentRel: ParentRelation) =
        abstract member Region : string
        default self.Region : string = lefttop.ToString() + ":" + rightbottom.ToString()
        abstract member ToGraphViz : int -> int*string
        abstract member ParentRelation : ParentRelation
        default self.ParentRelation = parentRel
        abstract member Z : int
        default self.Z : int = z

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

        static member MinEntropyPartition(lefttop: int*int)(rightbottom: int*int)(z: int)(fsc: FastSheetCounter)(vert: bool) : Rgn*Rgn=
            let (minX,minY) = lefttop
            let (maxX,maxY) = rightbottom

            let e =
                if vert then
                    (fun (vert)(xlo)(xhi)(ylo)(yhi)(i) ->
                        let entropy_one = fsc.EntropyFor z xlo (i - 1) ylo yhi
                        let entropy_two = fsc.EntropyFor z i xhi ylo yhi
                        entropy_one + entropy_two
                    )
                else
                    (fun (vert)(xlo)(xhi)(ylo)(yhi)(i) ->
                        let entropy_one = fsc.EntropyFor z xlo xhi ylo (i - 1)
                        let entropy_two = fsc.EntropyFor z xlo xhi i yhi
                        entropy_one + entropy_two
                    )

            let decomp =
                if vert then
                    (fun (vert)(xlo)(xhi)(ylo)(yhi)(part) ->
                        let region_one = (minX,minY),(part - 1,maxY)
                        let region_two = (part,minY),(maxX,maxY)
                        region_one,region_two
                    )
                else
                    (fun (vert)(xlo)(xhi)(ylo)(yhi)(part) ->
                        let region_one = (minX,minY),(maxX,part - 1)
                        let region_two = (minX,part),(maxX,maxY)
                        region_one,region_two
                    )

            let indices =
                if vert then
                    [| minX .. maxX |]
                else
                    [| minY .. maxY |]

            // which axis we use depends on whether we are
            // decomposing vertically or horizontally
            let parts = Array.map (fun i ->
                            i, e vert minX maxX minY maxY i
                        ) indices

            // argmin
            let mutable min_j = 0
            for j in [| 0 .. parts.Length - 1 |] do
                let (i,ent) = parts.[j]
                if ent < snd parts.[min_j] then
                    min_j <- j

            let part = fst parts.[min_j]

            decomp vert minX maxX minY maxY part

        static member Infer(fsc: FastSheetCounter)(z: int)(hb_inv: ROInvertedHistogram) : FasterBinaryMinEntropyTree =
            FasterBinaryMinEntropyTree.Decompose fsc z hb_inv

        static member private IsHomogeneous(c: AST.Address[])(ih: ROInvertedHistogram) : bool =
            let (_,_,rep) = ih.[c.[0]]
            c |>
            Array.forall (fun a ->
                let (_,_,co) = ih.[a]
                co = rep
            )

        static member DecomposeAt(fsc: FastSheetCounter)(z: int)(ih: ROInvertedHistogram)(initial_lt: int*int)(initial_rb: int*int) : FasterBinaryMinEntropyTree =
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
                    let leaf = FLeaf(lefttop, rightbottom, z, parentRelation) :> FasterBinaryMinEntropyTree
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
                    let e_vert_left = fsc.EntropyFor z (RUtil.XLo left) (RUtil.XHi left) (RUtil.YLo left) (RUtil.YHi left) 
                    let e_vert_right = fsc.EntropyFor z (RUtil.XLo right) (RUtil.XHi right) (RUtil.YLo right) (RUtil.YHi right)
                    let e_horz_top = fsc.EntropyFor z (RUtil.XLo top) (RUtil.XHi top) (RUtil.YLo top) (RUtil.YHi top)
                    let e_horz_bottom = fsc.EntropyFor z (RUtil.XLo bottom) (RUtil.XHi bottom) (RUtil.YLo bottom) (RUtil.YHi bottom)

                    let e_vert = e_vert_left + e_vert_right
                    let e_horz = e_horz_top + e_horz_bottom

                    // split vertically or horizontally (favor vert for ties)
                    let (entropy,p1,p2) =
                        if e_vert <= e_horz then
                            e_vert, left, right
                        else
                            e_horz, top, bottom

                    // base case 2: perfect decomposition & p1 values same as p2 values
                    let rep_p1 = fsc.ValueFor (RUtil.XLo p1) (RUtil.YLo p1) z
                    let rep_p2 = fsc.ValueFor (RUtil.XLo p2) (RUtil.YLo p2) z
                    if entropy = 0.0 && rep_p1 = rep_p2
                    then
                        let leaf = FLeaf(lefttop, rightbottom, z, parentRelation) :> FasterBinaryMinEntropyTree

                        // is this leaf the root?
                        match parentRelation with
                        | Root -> root_opt <- Some leaf
                        | _ -> ()

                        // add leaf to to link-up list
                        linkUp <- leaf :: linkUp
                    else
                        // "recursive" case
                        let node = FInner(lefttop, rightbottom, z, parentRelation)

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

                match node.ParentRelation with
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

        static member Decompose(fsc: FastSheetCounter)(z: int)(ih: ROInvertedHistogram) : FasterBinaryMinEntropyTree =
            let initial_lt = (fsc.MinXForWorksheet z, fsc.MinYForWorksheet z)
            let initial_rb = (fsc.MaxXForWorksheet z, fsc.MaxYForWorksheet z)
            FasterBinaryMinEntropyTree.DecomposeAt fsc z ih initial_lt initial_rb

        /// <summary>return the leaves of the tree, in order of smallest to largest region</summary>
        static member Regions(tree: FasterBinaryMinEntropyTree) : FLeaf[] =
            match tree with
            | :? FInner as i -> Array.append (FasterBinaryMinEntropyTree.Regions (i.Left)) (FasterBinaryMinEntropyTree.Regions (i.Right))
            | :? FLeaf as l -> [| l |]
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

        static member Coalesce(cs: ImmutableClustering)(ih: ROInvertedHistogram)(indivisibles: ImmutableClustering) : ImmutableClustering =
            // coalesce rectangular regions
            let cs = 
                try
                    FasterBinaryMinEntropyTree.RectangularCoalesce cs ih
                with
                | _ -> failwith "ugh"

            // merge indivisible clusters
            let cs' =
                try
                    FasterBinaryMinEntropyTree.MergeIndivisibles cs indivisibles
                with
                | _ -> failwith "ugh"

            cs'

        static member ClusterIsRectangular(c: HashSet<AST.Address>) : bool =
            let boundingbox = Utils.BoundingBoxHS c 0
            let diff = HashSetUtils.difference boundingbox c
            let isRect = diff.Count = 0
            isRect

        static member ImmClusterIsRectangular(c': ImmutableHashSet<AST.Address>) : bool =
            let c = new HashSet<AST.Address>(c')
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

        static member RectangularCoalesce(cs: ImmutableClustering)(hb_inv: ROInvertedHistogram) : ImmutableClustering =
            let mutable clusters = cs
            let mutable changed = true

            let mutable timesAround = 1

            // DEBUG
            try
                CommonFunctions.ReverseClusterLookup cs |> ignore
            with
            | e -> failwith "bad cluster"

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

        static member SheetAnalysesAreDistinct(regions: ImmutableClustering[]) : bool =
            let mutable ok = true
            let mutable i = 0
            while ok && i < regions.Length do
                let region = regions.[i]
                let addrs = region |> Seq.concat |> Seq.toArray
                let sheet = addrs.[0].WorksheetName // just grab the first one
                let oks = addrs |> Array.map (fun a -> a, a.WorksheetName = sheet)
                ok <- oks |> Array.forall (fun (_,ok') -> ok')
                let blerp =
                    if not ok then
                        "merp"
                    else
                        "blerp"
                i <- i + 1
            ok

    and FInner(lefttop: int*int, rightbottom: int*int, z: int, parentRel: ParentRelation) =
        inherit FasterBinaryMinEntropyTree(lefttop, rightbottom, z, parentRel)
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

    and FLeaf(lefttop: int*int, rightbottom: int*int, z: int, parentRel: ParentRelation) =
        inherit FasterBinaryMinEntropyTree(lefttop, rightbottom, z, parentRel)
        member self.Cells(ih: ROInvertedHistogram)(fsc: FastSheetCounter) : ImmutableHashSet<AST.Address> =
            let (ltX,ltY) = lefttop
            let (rbX,rbY) = rightbottom
            ih
            |> Seq.filter (fun kvp ->
                   let addr = kvp.Key
                   addr.X >= ltX && addr.X <= rbX &&
                   addr.Y >= ltY && addr.Y <= rbY &&
                   (fsc.ZForWorksheet addr.WorksheetName) = z
               )
            |> Seq.map (fun kvp -> kvp.Key)
            |> (fun xs -> xs.ToImmutableHashSet<AST.Address>())
        override self.ToGraphViz(i: int)=
            let start = "\"" + i.ToString() + "\""
            let node = start + " [label=\"" + self.Region + "\"]\n"
            i,node

    and ParentRelation =
    | LeftOf of FInner
    | RightOf of FInner
    | Root