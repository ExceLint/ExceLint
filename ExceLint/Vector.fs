namespace ExceLint
    open Depends
    open Feature
    open System
    open System.Collections.Generic
    open Utils

    module public Vector =
        type public Directory = string
        type public WorkbookName = string
        type public WorksheetName = string
        type public Path = Directory*WorkbookName*WorksheetName
        type public X = int    // i.e., column displacement
        type public Y = int    // i.e., row displacement
        type public Z = int    // i.e., worksheet displacement (0 if same sheet, 1 if different)

        // components for mixed vectors
        type public VectorComponent =
        | Abs of int
        | Rel of int
            override self.ToString() : string =
                match self with
                | Abs(i) -> "Abs(" + i.ToString() + ")"
                | Rel(i) -> "Rel(" + i.ToString() + ")"

        // the vector, relative to an origin
        type public Coordinates = (X*Y*Path)
        type public RelativeVector = (X*Y*Z)
        type public MixedVector = (VectorComponent*VectorComponent*Path)
        type public SquareVector(dx: double, dy: double, x: double, y: double) =
            member self.dx = dx
            member self.dy = dy
            member self.x = x
            member self.y = y
            member self.AsArray = [| dx; dy; x; y |]
            override self.Equals(o: obj) : bool =
                match o with
                | :? SquareVector as o' ->
                    (dx, dy, x, y) = (o'.dx, o'.dy, o'.x, o'.y)
                | _ -> false
            override self.GetHashCode() = hash (dx, dy, x, y)
            override self.ToString() =
                "(" +
                dx.ToString() + "," +
                dy.ToString() + "," +
                x.ToString() + "," +
                y.ToString() +
                ")"

        // handy datastructures
        type public Edge = SquareVector*SquareVector
        type public DistDict = Dictionary<Edge,double>

        // the first component is the tail (start) and the second is the head (end)
        type public FullyQualifiedVector =
        | MixedFQVector of Coordinates*MixedVector
        | AbsoluteFQVector of Coordinates*Coordinates
            override self.ToString() : string =
                match self with
                | MixedFQVector(tail,head) -> tail.ToString() + " -> " + head.ToString()
                | AbsoluteFQVector(tail,head) -> tail.ToString() + " -> " + head.ToString()

        let private fullPath(addr: AST.Address) : string*string*string =
            // portably create full path from components
            (addr.Path, addr.WorkbookName, addr.WorksheetName)

        let private vector(tail: AST.Address)(head: AST.Address)(mixed: bool) : FullyQualifiedVector =
            let tailXYP = (tail.X, tail.Y, fullPath tail)
            if mixed then
                let X = match head.XMode with
                        | AST.AddressMode.Absolute -> Abs(head.X)
                        | AST.AddressMode.Relative -> Rel(head.X)
                let Y = match head.YMode with
                        | AST.AddressMode.Absolute -> Abs(head.Y)
                        | AST.AddressMode.Relative -> Rel(head.Y)
                let headXYP = (X, Y, fullPath head)
                MixedFQVector(tailXYP, headXYP)
            else
                let headXYP = (head.X, head.Y, fullPath head)
                AbsoluteFQVector(tailXYP, headXYP)

        let private originPath(dag: DAG) : Path =
            (dag.getWorkbookDirectory(), dag.getWorkbookName(), dag.getWorksheetNames().[0]);

        let private vectorPathDiff(p1: Path)(p2: Path) : int =
            if p1 <> p2 then 1 else 0

        // the origin is defined as x = 0, y = 0, z = 0 if first sheet in the workbook else 1
        let private pathDiff(p: Path)(dag: DAG) : int =
            let op = originPath dag
            vectorPathDiff p op

        // represent the position of the head of the vector relative to the tail, (x1,y1,z1)
        // if the reference is off-sheet then optionally ignore X and Y vector components
        let private relativeToTail(absVect: FullyQualifiedVector)(dag: DAG)(offSheetInsensitive: bool) : RelativeVector =
            match absVect with
            | AbsoluteFQVector(tail,head) ->
                let (x1,y1,p1) = tail
                let (x2,y2,p2) = head
                if offSheetInsensitive && p1 <> p2 then
                    (0, 0, dag.getPathClosureIndex(p2))
                else
                    (x2-x1, y2-y1, vectorPathDiff p2 p1)
            | MixedFQVector(tail,head) ->
                let (x1,y1,p1) = tail
                let (x2,y2,p2) = head
                let x' = match x2 with
                            | Rel(x) -> x - x1
                            | Abs(x) -> x
                let y' = match y2 with
                            | Rel(y) -> y - y1
                            | Abs(y) -> y
                if offSheetInsensitive && p1 <> p2 then
                    (0, 0, dag.getPathClosureIndex(p2))
                else
                    (x', y', vectorPathDiff p2 p1)

        // represent the position of the the head of the vector relative to the origin, (0,0,0)
        let private relativeToOrigin(absVect: FullyQualifiedVector)(dag: DAG)(offSheetInsensitive: bool) : RelativeVector =
            match absVect with
            | AbsoluteFQVector(tail,head) ->
                let (_,_,tp) = tail
                let (x,y,p) = head
                if offSheetInsensitive && tp <> p then
                    (0, 0, dag.getPathClosureIndex(p))
                else
                    (x, y, pathDiff p dag)
            | MixedFQVector(tail,head) ->
                let (_,_,tp) = tail
                let (x,y,p) = head
                let x' = match x with | Abs(xa) -> xa | Rel(xr) -> xr
                let y' = match y with | Abs(ya) -> ya | Rel(yr) -> yr
                if offSheetInsensitive && tp <> p then
                    (0, 0, dag.getPathClosureIndex(p))
                else
                    (x', y', pathDiff p dag)

        let private L2Norm(X: double[]) : double =
            Math.Sqrt(
                Array.sumBy (fun x -> Math.Pow(x, 2.0)) X
            )

        let private relativeVectorToRealVectorArr(v: RelativeVector) : double[] =
            let (x,y,z) = v
            [|
                System.Convert.ToDouble(x);
                System.Convert.ToDouble(y);
                System.Convert.ToDouble(z);
            |]

        let private L2NormRV(v: RelativeVector) : double =
            L2Norm(relativeVectorToRealVectorArr(v))

        let private L2NormRVSum(vs: RelativeVector[]) : double =
            vs |> Array.map L2NormRV |> Array.sum

        let private SquareMatrix(origin: X*Y)(vs: RelativeVector[]) : X*Y*X*Y =
            let (x,y) = origin
            let xyoff = vs |> Array.fold (fun (xacc: X, yacc: Y)(x': X,y': Y,z': Z) -> xacc + x', yacc + y') (0,0)
            (fst xyoff, snd xyoff, x, y)

        let transitiveInputVectors(fCell: AST.Address)(dag : DAG)(depth: int option)(mixed: bool) : FullyQualifiedVector[] =
            let rec tfVect(tailO: AST.Address option)(head: AST.Address)(depth: int option) : FullyQualifiedVector list =
                let vlist = match tailO with
                            | Some tail -> [vector tail head mixed]
                            | None -> []

                match depth with
                | Some(0) -> vlist
                | Some(d) -> tfVect_b head (Some(d-1)) vlist
                | None -> tfVect_b head None vlist

            and tfVect_b(tail: AST.Address)(nextDepth: int option)(vlist: FullyQualifiedVector list) : FullyQualifiedVector list =
                if (dag.isFormula tail) then
                    // find all of the inputs for source
                    let heads_single = dag.getFormulaSingleCellInputs tail |> List.ofSeq
                    let heads_vector = dag.getFormulaInputVectors tail |>
                                            List.ofSeq |>
                                            List.map (fun rng -> rng.Addresses() |> Array.toList) |>
                                            List.concat
                    let heads = heads_single @ heads_vector
                    // recursively call this function
                    vlist @ (List.map (fun head -> tfVect (Some tail) head nextDepth) heads |> List.concat)
                else
                    vlist
    
            tfVect None fCell depth |> List.toArray

        let inputVectors(fCell: AST.Address)(dag : DAG)(mixed: bool) : FullyQualifiedVector[] =
            transitiveInputVectors fCell dag (Some 1) mixed

        let transitiveOutputVectors(dCell: AST.Address)(dag : DAG)(depth: int option)(mixed: bool) : FullyQualifiedVector[] =
            let rec tdVect(sourceO: AST.Address option)(sink: AST.Address)(depth: int option) : FullyQualifiedVector list =
                let vlist = match sourceO with
                            | Some source -> [vector sink source mixed]
                            | None -> []

                match depth with
                | Some(0) -> vlist
                | Some(d) -> tdVect_b sink (Some(d-1)) vlist
                | None -> tdVect_b sink None vlist

            and tdVect_b(sink: AST.Address)(nextDepth: int option)(vlist: FullyQualifiedVector list) : FullyQualifiedVector list =
                    // find all of the formulas that use sink
                    let outAddrs = dag.getFormulasThatRefCell sink
                                    |> Array.toList
                    let outAddrs2 = Array.map (dag.getFormulasThatRefVector) (dag.getVectorsThatRefCell sink)
                                    |> Array.concat |> Array.toList
                    let allFrm = outAddrs @ outAddrs2 |> List.distinct

                    // recursively call this function
                    vlist @ (List.map (fun sink' -> tdVect (Some sink) sink' nextDepth) allFrm |> List.concat)

            tdVect None dCell depth |> List.toArray

        let outputVectors(dCell: AST.Address)(dag : DAG)(mixed: bool) : FullyQualifiedVector[] =
            transitiveOutputVectors dCell dag (Some 1) mixed

        let getVectors(cell: AST.Address)(dag: DAG)(transitive: bool)(isForm: bool)(isRel: bool)(isMixed: bool)(isOSI: bool) : RelativeVector[] =
            let depth = if transitive then None else (Some 1)
            let vectors =
                if isForm then
                    transitiveInputVectors
                else
                    transitiveOutputVectors
            let rebase =
                if isRel then
                    Array.map (fun v -> relativeToTail v dag isOSI)
                else
                    Array.map (fun v -> relativeToOrigin v dag isOSI)
            let output = rebase (vectors cell dag depth isMixed)
            output

        let private oldAspect(data: (X*Y*X*Y)[]) : double =
            // compute aspect ratio
            let width_min = Array.map (fun (sdx, sdy, x, y) -> x) data |> Array.min
            let height_min = Array.map (fun (sdx, sdy, x, y) -> y) data |> Array.min
            let width_max = Array.map (fun (sdx, sdy, x, y) -> x) data |> Array.max
            let height_max = Array.map (fun (sdx, sdy, x, y) -> y) data |> Array.max
            let width = width_max - width_min + 1
            let height = height_max - height_min + 1
            if width < height then
                float(height) / float(width)
            else
                float(width) / float(height)

        let private normalizeColumn(data: double[]) : double[] =
            let min = Array.min data
            let max = Array.max data
            if max = min then
                Array.create data.Length 0.5
            else
                Array.map (fun x -> (x - min) / ( max - min)) data

        let SquareMatrixForCell(cell: AST.Address)(dag: DAG) : X*Y*X*Y =
            let debugfrm = dag.getFormulaAtAddress(cell)
            let vs = getVectors cell dag (*transitive*) false (*isForm*) true (*isRel*) true (*isMixed*) true (*isOffSheetInsensitive*) true
            SquareMatrix (cell.X, cell.Y) vs

        let column(i: int)(data: (X*Y*X*Y)[]) : double[] =
            Array.map (fun row ->
               let (x1,x2,x3,x4) = row
               let arr = [| x1; x2; x3; x4 |]
               double (arr.[i])
            ) data

        let combine(cols: double[][]) : SquareVector[] =
            let len = cols.[0].Length
            let mutable rows: SquareVector list = []
            for i in 0..len-1 do
                rows <- SquareVector(cols.[0].[i], cols.[1].[i], cols.[2].[i], cols.[3].[i]) :: rows
            List.rev rows |> List.toArray

        let SingleSquareVector(cell: AST.Address)(dag: DAG)(normalizeRefSpace: bool)(normalizeSSSpace: bool) : SquareVector =
            // yes, I am lazy; this is basically a copy of AllSquareVectors; TODO refactor
            let fs = [|cell|]
            let mats = Array.map (fun f -> SquareMatrixForCell f dag) fs

            let sdx_vect = if normalizeRefSpace then normalizeColumn (column 0 mats) else column 0 mats
            let sdy_vect = if normalizeRefSpace then normalizeColumn (column 1 mats) else column 1 mats
            let x_vect = if normalizeSSSpace then normalizeColumn (column 2 mats) else column 2 mats
            let y_vect = if normalizeSSSpace then normalizeColumn (column 3 mats) else column 3 mats

            let arr = combine([| sdx_vect; sdy_vect; x_vect; y_vect |])
            arr.[0]

        let AllSquareVectors(dag: DAG)(normalizeRefSpace: bool)(normalizeSSSpace: bool)(wsname: string) : SquareVector[] =
            let fs = dag.getAllFormulaAddrs() |> Array.filter (fun f -> f.WorksheetName = wsname) 
            let mats = Array.map (fun f -> SquareMatrixForCell f dag) fs

            let sdx_vect = if normalizeRefSpace then normalizeColumn (column 0 mats) else column 0 mats
            let sdy_vect = if normalizeRefSpace then normalizeColumn (column 1 mats) else column 1 mats
            let x_vect = if normalizeSSSpace then normalizeColumn (column 2 mats) else column 2 mats
            let y_vect = if normalizeSSSpace then normalizeColumn (column 3 mats) else column 3 mats

            combine([| sdx_vect; sdy_vect; x_vect; y_vect |])

        let DistDictToSVHashSet(dd: DistDict) : HashSet<SquareVector> =
            let hs = new HashSet<SquareVector>()
            for edge in dd do
                let (efrom,eto) = edge.Key
                hs.Add(efrom) |> ignore
                hs.Add(eto) |> ignore
            hs

        let dist(e: Edge) : double =
            let (p,p') = e

            let p_arr = p.AsArray
            let p'_arr = p'.AsArray

            let elts = Array.map (fun i ->
                           (p_arr.[i] - p'_arr.[i]) * (p_arr.[i] - p'_arr.[i])
                       ) [|0..3|]

            let res = Math.Sqrt(Array.sum elts)

            res

        let Nk(p: SquareVector)(k : int)(G: HashSet<SquareVector>)(DD: DistDict) : HashSet<SquareVector> =
            let subgraph = DD |>
                            Seq.filter (fun (kvp: KeyValuePair<Edge,double>) ->
                                let p' = fst kvp.Key
                                let o = snd kvp.Key
                                let b1 = p <> p'       // p must not be in Nk
                                let b2 = p <> o       // also, we don't care about dist(p,p)
                                let b3 = G.Contains(o)   // and G may also be a subset of points
                                b1 && b2 && b3
                            ) |>
                            Seq.toArray
            let subgraph_sorted = subgraph |> Array.sortBy (fun (kvp: KeyValuePair<Edge,double>) -> kvp.Value)
            let subgraph_sorted_k = subgraph_sorted |> Array.take k
            let kn = subgraph_sorted_k |> Array.map (fun (kvp: KeyValuePair<Edge,double>) -> snd kvp.Key)

            assert (Array.length kn <= k)

            new HashSet<SquareVector>(kn)

        let Nk_cells(p: AST.Address)(k: int)(dag: DAG)(normalizeRefSpace: bool)(normalizeSSSpace: bool)(DD: DistDict) : HashSet<AST.Address> =
            let p' = SingleSquareVector p dag normalizeRefSpace normalizeSSSpace
            let Gmap = Array.map (fun c ->
                           SingleSquareVector c dag normalizeRefSpace normalizeSSSpace, c
                       ) (dag.getAllFormulaAddrs() |> Array.filter (fun c -> c.WorksheetName = p.WorksheetName))
                       |> adict
            let G = Gmap.Keys |> Seq.toArray |> fun cs -> new HashSet<SquareVector>(cs)
            let nb = Nk p' k G DD
            Seq.map (fun n -> Gmap.[n]) nb
            |> fun ns -> new HashSet<AST.Address>(ns)

        let hsDiff(A: HashSet<'a>)(B: HashSet<'a>) : HashSet<'a> =
            let hs = new HashSet<'a>()
            for a in A do
                if not (B.Contains a) then
                    hs.Add a |> ignore
            hs
            
        let edges(G: SquareVector[]) : Edge[] =
            Array.map (fun i -> Array.map (fun j -> i,j) G) G |> Array.concat

        let pairwiseDistances(E: Edge[]) : DistDict =
            let d = new DistDict()
            for e in E do
                d.Add(e, dist e)
            d

        let SBNTrail(p: SquareVector)(G: HashSet<SquareVector>)(DD: DistDict) : Edge[] =
            let rec sbnt(path: Edge list) : Edge list =
                // make a hashset out of the path
                let E = new HashSet<SquareVector>()
                for e in path do
                    let (start,dest) = e
                    E.Add start |> ignore
                    E.Add dest |> ignore

                // compute G\E
                let G' = hsDiff G E

                if E.Count = 0 || G'.Count <> 0 then
                    // base case; inductive steps follow
                    if E.Count = 0 then
                        E.Add p |> ignore

                    // find min distance edges to points in G' from all points in E
                    let edges = Seq.map (fun dest ->
                                    // create candidate edges
                                    Seq.map (fun origin -> (origin,dest)) E
                                ) G' |> Seq.concat

                    // rank by distance, smallest to largest
                    let edges_ranked = Seq.sortBy (fun edge -> DD.[edge]) edges |> Seq.toList
                
                    // add smallest edge to path
                    sbnt(edges_ranked.Head :: path)
                else
                    // terminating case: G'.Count = 0
                    path

            sbnt([]) |> List.rev |> List.toArray

        let acDist(p: SquareVector)(es: Edge[])(DD: DistDict) : double =
            // there are r - 1 edges, so r = len(es) + 1
            let r = float es.Length + 1.0
            let cost_desc = Array.map (fun e -> DD.[e]) es

            let cost = Array.mapi (fun i e ->
                            let i' = float i + 1.0
                            assert (i' > 0.0)
                            if i' <= (r - 1.0) then
                                let expr = (2.0 * (r - i') * (DD.[e]))
                                            /
                                            (r * (r - 1.0))
                                expr
                            else
                                0.0
                        ) es
            let tot = Array.sum cost
            tot

        // compute a reasonable k
        let COFk(cell: AST.Address)(dag: DAG)(normalizeRefSpace: bool)(normalizeSSSpace: bool) : int =
            let wsname = cell.WorksheetName
            let wscells = dag.allInputs() |> Array.filter (fun c -> c.WorksheetName = wsname)
            let miny: int = Array.map (fun (c: AST.Address) -> c.Row) wscells |> Array.min
            let maxy: int = Array.map (fun (c: AST.Address) -> c.Row) wscells |> Array.max
            maxy - miny

        let COF(p: SquareVector)(k: int)(G: HashSet<SquareVector>)(DD: DistDict) : double =
            // get p's k nearest neighbors
            let kN = Nk p k G DD
            assert not (kN.Contains p)
            // compute SBN trail for p
            let es = SBNTrail p kN DD
            // compute ac-dist for kN union p
            let acDist_p = acDist p es DD
            // compute the average chaining distance for each point o in kN.
            let acs = kN |>
                      Seq.map (fun o ->
                          let o_kN = Nk o k G DD
                          assert not (o_kN.Contains o)
                          let o_es = SBNTrail o o_kN DD
                          acDist o o_es DD
                      ) |>
                      Seq.toArray
            // compute COF
            let res = acDist_p / ((1.0 / float k) * Array.sum acs)
            res

        type DeepInputVectorRelativeL2NormSum() = 
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag : DAG) : double = 
                L2NormRVSum (getVectors cell dag (*transitive*) true (*isForm*) true (*isRel*) true (*isMixed*) false (*isOffSheetInsensitive*) true)
            static member capability : string*Capability =
                (typeof<DeepInputVectorRelativeL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = DeepInputVectorRelativeL2NormSum.run } )

        type DeepOutputVectorRelativeL2NormSum() = 
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag : DAG) : double = 
                L2NormRVSum (getVectors cell dag (*transitive*) true (*isForm*) false (*isRel*) true (*isMixed*) false (*isOffSheetInsensitive*) true)
            static member capability : string*Capability =
                (typeof<DeepOutputVectorRelativeL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = DeepOutputVectorRelativeL2NormSum.run } )

        type DeepInputVectorAbsoluteL2NormSum() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : double =
                L2NormRVSum (getVectors cell dag (*transitive*) true (*isForm*) true (*isRel*) false (*isMixed*) false (*isOffSheetInsensitive*) true)
            static member capability : string*Capability =
                (typeof<DeepInputVectorAbsoluteL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = DeepInputVectorAbsoluteL2NormSum.run } )

        type DeepOutputVectorAbsoluteL2NormSum() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : double =
                L2NormRVSum (getVectors cell dag (*transitive*) true (*isForm*) false (*isRel*) false (*isMixed*) false (*isOffSheetInsensitive*) true)
            static member capability : string*Capability =
                (typeof<DeepOutputVectorAbsoluteL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = DeepOutputVectorAbsoluteL2NormSum.run } )

        type ShallowInputVectorRelativeL2NormSum() = 
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag : DAG) : double = 
                L2NormRVSum (getVectors cell dag (*transitive*) false (*isForm*) true (*isRel*) true (*isMixed*) false (*isOffSheetInsensitive*) true)
            static member capability : string*Capability =
                (typeof<ShallowInputVectorRelativeL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowInputVectorRelativeL2NormSum.run } )

        type ShallowOutputVectorRelativeL2NormSum() = 
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag : DAG) : double = 
                L2NormRVSum (getVectors cell dag (*transitive*) false (*isForm*) false (*isRel*) true (*isMixed*) false (*isOffSheetInsensitive*) true)
            static member capability : string*Capability =
                (typeof<ShallowOutputVectorRelativeL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowOutputVectorRelativeL2NormSum.run } )

        type ShallowInputVectorAbsoluteL2NormSum() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : double =
                L2NormRVSum (getVectors cell dag (*transitive*) false (*isForm*) true (*isRel*) false (*isMixed*) false (*isOffSheetInsensitive*) true)
            static member capability : string*Capability =
                (typeof<ShallowInputVectorAbsoluteL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowInputVectorAbsoluteL2NormSum.run } )

        type ShallowOutputVectorAbsoluteL2NormSum() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : double =
                L2NormRVSum (getVectors cell dag (*transitive*) false (*isForm*) false (*isRel*) false (*isMixed*) false (*isOffSheetInsensitive*) true)
            static member capability : string*Capability =
                (typeof<ShallowOutputVectorAbsoluteL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowOutputVectorAbsoluteL2NormSum.run } )

        type ShallowInputVectorMixedL2NormSum() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : double =
                L2NormRVSum (getVectors cell dag (*transitive*) false (*isForm*) true (*isRel*) true (*isMixed*) true (*isOffSheetInsensitive*) true)
            static member capability : string*Capability =
                (typeof<ShallowInputVectorMixedL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowInputVectorMixedL2NormSum.run } )

        type ShallowOutputVectorMixedL2NormSum() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : double =
                L2NormRVSum (getVectors cell dag (*transitive*) false (*isForm*) false (*isRel*) true (*isMixed*) true (*isOffSheetInsensitive*) true)
            static member capability : string*Capability =
                (typeof<ShallowOutputVectorMixedL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowOutputVectorMixedL2NormSum.run } )

        type DeepInputVectorMixedL2NormSum() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : double =
                L2NormRVSum (getVectors cell dag (*transitive*) true (*isForm*) true (*isRel*) true (*isMixed*) true (*isOffSheetInsensitive*) true)
            static member capability : string*Capability =
                (typeof<DeepInputVectorMixedL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = DeepInputVectorMixedL2NormSum.run } )

        type DeepOutputVectorMixedL2NormSum() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : double =
                L2NormRVSum (getVectors cell dag (*transitive*) true (*isForm*) false (*isRel*) true (*isMixed*) true (*isOffSheetInsensitive*) true)
            static member capability : string*Capability =
                (typeof<DeepOutputVectorMixedL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = DeepOutputVectorMixedL2NormSum.run } )

        type Product private () =
            let mutable state = 0
            static let instance = Product()
            static member Instance = instance
            member this.DoSomething() = 
                state <- state + 1
                printfn "Doing something for the %i time" state
                ()

        let private MutateCache(cache: Dictionary<WorkbookName,Dictionary<WorksheetName,DistDict>>)(dag: DAG)(normalizeRefSpace: bool)(normalizeSSSpace: bool) : unit =
                let wbname = dag.getWorkbookName()

                // get workbook-level cache
                let wbcache = if not (cache.ContainsKey wbname) then
                                  cache.Add(wbname, new Dictionary<WorksheetName,DistDict>())
                                  cache.[wbname]
                              else
                                  cache.[wbname]
                
                // loop through all worksheet names
                Array.iter (fun wsname ->
                    if not (wbcache.ContainsKey(wsname)) then
                        // get vectors
                        let vectors = AllSquareVectors dag normalizeRefSpace normalizeSSSpace wsname
                        // compute distances
                        let dists = pairwiseDistances (edges vectors)
                        // cache
                        wbcache.Add(wsname, dists)
                ) (dag.getWorksheetNames())

        type BaseCOFFeature() =
            inherit BaseFeature()

        type ShallowInputVectorMixedCOFRefUnnormSSNorm() =
            inherit BaseCOFFeature()
            let cache = new Dictionary<WorkbookName,Dictionary<WorksheetName,DistDict>>()
            static let instance = new ShallowInputVectorMixedCOFRefUnnormSSNorm()
            member self.BuildDistDict(dag: DAG) : Dictionary<WorksheetName,DistDict> =
                MutateCache cache dag false true
                cache.[dag.getWorkbookName()]
            static member normalizeRefSpace: bool = false
            static member normalizeSSSpace: bool = true
            static member Instance = instance
            static member run(cell: AST.Address)(dag: DAG) : double =
                let normRef = false
                let normSS = true
                
                let k = COFk cell dag normRef normSS
                let dd = (ShallowInputVectorMixedCOFRefUnnormSSNorm.Instance.BuildDistDict dag).[cell.WorksheetName]
                let vcell = SingleSquareVector cell dag normRef normSS
                let neighbors = DistDictToSVHashSet dd
                COF vcell k neighbors dd

            static member capability : string*Capability =
                (typeof<ShallowInputVectorMixedCOFRefUnnormSSNorm>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowInputVectorMixedCOFRefUnnormSSNorm.run } )