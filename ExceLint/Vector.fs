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
        type public C = int    // i.e., a constant (1 if reference is a constant, 0 otherwise) 

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
        type public RelativeVectorAndConstant = (X*Y*Z*C)
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
        type public RichVector =
        | MixedFQVector of Coordinates*MixedVector
        | AbsoluteFQVector of Coordinates*Coordinates
            override self.ToString() : string =
                match self with
                | MixedFQVector(tail,head) -> tail.ToString() + " -> " + head.ToString()
                | AbsoluteFQVector(tail,head) -> tail.ToString() + " -> " + head.ToString()

        type private VectorMaker = AST.Address -> AST.Address -> RichVector
        type private Rebaser = RichVector -> DAG -> bool -> RelativeVector

        let private fullPath(addr: AST.Address) : string*string*string =
            // portably create full path from components
            (addr.Path, addr.WorkbookName, addr.WorksheetName)

        let private vector(tail: AST.Address)(head: AST.Address)(mixed: bool) : RichVector =
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
        let private relativeToTail(absVect: RichVector)(dag: DAG)(offSheetInsensitive: bool) : RelativeVector =
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
        let private relativeToOrigin(absVect: RichVector)(dag: DAG)(offSheetInsensitive: bool) : RelativeVector =
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

        let private RVSum(v1: RelativeVector)(v2: RelativeVector) : RelativeVector =
            let (x1,y1,z1) = v1
            let (x2,y2,z2) = v2
            (x1 + x2, y1 + y2, z1 + z2)

        let private L2NormRVSum(vs: RelativeVector[]) : double =
            vs |> Array.map L2NormRV |> Array.sum

        let private Resultant(vs: RelativeVector[]) : RelativeVector =
            vs |> Array.fold (fun (acc: RelativeVector)(v: RelativeVector) -> RVSum acc v) (0,0,0)

        let private SquareMatrix(origin: X*Y)(vs: RelativeVector[]) : X*Y*X*Y =
            let (x,y) = origin
            let xyoff = vs |> Array.fold (fun (xacc: X, yacc: Y)(x': X,y': Y,z': Z) -> xacc + x', yacc + y') (0,0)
            (fst xyoff, snd xyoff, x, y)

        let transitiveInputVectors(fCell: AST.Address)(dag : DAG)(depth: int option)(vector_f: VectorMaker) : RichVector[] =
            let rec tfVect(tailO: AST.Address option)(head: AST.Address)(depth: int option) : RichVector list =
                let vlist = match tailO with
                            | Some tail -> [vector_f tail head]
                            | None -> []

                match depth with
                | Some(0) -> vlist
                | Some(d) -> tfVect_b head (Some(d-1)) vlist
                | None -> tfVect_b head None vlist

            and tfVect_b(tail: AST.Address)(nextDepth: int option)(vlist: RichVector list) : RichVector list =
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

        let inputVectors(fCell: AST.Address)(dag : DAG)(vector_f: VectorMaker) : RichVector[] =
            transitiveInputVectors fCell dag (Some 1) vector_f

        let transitiveOutputVectors(dCell: AST.Address)(dag : DAG)(depth: int option)(vector_f: VectorMaker) : RichVector[] =
            let rec tdVect(sourceO: AST.Address option)(sink: AST.Address)(depth: int option) : RichVector list =
                let vlist = match sourceO with
                            | Some source -> [vector_f sink source]
                            | None -> []

                match depth with
                | Some(0) -> vlist
                | Some(d) -> tdVect_b sink (Some(d-1)) vlist
                | None -> tdVect_b sink None vlist

            and tdVect_b(sink: AST.Address)(nextDepth: int option)(vlist: RichVector list) : RichVector list =
                    // find all of the formulas that use sink
                    let outAddrs = dag.getFormulasThatRefCell sink
                                    |> Array.toList
                    let outAddrs2 = Array.map (dag.getFormulasThatRefVector) (dag.getVectorsThatRefCell sink)
                                    |> Array.concat |> Array.toList
                    let allFrm = outAddrs @ outAddrs2 |> List.distinct

                    // recursively call this function
                    vlist @ (List.map (fun sink' -> tdVect (Some sink) sink' nextDepth) allFrm |> List.concat)

            tdVect None dCell depth |> List.toArray

        let outputVectors(dCell: AST.Address)(dag : DAG)(vector_f: VectorMaker) : RichVector[] =
            transitiveOutputVectors dCell dag (Some 1) vector_f

        let getVectors(cell: AST.Address)(dag: DAG)(vector_f: VectorMaker)(transitive: bool)(isForm: bool) : RichVector[] =
            let depth = if transitive then None else (Some 1)
            let vectors =
                if isForm then
                    transitiveInputVectors
                else
                    transitiveOutputVectors
            let output = vectors cell dag depth vector_f
            output

        // find normalization function that shifts and scales column values appropriately
        let private colNormFunc(column: double[]) : double -> double =
            let min = Array.min column
            let max = Array.max column
            if max = min then
                fun x -> 0.5
            else
                fun x -> (x - min) / (max - min)

        let private idf(x: double[]) : double -> double = fun x -> x

        let private zeroOneNormalization(worksheet: SquareVector[])(normalizeRefSpace: bool)(normalizeSSSpace: bool) : SquareVector[] =
            assert (worksheet.Length <> 0)

            // get normalization functions for each column
            let dx_norm_f =  Array.map (fun (v: SquareVector) -> v.dx) worksheet
                             |> fun xs -> if normalizeRefSpace then colNormFunc xs else idf xs
            let dy_norm_f = Array.map (fun (v: SquareVector) -> v.dy) worksheet
                            |> fun xs -> if normalizeRefSpace then colNormFunc xs else idf xs
            let x_norm_f = Array.map (fun (v: SquareVector) -> v.x) worksheet
                           |> fun xs -> if normalizeSSSpace then colNormFunc xs else idf xs
            let y_norm_f = Array.map (fun (v: SquareVector) -> v.y) worksheet
                           |> fun xs -> if normalizeSSSpace then colNormFunc xs else idf xs

            // normalize and reutrn new vector array
            Array.map (fun (v: SquareVector) ->
                let dx = dx_norm_f v.dx
                let dy = dy_norm_f v.dy
                let x = x_norm_f v.x
                let y = y_norm_f v.y
                SquareVector(dx, dy, x, y)
            ) worksheet

        let SquareVectorForCell(cell: AST.Address)(dag: DAG)(vector_f: VectorMaker) : SquareVector =
            let vs = getVectors cell dag vector_f (*transitive*) false (*isForm*) true
            let rvs = Array.map (fun rv -> relativeToTail rv dag (*isOffSheetInsensitive*) true) vs
            let sm = SquareMatrix (cell.X, cell.Y) rvs
            let (dx,dy,x,y) = sm
            SquareVector(float dx,float dy,float x,float y)

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

//        let AllSquareVectors(fs: AST.Address[])(dag: DAG)(normalizeRefSpace: bool)(normalizeSSSpace: bool) : Dictionary<AST.Address,SquareVector> =
//            let mats = Array.map (fun f -> SquareMatrixForCell f dag) fs
//
//            let sdx_vect = if normalizeRefSpace then normalizeColumn (column 0 mats) else column 0 mats
//            let sdy_vect = if normalizeRefSpace then normalizeColumn (column 1 mats) else column 1 mats
//            let x_vect = if normalizeSSSpace then normalizeColumn (column 2 mats) else column 2 mats
//            let y_vect = if normalizeSSSpace then normalizeColumn (column 3 mats) else column 3 mats
//
//            let sqvect = combine([| sdx_vect; sdy_vect; x_vect; y_vect |])
//
//            Array.mapi (fun i f square-> f,sqvect.[i]) fs |> adict

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

        let edges(G: SquareVector[]) : Edge[] =
            Array.map (fun i -> Array.map (fun j -> i,j) G) G |> Array.concat

        let pairwiseDistances(E: Edge[]) : DistDict =
            let d = new DistDict()
            for e in E do
                d.Add(e, dist e)
            d

        let Nk(p: SquareVector)(k : int)(G: HashSet<SquareVector>)(DD: DistDict) : HashSet<SquareVector> =
            let DDarr = DD |> Seq.toArray

            let k' = if G.Count - 1 < k then G.Count - 1 else k

            let subgraph = DDarr |>
                           Array.filter (fun (kvp: KeyValuePair<Edge,double>) ->
                                let p' = fst kvp.Key
                                let o = snd kvp.Key
                                let b1 = p = p'        // the starting edge must be p
                                let b2 = p <> o        // p itself must not be considered a neighbor
                                let b3 = G.Contains(o) // neighbor must be in the subgraph G
                                b1 && b2 && b3
                            )
            let subgraph_sorted = subgraph |> Array.sortBy (fun (kvp: KeyValuePair<Edge,double>) -> kvp.Value)

            let subgraph_sorted_k = subgraph_sorted |> Array.take k'
            let kn = subgraph_sorted_k |> Array.map (fun (kvp: KeyValuePair<Edge,double>) -> snd kvp.Key)

            assert (Array.length kn <= k')

            new HashSet<SquareVector>(kn)

        let Nk_cells(p: AST.Address)(k: int)(dag: DAG)(normalizeRefSpace: bool)(normalizeSSSpace: bool)(BDD: Dictionary<WorksheetName,Dictionary<AST.Address,SquareVector>>)(DD: DistDict) : HashSet<AST.Address> =
            let p' = BDD.[p.WorksheetName].[p]
            let Gmap = Array.map (fun (c: AST.Address) ->
                           BDD.[c.WorksheetName].[c], c
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
                    let edges_ranked = Seq.sortBy (fun edge ->
                                          try
                                            let d = DD.[edge]
                                            d
                                          with
                                          | e ->
                                            let m = e.Message
                                            raise e
                                       ) edges |> Seq.toList
                
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

        let L2NormSumMaker(cell: AST.Address)(dag: DAG)(isMixed: bool)(isTransitive: bool)(isFormula: bool)(isOffSheetInsensitive: bool)(rebase_f: Rebaser) =
            let vector_f =
                (fun (source: AST.Address)(sink: AST.Address) -> vector source sink isMixed)
            getVectors cell dag vector_f isTransitive isFormula
            |> Array.map (fun v -> rebase_f v dag isOffSheetInsensitive)
            |> L2NormRVSum
            |> Num

        let CountableVectorMaker(cell: AST.Address)(dag: DAG)(isMixed: bool)(isTransitive: bool)(isFormula: bool)(isOffSheetInsensitive: bool)(rebase_f: Rebaser) =
            let vector_f =
                    (fun (source: AST.Address)(sink: AST.Address) -> vector source sink isMixed)
            getVectors cell dag vector_f isTransitive isFormula
            |> Array.map (fun v -> rebase_f v dag isOffSheetInsensitive)
            |> Resultant
            |> (fun (x,y,z) -> Vector(double x, double y, double z))

        type DeepInputVectorRelativeL2NormSum() = 
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag : DAG) : Countable = 
                let isMixed = false
                let isTransitive = true
                let isFormula = true
                let isOffSheetInsensitive = true
                let rebase_f = relativeToTail
                L2NormSumMaker cell dag isMixed isTransitive isFormula isOffSheetInsensitive rebase_f

            static member capability : string*Capability =
                (typeof<DeepInputVectorRelativeL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = DeepInputVectorRelativeL2NormSum.run } )

        type DeepOutputVectorRelativeL2NormSum() = 
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag : DAG) : Countable = 
                let isMixed = false
                let isTransitive = true
                let isFormula = false
                let isOffSheetInsensitive = true
                let rebase_f = relativeToTail
                L2NormSumMaker cell dag isMixed isTransitive isFormula isOffSheetInsensitive rebase_f

            static member capability : string*Capability =
                (typeof<DeepOutputVectorRelativeL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = DeepOutputVectorRelativeL2NormSum.run } )

        type DeepInputVectorAbsoluteL2NormSum() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : Countable =
                let isMixed = false
                let isTransitive = true
                let isFormula = true
                let isOffSheetInsensitive = true
                let rebase_f = relativeToOrigin
                L2NormSumMaker cell dag isMixed isTransitive isFormula isOffSheetInsensitive rebase_f
            static member capability : string*Capability =
                (typeof<DeepInputVectorAbsoluteL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = DeepInputVectorAbsoluteL2NormSum.run } )

        type DeepOutputVectorAbsoluteL2NormSum() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : Countable =
                let isMixed = false
                let isTransitive = true
                let isFormula = false
                let isOffSheetInsensitive = true
                let rebase_f = relativeToOrigin
                L2NormSumMaker cell dag isMixed isTransitive isFormula isOffSheetInsensitive rebase_f
            static member capability : string*Capability =
                (typeof<DeepOutputVectorAbsoluteL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = DeepOutputVectorAbsoluteL2NormSum.run } )

        type ShallowInputVectorRelativeL2NormSum() = 
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag : DAG) : Countable = 
                let isMixed = false
                let isTransitive = false
                let isFormula = true
                let isOffSheetInsensitive = true
                let rebase_f = relativeToTail
                L2NormSumMaker cell dag isMixed isTransitive isFormula isOffSheetInsensitive rebase_f
            static member capability : string*Capability =
                (typeof<ShallowInputVectorRelativeL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowInputVectorRelativeL2NormSum.run } )

        type ShallowOutputVectorRelativeL2NormSum() = 
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag : DAG) : Countable = 
                let isMixed = false
                let isTransitive = false
                let isFormula = false
                let isOffSheetInsensitive = true
                let rebase_f = relativeToTail
                L2NormSumMaker cell dag isMixed isTransitive isFormula isOffSheetInsensitive rebase_f
            static member capability : string*Capability =
                (typeof<ShallowOutputVectorRelativeL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowOutputVectorRelativeL2NormSum.run } )

        type ShallowInputVectorAbsoluteL2NormSum() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : Countable =
                let isMixed = false
                let isTransitive = false
                let isFormula = true
                let isOffSheetInsensitive = true
                let rebase_f = relativeToOrigin
                L2NormSumMaker cell dag isMixed isTransitive isFormula isOffSheetInsensitive rebase_f
            static member capability : string*Capability =
                (typeof<ShallowInputVectorAbsoluteL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowInputVectorAbsoluteL2NormSum.run } )

        type ShallowOutputVectorAbsoluteL2NormSum() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : Countable =
                let isMixed = false
                let isTransitive = false
                let isFormula = false
                let isOffSheetInsensitive = true
                let rebase_f = relativeToOrigin
                L2NormSumMaker cell dag isMixed isTransitive isFormula isOffSheetInsensitive rebase_f
            static member capability : string*Capability =
                (typeof<ShallowOutputVectorAbsoluteL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowOutputVectorAbsoluteL2NormSum.run } )

        type ShallowInputVectorMixedL2NormSum() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : Countable =
                let isMixed = true
                let isTransitive = false
                let isFormula = true
                let isOffSheetInsensitive = true
                let rebase_f = relativeToTail
                L2NormSumMaker cell dag isMixed isTransitive isFormula isOffSheetInsensitive rebase_f
            static member capability : string*Capability =
                (typeof<ShallowInputVectorMixedL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowInputVectorMixedL2NormSum.run } )

        type ShallowInputVectorMixedResultant() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : Countable =
                let isMixed = true
                let isTransitive = false
                let isFormula = true
                let isOffSheetInsensitive = true
                let rebase_f = relativeToTail
                CountableVectorMaker cell dag isMixed isTransitive isFormula isOffSheetInsensitive rebase_f
            static member capability : string*Capability =
                (typeof<ShallowInputVectorMixedResultant>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowInputVectorMixedResultant.run } )

//        type ShallowInputVectorMixedResultantWithConstant() =
//            inherit BaseFeature()
//            static member run(cell: AST.Address)(dag: DAG) : Countable =
//                let (x,y,z,c) = 

        type ShallowOutputVectorMixedL2NormSum() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : Countable =
                let isMixed = true
                let isTransitive = false
                let isFormula = false
                let isOffSheetInsensitive = true
                let rebase_f = relativeToTail
                L2NormSumMaker cell dag isMixed isTransitive isFormula isOffSheetInsensitive rebase_f
            static member capability : string*Capability =
                (typeof<ShallowOutputVectorMixedL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowOutputVectorMixedL2NormSum.run } )

        type DeepInputVectorMixedL2NormSum() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : Countable =
                let isMixed = true
                let isTransitive = true
                let isFormula = true
                let isOffSheetInsensitive = true
                let rebase_f = relativeToTail
                L2NormSumMaker cell dag isMixed isTransitive isFormula isOffSheetInsensitive rebase_f
            static member capability : string*Capability =
                (typeof<DeepInputVectorMixedL2NormSum>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = DeepInputVectorMixedL2NormSum.run } )

        type DeepOutputVectorMixedL2NormSum() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : Countable =
                let isMixed = true
                let isTransitive = true
                let isFormula = false
                let isOffSheetInsensitive = true
                let rebase_f = relativeToTail
                L2NormSumMaker cell dag isMixed isTransitive isFormula isOffSheetInsensitive rebase_f
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

        let private MutateCache(bigcache: Dictionary<WorkbookName,Dictionary<WorksheetName,Dictionary<AST.Address,SquareVector>>>)(cache: Dictionary<WorkbookName,Dictionary<WorksheetName,DistDict>>)(dag: DAG)(normalizeRefSpace: bool)(normalizeSSSpace: bool)(vector_f: VectorMaker) : unit =
                let wbname = dag.getWorkbookName()

                // get workbook-level cache
                let wbcache = if not (cache.ContainsKey wbname) then
                                  cache.Add(wbname, new Dictionary<WorksheetName,DistDict>())
                                  cache.[wbname]
                              else
                                  cache.[wbname]

                let bwbcache = if not (bigcache.ContainsKey wbname) then
                                  bigcache.Add(wbname, new Dictionary<WorksheetName,Dictionary<AST.Address,SquareVector>>())
                                  bigcache.[wbname]
                               else
                                  bigcache.[wbname]
                
                // loop through all worksheet names
                Array.iter (fun wsname ->
                    if not (wbcache.ContainsKey(wsname)) then
                        // get formulas
                        let fs = dag.getAllFormulaAddrs() |> Array.filter (fun f -> f.WorksheetName = wsname) 

                        // don't do anything if the sheet has no formulas
                        if fs.Length <> 0 then
                            // get square vectors for every formula
                            let vmap = Array.map (fun cell -> cell,SquareVectorForCell cell dag vector_f) fs |> adict
                            let rev_vmap = Seq.map (fun (kvp: KeyValuePair<AST.Address,SquareVector>) -> kvp.Value, kvp.Key) vmap |> adict
                            let vectors = vmap.Values |> Seq.toArray

                            // normalize both data structures for this sheet
                            let vectors' = zeroOneNormalization vectors normalizeRefSpace normalizeSSSpace
                            let vmap' = Array.mapi (fun i v -> rev_vmap.[v],vectors'.[i]) vectors |> adict

                            // compute distances
                            let dists = pairwiseDistances (edges vectors')
                            // cache
                            wbcache.Add(wsname, dists)
                            bwbcache.Add(wsname, vmap')
                ) (dag.getWorksheetNames())

        type BaseCOFFeature() =
            inherit BaseFeature()

        type ShallowInputVectorMixedCOFNoAspect() =
            inherit BaseCOFFeature()
            let bigcache = new Dictionary<WorkbookName,Dictionary<WorksheetName,Dictionary<AST.Address,SquareVector>>>()
            let cache = new Dictionary<WorkbookName,Dictionary<WorksheetName,DistDict>>()
            static let instance = new ShallowInputVectorMixedCOFNoAspect()
            member self.BuildDistDict(dag: DAG) : Dictionary<WorksheetName,Dictionary<AST.Address,SquareVector>>*Dictionary<WorksheetName,DistDict> =
                let vector_f =
                    (fun (source: AST.Address)(sink: AST.Address) -> vector source sink true)
                MutateCache bigcache cache dag ShallowInputVectorMixedCOFNoAspect.normalizeRefSpace ShallowInputVectorMixedCOFNoAspect.normalizeSSSpace vector_f
                let c = cache.[dag.getWorkbookName()]
                let bc = bigcache.[dag.getWorkbookName()]
                bc,c
            static member normalizeRefSpace: bool = true
            static member normalizeSSSpace: bool = true
            static member Instance = instance
            static member run(cell: AST.Address)(dag: DAG) : Countable =
                let normRef = ShallowInputVectorMixedCOFNoAspect.normalizeRefSpace
                let normSS = ShallowInputVectorMixedCOFNoAspect.normalizeSSSpace
                
                let k = COFk cell dag normRef normSS
                let (BDD,DD) = ShallowInputVectorMixedCOFNoAspect.Instance.BuildDistDict dag
                let dd = DD.[cell.WorksheetName]
                let bdd = BDD.[cell.WorksheetName]

                let vcell = bdd.[cell]

                let neighbors = DistDictToSVHashSet dd
                COF vcell k neighbors dd
                |> Num

            static member capability : string*Capability =
                (typeof<ShallowInputVectorMixedCOFNoAspect>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowInputVectorMixedCOFNoAspect.run } )

        type ShallowInputVectorMixedCOFAspect() =
            inherit BaseCOFFeature()
            let bigcache = new Dictionary<WorkbookName,Dictionary<WorksheetName,Dictionary<AST.Address,SquareVector>>>()
            let cache = new Dictionary<WorkbookName,Dictionary<WorksheetName,DistDict>>()
            static let instance = new ShallowInputVectorMixedCOFAspect()
            member self.BuildDistDict(dag: DAG) : Dictionary<WorksheetName,Dictionary<AST.Address,SquareVector>>*Dictionary<WorksheetName,DistDict> =
                let vector_f =
                    (fun (source: AST.Address)(sink: AST.Address) -> vector source sink true)
                MutateCache bigcache cache dag ShallowInputVectorMixedCOFAspect.normalizeRefSpace ShallowInputVectorMixedCOFAspect.normalizeSSSpace vector_f
                let c = cache.[dag.getWorkbookName()]
                let bc = bigcache.[dag.getWorkbookName()]
                bc,c
            static member normalizeRefSpace: bool = true
            static member normalizeSSSpace: bool = true
            static member Instance = instance
            static member run(cell: AST.Address)(dag: DAG) : Countable =
                let normRef = ShallowInputVectorMixedCOFAspect.normalizeRefSpace
                let normSS = ShallowInputVectorMixedCOFAspect.normalizeSSSpace
                
                let k = COFk cell dag normRef normSS
                let (BDD,DD) = ShallowInputVectorMixedCOFAspect.Instance.BuildDistDict dag
                let dd = DD.[cell.WorksheetName]
                let bdd = BDD.[cell.WorksheetName]

                let vcell = bdd.[cell]

                let neighbors = DistDictToSVHashSet dd
                COF vcell k neighbors dd
                |> Num

            static member capability : string*Capability =
                (typeof<ShallowInputVectorMixedCOFAspect>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowInputVectorMixedCOFAspect.run } )