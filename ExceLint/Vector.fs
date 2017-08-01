namespace ExceLint
    open Depends
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
        type public C = double // i.e., a constant value

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
        type public RelativeVector =
        | NoConstant of X*Y*Z
        | NoConstantWithLoc of X*Y*Z*X*Y*Z
        | Constant of X*Y*Z*C
        | ConstantWithLoc of X*Y*Z*X*Y*Z*C
            member self.Zero =
                match self with
                | Constant(_,_,_,_) -> Constant(0,0,0,0.0)
                | ConstantWithLoc(x,y,z,_,_,_,_) -> ConstantWithLoc(x,y,z,0,0,0,0.0)
                | NoConstant(_,_,_) -> NoConstant(0,0,0)
                | NoConstantWithLoc(x,y,z,_,_,_) -> NoConstantWithLoc(x,y,z,0,0,0)
        type public MixedVector = (VectorComponent*VectorComponent*Path)
        type public MixedVectorWithConstant = (VectorComponent*VectorComponent*Path*C)
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
        | MixedFQVectorWithConstant of Coordinates*MixedVectorWithConstant
        | MixedFQVector of Coordinates*MixedVector
        | AbsoluteFQVector of Coordinates*Coordinates
            override self.ToString() : string =
                match self with
                | MixedFQVectorWithConstant(tail,head) -> tail.ToString() + " -> " + head.ToString()
                | MixedFQVector(tail,head) -> tail.ToString() + " -> " + head.ToString()
                | AbsoluteFQVector(tail,head) -> tail.ToString() + " -> " + head.ToString()
        type private KeepConstantValue =
        | Yes
        | No

        type private VectorMaker = AST.Address -> AST.Address -> RichVector
        type private ConstantVectorMaker = AST.Address -> AST.Expression -> RichVector list
        type private Rebaser = RichVector -> DAG -> bool -> bool -> RelativeVector

        let private fullPath(addr: AST.Address) : string*string*string =
            // portably create full path from components
            (addr.Path, addr.WorkbookName, addr.WorksheetName)

        let private vector(tail: AST.Address)(head: AST.Address)(mixed: bool)(include_constant: bool) : RichVector =
            let tailXYP = (tail.X, tail.Y, fullPath tail)
            if mixed then
                let X = match head.XMode with
                        | AST.AddressMode.Absolute -> Abs(head.X)
                        | AST.AddressMode.Relative -> Rel(head.X)
                let Y = match head.YMode with
                        | AST.AddressMode.Absolute -> Abs(head.Y)
                        | AST.AddressMode.Relative -> Rel(head.Y)
                
                if include_constant then
                    let headXYP = (X, Y, fullPath head, 0.0)
                    MixedFQVectorWithConstant(tailXYP, headXYP)
                else
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
        let private relativeToTail(absVect: RichVector)(dag: DAG)(offSheetInsensitive: bool)(includeLoc: bool) : RelativeVector =
            match absVect with
            | AbsoluteFQVector(tail,head) ->
                let (x1,y1,p1) = tail
                let (x2,y2,p2) = head
                if offSheetInsensitive && p1 <> p2 then
                    if includeLoc then
                        NoConstantWithLoc(x1, y1, dag.getPathClosureIndex(p1), 0, 0, dag.getPathClosureIndex(p2))
                    else
                        NoConstant(0, 0, dag.getPathClosureIndex(p2))
                else
                    if includeLoc then
                        NoConstantWithLoc(x1, y1, dag.getPathClosureIndex(p1), x2-x1, y2-y1, vectorPathDiff p2 p1)
                    else
                        NoConstant(x2-x1, y2-y1, vectorPathDiff p2 p1)
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
                    if includeLoc then
                        NoConstantWithLoc(x1, y1, dag.getPathClosureIndex(p1), 0, 0, dag.getPathClosureIndex(p2))
                    else
                        NoConstant(0, 0, dag.getPathClosureIndex(p2))
                else
                    if includeLoc then
                        NoConstantWithLoc(x1, y1, dag.getPathClosureIndex(p1), x', y', vectorPathDiff p2 p1)
                    else
                        NoConstant(x', y', vectorPathDiff p2 p1)
            | MixedFQVectorWithConstant(tail,head) ->
                let (x1,y1,p1) = tail
                let (x2,y2,p2,c) = head

                let x' = match x2 with
                            | Rel(x) -> x - x1
                            | Abs(x) -> x
                let y' = match y2 with
                            | Rel(y) -> y - y1
                            | Abs(y) -> y
                if offSheetInsensitive && p1 <> p2 then
                    if includeLoc then
                        ConstantWithLoc(x1, y1, dag.getPathClosureIndex(p1), 0, 0, dag.getPathClosureIndex(p2), c)
                    else
                        Constant(0, 0, dag.getPathClosureIndex(p2), c)
                else
                    if includeLoc then
                        ConstantWithLoc(x1, y1, dag.getPathClosureIndex(p1), x', y', vectorPathDiff p2 p1, c)
                    else
                        Constant(x', y', vectorPathDiff p2 p1, c)

        // represent the position of the the head of the vector relative to the origin, (0,0,0)
        let private relativeToOrigin(absVect: RichVector)(dag: DAG)(offSheetInsensitive: bool)(includeLoc: bool) : RelativeVector =
            match absVect with
            | AbsoluteFQVector(tail,head) ->
                let (tx,ty,tp) = tail
                let (x,y,p) = head
                if offSheetInsensitive && tp <> p then
                    if includeLoc then
                        NoConstantWithLoc(tx, ty, dag.getPathClosureIndex(tp), 0, 0, dag.getPathClosureIndex(p))
                    else
                        NoConstant(0, 0, dag.getPathClosureIndex(p))
                else
                    if includeLoc then
                        NoConstantWithLoc(tx, ty, dag.getPathClosureIndex(tp), x, y, pathDiff p dag)
                    else
                        NoConstant(x, y, pathDiff p dag)
            | MixedFQVector(tail,head) ->
                let (tx,ty,tp) = tail
                let (x,y,p) = head
                let x' = match x with | Abs(xa) -> xa | Rel(xr) -> xr
                let y' = match y with | Abs(ya) -> ya | Rel(yr) -> yr
                if offSheetInsensitive && tp <> p then
                    if includeLoc then
                        NoConstantWithLoc(tx, ty, dag.getPathClosureIndex(tp), 0, 0, dag.getPathClosureIndex(p))
                    else
                        NoConstant(0, 0, dag.getPathClosureIndex(p))
                else
                    if includeLoc then
                        NoConstantWithLoc(tx, ty, dag.getPathClosureIndex(tp), x', y', pathDiff p dag)
                    else
                        NoConstant(x', y', pathDiff p dag)
            | MixedFQVectorWithConstant(tail,head) ->
                let (tx,ty,tp) = tail
                let (x,y,p,c) = head
                let x' = match x with | Abs(xa) -> xa | Rel(xr) -> xr
                let y' = match y with | Abs(ya) -> ya | Rel(yr) -> yr
                if offSheetInsensitive && tp <> p then
                    if includeLoc then
                        ConstantWithLoc(tx, ty, dag.getPathClosureIndex(p), 0, 0, dag.getPathClosureIndex(p), c)
                    else
                        Constant(0, 0, dag.getPathClosureIndex(p), c)
                else
                    if includeLoc then
                        ConstantWithLoc(tx, ty, dag.getPathClosureIndex(p), x', y', pathDiff p dag, c)
                    else
                        Constant(x', y', pathDiff p dag, c)

        let private L2Norm(X: double[]) : double =
            Math.Sqrt(
                Array.sumBy (fun x -> Math.Pow(x, 2.0)) X
            )

        let private relativeVectorToRealVectorArr(v: RelativeVector) : double[] =
            match v with
            | NoConstant(x,y,z) ->
                [|
                    System.Convert.ToDouble(x);
                    System.Convert.ToDouble(y);
                    System.Convert.ToDouble(z);
                |]
            | Constant(x,y,z,c) ->
                [|
                    System.Convert.ToDouble(x);
                    System.Convert.ToDouble(y);
                    System.Convert.ToDouble(z);
                    c;
                |]
            | NoConstantWithLoc(x,y,z,dx,dy,dz) ->
                [|
                    System.Convert.ToDouble(x);
                    System.Convert.ToDouble(y);
                    System.Convert.ToDouble(z);
                    System.Convert.ToDouble(dx);
                    System.Convert.ToDouble(dy);
                    System.Convert.ToDouble(dz);
                |]
            | ConstantWithLoc(x,y,z,dx,dy,dz,dc) ->
                [|
                    System.Convert.ToDouble(x);
                    System.Convert.ToDouble(y);
                    System.Convert.ToDouble(z);
                    System.Convert.ToDouble(dx);
                    System.Convert.ToDouble(dy);
                    System.Convert.ToDouble(dz);
                    System.Convert.ToDouble(dc);
                |]

        let private L2NormRV(v: RelativeVector) : double =
            L2Norm(relativeVectorToRealVectorArr(v))

        let private RVSum(v1: RelativeVector)(v2: RelativeVector) : RelativeVector =
            match v1,v2 with
            | NoConstant(x1,y1,z1), NoConstant(x2,y2,z2) ->
                NoConstant(x1 + x2, y1 + y2, z1 + z2)
            | NoConstantWithLoc(x1,y1,z1,dx1,dy1,dz1), NoConstantWithLoc(x2,y2,z2,dx2,dy2,dz2) ->
                assert (x1 = x2 && y1 = y2 && z1 = z2)
                // we don't add reference sources, just reference destinations
                NoConstantWithLoc(x1, y1, z1, dx1 + dx2, dy1 + dy2, dz1 + dz2)
            | Constant(x1,y1,z1,c1), Constant(x2,y2,z2,c2) ->
                Constant(x1 + x2, y1 + y2, z1 + z2, c1 + c2)
            | ConstantWithLoc(x1,y1,z1,dx1,dy1,dz1,dc1), ConstantWithLoc(x2,y2,z2,dx2,dy2,dz2,dc2) ->
                assert (x1 = x2 && y1 = y2 && z1 = z2)
                // we don't add reference sources, just reference destinations
                ConstantWithLoc(x1, y1, z1, dx1 + dx2, dy1 + dy2, dz1 + dz2, dc1 + dc2)
            | _ -> failwith "Cannot sum RelativeVectors of different subtypes."

        let private L2NormRVSum(vs: RelativeVector[]) : double =
            vs |> Array.map L2NormRV |> Array.sum

        let private Resultant(vs: RelativeVector[]) : RelativeVector =
            vs |>
            Array.fold (fun (acc: RelativeVector option)(v: RelativeVector) ->
                match acc with
                | None -> Some (RVSum v.Zero v)
                | Some a -> Some (RVSum a v)
            ) None |>
            (fun rvopt ->
                match rvopt with
                | Some rv -> rv
                | None -> failwith "Empty resultant!"
            )

        let private SquareMatrix(origin: X*Y)(vs: RelativeVector[]) : X*Y*X*Y =
            let (x,y) = origin
            let xyoff = vs |>
                        Array.fold (fun (xacc: X, yacc: Y)(rv: RelativeVector) ->
                            let (x',y') =
                                match rv with
                                | Constant(x,y,_,_) -> x,y
                                | NoConstant(x,y,_) -> x,y
                                | _ -> failwith "not supported"
                            xacc + x', yacc + y'
                        ) (0,0)
            (fst xyoff, snd xyoff, x, y)

        let transitiveInputVectors(fCell: AST.Address)(dag : DAG)(depth: int option)(vector_f: VectorMaker)(cvector_f: ConstantVectorMaker) : RichVector[] =
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
                    try
                        // parse again, because the DAG treats repeated
                        // references to the same cell but with different
                        // address modes as the same reference; they are not.

                        // Sometimes idiots denote comments with '='.
                        let fexpr = Parcel.parseFormulaAtAddress tail (dag.getFormulaAtAddress tail)

                        // find all of the inputs for source
                        let heads_single = Parcel.addrReferencesFromExpr fexpr |> List.ofArray
                        let heads_vector = Parcel.rangeReferencesFromExpr fexpr |>
                                                List.ofArray |>
                                                List.map (fun rng -> rng.Addresses() |> Array.toList) |>
                                                List.concat

                        // find all constant inputs for source
                        let cvects = cvector_f tail fexpr

                        let heads = heads_single @ heads_vector
                        // recursively call this function
                        vlist @ cvects @ (List.map (fun head -> tfVect (Some tail) head nextDepth) heads |> List.concat)
                    with
                    | e -> vlist  // I guess we give up
                else
                    let value = dag.readCOMValueAtAddress(tail)
                    let mutable num = 0.0
                    num <- if Double.TryParse(value, &num) then
                               // a constant, i.e., references one thing
                               1.0
                           else if String.IsNullOrWhiteSpace(value) then
                               // it's blank, i.e., references nothing
                               0.0
                           else
                               // it's a string
                               -1.0  // pretty arbitrary... maybe we should have a "blank" dimension
                    let env = AST.Env(tail.Path, tail.WorkbookName, tail.WorksheetName)
                    let expr = AST.ReferenceExpr (AST.ReferenceConstant(env, num))
                    let dv = cvector_f tail expr
                    dv @ vlist

    
            tfVect None fCell depth |> List.toArray

        let transitiveOutputVectors(dCell: AST.Address)(dag : DAG)(depth: int option)(vector_f: VectorMaker)(cvector_f: ConstantVectorMaker) : RichVector[] =
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

        let private makeVector(isMixed: bool)(includeConstant: bool): VectorMaker =
            (fun (source: AST.Address)(sink: AST.Address) ->
                vector source sink isMixed includeConstant
            )

        let private nopConstantVector : ConstantVectorMaker =
            (fun (a: AST.Address)(e: AST.Expression) -> [])

        let private makeConstantVectorsFromConstants(k: KeepConstantValue) : ConstantVectorMaker =
            (fun (tail: AST.Address)(e: AST.Expression) ->
                // convert into RichVector form
                let tailXYP = (tail.X, tail.Y, fullPath tail)

                // the path for the head is the same as the path for the tail for constants
                let path = fullPath tail
                let constants = Parcel.constantsFromExpr e

                // the vectorcomponents for constants are Abs(0)
                let cvc = Abs(0)

                // make vectors
                let cf = match k with
                         | Yes -> (fun (rc: AST.ReferenceConstant) -> rc.Value)
                         | No -> (fun (rc: AST.ReferenceConstant) -> if rc.Value <> 0.0 then 1.0 else 0.0)

                let vs = Array.map (fun (c: AST.ReferenceConstant) ->
                             RichVector.MixedFQVectorWithConstant(tailXYP, (cvc, cvc, path, cf c))
                         ) constants |> Array.toList

                vs
            )

        let getVectors(cell: AST.Address)(dag: DAG)(vector_f: VectorMaker)(cvector_f: ConstantVectorMaker)(transitive: bool)(isForm: bool) : RichVector[] =
            let depth = if transitive then None else (Some 1)
            let vectors =
                if isForm then
                    transitiveInputVectors
                else
                    transitiveOutputVectors
            let output = vectors cell dag depth vector_f cvector_f
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

        let SquareVectorForCell(cell: AST.Address)(dag: DAG)(vector_f: VectorMaker)(cvector_f: ConstantVectorMaker) : SquareVector =
            let vs = getVectors cell dag vector_f cvector_f (*transitive*) false (*isForm*) true
            let rvs = Array.map (fun rv -> relativeToTail rv dag (*isOffSheetInsensitive*) true false) vs
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

        let getRebasedVectors(cell: AST.Address)(dag: DAG)(isMixed: bool)(isTransitive: bool)(isFormula: bool)(isOffSheetInsensitive: bool)(isRelative: bool)(includeConstant: bool) =
            getVectors cell dag (makeVector isMixed includeConstant) nopConstantVector isTransitive isFormula
            |> Array.map (fun v ->
                   (if isRelative then relativeToTail else relativeToOrigin) v dag isOffSheetInsensitive includeConstant
               )

        let L2NormSumMaker(cell: AST.Address)(dag: DAG)(isMixed: bool)(isTransitive: bool)(isFormula: bool)(isOffSheetInsensitive: bool)(rebase_f: Rebaser) =
            getVectors cell dag (makeVector isMixed false) nopConstantVector isTransitive isFormula
            |> Array.map (fun v -> rebase_f v dag isOffSheetInsensitive false)
            |> L2NormRVSum
            |> Num

        let L2NormSumMakerCS(cell: AST.Address)(dag: DAG)(isMixed: bool)(isTransitive: bool)(isFormula: bool)(isOffSheetInsensitive: bool)(rebase_f: Rebaser) =
            getVectors cell dag (makeVector isMixed false) nopConstantVector isTransitive isFormula
            |> Array.map (fun v -> rebase_f v dag isOffSheetInsensitive false)
            |> L2NormRVSum
            |> Num

        let ResultantMaker(cell: AST.Address)(dag: DAG)(isMixed: bool)(includeConstant: bool)(includeLoc: bool)(isTransitive: bool)(isFormula: bool)(isOffSheetInsensitive: bool)(constant_f: ConstantVectorMaker)(rebase_f: Rebaser) : Countable =
            let vs = getVectors cell dag (makeVector isMixed includeConstant) constant_f isTransitive isFormula
            let rebased_vs = vs |> Array.map (fun v -> rebase_f v dag isOffSheetInsensitive includeLoc)
            let resultant = rebased_vs |> Resultant
            let countable =
                resultant
                |> (fun rv ->
                        match rv with
                        | Constant(x,y,z,c) -> CVectorResultant(double x, double y, double z, double c)
                        | NoConstant(x,y,z) -> Vector(double x, double y, double z)
                        | ConstantWithLoc(x,y,z,dx,dy,dz,dc) -> FullCVectorResultant(double x, double y, double z, double dx, double dy, double dz, double dc)
                        | NoConstantWithLoc(x,y,z,dx,dy,dz) -> Countable.SquareVector(double dx, double dy, double dz, double x, double y, double z)
                   )
            countable

        type DeepInputVectorRelativeL2NormSum() = 
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : Countable = 
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
                let includeConstant = false
                let includeLoc = false
                let rebase_f = relativeToTail
                let constant_f = nopConstantVector
                ResultantMaker cell dag isMixed includeConstant includeLoc isTransitive isFormula isOffSheetInsensitive constant_f rebase_f 
            static member capability : string*Capability =
                (typeof<ShallowInputVectorMixedResultant>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowInputVectorMixedResultant.run } )

        type ShallowInputVectorMixedCVectorResultantNotOSI() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : Countable =
                let isMixed = true
                let isTransitive = false
                let isFormula = true
                let isOffSheetInsensitive = false
                let includeConstant = true
                let includeLoc = false
                let rebase_f = relativeToTail
                let constant_f = (makeConstantVectorsFromConstants KeepConstantValue.No)
                ResultantMaker cell dag isMixed includeConstant includeLoc isTransitive isFormula isOffSheetInsensitive constant_f rebase_f 
            static member capability : string*Capability =
                (typeof<ShallowInputVectorMixedCVectorResultantNotOSI>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowInputVectorMixedCVectorResultantNotOSI.run } )

        type ShallowInputVectorMixedFullCVectorResultantNotOSI() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : Countable =
                let isMixed = true
                let isTransitive = false
                let isFormula = true
                let isOffSheetInsensitive = false
                let includeConstant = true
                let includeLoc = true
                let rebase_f = relativeToTail
                let constant_f = (makeConstantVectorsFromConstants KeepConstantValue.No)
                ResultantMaker cell dag isMixed includeConstant includeLoc isTransitive isFormula isOffSheetInsensitive constant_f rebase_f 
            static member capability : string*Capability =
                (typeof<ShallowInputVectorMixedFullCVectorResultantNotOSI>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowInputVectorMixedFullCVectorResultantNotOSI.run } )

        type ShallowInputVectorMixedFullCVectorResultantOSI() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: DAG) : Countable =
                let isMixed = true
                let isTransitive = false
                let isFormula = true
                let isOffSheetInsensitive = true
                let includeConstant = true
                let includeLoc = true
                let rebase_f = relativeToTail
                let constant_f = makeConstantVectorsFromConstants KeepConstantValue.Yes  // note that we want to set constant values explicitly here
                ResultantMaker cell dag isMixed includeConstant includeLoc isTransitive isFormula isOffSheetInsensitive constant_f rebase_f 
            static member capability : string*Capability =
                (typeof<ShallowInputVectorMixedFullCVectorResultantOSI>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowInputVectorMixedFullCVectorResultantOSI.run } )

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

        let private MutateCache(bigcache: Dictionary<WorkbookName,Dictionary<WorksheetName,Dictionary<AST.Address,SquareVector>>>)(cache: Dictionary<WorkbookName,Dictionary<WorksheetName,DistDict>>)(dag: DAG)(normalizeRefSpace: bool)(normalizeSSSpace: bool)(vector_f: VectorMaker)(cvector_f: ConstantVectorMaker) : unit =
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
                            let vmap = Array.map (fun cell -> cell,SquareVectorForCell cell dag vector_f cvector_f) fs |> adict
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
                    (fun (source: AST.Address)(sink: AST.Address) -> vector source sink true false)
                MutateCache bigcache cache dag ShallowInputVectorMixedCOFNoAspect.normalizeRefSpace ShallowInputVectorMixedCOFNoAspect.normalizeSSSpace vector_f nopConstantVector
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
                    (fun (source: AST.Address)(sink: AST.Address) -> vector source sink true false)
                MutateCache bigcache cache dag ShallowInputVectorMixedCOFAspect.normalizeRefSpace ShallowInputVectorMixedCOFAspect.normalizeSSSpace vector_f nopConstantVector
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