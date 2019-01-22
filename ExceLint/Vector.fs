namespace ExceLint
    open FastDependenceAnalysis
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

        type ArityZero() =
            static let idx: Dictionary<string,int> =
                let mutable i = 1
                Grammar.Arity0Names |>
                Array.sort |>
                Array.fold (fun (acc: Dictionary<string,int>)(n: string) ->
                    acc.Add(n, -i)
                    i <- i + 1
                    acc
                ) (new Dictionary<string,int>()) 
            static member isZeroArity n = idx.ContainsKey n
            static member hasIndex n = idx.[n]

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
        type private Rebaser = RichVector -> Graph -> bool -> bool -> RelativeVector

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

        let private originPath(dag: Graph) : Path =
            (dag.Path, dag.Workbook, dag.Worksheet);

        let private vectorPathDiff(p1: Path)(p2: Path) : int =
            if p1 <> p2 then 1 else 0

        // the origin is defined as x = 0, y = 0, z = 0 if first sheet in the workbook else 1
        let private pathDiff(p: Path)(dag: Graph) : int =
            let op = originPath dag
            vectorPathDiff p op

        // represent the position of the head of the vector relative to the tail, (x1,y1,z1)
        // if the reference is off-sheet then optionally ignore X and Y vector components
        let private relativeToTail(absVect: RichVector)(dag: Graph)(offSheetInsensitive: bool)(includeLoc: bool) : RelativeVector =
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
        let private relativeToOrigin(absVect: RichVector)(dag: Graph)(offSheetInsensitive: bool)(includeLoc: bool) : RelativeVector =
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

        let zeroArityXYP(op: string)(tailPath: string*string*string)(cvc: C) : MixedVectorWithConstant option =
            if ArityZero.isZeroArity op then
                Some (VectorComponent.Abs (ArityZero.hasIndex op), VectorComponent.Abs (ArityZero.hasIndex op), tailPath, cvc)
            else
                None

        let refsForArityZeroOps(tail: AST.Address)(ops: string list) : RichVector list =
            if ops.Length = 0 then
                []
            else
                let tailPath = fullPath tail
                let tailXYP = tail.X, tail.Y, tailPath
                let c = 0.0    // no constant
                let heads = ops |> List.map (fun op -> zeroArityXYP op tailPath c) |> List.choose id
                heads |> List.map (fun head -> MixedFQVectorWithConstant(tailXYP, head))

        let transitiveInputVectors(fCell: AST.Address)(dag : Graph)(depth: int option)(vector_f: VectorMaker)(cvector_f: ConstantVectorMaker) : RichVector[] =
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

                        // Get references for zero-arity functions
                        let ops = Parcel.operatorNamesFromExpr fexpr
                        let zvects = refsForArityZeroOps tail ops

                        let heads = heads_single @ heads_vector
                        // recursively call this function
                        vlist @ cvects @ zvects @ (List.map (fun head -> tfVect (Some tail) head nextDepth) heads |> List.concat)
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
                         | No -> (fun (rc: AST.ReferenceConstant) -> if rc.Value = 0.0 then 0.0 else if rc.Value = -1.0 then -1.0 else 1.0)

                let vs = Array.map (fun (c: AST.ReferenceConstant) ->
                             RichVector.MixedFQVectorWithConstant(tailXYP, (cvc, cvc, path, cf c))
                         ) constants |> Array.toList

                vs
            )

        let getVectors(cell: AST.Address)(dag: Graph)(vector_f: VectorMaker)(cvector_f: ConstantVectorMaker)(transitive: bool)(isForm: bool) : RichVector[] =
            let depth = if transitive then None else (Some 1)
            let output = transitiveInputVectors cell dag depth vector_f cvector_f
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

        let SquareVectorForCell(cell: AST.Address)(dag: Graph)(vector_f: VectorMaker)(cvector_f: ConstantVectorMaker) : SquareVector =
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

        let hsDiff(A: HashSet<'a>)(B: HashSet<'a>) : HashSet<'a> =
            let hs = new HashSet<'a>()
            for a in A do
                if not (B.Contains a) then
                    hs.Add a |> ignore
            hs

        let getRebasedVectors(cell: AST.Address)(dag: Graph)(isMixed: bool)(isTransitive: bool)(isFormula: bool)(isOffSheetInsensitive: bool)(isRelative: bool)(includeConstant: bool) =
            getVectors cell dag (makeVector isMixed includeConstant) nopConstantVector isTransitive isFormula
            |> Array.map (fun v ->
                   (if isRelative then relativeToTail else relativeToOrigin) v dag isOffSheetInsensitive includeConstant
               )

        let ResultantMaker(cell: AST.Address)(dag: Graph)(isMixed: bool)(includeConstant: bool)(includeLoc: bool)(isTransitive: bool)(isFormula: bool)(isOffSheetInsensitive: bool)(constant_f: ConstantVectorMaker)(rebase_f: Rebaser) : Countable =
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

        // THIS IS THE IMPORTANT ONE USED IN THE PAPER
        type ShallowInputVectorMixedFullCVectorResultantOSI() =
            inherit BaseFeature()
            static member run(cell: AST.Address)(dag: Graph) : Countable =
                let isMixed = true
                let isTransitive = false
                let isFormula = true
                let isOffSheetInsensitive = true
                let includeConstant = true
                let includeLoc = true
                let keepConstantValues = KeepConstantValue.No
                let rebase_f = relativeToTail
                let constant_f = makeConstantVectorsFromConstants keepConstantValues
                ResultantMaker cell dag isMixed includeConstant includeLoc isTransitive isFormula isOffSheetInsensitive constant_f rebase_f 
            static member capability : string*Capability =
                (typeof<ShallowInputVectorMixedFullCVectorResultantOSI>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = ShallowInputVectorMixedFullCVectorResultantOSI.run } )
            // THIS FUNCTION GETS THE VECTOR SET FOR THE ANALYSIS ABOVE
            static member getPaperVectors(cell: AST.Address)(dag: Graph) : Countable[] =
                let isMixed = true
                let isTransitive = false
                let isFormula = true
                let isOffSheetInsensitive = true
                let includeConstant = true
                let includeLoc = true
                let keepConstantValues = KeepConstantValue.No
                let rebase_f = relativeToTail
                let constant_f = makeConstantVectorsFromConstants keepConstantValues
                let vs = getVectors cell dag (makeVector isMixed includeConstant) constant_f isTransitive isFormula
                let rebased_vs = vs |> Array.map (fun v -> rebase_f v dag isOffSheetInsensitive includeLoc)
                let rvarrs =
                    rebased_vs |>
                    Array.map (fun v ->
                        match v with
                        | ConstantWithLoc(x,y,z,x',y',z',c) -> Countable.FullCVectorResultant((double x, double y, double z, double x', double y', double z', double c))   
                        | _ -> failwith "this should never happen"
                    ) |>
                    Array.map (fun v -> v.LocationFree)
                rvarrs