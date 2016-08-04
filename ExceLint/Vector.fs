namespace ExceLint
    open Depends
    open Feature
    open System

    module Vector =
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
//            System.IO.Path.Combine([|addr.Path; addr.WorkbookName; addr.WorksheetName|])
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
//            System.IO.Path.Combine(dag.getWorkbookPath(), dag.getWorksheetNames().[0]);

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
