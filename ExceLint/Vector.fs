namespace ExceLint
    open COMWrapper
    open Depends
    open Feature
    open System

    module Vector =
        type Path = string
        type X = int    // i.e., column displacement
        type Y = int    // i.e., row displacement
        type Z = int    // i.e., worksheet displacement (0 if same sheet, 1 if different)

        // components for mixed vectors
        type VectorComponent =
        | Abs of int
        | Rel of int

        // the vector, relative to an origin
        type RelativeVector = (X*Y*Z)
        type MixedVector = (VectorComponent*VectorComponent*Path)

        // the first component is the tail (start) and the second is the head (end)
        type FullyQualifiedVector =
        | MixedFQVector of (X*Y*Path)*MixedVector
        | AbsoluteFQVector of (X*Y*Path)*(X*Y*Path)

        let private fullPath(addr: AST.Address) : string =
            // portably create full path from components
            System.IO.Path.Combine([|addr.Path; addr.WorkbookName; addr.WorksheetName|])

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
                failwith "nope"
            else
                let headXYP = (head.X, head.Y, fullPath head)
                AbsoluteFQVector(tailXYP, headXYP)

        let private originPath(dag: DAG) : Path =
            System.IO.Path.Combine(dag.getWorkbookPath(), dag.getWorksheetNames().[0]);

        let private vectorPathDiff(p1: Path)(p2: Path) : int =
            if p1 <> p2 then 1 else 0

        // the origin is defined as x = 0, y = 0, z = 0 if first sheet in the workbook else 1
        let private pathDiff(p: Path)(dag: DAG) : int =
            let op = originPath dag
            if p <> op then 1 else 0

        // represent the position of the head of the vector relative to the tail, (x1,y1,z1)
        let private relativeToTail(absVect: FullyQualifiedVector)(dag: DAG) : RelativeVector =
            match absVect with
            | AbsoluteFQVector(tail,head) ->
                let (x1,y1,p1) = tail
                let (x2,y2,p2) = head
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
                (x', y', vectorPathDiff p2 p1)

        // represent the position of the the head of the vector relative to the origin, (0,0,0)
        let private relativeToOrigin(absVect: FullyQualifiedVector)(dag: DAG) : RelativeVector =
            match absVect with
            | AbsoluteFQVector(tail,head) ->
                let (x,y,p) = head
                (x, y, pathDiff p dag)
            | MixedFQVector(tail,head) ->
                let (x,y,p) = head
                let x' = match x with | Abs(xa) -> xa | Rel(xr) -> xr
                let y' = match y with | Abs(ya) -> ya | Rel(yr) -> yr
                (x', y', pathDiff p dag)

        let private L2Norm(X: double[]) : double =
            Math.Sqrt(
                Array.sumBy (fun x -> Math.Pow(x, 2.0)) X
            )

        let private basisVectorToRealVectorArr(v: RelativeVector) : double[] =
            let (x,y,z) = v
            [|
                System.Convert.ToDouble(x);
                System.Convert.ToDouble(y);
                System.Convert.ToDouble(z);
            |]

        let private L2NormBV(v: RelativeVector) : double =
            L2Norm(basisVectorToRealVectorArr(v))

        let private L2NormBVSum(vs: RelativeVector[]) : double =
            vs |> Array.map L2NormBV |> Array.sum

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

        let dataVectors(dCell: AST.Address)(dag : DAG)(mixed: bool) : FullyQualifiedVector[] =
            transitiveOutputVectors dCell dag (Some 1) mixed

        let getVectors(cell: AST.Address)(dag: DAG)(transitive: bool)(isForm: bool)(isRel: bool)(isMixed: bool) : RelativeVector[] =
            let depth = if transitive then None else (Some 1)
            let vectors =
                if isForm then
                    transitiveInputVectors
                else
                    transitiveOutputVectors
            let rebase =
                if isRel then
                    Array.map (fun v -> relativeToTail v dag)
                else
                    Array.map (fun v -> relativeToOrigin v dag)
            rebase (vectors cell dag depth isMixed)

        type RelativeTransitiveInputL2NormSum() = 
            inherit BaseFeature()

            // fCell is the address of a formula here
            static member run(cell: AST.Address)(dag : DAG) : double = 
                L2NormBVSum (getVectors cell dag (*transitive*) true (*isForm*) true (*isRel*) true (*isMixed*) false)

        type TransitiveDataRelativeL2NormSum() = 
            inherit BaseFeature()

            // dCell is the address of a data cell here
            static member run(cell: AST.Address)(dag : DAG) : double = 
                L2NormBVSum (getVectors cell dag (*transitive*) true (*isForm*) false (*isRel*) true (*isMixed*) false)

        type TransitiveFormulaAbsoluteL2NormSum() =
            inherit BaseFeature()

            // fCell is the address of a formula here
            static member run(cell: AST.Address)(dag: DAG) : double =
                L2NormBVSum (getVectors cell dag (*transitive*) true (*isForm*) true (*isRel*) false (*isMixed*) false)

        type TransitiveDataAbsoluteL2NormSum() =
            inherit BaseFeature()

            // dCell is the address of a data cell here
            static member run(cell: AST.Address)(dag: DAG) : double =
                L2NormBVSum (getVectors cell dag (*transitive*) true (*isForm*) false (*isRel*) false (*isMixed*) false)

        type FormulaRelativeL2NormSum() = 
            inherit BaseFeature()

            // fCell is the address of a formula here
            static member run(cell: AST.Address)(dag : DAG) : double = 
                L2NormBVSum (getVectors cell dag (*transitive*) false (*isForm*) true (*isRel*) true (*isMixed*) false)

        type DataRelativeL2NormSum() = 
            inherit BaseFeature()

            // dCell is the address of a data cell here
            static member run(cell: AST.Address)(dag : DAG) : double = 
                L2NormBVSum (getVectors cell dag (*transitive*) false (*isForm*) false (*isRel*) true (*isMixed*) false)

        type FormulaAbsoluteL2NormSum() =
            inherit BaseFeature()

            // fCell is the address of a formula here
            static member run(cell: AST.Address)(dag: DAG) : double =
                L2NormBVSum (getVectors cell dag (*transitive*) false (*isForm*) true (*isRel*) false (*isMixed*) false)

        type DataAbsoluteL2NormSum() =
            inherit BaseFeature()

            // dCell is the address of a data cell here
            static member run(cell: AST.Address)(dag: DAG) : double =
                L2NormBVSum (getVectors cell dag (*transitive*) false (*isForm*) false (*isRel*) false (*isMixed*) false)