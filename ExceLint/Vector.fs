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
        // (tail)*(head); the vector, not relative
        type AbsoluteVector = (X*Y*Path)*(X*Y*Path)
        // the vector, relative to a basis
        type BasisVector = (X*Y*Z)

        let private fullPath(addr: AST.Address) : string =
            // portably create full path from components
            System.IO.Path.Combine([|addr.Path; addr.WorkbookName; addr.WorksheetName|])

        let private vector(tail: AST.Address)(head: AST.Address) : AbsoluteVector =
            let tailXYP = (tail.X, tail.Y, fullPath tail)
            let headXYP = (head.X, head.Y, fullPath head)
            (tailXYP, headXYP)

        let private originPath(dag: DAG) : Path =
            System.IO.Path.Combine(dag.getWorkbookPath(), dag.getWorksheetNames().[0]);

        let private vectorPathDiff(p1: Path)(p2: Path) : int =
            if p1 <> p2 then 1 else 0

        // the origin is defined as x = 0, y = 0, z = 0 if first sheet in the workbook else 1
        let private pathDiff(p: Path)(dag: DAG) : int =
            let op = originPath dag
            if p <> op then 1 else 0

        // represent the position of the head of the vector relative to the tail, (x1,y1,z1)
        let private rebaseVectorToTail(absVect: AbsoluteVector)(dag: DAG) : BasisVector =
            let ((x1,y1,p1),(x2,y2,p2)) = absVect
            (x2-x1, y2-y1, vectorPathDiff p2 p1)

        // represent the position of the the head of the vector relative to the origin, (0,0,0)
        let private rebaseVectorToOrigin(absVect: AbsoluteVector)(dag: DAG) : BasisVector =
            let (_,(x2,y2,p2)) = absVect
            let v = (x2, y2, pathDiff p2 dag)
            v

        let private L2Norm(X: double[]) : double =
            Math.Sqrt(
                Array.sumBy (fun x -> Math.Pow(x, 2.0)) X
            )

        let private basisVectorToRealVectorArr(v: BasisVector) : double[] =
            let (x,y,z) = v
            [|
                System.Convert.ToDouble(x);
                System.Convert.ToDouble(y);
                System.Convert.ToDouble(z);
            |]

        let private L2NormBV(v: BasisVector) : double =
            L2Norm(basisVectorToRealVectorArr(v))

        let private L2NormBVSum(vs: BasisVector[]) : double =
            vs |> Array.map L2NormBV |> Array.sum

        let transitiveFormulaVectors(fCell: AST.Address)(dag : DAG)(depth: int option) : AbsoluteVector[] =
            let rec tfVect(sinkO: AST.Address option)(source: AST.Address)(depth: int option) : AbsoluteVector list =
                let vlist = match sinkO with
                            | Some sink -> [vector sink source]
                            | None -> []

                match depth with
                | Some(0) -> vlist
                | Some(d) -> tfVect_b source (Some(d-1)) vlist
                | None -> tfVect_b source None vlist

            and tfVect_b(source: AST.Address)(nextDepth: int option)(vlist: AbsoluteVector list) : AbsoluteVector list =
                if (dag.isFormula source) then
                    // find all of the inputs for source
                    let sources_single = dag.getFormulaSingleCellInputs source |> List.ofSeq
                    let sources_vector = dag.getFormulaInputVectors source |>
                                            List.ofSeq |>
                                            List.map (fun rng -> rng.Addresses() |> Array.toList) |>
                                            List.concat
                    let sources' = sources_single @ sources_vector
                    // recursively call this function
                    vlist @ (List.map (fun source' -> tfVect (Some source) source' nextDepth) sources' |> List.concat)
                else
                    vlist
    
            tfVect None fCell depth |> List.toArray

        let formulaVectors(fCell: AST.Address)(dag : DAG) : AbsoluteVector[] =
            transitiveFormulaVectors fCell dag (Some 1)

        let transitiveDataVectors(dCell: AST.Address)(dag : DAG)(depth: int option) : AbsoluteVector[] =
            let rec tdVect(sourceO: AST.Address option)(sink: AST.Address)(depth: int option) : AbsoluteVector list =
                let vlist = match sourceO with
                            | Some source -> [vector sink source]
                            | None -> []

                match depth with
                | Some(0) -> vlist
                | Some(d) -> tdVect_b sink (Some(d-1)) vlist
                | None -> tdVect_b sink None vlist

            and tdVect_b(sink: AST.Address)(nextDepth: int option)(vlist: AbsoluteVector list) : AbsoluteVector list =
                    // find all of the formulas that use sink
                    let outAddrs = dag.getFormulasThatRefCell sink
                                    |> Array.toList
                    let outAddrs2 = Array.map (dag.getFormulasThatRefVector) (dag.getVectorsThatRefCell sink)
                                    |> Array.concat |> Array.toList
                    let allFrm = outAddrs @ outAddrs2 |> List.distinct

                    // recursively call this function
                    vlist @ (List.map (fun sink' -> tdVect (Some sink) sink' nextDepth) allFrm |> List.concat)

            tdVect None dCell depth |> List.toArray

        let dataVectors(dCell: AST.Address)(dag : DAG) : AbsoluteVector[] =
            transitiveDataVectors dCell dag (Some 1)

        let getVectors(cell: AST.Address)(dag: DAG)(transitive: bool)(isForm: bool)(isRel: bool) : BasisVector[] =
            let depth = if transitive then None else (Some 1)
            let vectors =
                if isForm then
                    transitiveFormulaVectors
                else
                    transitiveDataVectors
            let rebase =
                if isRel then
                    Array.map (fun v -> rebaseVectorToTail v dag)
                else
                    Array.map (fun v -> rebaseVectorToOrigin v dag)
            rebase (vectors cell dag depth)

        type TransitiveFormulaRelativeL2NormSum() = 
            inherit BaseFeature()

            // fCell is the address of a formula here
            static member run(cell: AST.Address)(dag : DAG) : double = 
                L2NormBVSum (getVectors cell dag (*transitive*) true (*isForm*) true (*isRel*) true)

        type TransitiveDataRelativeL2NormSum() = 
            inherit BaseFeature()

            // dCell is the address of a data cell here
            static member run(cell: AST.Address)(dag : DAG) : double = 
                L2NormBVSum (getVectors cell dag (*transitive*) true (*isForm*) false (*isRel*) true)

        type TransitiveFormulaAbsoluteL2NormSum() =
            inherit BaseFeature()

            // fCell is the address of a formula here
            static member run(cell: AST.Address)(dag: DAG) : double =
                L2NormBVSum (getVectors cell dag (*transitive*) true (*isForm*) true (*isRel*) false)

        type TransitiveDataAbsoluteL2NormSum() =
            inherit BaseFeature()

            // dCell is the address of a data cell here
            static member run(cell: AST.Address)(dag: DAG) : double =
                L2NormBVSum (getVectors cell dag (*transitive*) true (*isForm*) false (*isRel*) false)

        type FormulaRelativeL2NormSum() = 
            inherit BaseFeature()

            // fCell is the address of a formula here
            static member run(cell: AST.Address)(dag : DAG) : double = 
                L2NormBVSum (getVectors cell dag (*transitive*) false (*isForm*) true (*isRel*) true)

        type DataRelativeL2NormSum() = 
            inherit BaseFeature()

            // dCell is the address of a data cell here
            static member run(cell: AST.Address)(dag : DAG) : double = 
                L2NormBVSum (getVectors cell dag (*transitive*) false (*isForm*) false (*isRel*) true)

        type FormulaAbsoluteL2NormSum() =
            inherit BaseFeature()

            // fCell is the address of a formula here
            static member run(cell: AST.Address)(dag: DAG) : double =
                L2NormBVSum (getVectors cell dag (*transitive*) false (*isForm*) true (*isRel*) false)

        type DataAbsoluteL2NormSum() =
            inherit BaseFeature()

            // dCell is the address of a data cell here
            static member run(cell: AST.Address)(dag: DAG) : double =
                L2NormBVSum (getVectors cell dag (*transitive*) false (*isForm*) false (*isRel*) false)