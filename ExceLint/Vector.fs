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

        let transitiveFormulaVectors(fCell: AST.Address)(dag : DAG) : AbsoluteVector[] =
            let rec tfVect(sinkO: AST.Address option)(source: AST.Address) : AbsoluteVector list =
                let vlist = match sinkO with
                            | Some sink -> [vector sink source]
                            | None -> []

                if (dag.isFormula source) then
                    // find all of the inputs for source
                    let sources_single = dag.getFormulaSingleCellInputs source |> List.ofSeq
                    let sources_vector = dag.getFormulaInputVectors source |>
                                            List.ofSeq |>
                                            List.map (fun rng -> rng.Addresses() |> Array.toList) |>
                                            List.concat
                    let sources' = sources_single @ sources_vector
                    // recursively call this function
                    vlist @ (List.map (fun source' -> tfVect (Some source) source') sources' |> List.concat)
                else
                    vlist
    
            tfVect None fCell |> List.toArray

        let transitiveDataVectors(dCell: AST.Address)(dag : DAG) : AbsoluteVector[] =
            let rec tdVect(source0: AST.Address option)(sink: AST.Address) : AbsoluteVector list =
                let vlist = match source0 with
                            | Some source -> [vector sink source]
                            | None -> []

                // find all of the formulas that use sink
                let outAddrs = dag.getFormulasThatRefCell sink
                                |> Array.toList
                let outAddrs2 = Array.map (dag.getFormulasThatRefVector) (dag.getVectorsThatRefCell sink)
                                |> Array.concat |> Array.toList
                let allFrm = outAddrs @ outAddrs2 |> List.distinct

                // recursively call this function
                vlist @ (List.map (fun sink' -> tdVect (Some sink) sink') allFrm |> List.concat)

            tdVect None dCell |> List.toArray

        // the transitive set of vectors, starting at a given formula cell,
        // all relative to their tails
        let transitiveFormulaRelativeVectors(fCell: AST.Address)(dag: DAG) : BasisVector[] =
            transitiveFormulaVectors fCell dag |>
            Array.map (fun v -> rebaseVectorToTail v dag)

        // the transitive set of vectors, starting at a given data cell,
        // all relative to their tails
        let transitiveDataRelativeVectors(dCell: AST.Address)(dag: DAG) : BasisVector[] =
            transitiveDataVectors dCell dag |>
            Array.map (fun v -> rebaseVectorToTail v dag)

        // the transitive set of vectors, starting at a given formula cell,
        // all relative to the origin
        let transitiveFormulaOriginVectors(fCell: AST.Address)(dag: DAG) : BasisVector[] =
            transitiveFormulaVectors fCell dag |>
            Array.map (fun v -> rebaseVectorToOrigin v dag)

        // the transitive set of vectors, starting at a given data cell,
        // all relative to the origin
        let transitiveDataOriginVectors(fCell: AST.Address)(dag: DAG) : BasisVector[] =
            transitiveDataVectors fCell dag |>
            Array.map (fun v -> rebaseVectorToOrigin v dag)

        type FormulaRelativeL2NormSum() = 
            inherit BaseFeature()

            // fCell is the address of a formula here
            static member run(fCell: AST.Address)(dag : DAG) : double = 
                L2NormBVSum (transitiveFormulaRelativeVectors fCell dag)

        type DataRelativeL2NormSum() = 
            inherit BaseFeature()

            // dCell is the address of a data cell here
            static member run(dCell: AST.Address)(dag : DAG) : double = 
                L2NormBVSum (transitiveDataRelativeVectors dCell dag)

        type FormulaAbsoluteL2NormSum() =
            inherit BaseFeature()

            // fCell is the address of a formula here
            static member run(fCell: AST.Address)(dag: DAG) : double =
                let vs = transitiveFormulaOriginVectors fCell dag
                L2NormBVSum vs

        type DataAbsoluteL2NormSum() =
            inherit BaseFeature()

            // dCell is the address of a data cell here
            static member run(dCell: AST.Address)(dag: DAG) : double =
                let vs = transitiveDataOriginVectors dCell dag
                L2NormBVSum vs