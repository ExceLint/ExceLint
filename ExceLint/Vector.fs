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
        type AbsoluteVector = (X*Y*Path)*(X*Y*Path)
        // the origin is defined as x = 0, y = 0, z = 0 if first sheet in the workbook else 1
        type OriginVector = (X*Y*Z)

        let private vector(sink: AST.Address)(source: AST.Address) : AbsoluteVector =
            let sinkXYP = (sink.X, sink.Y, sink.Path)
            let sourceXYP = (source.X, source.Y, source.Path)
            (sinkXYP, sourceXYP)

        let private originPath(dag: DAG) : Path =
            dag.getWorkbookPath()

        let private pathDiff(p: Path)(dag: DAG) : int =
            if p = (originPath dag) then 1 else 0

        let private rebaseVector(absVect: AbsoluteVector)(dag: DAG) : OriginVector =
            let ((x1,y1,p1),(x2,y2,p2)) = absVect
            (x2-x1, y2-y1, (pathDiff p2 dag)-(pathDiff p1 dag))

        let private L2Norm(X: double[]) : double =
            Math.Sqrt(
                Array.sumBy (fun x -> Math.Pow(x, 2.0)) X
            )

        let private originVectorToRealVector(v: OriginVector) : double[] =
            let (x,y,z) = v
            [|
                System.Convert.ToDouble(x);
                System.Convert.ToDouble(y);
                System.Convert.ToDouble(z);
            |]

        let private L2NormOV(v: OriginVector) : double =
            L2Norm(originVectorToRealVector(v))

        let private L2NormOVSum(vs: OriginVector[]) : double =
            vs |> Array.map L2NormOV |> Array.sum

        let transitiveFormulaVectors(cell: AST.Address)(dag : DAG) : AbsoluteVector[] =
            let rec tVect(sinkO: AST.Address option)(source: AST.Address) : AbsoluteVector list =
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
                    vlist @ (List.map (fun source' -> tVect (Some source) source') sources' |> List.concat)
                else
                    vlist
    
            tVect None cell |> List.toArray

        let transitiveFormulaRelativeVectors(cell: AST.Address)(dag: DAG) : OriginVector[] =
            transitiveFormulaVectors cell dag |>
            Array.map (fun v -> rebaseVector v dag)

        type FormulaRelativeL2NormSum() = 
            inherit BaseFeature()

            static member run(cell: AST.Address)(dag : DAG) : double = 
                L2NormOVSum (transitiveFormulaRelativeVectors cell dag)

        type FormulaRelativeAngleSum() = 
            inherit BaseFeature()

            static member run(cell: AST.Address)(dag : DAG) = 
                failwith "not implemented"