namespace ExceLint
    open COMWrapper
    open Depends
    open Feature
    open System

    module Vector =
        type AbsoluteVector = (int*int)*(int*int)
        type OriginVector = (int*int)

        let private vector(sink: AST.Address)(source: AST.Address) : AbsoluteVector =
            let sinkXY = (sink.X, sink.Y)
            let sourceXY = (source.X, source.Y)
            (sinkXY, sourceXY)

        let private rebaseVector(absVect: AbsoluteVector) : OriginVector =
            let ((x1,y1),(x2,y2)) = absVect
            (x2-x1, y2-y1)

        let private L2Norm(origVect: OriginVector) : double =
            let (x,y) = origVect
            Math.Sqrt(
                Math.Pow(System.Convert.ToDouble(x),2.0) +
                Math.Pow(System.Convert.ToDouble(y),2.0)
            )

        let private L2NormSum(origVects: OriginVector[]) : double =
            origVects |> Array.map L2Norm |> Array.sum

        let transitiveSourceVectors(cell: AST.Address, dag : DAG) : AbsoluteVector[] =
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

        type OriginStructureVector() = 
            inherit BaseFeature()

            static member run cell (dag : DAG) = 
                failwith "not implemented"

        type RelativeStructureVector() = 
            inherit BaseFeature()

            static member run cell (dag : DAG) = 
                failwith "not implemented"