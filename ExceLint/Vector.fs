module Vector

open COMWrapper
open Depends
open Feature

type AbsoluteVector = (int*int)*(int*int)
type OriginVector = (int*int)

type OriginStructureVector() = 
    inherit BaseFeature()

    static member vector(fromCell: AST.Address)(toCell: AST.Address) : AbsoluteVector =
        failwith "not implemented"

    static member rebaseVector(absVect: AbsoluteVector) : OriginVector =
        failwith "not implemented"

    static member vectorSum(origVects: OriginVector[]) : double =
        failwith "not implemented"

    static member run cell (dag : DAG) = 
        let fStr = dag.getFormulaAtAddress cell

        // get all referents of the formula
        let refs = dag.getFormulaSingleCellInputs cell |> Array.ofSeq

        // for every referent, get its vector
        let vectors = Array.map (fun r -> OriginStructureVector.vector cell r) refs

        // rebase vectors to origin
        let vectors' = Array.map OriginStructureVector.rebaseVector vectors

        // recursive procedure:
        // if no children, compute summary statistic (vector sum)
        // else recurse and compute summary statistics for children

        failwith "not implemented"

type RelativeStructureVector() = 
    inherit BaseFeature()

    static member run cell (dag : DAG) = 
        failwith "not implemented"