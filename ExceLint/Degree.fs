module Degree
    open COMWrapper
    open Depends

    /// <summary>Gets the number of inputs referenced by fAddr.</summary>
    /// <param name="fAddr">the address of a formula cell</param>
    /// <param name="dag">a DAG</param>
    /// <returns>the number of inputs referenced by fAddr</returns>
    let getIndegreeForFormulaCell(fAddr: AST.Address)(dag: DAG) : int =
        let allInputAddresses = (dag.getFormulaSingleCellInputs fAddr |> List.ofSeq)
                                @ (dag.getFormulaInputVectors fAddr
                                    |> List.ofSeq
                                    |> List.map (fun rng -> rng.Addresses() |> List.ofArray)
                                    |> List.concat)
        allInputAddresses.Length

    /// <summary>Gets the number of inputs referenced by fAddr.</summary>
    /// <param name="cell">the address of an arbitrary cell</param>
    /// <param name="dag">a DAG</param>
    /// <returns>the number of inputs referenced by cell</returns>
    let getIndegreeForCell(cell: AST.Address)(dag: DAG) =
        if dag.isFormula(cell) then
            getIndegreeForFormulaCell cell dag
        else
            0

    /// <summary>Gets the number of formulas that reference a cell.</summary>
    /// <param name="cell">the address of an arbitrary cell</param>
    /// <param name="dag">a DAG</param>
    /// <returns>the number of formulas that reference a cell</returns>
    let getOutdegreeForCell(cell: AST.Address)(dag: DAG) =
        let referencingFormulas = dag.getFormulasThatRefCell cell
        let formulasThatRefRanges =
            Array.map (fun rng ->
                dag.getFormulasThatRefVector(rng)
            ) (dag.getVectorsThatRefCell cell) |> Array.concat

        referencingFormulas.Length + formulasThatRefRanges.Length