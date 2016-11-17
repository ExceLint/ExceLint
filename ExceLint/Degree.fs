module Degree
    open Depends
    open Feature

    type InDegree() = 
         inherit BaseFeature()
         /// <summary>Gets the number of inputs referenced by fAddr.</summary>
         /// <param name="fAddr">the address of a formula cell</param>
         /// <param name="dag">a DAG</param>
         /// <returns>the number of inputs referenced by fAddr</returns>
         static member private getIndegreeForFormulaCell (fAddr: AST.Address) (dag : DAG) =
            let allInputAddresses = 
                (dag.getFormulaSingleCellInputs fAddr |> List.ofSeq)
                @ (dag.getFormulaInputVectors fAddr
                    |> List.ofSeq
                    |> List.map (fun rng -> rng.Addresses() |> List.ofArray)
                    |> List.concat)
            System.Convert.ToDouble(allInputAddresses.Length)

         /// <summary>Gets the number of inputs referenced by fAddr.</summary>
         /// <param name="cell">the address of an arbitrary cell</param>
         /// <param name="dag">a DAG</param>
         /// <returns>the number of inputs referenced by cell</returns>
         static member run cell (dag : DAG) = 
             if dag.isFormula(cell) then
                 InDegree.getIndegreeForFormulaCell cell dag |> Num
             else
                 0.0 |> Num

         static member capability : string*Capability =
             (typeof<InDegree>.Name,
                 { enabled = false; kind = ConfigKind.Feature; runner = InDegree.run } )

    type OutDegree() =
        inherit BaseFeature()

            /// <summary>Gets the number of formulas that reference a cell.</summary>
            /// <param name="cell">the address of an arbitrary cell</param>
            /// <param name="dag">a DAG</param>
            /// <returns>the number of formulas that reference a cell</returns>
            static member run cell (dag: DAG) =
                let referencingFormulas = dag.getFormulasThatRefCell cell
                let formulasThatRefRanges =
                    Array.map (fun rng ->
                        dag.getFormulasThatRefVector(rng)
                    ) (dag.getVectorsThatRefCell cell) |> Array.concat

                System.Convert.ToDouble(referencingFormulas.Length + formulasThatRefRanges.Length)

            static member capability : string*Capability =
                (typeof<OutDegree>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = InDegree.run } )

    type CombinedDegree() =
        inherit BaseFeature()

            static member run cell (dag: DAG) =
                let idnum = match (InDegree.run cell dag) with
                            | Num n -> n
                            | _ -> failwith "impossible"
                (idnum + OutDegree.run cell dag) |> Num

            static member capability : string*Capability =
                (typeof<CombinedDegree>.Name,
                    { enabled = false; kind = ConfigKind.Feature; runner = InDegree.run } )