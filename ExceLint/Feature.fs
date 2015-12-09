module Feature

    open COMWrapper
    open Depends


    type BaseFeature() =
        static member run (cell: AST.Address) (dag: DAG) : double = 0.0