module Feature
    open Depends

    type ConfigKind =
    | Feature
    | Scope
    | Misc

    type Capability = { enabled : bool; kind: ConfigKind; runner: AST.Address -> Depends.DAG -> double; }

    type BaseFeature() =
        static member run (cell: AST.Address) (dag: DAG): double = failwith "Feature must provide run method."
        static member capability : string*Capability = failwith "Feature must provide capability."
