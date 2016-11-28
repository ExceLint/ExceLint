module Feature
    open Depends

    type ConfigKind =
    | Feature
    | Scope
    | Misc

    type Countable =
    | Num of double
    | Vector of int*int*int
        override self.ToString() : string =
            match self with
            | Num d -> d.ToString()
            | Vector(x,y,z) ->
                "<" + x.ToString() + "," + y.ToString() + "," + z.ToString() + ">"

    type Capability = { enabled : bool; kind: ConfigKind; runner: AST.Address -> Depends.DAG -> Countable; }

    type BaseFeature() =
        static member run (cell: AST.Address) (dag: DAG): Countable = failwith "Feature must provide run method."
        static member capability : string*Capability = failwith "Feature must provide capability."
