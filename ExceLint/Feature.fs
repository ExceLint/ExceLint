module Feature
    open Depends

    type ConfigKind =
    | Feature
    | Scope
    | Misc

    type Countable =
    | Num of double
    | Vector of double*double*double
    | SquareVector of double*double*double*double*double*double
        override self.ToString() : string =
            match self with
            | Num d -> d.ToString()
            | Vector(x,y,z) ->
                "<" + x.ToString() + "," + y.ToString() + "," + z.ToString() + ">"
            | SquareVector(dx,dy,dz,x,y,z) ->
                "<" + dx.ToString() + "," + dy.ToString() + "," + dz.ToString() + "," + x.ToString() + "," + y.ToString() + "," + z.ToString() +  ">"
        member self.MeanFoldDefault : Countable =
            match self with
            | Num _ -> Num(0.0)
            | Vector _ -> Vector(0.0,0.0,0.0)
            | SquareVector _ -> SquareVector(0.0,0.0,0.0,0.0,0.0,0.0)
        member self.Add(co: Countable) : Countable =
            match (self,co) with
            | Num d1, Num d2 -> Num(d1 + d2)
            | Vector(x1,y1,z1), Vector(x2,y2,z2) -> Vector(x1 + x2, y1 + y2, z1 + z2)
            | SquareVector(dx1,dy1,dz1,x1,y1,z1), SquareVector(dx2,dy2,dz2,x2,y2,z2) ->
                SquareVector(dx1 + dx2, dy1 + dy2, dz1 + dz2, x1 + x2, y1 + y2, z1 + z2)
            | _ -> failwith "Invalid operation.  Both Countables must be of the same type."
        member self.ScalarDivide(d: double) : Countable =
            match self with
            | Num n -> Num(n / d)
            | Vector(x,y,z) -> Vector(x / d, y / d, z / d)
            | SquareVector(dx,dy,dz,x,y,z) ->
                SquareVector(dx / d, dy / d, dz / d, x / d, y / d, z / d)
            | _ -> failwith "Invalid operation.  Both Countables must be of the same type."

    type Capability = { enabled : bool; kind: ConfigKind; runner: AST.Address -> Depends.DAG -> Countable; }

    type FeatureLambda = AST.Address -> DAG -> Countable

    type BaseFeature() =
        static member run (cell: AST.Address) (dag: DAG): Countable = failwith "Feature must provide run method."
        static member capability : string*Capability = failwith "Feature must provide capability."
