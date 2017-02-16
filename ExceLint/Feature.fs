module Feature
    open Depends
    open System

    type ConfigKind =
    | Feature
    | Scope
    | Misc

//    type ICountable<'T>=
//        abstract member Add: 'T -> 'T
//
//    type Num(d: double) =
//        member self.d = d
//        interface ICountable<Num> with
//            member self.Add(n: Num) = Num(d + n.d)

    type Countable =
    | Num of double
    | Vector of double*double*double
    | ResultantVectorWithConstant of double*double*double*double
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
        member self.Negate : Countable =
            match self with
            | Num n -> Num(-n)
            | Vector(x,y,z) -> Vector(-x, -y, -z)
            | SquareVector(dx,dy,dz,x,y,z) ->
                SquareVector(-dx,-dy,-dz,-x,-y,-z)
        member self.ScalarDivide(d: double) : Countable =
            match self with
            | Num n -> Num(n / d)
            | Vector(x,y,z) -> Vector(x / d, y / d, z / d)
            | SquareVector(dx,dy,dz,x,y,z) ->
                SquareVector(dx / d, dy / d, dz / d, x / d, y / d, z / d)
        member self.VectorMultiply(co: Countable) : double =
            match (self,co) with
            | Num d1, Num d2 -> d1 * d2
            | Vector(x1,y1,z1), Vector(x2,y2,z2) -> x1 * x2 + y1 * y2 + z1 * z2
            | SquareVector(dx1,dy1,dz1,x1,y1,z1), SquareVector(dx2,dy2,dz2,x2,y2,z2) ->
                dx1 * dx2 + dy1 * dy2 + dz1 * dz2 + x1 * x2 + y1 * y2 + z1 * z2
            | _ -> failwith "Invalid operation.  Both Countables must be of the same type."
        member self.Sqrt : Countable =
            match self with
            | Num n -> Num(Math.Sqrt(n))
            | Vector(x,y,z) -> Vector(Math.Sqrt(x), Math.Sqrt(y), Math.Sqrt(z))
            | SquareVector(dx,dy,dz,x,y,z) ->
                SquareVector(Math.Sqrt(dx), Math.Sqrt(dy), Math.Sqrt(dz), Math.Sqrt(x), Math.Sqrt(y), Math.Sqrt(z))
        member self.Abs : Countable =
            match self with
            | Num n -> Num(Math.Abs(n))
            | Vector(x,y,z) -> Vector(Math.Abs(x),Math.Abs(y),Math.Abs(z))
            | SquareVector(dx,dy,dz,x,y,z) ->
                SquareVector(Math.Abs(dx),Math.Abs(dy),Math.Abs(dz),Math.Abs(x),Math.Abs(y),Math.Abs(z))

    type Capability = { enabled : bool; kind: ConfigKind; runner: AST.Address -> Depends.DAG -> Countable; }

    type FeatureLambda = AST.Address -> DAG -> Countable

    type BaseFeature() =
        static member run (cell: AST.Address) (dag: DAG): Countable = failwith "Feature must provide run method."
        static member capability : string*Capability = failwith "Feature must provide capability."

    // the following are for C# interop because
    // discriminated union types are not exported
    let makeNum(d: double) = Num d
    let makeVector(x: double, y: double, z: double) = Vector(x,y,z)
    let makeSpatialVector(dx: double, dy: double, dz: double, x: double, y: double, z: double) = SquareVector(dx,dy,dz,x,y,z)
