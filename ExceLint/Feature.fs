namespace ExceLint
    open Depends
    open System

    type ConfigKind =
    | Feature
    | Scope
    | Misc
    
    type Countable =
    | Num of double
    | Vector of double*double*double
    | SquareVector of double*double*double*double*double*double
    | CVectorResultant of double*double*double*double
    | FullCVectorResultant of double*double*double*double*double*double*double
        override self.ToString() : string =
            match self with
            | Num d -> d.ToString()
            | Vector(x,y,z) ->
                "<" + x.ToString() + "," + y.ToString() + "," + z.ToString() + ">"
            | SquareVector(dx,dy,dz,x,y,z) ->
                "<" + dx.ToString() + "," + dy.ToString() + "," + dz.ToString() + "," + x.ToString() + "," + y.ToString() + "," + z.ToString() +  ">"
            | CVectorResultant(x,y,z,c) ->
                "<" + x.ToString() + "," + y.ToString() + "," + z.ToString() + "," + c.ToString() + ">"
            | FullCVectorResultant(x,y,z,dx,dy,dz,dc) ->
                "<" + x.ToString() + "," + y.ToString() + "," + z.ToString() + "," + dx.ToString() + "," + dy.ToString() + "," + dz.ToString() + "," + dc.ToString() + ">"
        member self.LocationFree : Countable =
            match self with
            | FullCVectorResultant(x,y,z,dx,dy,dz,dc) -> CVectorResultant(dx,dy,dz,dc)
            | _ -> self
        member self.Zero : Countable =
            match self with
            | Num _ -> Num(0.0)
            | Vector _ -> Vector(0.0,0.0,0.0)
            | SquareVector _ -> SquareVector(0.0,0.0,0.0,0.0,0.0,0.0)
            | CVectorResultant _ -> CVectorResultant(0.0,0.0,0.0,0.0)
            | FullCVectorResultant _ -> FullCVectorResultant(0.0,0.0,0.0,0.0,0.0,0.0,0.0)
        member self.ToVector : double[] =
            match self with
            | Num n -> [| n |]
            | Vector(x,y,z) -> [| x; y; z |]
            | SquareVector(dx,dy,dz,x,y,z) -> [| dx; dy; dz; x; y; z |]
            | CVectorResultant(x1,y1,z1,c1) -> [| x1; y1; z1; c1 |]
            | FullCVectorResultant(x,y,z,dx,dy,dz,dc) -> [| x; y; z; dx; dy; dz; dc |]
        member self.Add(co: Countable) : Countable =
            match (self,co) with
            | Num d1, Num d2 -> Num(d1 + d2)
            | Vector(x1,y1,z1), Vector(x2,y2,z2) -> Vector(x1 + x2, y1 + y2, z1 + z2)
            | SquareVector(dx1,dy1,dz1,x1,y1,z1), SquareVector(dx2,dy2,dz2,x2,y2,z2) ->
                SquareVector(dx1 + dx2, dy1 + dy2, dz1 + dz2, x1 + x2, y1 + y2, z1 + z2)
            | CVectorResultant(x1,y1,z1,c1), CVectorResultant(x2,y2,z2,c2) ->
                CVectorResultant(x1 + x2, y1 + y2, z1 + z2, c1 + c2)
            | FullCVectorResultant(x1,y1,z1,dx1,dy1,dz1,dc1),FullCVectorResultant(x2,y2,z2,dx2,dy2,dz2,dc2) ->
                FullCVectorResultant(x1 + x2, y1 + y2, z1 + z2, dx1 + dx2, dy1 + dy2, dz1 + dz2, dc1 + dc2)
            | _ -> failwith "Invalid operation.  Both Countables must be of the same type."
        member self.Negate : Countable =
            match self with
            | Num n -> Num(-n)
            | Vector(x,y,z) -> Vector(-x, -y, -z)
            | SquareVector(dx,dy,dz,x,y,z) ->
                SquareVector(-dx,-dy,-dz,-x,-y,-z)
            | CVectorResultant(x1,y1,z1,c1) ->
                CVectorResultant(-x1,-y1,-z1,-c1)
            | FullCVectorResultant(x,y,z,dx,dy,dz,dc) ->
                FullCVectorResultant(-x,-y,-z,-dx,-dy,-dz,-dc)
        member self.Sub(co: Countable) : Countable =
            self.Add(co.Negate)
        member self.ScalarMultiply(d: double) : Countable =
            match self with
            | Num n -> Num(n * d)
            | Vector(x,y,z) -> Vector(x * d, y * d, z * d)
            | SquareVector(dx,dy,dz,x,y,z) ->
                SquareVector(dx * d, dy * d, dz * d, x * d, y * d, z * d)
            | CVectorResultant(x1,y1,z1,c1) ->
                CVectorResultant(x1 * d, y1 * d, z1 * d, c1 * d)
            | FullCVectorResultant(x,y,z,dx,dy,dz,dc) ->
                FullCVectorResultant(x * d, y * d, z * d, dx * d, dy * d, dz * d, dc * d)
        member self.Location : Countable =
            match self with
            | FullCVectorResultant(x,y,z,_,_,_,_) -> Vector(x,y,z)
            | _ -> failwith "undefined"
        member self.ScalarDivide(d: double) : Countable =
            match self with
            | Num n -> Num(n / d)
            | Vector(x,y,z) -> Vector(x / d, y / d, z / d)
            | SquareVector(dx,dy,dz,x,y,z) ->
                SquareVector(dx / d, dy / d, dz / d, x / d, y / d, z / d)
            | CVectorResultant(x1,y1,z1,c1) ->
                CVectorResultant(x1 / d, y1 / d, z1 / d, c1 / d)
            | FullCVectorResultant(x,y,z,dx,dy,dz,dc) ->
                FullCVectorResultant(x / d, y / d, z / d, dx / d, dy / d, dz / d, dc / d)
        // this simulates the effect of "fixing" a cell
        // copy the non-location part of the resultant from co to self
        member self.UpdateResultant(co: Countable) : Countable =
            match (self,co) with
            | FullCVectorResultant(x,y,z,dx,dy,dz,dc),CVectorResultant(dx2,dy2,dz2,dc2) -> FullCVectorResultant(x,y,z,dx2,dy2,dz2,dc2)
            | FullCVectorResultant(x,y,z,_,_,_,_),FullCVectorResultant(_,_,_,dx2,dy2,dz2,dc2) -> FullCVectorResultant(x,y,z,dx2,dy2,dz2,dc2)
            | _ -> failwith "Operation not supported."
        member self.VectorMultiply(co: Countable) : double =
            match (self,co) with
            | Num d1, Num d2 -> d1 * d2
            | Vector(x1,y1,z1), Vector(x2,y2,z2) -> x1 * x2 + y1 * y2 + z1 * z2
            | SquareVector(dx1,dy1,dz1,x1,y1,z1), SquareVector(dx2,dy2,dz2,x2,y2,z2) ->
                dx1 * dx2 + dy1 * dy2 + dz1 * dz2 + x1 * x2 + y1 * y2 + z1 * z2
            | CVectorResultant(x1,y1,z1,c1), CVectorResultant(x2,y2,z2,c2) ->
                x1 * x2 + y1 * y2 + z1 * z2 + c1 * c2
            | FullCVectorResultant(x1,y1,z1,dx1,dy1,dz1,dc1),FullCVectorResultant(x2,y2,z2,dx2,dy2,dz2,dc2) ->
                x1 * x2 + y1 * y2 + z1 * z2 + dx1 * dx2 + dy1 * dy2 + dz1 * dz2 + dc1 * dc2
            | _ -> failwith "Invalid operation.  Both Countables must be of the same type."
        member self.Sqrt : Countable =
            match self with
            | Num n -> Num(Math.Sqrt(n))
            | Vector(x,y,z) -> Vector(Math.Sqrt(x), Math.Sqrt(y), Math.Sqrt(z))
            | SquareVector(dx,dy,dz,x,y,z) ->
                SquareVector(Math.Sqrt(dx), Math.Sqrt(dy), Math.Sqrt(dz), Math.Sqrt(x), Math.Sqrt(y), Math.Sqrt(z))
            | CVectorResultant(x,y,z,c) ->
                CVectorResultant(Math.Sqrt(x), Math.Sqrt(y), Math.Sqrt(z), Math.Sqrt(c))
            | FullCVectorResultant(x,y,z,dx,dy,dz,dc) ->
                FullCVectorResultant(Math.Sqrt(x), Math.Sqrt(y), Math.Sqrt(z), Math.Sqrt(dx), Math.Sqrt(dy), Math.Sqrt(dz), Math.Sqrt(dc))
        member self.Abs : Countable =
            match self with
            | Num n -> Num(Math.Abs(n))
            | Vector(x,y,z) -> Vector(Math.Abs(x),Math.Abs(y),Math.Abs(z))
            | SquareVector(dx,dy,dz,x,y,z) ->
                SquareVector(Math.Abs(dx),Math.Abs(dy),Math.Abs(dz),Math.Abs(x),Math.Abs(y),Math.Abs(z))
            | CVectorResultant(x,y,z,c) ->
                CVectorResultant(Math.Abs(x), Math.Abs(y), Math.Abs(z), Math.Abs(c))
            | FullCVectorResultant(x,y,z,dx,dy,dz,dc) ->
                FullCVectorResultant(Math.Abs(x), Math.Abs(y), Math.Abs(z), Math.Abs(dx), Math.Abs(dy), Math.Abs(dz), Math.Abs(dc))
        member self.ToCVectorResultant : Countable =
            match self with
            | FullCVectorResultant(x,y,z,dx,dy,dz,dc) -> CVectorResultant(dx,dy,dz,dc)
            | _ -> failwith "Unsupported conversion."
        member self.ElementwiseMin(co: Countable) : Countable =
            self.ElementwiseOp co (fun x1 x2 -> if x1 < x2 then x1 else x2)
        member self.ElementwiseMax(co: Countable) : Countable =
            self.ElementwiseOp co (fun x1 x2 -> if x1 > x2 then x1 else x2)
        member self.ElementwiseDivide(co: Countable) : Countable =
            self.ElementwiseOp co (fun x1 x2 -> x1 / x2)
        // Hadamard product for two vectors
        member self.ElementwiseMultiply(co: Countable) : Countable =
            self.ElementwiseOp co (fun x1 x2 -> x1 * x2)
        member self.ElementwiseOp(co: Countable)(op: double -> double -> double) : Countable =
            match self,co with
            | Num n1,Num n2 -> Num(op n1 n2)
            | Vector(x1,y1,z1),Vector(x2,y2,z2) ->
                Vector(
                    op x1 x2,
                    op y1 y2,
                    op z1 z2
                )
            | SquareVector(dx1,dy1,dz1,x1,y1,z1),SquareVector(dx2,dy2,dz2,x2,y2,z2) ->
                SquareVector(
                    op dx1 dx2, 
                    op dy1 dy2, 
                    op dz1 dz2, 
                    op x1 x2,
                    op y1 y2,
                    op z1 z2
                )
            | CVectorResultant(x1,y1,z1,c1), CVectorResultant(x2,y2,z2,c2) ->
                CVectorResultant(
                    op x1 x2,
                    op y1 y2,
                    op z1 z2,
                    op c1 c2
                )
            | FullCVectorResultant(x1,y1,z1,dx1,dy1,dz1,dc1), FullCVectorResultant(x2,y2,z2,dx2,dy2,dz2,dc2) ->
                FullCVectorResultant(
                    op x1 x2,
                    op y1 y2,
                    op z1 z2,
                    op dx1 dx2,
                    op dy1 dy2,
                    op dz1 dz2,
                    op dc1 dc2
                )
            | _ -> failwith "Cannot do operation on vectors of different lengths."
        member self.toArray : double[] =
            match self with
            | Num n -> [| n |]
            | Vector(x1, x2, x3) -> [| x1; x2; x3 |]
            | SquareVector(x1, x2, x3, x4, x5, x6) -> [| x1; x2; x3; x4; x5; x6 |]
            | CVectorResultant(x1, x2, x3, x4) -> [| x1; x2; x3; x4 |]
            | FullCVectorResultant(x1, x2, x3, x4, x5, x6, x7) -> [| x1; x2; x3; x4; x5; x6; x7 |]
        member self.L2Norm : double =
            Math.Sqrt(
                Array.sumBy (fun x -> Math.Pow(x, 2.0)) (self.toArray)
            )
        member self.EuclideanDistance(co: Countable) : double =
            let diff = self.Sub co
            (diff.ElementwiseMultiply diff).ToVector
            |> Array.sum
            |> Math.Sqrt
        member self.IsOffSheet : bool =
            match self with
            | Num n -> false
            | Vector(x,y,z) -> z <> 0.0
            | SquareVector(dx,dy,dz,x,y,z) -> dz <> 0.0
            | CVectorResultant(x1,y1,z1,c1) -> z1 <> 0.0
            | FullCVectorResultant(x,y,z,dx,dy,dz,dc) -> dz <> 0.0
        member self.SameSheet(co: Countable) : bool =
            match (self,co) with
            | Num _, Num _ -> failwith "Cannot check same sheet for two numbers."
            | Vector(_,_,z1), Vector(_,_,z2) -> z1 = z2
            | SquareVector(_,_,_,_,_,z1), SquareVector(_,_,_,_,_,z2) -> z1 = z2
            | CVectorResultant(_,_,z1,_), CVectorResultant(_,_,z2,_) -> z1 = z2
            | FullCVectorResultant(_,_,z1,_,_,_,_), FullCVectorResultant(_,_,z2,_,_,_,_) -> z1 = z2
        member self.IsFormula : bool =
            // NOTE: this does not work for formulas that normally have no references, e.g =RAND()
            match self with
            | Num n -> failwith "unknowable"
            | Vector(x,y,z) -> failwith "unknowable"
            | SquareVector(dx,dy,dz,x,y,z) -> failwith "unknowable"
            | CVectorResultant(x1,y1,z1,c1) -> not (x1 = 0.0 && y1 = 0.0 && z1 = 0.0 && c1 = 0.0) // i.e., no refs at all
            // i.e., no refs at all
            // note that dc <= 0.0 is in keeping with the arbitrary choice to make strings dc = -1.0 and no constants dc = 0.0
            | FullCVectorResultant(x,y,z,dx,dy,dz,dc) -> not (dx = 0.0 && dy = 0.0 && dz = 0.0 && dc <= 0.0)
        static member Normalize(X: Countable[]) : Countable[] =
            let min = Array.fold (fun (a:Countable)(x: Countable) -> a.ElementwiseMin x) X.[0] X 
            let max = Array.fold (fun (a:Countable)(x: Countable) -> a.ElementwiseMax x) X.[0] X 
            let rng = max.Sub(min)
            let div = (fun x1 x2 -> if x1 = 0.0 && x2 = 0.0 then 0.0 else x1 / x2)
            Array.map (fun (x:Countable) ->
                let diff = x.Sub(min)
                let scaled = diff.ElementwiseOp rng div
                scaled
            ) X
        static member Mean(C: seq<Countable>) : Countable =
            let mutable i = 1.0
            let mutable sum = Seq.head C
            for c in (Seq.tail C) do
                sum <- sum.Add c
                i <- i + 1.0
            sum.ScalarDivide i

    type Capability = { enabled : bool; kind: ConfigKind; runner: AST.Address -> Depends.DAG -> Countable; }

    type FeatureLambda = AST.Address -> DAG -> Countable

    type BaseFeature() =
        static member run (cell: AST.Address) (dag: DAG): Countable = failwith "Feature must provide run method."
        static member capability : string*Capability = failwith "Feature must provide capability."

    // the following are for C# interop because
    // discriminated union types are not exported
    type FeatureUtil =
        static member makeNum(d: double) = Num d
        static member makeVector(x: double, y: double, z: double) = Vector(x,y,z)
        static member makeSpatialVector(dx: double, dy: double, dz: double, x: double, y: double, z: double) = SquareVector(dx,dy,dz,x,y,z)
        static member makeFullCVR(x: double, y: double, z: double, dx: double, dy: double, dz: double, dc: double) = FullCVectorResultant(x,y,z,dx,dy,dz,dc)