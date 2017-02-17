namespace Feature
  open System

  type SquareVector(dx: double, dy: double, dz: double, x: double, y: double, z: double) =
      member self.dx = dx
      member self.dy = dy
      member self.dz = dz
      member self.x = x
      member self.y = y
      member self.z = z
      override self.ToString() : string =
          "<" + dx.ToString() + "," +
                dy.ToString() + "," +
                dz.ToString() + "," +
                 x.ToString() + "," +
                 y.ToString() + "," +
                 z.ToString() +  ">"
      interface ICountable<SquareVector> with
          member self.Add(v: SquareVector) =
            SquareVector(dx + v.dx, dy + v.dy, dz + v.dz, x + v.x, y + v.y, z + v.z)
          member self.MeanFoldDefault = SquareVector(0.0,0.0,0.0,0.0,0.0,0.0)
          member self.Negate = SquareVector(-dx,-dy,-dz,-x,-y,-z)
          member self.ScalarDivide(div: double) =
            SquareVector(dx / div, dy / div, dz / div, x / div, y / div, z / div)
          member self.VectorMultiply(v: SquareVector) =
            dx * v.dx + dy * v.dy + dz * v.dz + x * v.x + y * v.y + z * v.z
          member self.Sqrt =
            SquareVector(Math.Sqrt(dx), Math.Sqrt(dy), Math.Sqrt(dz), Math.Sqrt(x), Math.Sqrt(y), Math.Sqrt(z))
          member self.Abs =
            SquareVector(Math.Abs(dx),Math.Abs(dy),Math.Abs(dz),Math.Abs(x),Math.Abs(y),Math.Abs(z))