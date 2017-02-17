namespace Feature
  open System

  type Vector(x: double, y: double, z: double) =
      member self.x = x
      member self.y = y
      member self.z = z
      override self.ToString() : string =
          "<" + x.ToString() + "," + y.ToString() + "," + z.ToString() + ">"
      interface ICountable<Vector> with
          member self.Add(v: Vector) = Vector(x + v.x, y + v.y, z + v.z)
          member self.MeanFoldDefault = Vector(0.0,0.0,0.0)
          member self.Negate = Vector(-x, -y, -z)
          member self.ScalarDivide(div: double) = Vector(x / div, y / div, z / div)
          member self.VectorMultiply(v: Vector) = x * v.x + y * v.y + z * v.z
          member self.Sqrt = Vector(Math.Sqrt(x), Math.Sqrt(y), Math.Sqrt(z))
          member self.Abs = Vector(Math.Abs(x),Math.Abs(y),Math.Abs(z))