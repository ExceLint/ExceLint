namespace Feature
  open System

  type Num(d: double) =
    member self.d = d
    override self.ToString() : string = d.ToString()
    interface ICountable<Num> with
      member self.Add(n: Num) = Num(d + n.d)
      member self.MeanFoldDefault = Num(0.0)
      member self.Negate = Num(-d)
      member self.ScalarDivide(div: double) = Num(d / div)
      member self.VectorMultiply(n: Num) = d * n.d
      member self.Sqrt = Num(Math.Sqrt(d))
      member self.Abs = Num(Math.Abs(d))