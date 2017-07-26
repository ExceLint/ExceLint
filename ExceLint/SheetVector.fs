namespace ExceLint
    open System
    open System.Numerics

    type SheetVector(ci: BigInteger option, topleft: int*int, bottomright: int*int) =
        // topleft and bottomright are one-based
        let _width = fst(bottomright) - fst(topleft) + 1
        let _height = snd(bottomright) - snd(topleft) + 1
        let _len = _width * _height
        let mutable _cells = match ci with
                             | None -> BigInteger.op_LeftShift(BigInteger.One, _len) - BigInteger.One
                             | Some(ci) -> ci
        member self.SetAll() =
            _cells <- BigInteger.op_LeftShift(BigInteger.One, _len) - BigInteger.One
        member self.UnsetAll() =
            let and_mask = BigInteger.Zero
            _cells <- BigInteger.op_BitwiseAnd(_cells, and_mask)
        member self.Width = _width
        // the internal, zero-based index for a set of coordinates
        static member Index(xy: int*int, width: int, debug_from: string) : int = SheetVector.Index(xy, width)
        static member Index(xy: int*int, width: int) : int =
            let x = (fst xy)
            let y = (snd xy)
            // this is the location of the bit in the bitmap
            let xyint = (y - 1) * width + x
            // subtract 1 from xyint since bitmap is zero-based
            xyint - 1
        member self.Coords(i: int) =
            let i' = i + 1
            let y = int(Math.Ceiling(float(i')/float(_width)))
            let x = i' - ((y-1) * _width)
            x,y
        // sets a bit located at coordinates x,y
        member self.Set(xy: int*int) =
            // this is the OR mask for that bit
            let or_mask = BigInteger.op_LeftShift(BigInteger.One, SheetVector.Index(xy, self.Width))
            _cells <- BigInteger.op_BitwiseOr(_cells, or_mask)
        member self.Unset(xy: int*int) =
            // this is the AND mask for that bit
            let and_mask = BigInteger.op_OnesComplement(BigInteger.op_LeftShift(BigInteger.One, SheetVector.Index(xy, self.Width)))
            _cells <- BigInteger.op_BitwiseAnd(_cells, and_mask)
        member self.Copy : SheetVector =
            SheetVector(Some(_cells), topleft, bottomright)
        member self.BitwiseAnd(c: SheetVector) : SheetVector = SheetVector(Some(BigInteger.op_BitwiseAnd(_cells, c.Bitmap)), topleft, bottomright)
        member self.BitwiseOr(c: SheetVector) : SheetVector = SheetVector(Some(BigInteger.op_BitwiseOr(_cells, c.Bitmap)), topleft, bottomright)
        member self.BitwiseXOr(c: SheetVector) : SheetVector = SheetVector(Some(BigInteger.op_ExclusiveOr(_cells, c.Bitmap)), topleft, bottomright)
        member self.Negate() : SheetVector =
            let n = BigInteger.op_OnesComplement(_cells)
            SheetVector(Some(n), topleft, bottomright)
        member self.TopLeft = topleft
        member self.BottomRight = bottomright
        member self.Bitmap = _cells
        override self.ToString() =
            String.Join("; ", Array.map (fun bitidx -> self.Coords(bitidx).ToString()) (self.whichBits))

        // based on Kernighan's bit-counting algorithm
        // tells you the indices of the bits in the bitvector
        // O(# of bits set)
        //
        // unsigned int v; // count the number of bits set in v
        // unsigned int c; // c accumulates the total bits set in v
        // for (c = 0; v; c++)
        // {
        //   v &= v - 1; // clear the least significant bit set
        // }
        member self.whichBits : int[] =
            let rec getBits(v: BigInteger)(xs: int list) : int list =
                if v = BigInteger.Zero then
                    xs
                else
                    let v_new = BigInteger.op_BitwiseAnd(v, v - BigInteger.One)
                    let v_minus_v_new = v - v_new
                    let idx = SheetVector.ilog2(v_minus_v_new.ToByteArray())
                    getBits v_new (idx :: xs)
            getBits _cells [] |> Array.ofList
        member self.countBits : int =
            let rec cBits(v: BigInteger)(i: int) : int =
                if v = BigInteger.Zero then
                    i
                else
                    let v_new = BigInteger.op_BitwiseAnd(v, v - BigInteger.One)
                    let v_minus_v_new = v - v_new
                    let idx = SheetVector.ilog2(v_minus_v_new.ToByteArray())
                    cBits v_new (i + 1)
            cBits _cells 0

        member self.SameAs(c: SheetVector) : bool = BigInteger.op_BitwiseAnd(c.Bitmap, _cells) = c.Bitmap
        member self.Length = _len
        member self.IsSet(xy: int*int) : bool =
            let and_mask = BigInteger.op_LeftShift(BigInteger.One, SheetVector.Index(xy, self.Width))
            BigInteger.op_BitwiseAnd(and_mask, _cells).Equals(and_mask)
        member self.EmptyIntersect(c: SheetVector) : bool = BigInteger.op_BitwiseAnd(c.Bitmap, _cells) = BigInteger.Zero
        member self.BitwiseEquals(c: SheetVector) : bool = _cells.Equals(c.Bitmap)
        override self.Equals(obj: obj) : bool = _cells.Equals((obj :?> SheetVector).Bitmap)
        override self.GetHashCode() : int = System.Convert.ToInt32(_cells)
        static member Empty(topleft: int*int, bottomright: int*int) = SheetVector(Some(BigInteger.Zero), topleft, bottomright)
        static member OneBit(xy: int*int, topleft: int*int, bottomright: int*int) =
            let width = fst(bottomright) - fst(topleft) + 1
            let bi = BigInteger.op_LeftShift(BigInteger.One, SheetVector.Index(xy, width))
            SheetVector(Some(bi), topleft, bottomright)
        static member ilog2(bytes: byte[]) : int =
            if bytes.[bytes.Length - 1] >= 128uy then
                -1  // -ve bigint (invalid - cannot take log of -ve number)
            else
                let mutable log = 0
                while (bytes.[bytes.Length - 1] >>> log) > 0uy do
                    log <- log + 1
                log + bytes.Length * 8 - 9