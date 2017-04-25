namespace ExceLint
    open System.Numerics

    module UInt128Ops =
        // from constant-time bit-counting algorithm here:
        // https://blogs.msdn.microsoft.com/jeuge/2005/06/08/bit-fiddling-3/
        let CountOnes32(a: uint32) : int =
            let c = a - ((a >>> 1) &&& uint32 0o33333333333) - ((a >>> 2) &&& uint32 0o11111111111)
            let sum = ((c + (c >>> 3)) &&& uint32 0o30707070707) % (uint32 63)
            int32 sum
        let CountOnes64(a: uint64) : int =
            let low = uint32 a
            let high = uint32 (a >>> 32)
            CountOnes32 high + CountOnes32 low
        let prettyPrint64(v: uint64) : string =
            let mutable num = v
            let mutable stack = []

            let zero = 0UL
            let one = 1UL
            let two = 2UL

            while num > zero do
                let rem = num % two
                stack <- rem :: stack
                num <- num / two

            List.fold (fun (a: string)(e: uint64) ->
                a + (if one = e then "1" else "0")
            ) "" stack
        let prettyPrint32(v: uint32) : string =
            let mutable num = v
            let mutable stack = []

            let zero = 0ul
            let one = 1ul
            let two = 2ul

            while num > zero do
                let rem = num % two
                stack <- rem :: stack
                num <- num / two

            List.fold (fun (a: string)(e: uint32) ->
                a + (if one = e then "1" else "0")
            ) "" stack

        type BinarySearchOutcome =
        | SearchLow
        | SearchHigh
        | KeyFound

        type GenBinarySearchOutcome =
        | Found of int
        | Last of int

        let midpoint(min: int)(max: int) : int = 
            min + int32 (System.Math.Ceiling((double (System.Math.Abs(max - min))) / 2.0))

        let GeneralizedBinarySearch(low: int)(high: int)(testFn: int -> BinarySearchOutcome)(saveWhen: BinarySearchOutcome) : GenBinarySearchOutcome =
            let mutable min = low
            let mutable max = high
            let mutable last = -1
            let mutable found = -1

            let mutable test = None

            // continue searching while [imin,imax] is not empty
            while max >= min && test.IsNone do
                let mid = midpoint min max

                test <- match testFn mid with
                           | BinarySearchOutcome.KeyFound ->
                                // we found it; return the index
                                Some (GenBinarySearchOutcome.Found mid)
                           | BinarySearchOutcome.SearchLow ->
                                // change max index to search lower subarray
                                max <- mid - 1
                                // if we're looking for a high bound, save whenever we search low
                                if saveWhen = BinarySearchOutcome.SearchLow then
                                    found <- mid
                                last <- mid
                                None
                           | BinarySearchOutcome.SearchHigh ->
                                // change min index to search upper subarray
                                min <- mid + 1
                                // if we're looking for a low bound, save whenever we search high
                                if saveWhen = BinarySearchOutcome.SearchHigh then
                                    found <- mid
                                last <- mid
                                None

            match test with
            | Some(f) -> f
            | None ->
                if found <> -1 then
                    GenBinarySearchOutcome.Found found
                else
                    GenBinarySearchOutcome.Last last

    [<CustomEquality; CustomComparison>]
    type UInt128 =
        struct
            val High: uint64
            val Low: uint64
            new(high: uint64, low: uint64) = { High = high; Low = low }
            new(n: int) = { High = 0UL; Low = uint64 n }

            static member Zero = UInt128(0UL,0UL)
            static member One = UInt128(0UL,1UL)
            static member MaxValue = UInt128(System.UInt64.MaxValue, System.UInt64.MaxValue)
            static member FromBigInteger(a: BigInteger) : UInt128 =
                let lmask = BigInteger.op_LeftShift(BigInteger.One, 64) - BigInteger.One
                let hmask = BigInteger.op_LeftShift(lmask, 64)
                let low = uint64 (BigInteger.op_BitwiseAnd(a, lmask))
                let highbi = BigInteger.op_RightShift(BigInteger.op_BitwiseAnd(a, hmask), 64)
                let high = uint64 highbi
                UInt128(high, low)
            static member FromBinaryString(a: string) : UInt128 =
                let cs = a.ToCharArray() |> Array.rev

                seq { 0 .. 127 }
                |> Seq.toArray
                |> Array.fold (fun acc i ->
                        if i < cs.Length && cs.[i] = '1' then
                            acc.BitwiseOr (UInt128.One.LeftShift i)
                        else
                            acc
                    ) UInt128.Zero
            // this function lets you supply a binary prefix string
            // which is then zero-filled to yield a UInt128
            static member FromZeroFilledPrefix(prefix: string) : UInt128 =
                assert (prefix.Length <= 128)
                let b = UInt128.FromBinaryString prefix
                let lshft = 128 - prefix.Length
                b.LeftShift lshft
            static member nBitMask(n: int) : UInt128 =
                (UInt128.One.LeftShift n).Sub UInt128.One
            /// <summary>
            /// Calculate a UInt128 bitmask starting at startpos (inclusive) and ending at endpos (inclusive).
            /// </summary>
            static member calcMask(startpos: int)(endpos: int) : UInt128 =
                let numbits = endpos - startpos + 1
                (UInt128.nBitMask numbits).LeftShift (128 - numbits - startpos)

            member self.ToBigInteger : BigInteger =
                let l = BigInteger(self.Low)
                let h = BigInteger.op_LeftShift(BigInteger(self.High),64)
                BigInteger.op_BitwiseOr(h, l)
            member self.BitwiseNot : UInt128 =
                UInt128(~~~ self.High, ~~~ self.Low)
            member self.BitwiseOr(b: UInt128) : UInt128 =
                UInt128(self.High ||| b.High, self.Low ||| b.Low)
            member self.BitwiseAnd(b: UInt128) : UInt128 =
                UInt128(self.High &&& b.High, self.Low &&& b.Low)
            member self.BitwiseNand(b: UInt128) : UInt128 =
                self.BitwiseNot.BitwiseOr b.BitwiseNot
            member self.BitwiseXor(b: UInt128) : UInt128 =
                UInt128(self.High ^^^ b.High, self.Low ^^^ b.Low)
            member self.Add(b: UInt128) : UInt128 =
                let mutable carry = self.BitwiseAnd b
                let mutable result = self.BitwiseXor b
                while not (carry = UInt128.Zero) do
                    let shiftedcarry = carry.LeftShift 1
                    carry <- result.BitwiseAnd shiftedcarry
                    result <- result.BitwiseXor shiftedcarry
                result
            member self.Sub(b: UInt128) : UInt128 =
                let mutable a' = self
                let mutable b' = b
                while not (b' = UInt128.Zero) do
                    let borrow = a'.BitwiseNot.BitwiseAnd b'
                    a' <- a'.BitwiseXor b'
                    b' <- borrow.LeftShift 1
                a'
            member self.LeftShift(shf: int) : UInt128 =
                if shf > 127 then
                    UInt128(0UL,0UL)
                else if shf > 63 then
                    let ushf = shf - 64
                    let hi = self.Low <<< ushf
                    UInt128(hi,0UL)
                else
                    // get uppermost shf bits in low
                    let lmask = ((1UL <<< shf) - 1UL) <<< (64 - shf)
                    let lup = lmask &&& self.Low
                    // right shift lup to be shf low order bits of high dword
                    let lupl = lup >>> (64 - shf)
                    // compute new high
                    let hi = (self.High <<< shf) ||| lupl
                    let low = self.Low <<< shf
                    UInt128(hi, low)
            member self.RightShift(shf: int) : UInt128 =
                if shf > 127 then 
                    UInt128(0UL,0UL)
                else if shf > 63 then
                    let lshf = shf - 64
                    let low = self.High >>> lshf
                    UInt128(0UL,low)
                else
                    // get lowermost shf bits in high
                    let hmask = (1UL <<< shf) - 1UL
                    let hlow = hmask &&& self.High
                    // left shift hlow to be shf high order bits of low dword
                    let luph = hlow <<< (64 - shf)
                    // compute new low & high
                    let low = (self.Low >>> shf) ||| luph
                    let hi = self.High >>> shf
                    UInt128(hi, low)
            member self.GreaterThan(b: UInt128) : bool =
                self.High > b.High || (self.High = b.High && self.Low > b.Low)
            member self.Divide(b: UInt128) : UInt128 =
                let a' = self.ToBigInteger
                let b' = b.ToBigInteger
                let c = BigInteger.Divide(a', b')
                UInt128.FromBigInteger(c)
            member self.Modulus(b: UInt128) : UInt128 =
                let a' = self.ToBigInteger
                let b' = b.ToBigInteger
                let c = BigInteger.op_Modulus(a', b')
                UInt128.FromBigInteger(c)
            member self.CountOnes : int =
                UInt128Ops.CountOnes64 self.High + UInt128Ops.CountOnes64 self.Low
            member self.CountZeroes : int =
                128 - self.CountOnes
            /// <summary>
            /// Find the longest common prefix (starting from high bits)
            /// and return its length.
            /// </summary>
            /// <param name="b">
            /// Another UInt128 to compare against this UInt128.
            /// </param>
            member self.LongestCommonPrefix(b: UInt128) : int =
                let a = self
                let testFn = (fun i ->
                                let mask = UInt128.calcMask 0 i
                                let a' = a.BitwiseAnd mask
                                let b' = b.BitwiseAnd mask
                                if a' = b' then
                                    UInt128Ops.BinarySearchOutcome.SearchHigh
                                else
                                    UInt128Ops.BinarySearchOutcome.SearchLow
                             )
                match UInt128Ops.GeneralizedBinarySearch 0 127 testFn UInt128Ops.BinarySearchOutcome.SearchLow with
                | UInt128Ops.GenBinarySearchOutcome.Found i -> i
                | UInt128Ops.GenBinarySearchOutcome.Last i -> i

            override self.GetHashCode() : int = int32 self.Low
            override self.Equals(o: obj) : bool =
                match o with
                | :? UInt128 as other ->
                    self.High = other.High &&
                    self.Low = other.Low
                | _ -> invalidArg "o" "cannot compare values of different types"
            override self.ToString() : string =
                let mutable num = self
                let stack : UInt128[] = Array.create 128 UInt128.Zero

                let zero = UInt128.Zero
                let one = UInt128.One
                let two = UInt128(0UL,2UL)

                let mutable i = 127

                while num.GreaterThan zero do
                    let rem = num.Modulus two
                    stack.[i] <- rem
                    num <- num.Divide two
                    i <- i - 1

                Array.fold (fun (a: string)(e: UInt128) ->
                    a + (if one = e then "1" else "0")
                ) "" stack
            member self.MaskedBitsAsString(mask: UInt128) : string =
                let maskchars = mask.ToString().ToCharArray()
                let bvchars = self.ToString().ToCharArray()

                let mutable outchars = []
                for i in 0..maskchars.Length - 1 do
                    if maskchars.[i] = '1' then
                        outchars <- bvchars.[i] :: outchars

                let s = System.String.Join("", List.rev outchars)

                // compress long runs of 0s and 1s
                let r0 = System.Text.RegularExpressions.Regex("(00000*)")
                let r1 = System.Text.RegularExpressions.Regex("(11111*)")
                let mutable s' = s
                while r0.IsMatch(s') do
                    let m = r0.Match(s').Value
                    s' <- r0.Replace(s', "0[x" + m.Length.ToString() + "]")
                while r1.IsMatch(s') do
                    let m = r1.Match(s').Value
                    s' <- r1.Replace(s', "1[x" + m.Length.ToString() + "]")
                s'
            interface System.IComparable with
                member self.CompareTo(o: obj) =
                    match o with
                    | :? UInt128 as other ->
                        if self.High > other.High then
                            1
                        else if self.High = other.High then
                            if self.Low > other.Low then
                                1
                            else if self.Low = other.Low then
                                0
                            else
                                -1
                        else
                            -1
                    | _ -> invalidArg "o" "cannot compare values of different types"
        end