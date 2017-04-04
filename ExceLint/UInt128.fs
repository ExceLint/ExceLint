namespace ExceLint
    module UInt128 =
        open System.Numerics

        type UInt128 =
            struct
               val High: uint64
               val Low: uint64
               new(high: uint64, low: uint64) = { High = high; Low = low }
               new(n: int) = { High = 0UL; Low = uint64 n }
            end

        let rec ToBigInteger(a: UInt128) : BigInteger =
            let alow = BigInteger(a.Low)
            let ahigh = BigInteger.op_LeftShift(BigInteger(a.High),64)
            BigInteger.op_BitwiseOr(ahigh, alow)

        and FromBigInteger(a: BigInteger) : UInt128 =
            let lmask = BigInteger.op_LeftShift(BigInteger.One, 64) - BigInteger.One
            let hmask = BigInteger.op_LeftShift(lmask, 64)
            let low = uint64 (BigInteger.op_BitwiseAnd(a, lmask))
            let highbi = BigInteger.op_RightShift(BigInteger.op_BitwiseAnd(a, hmask), 64)
            let high = uint64 highbi
            UInt128(high, low)

        and FromBinaryString(a: string) : UInt128 =
            let cs = a.ToCharArray() |> Array.rev

            seq { 0 .. 127 }
            |> Seq.toArray
            |> Array.fold (fun acc i ->
                    if i < cs.Length && cs.[i] = '1' then
                        BitwiseOr acc (LeftShift One i)
                    else
                        acc
               ) Zero

        and Zero = UInt128(0UL,0UL)
        and One = UInt128(0UL,1UL)
        and MaxValue = UInt128(System.UInt64.MaxValue, System.UInt64.MaxValue)

        and BitwiseNot(a: UInt128) : UInt128 =
            UInt128(~~~ a.High, ~~~ a.Low)

        and BitwiseOr(a: UInt128)(b: UInt128) : UInt128 =
            UInt128(a.High ||| b.High, a.Low ||| b.Low)

        and BitwiseAnd(a: UInt128)(b: UInt128) : UInt128 =
            UInt128(a.High &&& b.High, a.Low &&& b.Low)

        and BitwiseNand(a: UInt128)(b: UInt128) : UInt128 =
            BitwiseOr (BitwiseNot a) (BitwiseNot b)

        and BitwiseXor(a: UInt128)(b: UInt128) : UInt128 =
            UInt128(a.High ^^^ b.High, a.Low ^^^ b.Low)

        // from constant-time bit-counting algorithm here:
        // https://blogs.msdn.microsoft.com/jeuge/2005/06/08/bit-fiddling-3/
        and CountOnes32(a: uint32) : int =
            let c = a - ((a >>> 1) &&& uint32 0o33333333333) - ((a >>> 2) &&& uint32 0o11111111111)
            let sum = ((c + (c >>> 3)) &&& uint32 0o30707070707) % (uint32 63)
            int32 sum

        and CountOnes64(a: uint64) : int =
            let low = uint32 a
            let high = uint32 (a >>> 32)

            let ppa = prettyPrint64 a
            let pplow = prettyPrint32 low
            let pphigh = prettyPrint32 high

            CountOnes32 high + CountOnes32 low

        and CountOnes(a: UInt128) : int =
            CountOnes64 a.High + CountOnes64 a.Low

        and CountZeroes(a: UInt128) : int =
            128 - CountOnes a

        and longestCommonPrefix(a: UInt128)(b: UInt128) : int =
            let o = BitwiseOr a b
            let n = BitwiseNand a b
            let x = BitwiseXor o n

            CountOnes x

        and Add(a: UInt128)(b: UInt128) : UInt128 =
            let mutable carry = BitwiseAnd a b
            let mutable result = BitwiseXor a b
            while not (Equals carry Zero) do
                let shiftedcarry = LeftShift carry 1
                carry <- BitwiseAnd result shiftedcarry
                result <- BitwiseXor result shiftedcarry
            result

        and Sub(a: UInt128)(b: UInt128) : UInt128 =
            let mutable a' = a
            let mutable b' = b
            while not (Equals b' Zero) do
                let borrow = BitwiseAnd (BitwiseNot a') b'
                a' <- BitwiseXor a' b'
                b' <- LeftShift borrow 1
            a'

        and LeftShift(a: UInt128)(shf: int) : UInt128 =
            if shf > 127 then
                UInt128(0UL,0UL)
            else if shf > 63 then
                let ushf = shf - 64
                let hi = a.Low <<< ushf
                UInt128(hi,0UL)
            else
                // get uppermost shf bits in low
                let lmask = ((1UL <<< shf) - 1UL) <<< (64 - shf)
                let lup = lmask &&& a.Low
                // right shift lup to be shf low order bits of high dword
                let lupl = lup >>> (64 - shf)
                // compute new high
                let hi = (a.High <<< shf) ||| lupl
                let low = a.Low <<< shf
                UInt128(hi, low)

        and RightShift(a: UInt128)(shf: int) : UInt128 =
            if shf > 127 then 
                UInt128(0UL,0UL)
            else if shf > 63 then
                let lshf = shf - 64
                let low = a.High >>> lshf
                UInt128(0UL,low)
            else
                // get lowermost shf bits in high
                let hmask = (1UL <<< shf) - 1UL
                let hlow = hmask &&& a.High
                // left shift hlow to be shf high order bits of low dword
                let luph = hlow <<< (64 - shf)
                // compute new low & high
                let low = (a.Low >>> shf) ||| luph
                let hi = a.High >>> shf
                UInt128(hi, low)

        and GreaterThan(a: UInt128)(b: UInt128) : bool =
            a.High > b.High || (a.High = b.High && a.Low > b.Low)

        and Equals(a: UInt128)(b: UInt128) : bool =
            a.High = b.High && a.Low = b.Low

        and Divide(a: UInt128)(b: UInt128) : UInt128 =
            let a' = ToBigInteger(a)
            let b' = ToBigInteger(b)
            let c = BigInteger.Divide(a', b')
            FromBigInteger(c)

        and Modulus(a: UInt128)(b: UInt128) : UInt128 =
            let a' = ToBigInteger(a)
            let b' = ToBigInteger(b)
            let c = BigInteger.op_Modulus(a', b')
            FromBigInteger(c)

        and prettyPrint64(v: uint64) : string =
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

        and prettyPrint32(v: uint32) : string =
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

        and prettyPrint(v: UInt128) : string =
            let mutable num = v
            let stack : UInt128[] = Array.create 128 Zero

            let zero = Zero
            let one = One
            let two = UInt128(0UL,2UL)

            let mutable i = 127

            while GreaterThan num zero do
                let rem = Modulus num two
                stack.[i] <- rem
                num <- Divide num two
                i <- i - 1

            Array.fold (fun (a: string)(e: UInt128) ->
                a + (if Equals one e then "1" else "0")
            ) "" stack

