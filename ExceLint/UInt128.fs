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

        and Zero = UInt128(0UL,0UL)
        and One = UInt128(0UL,1UL)

        and BitwiseOr(a: UInt128)(b: UInt128) : UInt128 =
            UInt128(a.High ||| b.High, a.Low ||| b.Low)

        and BitwiseAnd(a: UInt128)(b: UInt128) : UInt128 =
            UInt128(a.High &&& b.High, a.Low &&& b.Low)

        and LeftShift(a: UInt128)(shf: int) : UInt128 =
            if shf > 128 then
                UInt128(0UL,0UL)
            else if shf > 64 then
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

        and prettyPrint(v: UInt128) : string =
            let mutable num = v
            let mutable stack = []

            let zero = Zero
            let one = One
            let two = UInt128(0UL,2UL)

            while GreaterThan num zero do
                let rem = Modulus num two
                stack <- rem :: stack
                num <- Divide num two

            List.fold (fun (a: string)(e: UInt128) ->
                a + (if Equals one e then "1" else "0")
            ) "" stack

