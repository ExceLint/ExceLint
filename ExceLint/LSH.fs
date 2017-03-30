namespace ExceLint

    open System.Numerics

    module LSHCalc =
        let XBITS = 20
        let ZBITS = 64 - (2 * XBITS)

        let BXBITS = 20
        let BZBITS = 10

        let hash(x: uint64)(y: uint64)(z: uint64) : uint64 =
            // interleave x and y in lower XYBITS bits; z in upper ZBITS
            // [... ZBITS ... | ... 2 x XBITS ... ]

            // x,y region of bitvector
            let xyreg = Seq.fold (fun vect i ->
                            let xm = x &&& (1UL <<< i)
                            let ym = y &&& (1UL <<< i)
                            let vect'  = vect  ||| (xm <<< i)
                            let vect'' = vect' ||| (ym <<< i + 1)
                            vect''
                          ) (0UL) (seq { 0 .. XBITS - 1 })

            // z region of bitvector
            let zm = z &&& ((1UL <<< ZBITS) - 1UL)
            let zreg = zm <<< 2 * XBITS

            // combine
            let bv = xyreg ||| zreg
            bv

        let hashi(x: int)(y: int)(z: int) : uint64 =
            hash (uint64 x) (uint64 y) (uint64 z)

        let addr2Hash(a: AST.Address)(dag: Depends.DAG) : uint64 =
            let z = dag.getPathClosureIndex(a.Path,a.WorkbookName,a.WorksheetName)
            hashi a.X a.Y z

        let bimask(v: BigInteger)(len: int) : BigInteger =
            BigInteger.op_BitwiseAnd(v,BigInteger.op_LeftShift(BigInteger.One, len))

        let orShiftMask(t: BigInteger)(s: BigInteger)(shf: int) : BigInteger =
            BigInteger.op_BitwiseOr(t, BigInteger.op_LeftShift(s, shf))

        (*
            def divideBy2(decNumber):
                remstack = Stack()

                while decNumber > 0:
                    rem = decNumber % 2
                    remstack.push(rem)
                    decNumber = decNumber // 2

                binString = ""
                while not remstack.isEmpty():
                    binString = binString + str(remstack.pop())

                return binString
        *)

        let ppbi(v: BigInteger) : string =
            let mutable num = v
            let mutable stack = []

            let one = BigInteger.One
            let two = BigInteger(2)

            while BigInteger.op_GreaterThan(num, BigInteger.Zero) do
                let rem = BigInteger.op_Modulus(num, two)
                stack <- rem :: stack
                num <- BigInteger.Divide(num, two)

            List.fold (fun (a: string)(e: BigInteger) ->
                a + (if BigInteger.op_Equality(one,e) then "1" else "0")
            ) "" stack

        let h7(co: Countable) : BigInteger =
            let (x,y,z,x',y',z',c) =
                match co with
                | FullCVectorResultant(x,y,z,x',y',z',c) -> x,y,z,x',y',z',c
                | _ -> failwith "wrong vector type"
            // quantize
            let bx  = BigInteger(int x)
            let by  = BigInteger(int y)
            let bz  = BigInteger(int z)
            let bx' = BigInteger(int x')
            let by' = BigInteger(int y')
            let bz' = BigInteger(int z')
            let bc  = BigInteger(int c)

            // interleave low order bits
            let low = Seq.fold (fun v0 i ->
                                   let xm =  bimask bx  i
                                   let ym =  bimask by  i
                                   let xm' = bimask bx' i
                                   let ym' = bimask by' i
                                   let cm =  bimask bc  i

                                   let v1 = orShiftMask v0 xm   i
                                   let v2 = orShiftMask v1 ym  (i + 1)
                                   let v3 = orShiftMask v2 xm' (i + 2)
                                   let v4 = orShiftMask v3 ym' (i + 3)
                                   let v5 = orShiftMask v4 cm  (i + 4)

                                   v5
                               ) (BigInteger.Zero) (seq { 0 .. BXBITS - 1 })

            // interleave high order bits
            let highu = Seq.fold (fun v0 i ->
                                    let zm  = bimask bz  i
                                    let zm' = bimask bz' i

                                    let v1 = orShiftMask v0 zm   i
                                    let v2 = orShiftMask v1 zm' (i + 1)

                                    v2
                                ) (BigInteger.Zero) (seq { 0 .. BZBITS - 1})

            // shift high
            let high = BigInteger.op_LeftShift(highu, 5 * BXBITS)

            // OR low and high
            BigInteger.op_BitwiseOr(high, low)
            
    type LSH(cells: seq<AST.Address>) = class end