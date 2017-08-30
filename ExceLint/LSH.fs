namespace ExceLint

    open System.Collections.Generic
    open System.Numerics
    open Utils

    module LSHCalc =
        let private XBITS = 20
        let private ZBITS = 64 - (2 * XBITS)
        let private BXBITS = 20
        let private BZBITS = 10

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

        let bitAt(v: UInt128)(idx: int) : UInt128 =
            v.BitwiseAnd (UInt128.One.LeftShift idx)

        let orShiftMask(t: UInt128)(s: UInt128)(shf: int) : UInt128 =
            t.BitwiseOr (s.LeftShift shf)

        let h7(co: Countable) : UInt128 =
            let (x,y,z,dx,dy,dz,dc) =
                match co with
                | FullCVectorResultant(x,y,z,x',y',z',c) -> x,y,z,x',y',z',c
                | _ -> failwith "wrong vector type"
            // quantize
            let bx  = UInt128(int x)
            let by  = UInt128(int y)
            let bz  = UInt128(int z)
            let bdx = UInt128(int dx)
            let bdy = UInt128(int dy)
            let bdz = UInt128(int dz)
            let bdc = UInt128(int dc)

            // interleave low order bits
            let low = Seq.fold (fun v0 i ->
                                   let xm =  (bitAt bx  i).RightShift i
                                   let ym =  (bitAt by  i).RightShift i
                                   let dxm = (bitAt bdx i).RightShift i
                                   let dym = (bitAt bdy i).RightShift i
                                   let dcm = (bitAt bdc i).RightShift i

                                   let j = i * 5
                                   let v1 = orShiftMask v0 xm   j
                                   let v2 = orShiftMask v1 ym  (j + 1)
                                   let v3 = orShiftMask v2 dxm (j + 2)
                                   let v4 = orShiftMask v3 dym (j + 3)
                                   let v5 = orShiftMask v4 dcm (j + 4)

                                   v5
                               ) (UInt128.Zero) (seq { 0 .. BXBITS - 1 })

            // interleave high order bits
            let highu = Seq.fold (fun v0 i ->
                                    let zm  = bitAt bz  i
                                    let zm' = bitAt bdz i

                                    let v1 = orShiftMask v0 zm   i
                                    let v2 = orShiftMask v1 zm' (i + 1)

                                    v2
                                 ) (UInt128.Zero) (seq { 0 .. BZBITS - 1})

            // shift high
            let high = highu.LeftShift (5 * BXBITS)

            // OR low and high
            high.BitwiseOr low

        let h7unmasker(mask: UInt128) : UInt128 =
            // count the number of ones in current mask
            let numones = mask.CountOnes

            if numones < 5 * BXBITS then
                // unmask 5 more bits
                UInt128.calcMask 0 (numones - 5)
            else
                // unmask 2 more bits
                UInt128.calcMask 0 (numones - 2)