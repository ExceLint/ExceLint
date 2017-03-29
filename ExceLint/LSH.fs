namespace ExceLint
    module LSHCalc =
        let XBITS = 20
        let ZBITS = 64 - (2 * XBITS)

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
            
    type LSH(cells: seq<AST.Address>) = class end