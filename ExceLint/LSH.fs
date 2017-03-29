namespace ExceLint
    module private Calc =
        let XYBITS = 40
        let ZBITS = 64 - XYBITS

        let hash(x: uint64)(y: uint64)(z: uint64) : uint64 =
            // interleave x and y in lower XYBITS bits; z in upper ZBITS
            let xyz = Seq.fold (fun vect i ->
                        if i < ZBITS then
                            let xm = (x &&& (1UL <<< i))
                            let ym = (y &&& (1UL <<< i))
                            let vect'  = vect  ||| (xm <<< i)
                            let vect'' = vect' ||| (ym <<< i)
                            vect''
                        else
                            let zm = (z &&& (1UL <<< i - XYBITS))
                            let vect' = vect ||| (zm <<< i)
                            vect'
                      ) (0UL) (seq { 0 .. (XYBITS + ZBITS) - 1 })
            xyz

        let addr2Hash(a: AST.Address)(dag: Depends.DAG) : uint64 =
            let x = System.Convert.ToUInt64(a.X);
            let y = System.Convert.ToUInt64(a.Y);
            let z = System.Convert.ToUInt64(dag.getPathClosureIndex(a.Path,a.WorkbookName,a.WorksheetName))
            hash x y z
            
    type LSH(cells: seq<AST.Address>) = class end