namespace ExceLint
open System.Collections.Generic
open System.Collections.Immutable
    module Utils =
        // we're using C# Dictionary instead of F# map
        // for debugging (it's inspectable) and speed purposes
        type Dict<'a,'b> = Dictionary<'a,'b>
        type ImmutDict<'a,'b> = ImmutableDictionary<'a,'b>

        let adict(a: seq<('a*'b)>) = new Dict<'a,'b>(a |> dict)
        let argwhatever(f: 'a -> double)(xs: seq<'a>)(whatev: double -> double -> bool) : 'a =
                Seq.reduce (fun arg x -> if whatev (f arg) (f x) then arg else x) xs
        let argmax(f: 'a -> double)(xs: seq<'a>) : 'a =
            argwhatever f xs (fun a b -> a > b)
        let argmin(f: 'a -> double)(xs: seq<'a>) : 'a =
            argwhatever f xs (fun a b -> a < b)

        let argindexmin(f: 'a -> int)(xs: 'a[]) : int =
            let fx = Array.map f xs

            Array.mapi (fun i res -> (i, res)) fx |>
            Array.fold (fun arg (i, res) ->
                if arg = -1 || res < fx.[arg] then
                    i
                else
                    arg
            ) -1 

        let HsUnion<'a>(h1: HashSet<'a>)(h2: HashSet<'a>) : HashSet<'a> =
            let h3 = new HashSet<'a>(h1)
            h3.UnionWith h2
            h3

        let isSane(addr: AST.Address) : bool =
            addr.Col > 0 && addr.Row > 0

        let Above(addr: AST.Address) : AST.Address =
            AST.Address.fromR1C1withMode(
                addr.Row - 1,
                addr.Col,
                addr.RowMode,
                addr.ColMode,
                addr.WorksheetName,
                addr.WorkbookName,
                addr.Path)

        let Below(addr: AST.Address) : AST.Address =
            AST.Address.fromR1C1withMode(
                addr.Row + 1,
                addr.Col,
                addr.RowMode,
                addr.ColMode,
                addr.WorksheetName,
                addr.WorkbookName,
                addr.Path);

        let Left(addr: AST.Address) : AST.Address =
            AST.Address.fromR1C1withMode(
                addr.Row,
                addr.Col - 1,
                addr.RowMode,
                addr.ColMode,
                addr.WorksheetName,
                addr.WorkbookName,
                addr.Path);

        let Right(addr: AST.Address) : AST.Address =
            AST.Address.fromR1C1withMode(
                addr.Row,
                addr.Col + 1,
                addr.RowMode,
                addr.ColMode,
                addr.WorksheetName,
                addr.WorkbookName,
                addr.Path);

        let AdjacentCells(addr: AST.Address) : HashSet<AST.Address> =
            let n  = Above(addr);
            let ne = Above(Right(addr));
            let e  = Right(addr);
            let se = Below(Right(addr));
            let s  = Below(addr);
            let sw = Below(Left(addr));
            let w  = Left(addr);
            let nw = Above(Left(addr));

            let addrs = [| n; ne; e; se; s; sw; w; nw |]

            new HashSet<AST.Address>(addrs |> Array.filter isSane)

        let isAdjacent(addr1: AST.Address)(addr2: AST.Address) : bool =
            Seq.contains addr2 (AdjacentCells addr1)

        let Adjust(epsilon: int)(i: int) : int =
            if (i + epsilon) > 0 then i + epsilon else i

        /// <summary>returns the bounding range for the given set of
        /// cells, with an additional (positive-valued) epsilon</summary>
        let BoundingRegion(cells: seq<AST.Address>)(epsilon: int) : AST.Address * AST.Address =
            assert (Seq.length cells > 0)
            assert (epsilon >= 0)
            
            let fcell = Seq.head cells
            let wsname = fcell.WorksheetName
            let wbname = fcell.WorkbookName
            let path = fcell.Path

            let xmin = cells |> Seq.map (fun addr -> addr.X) |> Seq.min |> Adjust -epsilon
            let xmax = cells |> Seq.map (fun addr -> addr.X) |> Seq.max |> Adjust epsilon
            let ymin = cells |> Seq.map (fun addr -> addr.Y) |> Seq.min |> Adjust -epsilon
            let ymax = cells |> Seq.map (fun addr -> addr.Y) |> Seq.max |> Adjust epsilon

            let lefttop = AST.Address.fromR1C1withMode(ymin, xmin, AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path)
            let rightbottom = AST.Address.fromR1C1withMode(ymax, xmax, AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path)

            lefttop, rightbottom

        /// <summary>returns the set of cells in the bounding box around the given set of
        /// cells, with an additional (positive-valued) epsilon</summary>
        let BoundingBox(cells: seq<AST.Address>)(epsilon: int) : seq<AST.Address> =
            let (lefttop,rightbottom) = BoundingRegion cells epsilon
            let bregion = new AST.Range(lefttop, rightbottom)
            seq (bregion.Addresses())

        let BoundingBoxHS(cells: HashSet<AST.Address>)(epsilon: int) : HashSet<AST.Address> =
            new HashSet<AST.Address>(BoundingBox cells epsilon)
