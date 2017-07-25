namespace ExceLint
    open System.Collections.Generic
    open CommonTypes
    open Utils

    type InitGrids = Dict<Countable,int[][][]>
    type Grids = int[][][][]    // c, z, x, y; i.e., countable num, worksheet num, adjusted x coord, adjusted y coord
    type Dimensions = int[][]
    type Values = Dict<int*int*int, Countable>

    type FastSheetCounter(grids: Grids, dimensions: Dimensions, zNum: Dict<string,int>, countableMap: Dict<Countable,int>, valueMap: Values) = 
        let numCountables = grids.Length

        static member MinUpdateGrids(grids: Grids)(x: int)(y: int)(z: int)(old_cn: int)(new_cn: int) : Grids =
            // do the minimum amount of array cloning to
            // update old countable and new countable counts
            let new_grid = Array.copy grids
            // old
            let new_grid_c = Array.copy new_grid.[old_cn]
            new_grid.[old_cn] <- new_grid_c
            let new_grid_z = Array.copy new_grid_c.[z]
            new_grid_c.[z] <- new_grid_z
            let new_grid_x = Array.copy new_grid_z.[x]
            new_grid_z.[x] <- new_grid_x
            // new
            let new_grid_c2 = Array.copy new_grid.[new_cn]
            new_grid.[new_cn] <- new_grid_c2
            let new_grid_z2 = Array.copy new_grid_c2.[z]
            new_grid_c2.[z] <- new_grid_z2
            let new_grid_x2 = Array.copy new_grid_z2.[x]
            new_grid_z2.[x] <- new_grid_x2

            // change old count to 0
            new_grid_x.[y] <- 0

            // change new count to 1
            new_grid_x2.[y] <- 1

            new_grid

        member self.Fix(addr: AST.Address)(oldc: Countable)(c: Countable) : FastSheetCounter =
            let z = self.ZForWorksheet addr.WorksheetName
            let x = addr.X - dimensions.[z].[0] 
            let y = addr.Y - dimensions.[z].[1]

            // get countable num
            let old_cn = countableMap.[oldc]
            let new_cn = countableMap.[c]

            // update grids
            let grids' = FastSheetCounter.MinUpdateGrids grids x y z old_cn new_cn

            // update value map
            let valueMap' = new Values(valueMap)
            valueMap'.[(x,y,z)] <- c

            new FastSheetCounter(grids', dimensions, zNum, countableMap, valueMap')

        // all coordinates are inclusive
        member self.CountsForZ(z: int)(x_lo: int)(x_hi: int)(y_lo: int)(y_hi: int) : int[] =
            // adjust indices
            let ltX = dimensions.[z].[0]
            let ltY = dimensions.[z].[1]

            let x_lo' = x_lo - ltX
            let x_hi' = x_hi - ltX
            let y_lo' = y_lo - ltY
            let y_hi' = y_hi - ltY

            // count
            let counts : int[] = Array.zeroCreate numCountables
            for cNum in [| 0 .. grids.Length - 1 |] do
                let grid2d = grids.[cNum].[z]
                let mutable i = 0
                for x in [| x_lo' .. x_hi' |] do
                    for y in [| y_lo' .. y_hi' |] do
                        // each grid element stores either a 0 or a 1
                        if grid2d.[x].[y] = 1 then
                            i <- i + 1
//                        counts.[cNum] <- counts.[cNum] + grid2d.[x].[y]
                counts.[cNum] <- i

            counts

        member self.EntropyFor(z: int)(x_lo: int)(x_hi: int)(y_lo: int)(y_hi: int) : double =
            if x_lo > x_hi || y_lo > y_hi then
                System.Double.PositiveInfinity
            else
                // get counts, omitting zeroes
                let cs = self.CountsForZ z x_lo x_hi y_lo y_hi
                         |> Array.filter (fun i -> i > 0)

                // compute probability vector
                let ps = BasicStats.empiricalProbabilities cs

                // compute entropy
                BasicStats.entropy ps

        member self.ZForWorksheet(ws: string) = zNum.[ws]

        member self.MinXForWorksheet(z: int) =
            dimensions.[z].[0]

        member self.MinYForWorksheet(z: int) =
            dimensions.[z].[1]

        member self.MaxXForWorksheet(z: int) =
            dimensions.[z].[2]

        member self.MaxYForWorksheet(z: int) =
            dimensions.[z].[3]

        member self.ValueFor(x: int)(y: int)(z: int) = valueMap.[x,y,z]

        // layout of grid is [x][y]
        static member NewZeroGrid(width: int)(height: int) : int[][] =
            [| 1 .. width |]
            |> Array.map (fun x ->
                Array.zeroCreate height
            )

        static member Initialize(ih: ROInvertedHistogram) : FastSheetCounter =
            let mutable zMax = 0
            let zNum = new Dict<string,int>()
            let initdim = new Dict<int,int[]>()
            let values = new Values()

            // init grid dict
            let initgrids = new InitGrids()

            // iterate through all elements in the inverted histogram
            for kvp in ih do
                let x = kvp.Key.X
                let y = kvp.Key.Y
                let z = if zNum.ContainsKey kvp.Key.WorksheetName then
                            zNum.[kvp.Key.WorksheetName]
                        else
                            let zn = zMax
                            zNum.[kvp.Key.WorksheetName] <- zn
                            zMax <- zMax + 1
                            zn

                // initialize dimensions, if necessary
                if not (initdim.ContainsKey z) then
                    initdim.Add(z,[| x; y; x; y; |])

                // min x
                if x < initdim.[z].[0] then
                    initdim.[z].[0] <- x
                // min y
                if y < initdim.[z].[1] then
                    initdim.[z].[1] <- y
                // max x
                if x > initdim.[z].[2] then
                    initdim.[z].[2] <- x
                // max y
                if y > initdim.[z].[3] then
                    initdim.[z].[3] <- y

            // compute width and height and index by z
            let widthAndHeight =
                Array.init zMax (fun z ->
                    // add one because min and max are both inclusive
                    let width = initdim.[z].[2] - initdim.[z].[0] + 1
                    let height = initdim.[z].[3] - initdim.[z].[1] + 1
                    width, height
                )

            // convert initdim to jagged array
            let dimensions =
                Array.init zMax (fun z ->
                    Array.init 4 (fun i ->
                        initdim.[z].[i]
                    )
                )

            // go through a second time to populate grids
            for kvp in ih do
                let z = zNum.[kvp.Key.WorksheetName]

                // adjust x and y for array dimensions
                let x = kvp.Key.X - initdim.[z].[0] 
                let y = kvp.Key.Y - initdim.[z].[1]

                // get countable
                let (_,_,c) = kvp.Value
                let res = c.ToCVectorResultant

                // store x,y,z -> Countable map
                // use the real x, y, and z here; not adjusted values
                values.Add((kvp.Key.X,kvp.Key.Y,z), res)
                
                // create grids, if necessary
                if not (initgrids.ContainsKey res) then
                    // make all the grids for every z
                    let gs = Array.init zMax (fun iz ->
                                let (width,height) = widthAndHeight.[iz]
                                FastSheetCounter.NewZeroGrid width height
                             )

                    // add to dictionary
                    initgrids.Add(res, gs)

                // add count to grid
                initgrids.[res].[z].[x].[y] <- 1

            // assign numbers to countables
            let (cMax,cTups) =
                initgrids
                |> Seq.fold (fun (i,xs) kvp ->
                       i + 1, (kvp.Key,i) :: xs
                   ) (0,[])
            let countableMap = cTups |> adict
            let invCountableMap = cTups |> Seq.map (fun (c,i) -> (i,c)) |> adict

            // convert initgrids into grids
            let grids =
                Array.init (cMax - 1) (fun cNum ->
                    let c = invCountableMap.[cNum]
                    let gs = initgrids.[c]
                    gs
                )

            FastSheetCounter(grids, dimensions, zNum, countableMap, values)