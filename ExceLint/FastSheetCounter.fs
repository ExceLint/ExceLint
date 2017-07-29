namespace ExceLint
    open System.Collections.Generic
    open CommonTypes
    open Utils

    type SheetVectors = SheetVector[][]  // SheetVector[z][c]: c = countable num, z = worksheet num
    type Values = Dict<int*int*int, Countable>
    type Dimensions = Dict<int,int[]>

    type FastSheetCounter(grids: SheetVectors, dimensions: Dimensions, zNum: Dict<string,int>, countableMap: Dict<Countable,int>, valueMap: Values) = 
        let numCountables = grids.Length

        static member MinUpdateGrids(grids: SheetVectors)(x: int)(y: int)(z: int)(old_cn: int)(new_cn: int) : SheetVectors =
            // copy array
            let new_grid = Array.copy grids
            let svs = Array.copy new_grid.[z]
            new_grid.[z] <- svs

            // copy value & unset
            new_grid.[z].[old_cn] <- new_grid.[z].[old_cn].Copy
            new_grid.[z].[old_cn].Unset(x,y)

            // copy value & set
            new_grid.[z].[new_cn] <- new_grid.[z].[new_cn].Copy
            new_grid.[z].[new_cn].Set(x,y)

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
            let svs = grids.[z]

            svs
            |> Array.map (fun sv ->
                // create bitmask for range of interest
                let mask = SheetVector.Empty(sv.TopLeft, sv.BottomRight)
                for x = x_lo to x_hi do
                    for y = y_lo to y_hi do
                        mask.Set(x,y)

                // AND sheetvector and mask
                let sv_in_rng = sv.BitwiseAnd mask

                // count
                sv_in_rng.countBits
            )

        member self.NumWorksheets = zNum.Count

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
                let e = BasicStats.entropy ps

                e

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
        static member NewFalseGrid(width: int)(height: int) : bool[][] =
            [| 1 .. width |]
            |> Array.map (fun x ->
                Array.zeroCreate height
            )

        static member Initialize(ih: ROInvertedHistogram) : FastSheetCounter =
            let mutable zMax = 0
            let zNum = new Dict<string,int>()
            let dimensions = new Dict<int,int[]>()
            let values = new Values()
            let mutable cMax = 0
            let cnums = new Dict<Countable,int>()

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

                // save location-independent resultant vector
                let (_,_,c) = kvp.Value
                let cvr = c.ToCVectorResultant
                values.Add((x,y,z), cvr)

                // assign number to all distinct countables
                if not (cnums.ContainsKey cvr) then
                    cnums.Add(cvr, cMax)
                    cMax <- cMax + 1

                // initialize dimensions, if necessary
                if not (dimensions.ContainsKey z) then
                    dimensions.Add(z,[| x; y; x; y; |])

                // min x
                if x < dimensions.[z].[0] then
                    dimensions.[z].[0] <- x
                // min y
                if y < dimensions.[z].[1] then
                    dimensions.[z].[1] <- y
                // max x
                if x > dimensions.[z].[2] then
                    dimensions.[z].[2] <- x
                // max y
                if y > dimensions.[z].[3] then
                    dimensions.[z].[3] <- y

            // get distinct countables
            let cDistinct = cnums.Keys |> Seq.toArray

            // make quick-counting data-structure
            let sheetvects =
                Array.map (fun z ->
                    // get dimensions
                    let (ltX,ltY) = dimensions.[z].[0], dimensions.[z].[1]
                    let (rbX,rbY) = dimensions.[z].[2], dimensions.[z].[3]

                    Array.mapi (fun cNum c ->
                        // init empty sheet vector
                        let sv = SheetVector.Empty((ltX,ltY), (rbX,rbY))

                        // set all relevant bits
                        for x = ltX to rbX do
                            for y = ltY to rbY do
                                if values.ContainsKey (x,y,z) && values.[(x,y,z)] = c then
                                    sv.Set(x,y)
                        sv
                    ) cDistinct
                ) [| 0 .. zMax - 1 |]

            FastSheetCounter(sheetvects, dimensions, zNum, cnums, values)