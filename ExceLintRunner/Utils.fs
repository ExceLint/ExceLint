module Utils
    open ExceLint.CommonTypes
    open ExceLint.Utils
    open System.Collections.Generic

    type ClusterIDs = Dict<HashSet<AST.Address>,int>

    let getClusterIDs(clustering: Clustering) : ClusterIDs =
        let d = new Dict<HashSet<AST.Address>,int>()
        let mutable i = 0
        Seq.iter (fun c ->
            if not (d.ContainsKey c) then
                d.Add(c, i)
                i <- i + 1
        ) clustering
        d

    let clusteringToCSVRows(clustering: Clustering)(ids: ClusterIDs) : ExceLintFileFormats.ClusteringRow[] =
        Seq.map (fun c ->
            Seq.map (fun (addr: AST.Address) ->
                let row = new ExceLintFileFormats.ClusteringRow()
                row.Path <- addr.A1Path()
                row.Workbook <- addr.A1Workbook()
                row.Worksheet <- addr.A1Worksheet()
                row.Address <- addr.A1Local()
                row.Cluster <- ids.[c]
                row
            ) c
        ) clustering
        |> Seq.concat
        |> Seq.sortBy (fun row -> (row.Path, row.Workbook, row.Worksheet, row.Address))
        |> Seq.toArray