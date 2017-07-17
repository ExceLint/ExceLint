namespace ExceLint
    open System
    open System.Collections.Generic
    open Utils
    open CommonTypes
    open CommonFunctions

    module Distances =
        // define distance
        let min_dist(hb_inv: InvertedHistogram) : DistanceF =
            (fun (source: HashSet<AST.Address>)(target: HashSet<AST.Address>) ->
                let pairs = cartesianProduct source target

                let compute = (fun (a,b) ->
                                let (_,_,ac) = hb_inv.[a]
                                let (_,_,bc) = hb_inv.[b]
                                ac.EuclideanDistance bc
                                )
                let f = (fun (a,b) ->
                            compute (a,b)
                        )
                let minpair = argmin f pairs
                let dist = f minpair
                dist * (double source.Count)
            )

        let min_dist_ro(hb_inv: ROInvertedHistogram) : DistanceF =
            (fun (source: HashSet<AST.Address>)(target: HashSet<AST.Address>) ->
                let pairs = cartesianProduct source target

                let compute = (fun (a,b) ->
                                let (_,_,ac) = hb_inv.[a]
                                let (_,_,bc) = hb_inv.[b]
                                ac.EuclideanDistance bc
                                )
                let f = (fun (a,b) ->
                            compute (a,b)
                        )
                let minpair = argmin f pairs
                let dist = f minpair
                dist * (double source.Count)
            )

        // this is kinda-sorta EMD; it has no notion of flows because I have
        // no idea what that means in terms of spreadsheet formula fixes
        let earth_movers_dist(hb_inv: InvertedHistogram) : DistanceF =
            (fun (source: HashSet<AST.Address>)(target: HashSet<AST.Address>) ->
                let compute = (fun (a,b) ->
                                let (_,_,ac) = hb_inv.[a]
                                let (_,_,bc) = hb_inv.[b]
                                ac.EuclideanDistance bc
                                )

                let f = (fun (a,b) ->
                            compute (a,b)
                        )

                // for every point in source, find the closest point in target
                let ds = source
                            |> Seq.map (fun addr ->
                                let pairs = target |> Seq.map (fun t -> addr,t)
                                let min_dist: double = pairs |> Seq.map f |> Seq.min
                                min_dist
                            )

                Seq.sum ds
            )

                // this is kinda-sorta EMD; it has no notion of flows because I have
        // no idea what that means in terms of spreadsheet formula fixes
        let earth_movers_dist_ro(hb_inv: ROInvertedHistogram) : DistanceF =
            (fun (source: HashSet<AST.Address>)(target: HashSet<AST.Address>) ->
                let compute = (fun (a,b) ->
                                let (_,_,ac) = hb_inv.[a]
                                let (_,_,bc) = hb_inv.[b]
                                ac.EuclideanDistance bc
                                )

                let f = (fun (a,b) ->
                            compute (a,b)
                        )

                // for every point in source, find the closest point in target
                let ds = source
                            |> Seq.map (fun addr ->
                                let pairs = target |> Seq.map (fun t -> addr,t)
                                let min_dist: double = pairs |> Seq.map f |> Seq.min
                                min_dist
                            )

                Seq.sum ds
            )

        // define distance (min distance between clusters)
        let cent_dist_ro(hb_inv: ROInvertedHistogram) : DistanceF =
            (fun (source: HashSet<AST.Address>)(target: HashSet<AST.Address>) ->
                // Euclidean distance with a small twist:
                // The distance between any two cells on different
                // sheets is defined as infinity.
                // This ensures that we always agglomerate intra-sheet
                // before agglomerating inter-sheet.
                let s_centroid = centroid_ro source hb_inv
                let t_centroid = centroid_ro target hb_inv
                let dist = if s_centroid.SameSheet t_centroid then
                                s_centroid.EuclideanDistance t_centroid
                            else
                                Double.PositiveInfinity
                dist * (double source.Count)
            )


        // define distance (min distance between clusters)
        let cent_dist(hb_inv: InvertedHistogram) : DistanceF =
            (fun (source: HashSet<AST.Address>)(target: HashSet<AST.Address>) ->
                // Euclidean distance with a small twist:
                // The distance between any two cells on different
                // sheets is defined as infinity.
                // This ensures that we always agglomerate intra-sheet
                // before agglomerating inter-sheet.
                let s_centroid = centroid source hb_inv
                let t_centroid = centroid target hb_inv
                let dist = if s_centroid.SameSheet t_centroid then
                                s_centroid.EuclideanDistance t_centroid
                            else
                                Double.PositiveInfinity
                dist * (double source.Count)
            )

