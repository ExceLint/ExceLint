namespace ExceLint
    open System
    open System.Collections.Generic
    open System.Collections.Immutable
    open Utils
    open CommonTypes
    open CommonFunctions
    open Distances

    type ImmutableDistanceF = ROInvertedHistogram -> DistanceF

    module EntropyModelBuilder =
        let ScoreForCell(addr: AST.Address)(ih: ROInvertedHistogram) : Countable =
            let (_,_,v) = ih.[addr]
            v

        let AddressIsFormulaValued(a: AST.Address)(ih: ROInvertedHistogram)(graph: Depends.DAG) : bool =
            // have to check both because, e.g., =RAND() has a whitespace vector
            // and "fixes" may not be formulas in the dependence graph
            if ih.ContainsKey a then
                let score = ScoreForCell a ih
                let isfv = score.IsFormula
                let isf = graph.isFormula a
                let outcome = isf || isfv
                outcome
            else
                false

        let AddressIsNumericValued(a: AST.Address)(ih: ROInvertedHistogram)(graph: Depends.DAG) : bool =
            if graph.Values.ContainsKey a then
                let s = graph.Values.[a]
                let mutable d: double = 0.0
                let b = Double.TryParse(s, &d)
                let nf = not (AddressIsFormulaValued a ih graph)
                b && nf
            else
                false

        let AddressIsWhitespaceValued(a: AST.Address)(ih: ROInvertedHistogram)(graph: Depends.DAG) : bool =
            if graph.Values.ContainsKey a then
                let s = graph.Values.[a]
                let b = String.IsNullOrWhiteSpace(s)
                let nf = not (AddressIsFormulaValued a ih graph)
                b && nf
            else
                true

        let AddressIsStringValued(a: AST.Address)(ih: ROInvertedHistogram)(graph: Depends.DAG) : bool =
            let nn = not (AddressIsNumericValued a ih graph)
            let nf = not (AddressIsFormulaValued a ih graph)
            let nws = not (AddressIsWhitespaceValued a ih graph)
            nn & nf & nws

        type EntropyModel(graph: Depends.DAG, ih: ROInvertedHistogram, d: ImmutableDistanceF, indivisibles: ImmutableClustering) =
            // do region inference
            let tree = BinaryMinEntropyTree.Infer ih
            let regions = BinaryMinEntropyTree.Clustering tree ih indivisibles

            // save the reverse lookup for later use
            let revLookup = ReverseClusterLookup regions

            member self.InvertedHistogram : ROInvertedHistogram = ih

            member self.Clustering : ImmutableClustering = regions

            member self.Tree : BinaryMinEntropyTree = tree

            member private self.MergeCluster(source: ImmutableHashSet<AST.Address>)(target: ImmutableHashSet<AST.Address>) : EntropyModel =
                // get representative score from target
                let rep_score = target |> Seq.head |> (fun a -> ScoreForCell a ih)

                // update scores in histogram copy
                let ih' =
                    source
                    |> Seq.fold (fun (acc: ROInvertedHistogram)(addr: AST.Address) ->
                           let (a,b,oldscore) = ih.[addr]
                           let score' = oldscore.UpdateResultant rep_score
                           let bin = (a,b,score')
                           let acc' = acc.Remove addr
                           acc'.Add(addr, bin)
                       ) ih

                // update indivisibles
                let indivisibles' = indivisibles.Add (source.ToImmutableHashSet())

                new EntropyModel(graph, ih', d, indivisibles')

            member self.MergeCell(source: AST.Address)(target: AST.Address) : EntropyModel =
                // find the cluster of the target cell
                let target' = revLookup.[target]

                // is the cell a formula?
                if AddressIsFormulaValued source ih graph then
                    // find the equivalence class
                    let source' = revLookup.[source]
                    self.MergeCluster source' target'
                else
                    // otherwise, this is an ad-hoc fix
                    let source' = (new HashSet<AST.Address>([| source |])).ToImmutableHashSet()
                    self.MergeCluster source' target'

            member self.EntropyDiff(target: EntropyModel) : double =
                let c1 = self.Clustering
                let c2 = target.Clustering
                BinaryMinEntropyTree.ClusteringEntropyDiff c1 c2

            member self.Ranking : Ranking =
                // get adjacencies
                let adjs =
                    self.Clustering
                    |> Seq.map (fun target ->
                           // get adjacencies
                           let adj = HSAdjacentCellsImm target

                           // do not include adjacencies that are not in our histogram
                           let adj' = adj |> Seq.filter (fun addr -> ih.ContainsKey addr)

                           // flatten
                           adj'
                           |> Seq.map (fun a -> target, a)
                       )
                    |> Seq.concat
                    |> Seq.toArray

                // get models
                let models = 
                    adjs
                    // produce one model for each adjacency
                    |> Array.Parallel.map (fun (target, addr) ->
                            // put each address in its own cluster
                            let source = [| addr |].ToImmutableHashSet()

                            // produce a new model for each adjacency
                            let numodel = self.MergeCluster source target

                            // compute entropy difference
                            let entropy = self.EntropyDiff numodel
                            (addr, numodel, entropy)
                       )

                // eliminate duplicates
                let modelsNoDupes =
                    models
                    |> Array.groupBy (fun (k,_,_) -> k)
                    |> Array.map (fun (grp_k,fixes) ->
                         // eliminate all positive entropy fixes
                         let fixes' = fixes 
                                      |> Array.map (fun (k,m,e) -> if e < 0.0 then Some (k,m,e) else None )
                                      |> Array.choose id
                         if fixes'.Length = 0 then
                            None
                         else
                            Some(fixes' |> Array.sortBy (fun (_,_,e) -> e) |> Seq.head)
                       )
                    |> Array.choose id

                // convert to ranking and sort from highest entropy delta to lowest
                // (entropy deltas should be negative)
                let ranking =
                    modelsNoDupes
                    |> Array.map (fun (k,_,e) -> new KeyValuePair<AST.Address,double>(k,e))
                    |> Array.sortByDescending (fun kvp -> kvp.Value)

                ranking

            member self.MostLikelyAnomaly: AST.Address =
                let ranking = self.Ranking

                // return best address to user
                ranking.[0].Key

            static member Initialize(input: Input) : EntropyModel =
                // determine the set of cells to be analyzed
                let cells = analysisBase input.config input.dag

                // get all NLFRs for every formula cell
                let _runf = fun () -> runEnabledFeatures cells input.dag input.config input.progress
                let (ns: ScoreTable,feat_time: int64) = PerfUtils.runMillis _runf ()

                // scale
                let _runscale = fun () -> ScaleBySheet ns
                let (nlfrs: ScoreTable,scale_time: int64) = PerfUtils.runMillis _runscale () 

                // make HistoBin lookup by address
                let _runhisto = fun () -> invertedHistogram nlfrs input.dag input.config
                let (ih: ROInvertedHistogram,invert_time: int64) = PerfUtils.runMillis _runhisto ()

                // do something with this eventually
                let score_time = feat_time + scale_time + invert_time

                // define distance function
                let distance_f(invertedHistogram: ROInvertedHistogram) =
                    match input.config.DistanceMetric with
                    | DistanceMetric.NearestNeighbor -> min_dist_ro invertedHistogram
                    | DistanceMetric.EarthMover -> earth_movers_dist_ro invertedHistogram
                    | DistanceMetric.MeanCentroid -> cent_dist_ro invertedHistogram

                new EntropyModel(input.dag, ih, distance_f, ToImmutableClustering (new Clustering()))