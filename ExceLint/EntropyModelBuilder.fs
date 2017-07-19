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
        [<Struct>]
        type ProposedFix(source: ImmutableHashSet<AST.Address>, target: ImmutableHashSet<AST.Address>, entropyDelta: double, dotproduct: double) =
            member self.Source = source
            member self.Target = target
            member self.EntropyDelta = entropyDelta
            member self.DotProduct = dotproduct

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

        let ClusterIsFormulaValued(c: ImmutableHashSet<AST.Address>)(ih: ROInvertedHistogram)(graph: Depends.DAG) : bool =
            c |> Seq.forall (fun addr -> AddressIsFormulaValued addr ih graph)

        let ClusterDirectionVector(c1: ImmutableHashSet<AST.Address>) : Countable =
            let (lt,rb) = Utils.BoundingRegion c1 0
            Vector(double (rb.X - lt.X), double (rb.Y - lt.Y), 0.0)

        type EntropyModel(graph: Depends.DAG, ih: ROInvertedHistogram, d: ImmutableDistanceF, indivisibles: ImmutableClustering) =
            // do region inference
            let tree = BinaryMinEntropyTree.Infer ih
            let regions = BinaryMinEntropyTree.Clustering tree ih indivisibles

            // save the reverse lookup for later use
            let revLookup = ReverseClusterLookup regions

            member self.InvertedHistogram : ROInvertedHistogram = ih

            member self.Clustering : ImmutableClustering = regions

            member self.Tree : BinaryMinEntropyTree = tree

            member private self.UpdateHistogram(source: ImmutableHashSet<AST.Address>)(target: ImmutableHashSet<AST.Address>) : ROInvertedHistogram =
                // get representative score from target
                let rep_score = target |> Seq.head |> (fun a -> ScoreForCell a ih)

                // update scores in histogram copy
                source
                |> Seq.fold (fun (acc: ROInvertedHistogram)(addr: AST.Address) ->
                        let (a,b,oldscore) = ih.[addr]
                        let score' = oldscore.UpdateResultant rep_score
                        let bin = (a,b,score')
                        let acc' = acc.Remove addr
                        acc'.Add(addr, bin)
                    ) ih

            member private self.MergeCluster(source: ImmutableHashSet<AST.Address>)(target: ImmutableHashSet<AST.Address>) : EntropyModel =
                let ih' = self.UpdateHistogram source target

                // update indivisibles
                let indivisibles' = indivisibles.Add (source.ToImmutableHashSet())

                new EntropyModel(graph, ih', d, indivisibles')

            // does a merge like the other MergeCell call, but does not
            // build an entire model
            member self.FastMergeCluster(source: ImmutableHashSet<AST.Address>)(target: ImmutableHashSet<AST.Address>) : ImmutableClustering =
                // remove old source and target
                let cs = (self.Clustering.Remove source).Remove target

                // merge source and target
                let unioned = target.Union source

                // put union back into clustering
                let cs' = cs.Add unioned

                // update inverted histogram
                let ih' = self.UpdateHistogram source target

                // coalesce
                BinaryMinEntropyTree.RectangularCoalesce cs' ih'
                
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

            // returns the source cluster and the new clustering
            member self.FastMerge(source: AST.Address)(target: ImmutableHashSet<AST.Address>) : ImmutableHashSet<AST.Address>*ImmutableClustering =
                // find the cluster of the source cell
                let source' = revLookup.[source]

                // is the cell a formula or string?
                if AddressIsFormulaValued source ih graph ||
                   AddressIsStringValued source ih graph then
                    // merge the equivalence class
                    source',self.FastMergeCluster source' target
                // the cell is a number or whitespace
                else
                    // remove the cell from its cluster
                    let source'' = source'.Remove source

                    // remove the original source cluster
                    let cs = self.Clustering.Remove source'

                    // put the new source cluster back
                    let cs' = cs.Add source''

                    // remove the target
                    let cs' = cs.Remove target

                    // merge the single cell with the target
                    let target' = target.Add source

                    // put the new target back
                    let cs'' = cs'.Add target'

                    // update histogram
                    let source_as_cl = (new HashSet<AST.Address>([| source |])).ToImmutableHashSet()
                    let ih' = self.UpdateHistogram source_as_cl target

                    // do rectangular coalesce
                    let cs''' = BinaryMinEntropyTree.RectangularCoalesce cs'' ih'

                    // and we're done
                    source_as_cl,cs''

            member self.EntropyDiff(target: EntropyModel) : double =
                let c1 = self.Clustering
                let c2 = target.Clustering
                BinaryMinEntropyTree.ClusteringEntropyDiff c1 c2

            member self.Ranking : ProposedFix[] =
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

                // get models & prune
                let models = 
                    adjs
                    // produce one model for each adjacency
                    |> Array.map (fun (target, addr) ->
                            // is the merge rectangular?
                            if ClusterIsFormulaValued target ih graph &&
                               BinaryMinEntropyTree.CellMergeIsRectangular addr target then

                                // Fix equivalence class?
                                // Formuals and strings must all be fixed as a cluster together;
                                // Whitespace and numbers may be 'borrowed' from parent cluster
                                let source = if AddressIsFormulaValued addr ih graph
                                                || AddressIsStringValued addr ih graph then
                                                // get equivalence class
                                                revLookup.[addr]
                                             else
                                                // put ad-hoc fix in its own cluster
                                                [| addr |].ToImmutableHashSet()

                                // produce a new model for each adjacency
                                let numodel = self.MergeCluster source target

                                // compute entropy difference
                                let entropy = self.EntropyDiff numodel

                                // compute dot product
                                // we always use the prevailing direction of
                                // the parent cluster even for ad-hoc fixes
                                let source_v = ClusterDirectionVector revLookup.[addr]
                                let target_v = ClusterDirectionVector target
                                let dotproduct =
                                    // special case for null vectors, which
                                    // correspond to single cells
                                    if source_v.IsZero then
                                        1.0
                                    else
                                        source_v.DotProduct target_v

                                // eliminate all fixes that don't decrease entropy
                                // or cancel out because of dot product weight
                                if entropy * dotproduct >= 0.0 then
                                    None
                                else
                                    Some (ProposedFix(source, target, entropy, dotproduct))
                            else
                                None
                       )
                    |> Array.choose id

                // convert to ranking and sort from highest entropy delta to lowest
                // (entropy deltas should be negative)
                let ranking =
                    models
                    |> Array.sortByDescending (fun f -> f.EntropyDelta * f.DotProduct)

                ranking

            member self.FastRanking : ProposedFix[] =
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

                // get models & prune
                let proposals = 
                    adjs
                    // produce one model for each adjacency
                    |> Array.map (fun (target, addr) ->
                            // is the merge rectangular?
                            if ClusterIsFormulaValued target ih graph &&
                               BinaryMinEntropyTree.CellMergeIsRectangular addr target then

                                // generate fix clustering
                                let (source,proposal) = self.FastMerge addr target

                                // compute entropy difference
                                let entropy = BinaryMinEntropyTree.ClusteringEntropyDiff self.Clustering proposal

                                // compute dot product
                                // we always use the prevailing direction of
                                // the parent cluster even for ad-hoc fixes
                                let source_v = ClusterDirectionVector revLookup.[addr]
                                let target_v = ClusterDirectionVector target
                                let dotproduct =
                                    // special case for null vectors, which
                                    // correspond to single cells
                                    if source_v.IsZero then
                                        1.0
                                    else
                                        source_v.DotProduct target_v

                                // eliminate all fixes that don't decrease entropy
                                // or cancel out because of dot product weight
                                if entropy * dotproduct >= 0.0 then
                                    None
                                else
                                    Some (ProposedFix(source, target, entropy, dotproduct))
                            else
                                None
                       )
                    |> Array.choose id

                // convert to ranking and sort from highest entropy delta to lowest
                // (entropy deltas should be negative)
                let ranking =
                    proposals
                    |> Array.sortByDescending (fun f -> f.EntropyDelta * f.DotProduct)

                ranking

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