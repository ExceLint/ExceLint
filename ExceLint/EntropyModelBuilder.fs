namespace ExceLint
    open System
    open System.Collections.Generic
    open System.Collections.Immutable
    open Utils
    open CommonTypes
    open CommonFunctions
    open Distances

    type ImmDistanceFMaker = ROInvertedHistogram -> ImmDistanceF

    module EntropyModelBuilder =
        [<Struct>]
        type ProposedFix(source: ImmutableHashSet<AST.Address>, target: ImmutableHashSet<AST.Address>, entropyDelta: double, weighted_dp: double, distance: double) =
            member self.Source = source
            member self.Target = target
            member self.EntropyDelta = entropyDelta
            member self.E = 1.0 / -self.EntropyDelta
            member self.Distance = distance
            member self.WeightedDotProduct = weighted_dp
            member self.Score = (self.E * self.WeightedDotProduct) / self.Distance

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

        type EntropyModel(graph: Depends.DAG, ih: ROInvertedHistogram, d: ImmDistanceFMaker, indivisibles: ImmutableClustering) =
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

            member private self.Adjacencies(onlyFormulaTargets: bool) : (ImmutableHashSet<AST.Address>*AST.Address)[] =
                // get adjacencies
                self.Clustering
                |> Seq.filter (fun c -> if onlyFormulaTargets then ClusterIsFormulaValued c ih graph else true)
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

            member self.Ranking : ProposedFix[] =
                // get models & prune
                let fixes =
                    self.Adjacencies true
                    |> Array.map (fun (target, addr) ->
                           let source_class = revLookup.[addr]
                           // Fix equivalence class?
                           // Formuals and strings must all be fixed as a cluster together;
                           // Whitespace and numbers may be 'borrowed' from parent cluster
                           let source = if AddressIsFormulaValued addr ih graph
                                           || AddressIsStringValued addr ih graph then
                                           // get equivalence class
                                           source_class
                                           else
                                           // put ad-hoc fix in its own cluster
                                           [| addr |].ToImmutableHashSet()
                           
                           source, source_class, target
                       )

                // no converse fixes or duplicates
                let fhs = new HashSet<ImmutableHashSet<AST.Address>*ImmutableHashSet<AST.Address>>(fixes |> Array.map (fun (s,_,t) -> s,t))
                let fixes' =
                    fixes
                    |> Array.map (fun (s,sc,t) ->
                        // fix-normal-form:
                        // smallest number of things to change is the "source"
                        if s.Count < t.Count then
                            s,sc,t
                        else if s.Count = t.Count then
                            // same size; compare position; favor upperlefter
                            let (s_lt,_) = Utils.BoundingRegion s 0
                            let (t_lt,_) = Utils.BoundingRegion t 0
                            if s_lt.Y < t_lt.Y then
                                s,sc,t
                            else if s_lt.Y = t_lt.Y && s_lt.X < t_lt.X then
                                s,sc,t
                            else
                                t,sc,s
                        else
                            t,sc,s
                       )
                    |> Array.distinctBy (fun (s,_,t) -> s,t)
                    |> Array.map (fun (s,sc,t) ->
                           // compute dot product
                           // we always use the prevailing direction of
                           // the parent cluster even for ad-hoc fixes
                           let source_v = ClusterDirectionVector sc
                           let target_v = ClusterDirectionVector t

                           // if we're borrowing an element from its parent
                           // then scale the dot product accordingly
                           let dp_weight = if s = sc then 1.0 else 1.0 / double sc.Count                                   
                           let wdotproduct =
                                   // special case for null vectors, which
                                   // correspond to single cells
                                   if source_v.IsZero then
                                       1.0
                                   else
                                       dp_weight * (source_v.DotProduct target_v)
                           (s,t,wdotproduct)
                       )
                    // keep only fixes with a similar "prevailing direction"
                    |> Array.filter (fun (_,_,dp) -> dp > 0.0)

                // produce one model for each adjacency;
                // this is somewhat expensive to compute
                let models = 
                    fixes'
                    |> Array.Parallel.map (fun (source, target, wdotproduct) ->
                            // is the potential merge rectangular?
                            if not (BinaryMinEntropyTree.ImmMergeIsRectangular source target) then
                                None
                            else
                                // produce a new model for each adjacency
                                let numodel = self.MergeCluster source target

                                // compute entropy difference
                                let entropyDelta = self.EntropyDiff numodel

                                // compute inverse cluster distance
                                let dist = (d ih) source target

                                // eliminate all zero or negative-scored fixes
                                let pf = ProposedFix(source, target, entropyDelta, wdotproduct, dist)
                                if pf.Score <= 0.0 then
                                    None
                                else
                                    Some (pf)
                       )
                    |> Array.choose id
                    // no duplicates (happens for whole-cluster merges)
                    |> Array.distinctBy (fun fix -> fix.Source, fix.Target)

                // convert to ranking and sort from highest score to highest
                let ranking =
                    models
                    |> Array.sortByDescending (fun f -> f.Score)

                ranking

            static member Initialize(input: Input) : EntropyModel =
                // determine the set of cells to be analyzed
                let cells = analysisBase input.config input.dag

                // get all NLFRs for every formula cell
                let _runf = fun () -> runEnabledFeatures cells input.dag input.config input.progress
                let (ns: ScoreTable,feat_time: int64) = PerfUtils.runMillis _runf ()

                // scale
//                let _runscale = fun () -> ScaleBySheet ns
//                let (nlfrs: ScoreTable,scale_time: int64) = PerfUtils.runMillis _runscale () 

                // make HistoBin lookup by address
                let _runhisto = fun () -> invertedHistogram ns input.dag input.config
                let (ih: ROInvertedHistogram,invert_time: int64) = PerfUtils.runMillis _runhisto ()

                // do something with this eventually
//                let score_time = feat_time + scale_time + invert_time
                let score_time = feat_time + invert_time

                // define distance function
                let distance_f(invertedHistogram: ROInvertedHistogram) =
                    match input.config.DistanceMetric with
                    | DistanceMetric.NearestNeighbor -> min_dist_ro invertedHistogram
                    | DistanceMetric.EarthMover -> earth_movers_dist_ro invertedHistogram
                    | DistanceMetric.MeanCentroid -> cent_dist_ro invertedHistogram

                new EntropyModel(input.dag, ih, distance_f, ToImmutableClustering (new Clustering()))