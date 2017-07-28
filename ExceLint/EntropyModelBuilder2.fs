namespace ExceLint
    open System
    open System.Collections.Generic
    open System.Collections.Immutable
    open Utils
    open CommonTypes
    open CommonFunctions
    open Distances

    module EntropyModelBuilder2 =
        let PCT_TO_FLAG = 5


        [<Struct>]
        type ProposedFix(source: ImmutableHashSet<AST.Address>, target: ImmutableHashSet<AST.Address>, entropyDelta: double, weighted_dp: double, distance: double) =
            member self.Source = source
            member self.Target = target
            member self.EntropyDelta = entropyDelta
            member self.E = 1.0 / -self.EntropyDelta
            member self.Distance = distance
            member self.WeightedDotProduct = weighted_dp
            member self.Score = (self.E * self.WeightedDotProduct) / self.Distance

        [<Struct>]
        type Stats(feat_ms: int64, scale_ms: int64, invert_ms: int64, fsc_ms: int64, infer_ms: int64) =
            member self.FeatureTimeMS = feat_ms
            member self.ScaleTimeMS = scale_ms
            member self.InvertTimeMS = invert_ms
            member self.FastSheetCounterMS = fsc_ms
            member self.InferTimeMS = infer_ms

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
            nn && nf && nws

        let ClusterIsFormulaValued(c: ImmutableHashSet<AST.Address>)(ih: ROInvertedHistogram)(graph: Depends.DAG) : bool =
            c |> Seq.forall (fun addr -> AddressIsFormulaValued addr ih graph)

        let ClusterDirectionVector(c1: ImmutableHashSet<AST.Address>) : Countable =
            let (lt,rb) = Utils.BoundingRegion c1 0
            Vector(double (rb.X - lt.X), double (rb.Y - lt.Y), 0.0)

        type EntropyModel2(graph: Depends.DAG, regions: ImmutableClustering[], ih: ROInvertedHistogram, fsc: FastSheetCounter, d: ImmDistanceFMaker, indivisibles: ImmutableClustering[], stats: Stats) =
            // save the reverse lookup for later use
            let revLookups = Array.map (fun r -> ReverseClusterLookup r) regions

            member self.InvertedHistogram : ROInvertedHistogram = ih

            member self.Clustering(z: int) : ImmutableClustering =
                regions.[z]
                |> (fun s -> CommonTypes.makeImmutableGenericClustering s)

//            member self.Trees : FasterBinaryMinEntropyTree[] = trees

            member self.ZForWorksheet(sheet: string) : int = fsc.ZForWorksheet sheet

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

            member private self.MergeCluster(source: ImmutableHashSet<AST.Address>)(target: ImmutableHashSet<AST.Address>) : EntropyModel2 =
                // sanity check
                let z = fsc.ZForWorksheet (Seq.head source).WorksheetName
                assert (z = fsc.ZForWorksheet (Seq.head target).WorksheetName)
                
                // update histogram
                let ih' = self.UpdateHistogram source target


                // copy indivisibles & update
                let indivisibles' = Array.copy indivisibles
                indivisibles'.[z] <- indivisibles'.[z].Add (source.ToImmutableHashSet())

                // update fsc
                let fsc' =
                    source
                    |> Seq.fold (fun (accfsc: FastSheetCounter)(saddr: AST.Address) ->
                            let oldc = (ScoreForCell saddr ih).ToCVectorResultant
                            let newc = (ScoreForCell saddr ih').ToCVectorResultant
                            accfsc.Fix saddr oldc newc
                        ) fsc


                // update clustering
                let merged = source |> Seq.fold (fun (acc: ImmutableHashSet<AST.Address>)(a: AST.Address) -> acc.Add a) target

                // get source cluster for each source address
                let source_clusters = source |> Seq.map (fun a -> revLookups.[z].[a]) |> Seq.distinct

                // remove source addresses from source clusters
                let source_clusters' =
                    source_clusters
                    |> Seq.map (fun cs ->
                           source
                           |> Seq.fold (fun (acc: ImmutableHashSet<AST.Address>)(a: AST.Address) ->
                                  if acc.Contains a then
                                      acc.Remove a
                                  else
                                      acc
                              ) cs
                       )
                    |> Seq.filter (fun cs -> cs.Count > 0)

                // remove sources
                let a = source_clusters |> Seq.fold (fun (acc: ImmutableClustering)(sc: ImmutableHashSet<AST.Address>) -> acc.Remove sc) regions.[z]

                // add modified sources
                let b = source_clusters' |> Seq.fold (fun (acc: ImmutableClustering)(sc: ImmutableHashSet<AST.Address>) -> acc.Add sc) a

                // remove target
                let c = b.Remove target

                // add merged
                let cs' = c.Add merged

                // re-coalesce the sheet that changed
                let region' = FasterBinaryMinEntropyTree.Coalesce cs' ih' indivisibles'.[z]

                // update region in question
                let regions' = Array.copy regions
                regions'.[z] <- region'

                new EntropyModel2(graph, regions', ih', fsc', d, indivisibles', stats)
                
            member self.MergeCell(source: AST.Address)(target: AST.Address) : EntropyModel2 =
                // get z for s and t worksheet
                let z = fsc.ZForWorksheet source.WorksheetName

                // ensure that source and target are on the same sheet
                assert (z = fsc.ZForWorksheet target.WorksheetName)

                // find the cluster of the target cell
                let target' = revLookups.[z].[target]

                // is the cell a formula?
                if AddressIsFormulaValued source ih graph then
                    // find the equivalence class
                    let source' = revLookups.[z].[source]
                    self.MergeCluster source' target'
                else
                    // otherwise, this is an ad-hoc fix
                    let source' = (new HashSet<AST.Address>([| source |])).ToImmutableHashSet()
                    self.MergeCluster source' target'
                     
            member self.EntropyDiff(z: int)(target: EntropyModel2) : double =
                let c1 = self.Clustering z
                let c2 = target.Clustering z
                BinaryMinEntropyTree.ClusteringEntropyDiff c1 c2

//            member private self.Adjacencies(onlyFormulaTargets: bool) : (ImmutableHashSet<AST.Address>*AST.Address)[] =
//                // get adjacencies
//                self.Clustering
//                |> Seq.filter (fun c -> if onlyFormulaTargets then ClusterIsFormulaValued c ih graph else true)
//                |> Seq.map (fun target ->
//                        // get adjacencies
//                        let adj = HSAdjacentCellsImm target
//
//                        // do not include adjacencies that are not in our histogram
//                        let adj' = adj |> Seq.filter (fun addr -> ih.ContainsKey addr)
//
//                        // flatten
//                        adj'
//                        |> Seq.map (fun a -> target, a)
//                    )
//                |> Seq.concat
//                |> Seq.toArray

            member private self.PrevailingDirectionAdjacencies(z: int)(onlyFormulaTargets: bool) : (ImmutableHashSet<AST.Address>*AST.Address)[] =
                self.Clustering z
                |> Seq.filter (fun c -> if onlyFormulaTargets then ClusterIsFormulaValued c ih graph else true)
                |> Seq.map (fun target ->
                       // get prevailing direction of target cluster
                       let dir = ClusterDirectionVector target
                       let sameX = dir.X = 0.0
                       let sameY = dir.Y = 0.0

                       // get representative X and Y
                       let x = (Seq.head target).X
                       let y = (Seq.head target).Y

                       // get adjacencies
                       let adj = HSAdjacentCellsImm target

                       let adj' =
                           adj
                           // exclude adjacencies if cluster is not strongly linear
                           |> Seq.filter (fun addr ->
                                  if sameX && sameY then
                                      // target is a singleton; all adjacencies OK
                                      true
                                  else if sameX then
                                      // target is strongly vertical; only adjacencies that share x
                                      addr.X = x
                                  else if sameY then
                                      // target is strongly horizontal; only adjacencies that share y
                                      addr.Y = y
                                  else
                                      // cluster changes in both x and y directions-- not strongly linear
                                      false
                              )
                           // do not include adjacencies that are not in our histogram
                           |> Seq.filter (fun addr -> ih.ContainsKey addr)

                       // flatten
                       adj'
                       |> Seq.map (fun a -> target, a)
                   )
                |> Seq.concat
                |> Seq.toArray

            member self.Fixes(z: int) : ProposedFix[] =
                // get models & prune
                let fixes =
                    self.PrevailingDirectionAdjacencies z true
                    |> Array.map (fun (target, addr) ->
                           // get z for s and t worksheet
                           let z = fsc.ZForWorksheet addr.WorksheetName

                           // ensure that source and target are on the same sheet
                           assert (z = fsc.ZForWorksheet (Seq.head target).WorksheetName)

                           let source_class = revLookups.[z].[addr]
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
                    // no duplicates
                    |> Array.distinctBy (fun (s,sc,t) -> s,t)

                // sort so that small fixes are favored when deduping
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

                // no converse fixes
                let fhs = new HashSet<ImmutableHashSet<AST.Address>*ImmutableHashSet<AST.Address>>()
                let mutable keep = []
                for fix in fixes' do
                    let (s,sc,t) = fix
                    let key1 = (s,t)
                    let key2 = (t,s)
                    if not (fhs.Contains(key1)) && not (fhs.Contains key2) then
                        fhs.Add(key1) |> ignore
                        fhs.Add(key2) |> ignore
                        keep <- (s,sc,t) :: keep
                
                let fixes'' = List.toArray keep

                let fixes''' =
                    fixes''
                    // don't propose fixes where source and target have the same resultant
                    |> Array.filter (fun (s,_,t) ->
                        let s_co = ScoreForCell (Seq.head s) ih
                        let t_co = ScoreForCell (Seq.head t) ih
                        s_co <> t_co
                    )

                let dps =
                    fixes'''
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
                    |> Array.filter (fun (_,_,dp) -> dp > 0.0 && not (System.Double.IsNaN dp))
                    |> Array.sortByDescending (fun (_,_,dp) -> dp)

                // merge must be rectangular
                let dps' =
                    dps
                    |> Array.filter (fun (s,t,_) ->
                           BinaryMinEntropyTree.ImmMergeIsRectangular s t
                       )

                // produce one model for each adjacency;
                // this is somewhat expensive to compute
                let models = 
                    dps'
                    // TODO PUT PARALLEL BACK
                    |> Array.map (fun (source, target, wdotproduct) ->
                            // produce a new model for each adjacency
                            let numodel = self.MergeCluster source target

                            // compute entropy difference
                            let entropyDelta = self.EntropyDiff z numodel

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

                // sort from highest score to highest
                let ranking =
                    models
                    |> Array.sortByDescending (fun f -> f.Score)

                ranking

            member self.Scores : ScoreTable = ROInvertedHistogramToScoreTable ih

            member self.ScoreTimeMs = stats.FeatureTimeMS + stats.ScaleTimeMS + stats.InvertTimeMS + stats.FastSheetCounterMS + stats.InferTimeMS

            member self.CacheHits = fsc.Hits

            member self.CacheLookups = fsc.Lookups

            // compute the cutoff based on a percentage of the number of formulas,
            // by default PCT_TO_FLAG %
            member self.Cutoff : int =
                // this is a workbook-wide cutoff
                let num_formulas = graph.getAllFormulaAddrs().Length
                let frac = (double PCT_TO_FLAG) / 100.0
                int (Math.Floor((double num_formulas) * frac))

            static member InitialSetup(z: int)(ih: ROInvertedHistogram)(fsc: FastSheetCounter)(indivisibles: ImmutableClustering[]) : ImmutableClustering =
                // get the initial tree
                let tree = FasterBinaryMinEntropyTree.Infer fsc z ih

                // get clusters
                let regs = FasterBinaryMinEntropyTree.Regions tree
                let clusters = regs |> Array.map (fun leaf -> leaf.Cells ih fsc) |> makeImmutableGenericClustering
                
                // coalesce all cells that have the same cvector,
                // ensuring that all merged clusters remain rectangular
                let clusters' = FasterBinaryMinEntropyTree.Coalesce clusters ih indivisibles.[z]
                
                clusters'

            static member CombineFxes(fixeses: ProposedFix[][]) : int64*ProposedFix[] =
                let sw = System.Diagnostics.Stopwatch.StartNew()

                // combine
                let fixes = fixeses |> Array.concat

                // sort
                let fixes' = fixes |> Array.sortByDescending (fun pf -> pf.Score)

                sw.Stop()

                sw.ElapsedMilliseconds, fixes'

            static member Ranking(fixes: ProposedFix[]) : int64*Ranking =
                let sw = new System.Diagnostics.Stopwatch()
                sw.Start()

                // convert proposed fixes to ordinary 'cell flags'
                let rs = fixes
                         // combine all fixes across sheets
                         |> Array.map (fun pf ->
                                pf.Source
                                |> Seq.map (fun addr ->
                                       new KeyValuePair<AST.Address,double>(addr, pf.Score)
                                   )
                                |> Seq.toArray
                            )
                         |> Array.concat

                // no duplicate sources
                let rs' = rs |> Array.distinctBy (fun kvp -> kvp.Key)

                sw.Stop()
                let ranking_time_ms = sw.ElapsedMilliseconds

                ranking_time_ms,rs'

            static member Weights(fixeses: ProposedFix[][]) : Dict<AST.Address,double> =
                let d = new Dict<AST.Address,double>()
                fixeses
                |> Seq.concat
                |> Seq.iter (fun pf ->
                       pf.Source
                       |> Seq.iter (fun addr ->
                            if not (d.ContainsKey addr) then
                                d.Add(addr, pf.EntropyDelta)
                          )
                   )
                d

            static member RankingToClusters(fixes: ProposedFix[]) : ImmutableClustering =
                fixes
                |> Array.map (fun pf -> pf.Source)
                |> Array.distinctBy (fun s ->
                    let xs = s |> Seq.map (fun a -> a.A1Local().ToString())
                    String.Join(",", xs |> Seq.sort)
                   )
                |> (fun arr -> CommonTypes.makeImmutableGenericClustering arr)
                
            static member runClusterModel(input: Input) : AnalysisOutcome =
                try
                    if (analysisBase input.config input.dag).Length <> 0 then
                        let numSheets = input.dag.getWorksheetNames().Length
                        let m = EntropyModel2.Initialize input
                        let sw = System.Diagnostics.Stopwatch.StartNew()
                        let fixeses = [| 0 .. numSheets - 1|] |> Array.map m.Fixes
                        sw.Stop()
                        let fixtime = sw.ElapsedMilliseconds
                        let (combtime,fixes) = EntropyModel2.CombineFxes fixeses
                        let (rtime,ranking) = EntropyModel2.Ranking fixes

                        Success(Cluster
                            {
                                numcells = input.dag.allCells().Length;
                                numformulas = input.dag.getAllFormulaAddrs().Length;
                                scores = m.Scores;
                                ranking = ranking
                                score_time = m.ScoreTimeMs;
                                ranking_time = fixtime + combtime + rtime;
                                sig_threshold_idx = 0;
                                cutoff_idx = m.Cutoff;
                                weights = EntropyModel2.Weights fixeses;    // this just returns entropy delta for now
                                clustering = CommonFunctions.ToMutableClustering (EntropyModel2.RankingToClusters fixes);
                            }
                        )
                    else
                        CantRun "Cannot perform analysis. This worksheet contains no formulas."
                with
                | AnalysisCancelled -> Cancellation

            static member private InitIndivisibles(numSheets: int) : ImmutableClustering[] =
                Array.init numSheets (fun i -> ToImmutableClustering (new Clustering()))

            static member Initialize(input: Input) : EntropyModel2 =
                // determine the set of cells to be analyzed
                let cells = analysisBase input.config input.dag

                let sheets = cells |> Array.map (fun a -> a.WorksheetName) |> Array.distinct

                // get all NLFRs for every formula cell
                let _runf = fun () -> runEnabledFeatures cells input.dag input.config input.progress
                let (ns: ScoreTable,feat_time: int64) = PerfUtils.runMillis _runf ()

                // scale
                let _runscale = fun () -> ScaleBySheet ns
                let (nlfrs: ScoreTable,scale_time: int64) = PerfUtils.runMillis _runscale () 

                // make HistoBin lookup by address
                let _runhisto = fun () -> invertedHistogram nlfrs input.dag input.config
                let (ih: ROInvertedHistogram,invert_time: int64) = PerfUtils.runMillis _runhisto ()

                // make fsc
                let _runfsc = fun () -> FastSheetCounter.Initialize ih
                let (fsc: FastSheetCounter, fsc_time: int64) = PerfUtils.runMillis _runfsc ()

                // define distance function
                let distance_f(invertedHistogram: ROInvertedHistogram) =
                    match input.config.DistanceMetric with
                    | DistanceMetric.NearestNeighbor -> min_dist_ro invertedHistogram
                    | DistanceMetric.EarthMover -> earth_movers_dist_ro invertedHistogram
                    | DistanceMetric.MeanCentroid -> cent_dist_ro invertedHistogram

                // initialize indivisible set
                let indivisibles = EntropyModel2.InitIndivisibles sheets.Length

                // init all sheets
                let sw = System.Diagnostics.Stopwatch.StartNew()
                let regions = [| 0 .. (fsc.NumWorksheets - 1) |] |> Array.map (fun z -> EntropyModel2.InitialSetup z ih fsc indivisibles)
                sw.Stop()
                assert (FasterBinaryMinEntropyTree.SheetAnalysesAreDistinct regions)

                // collate stats
                let times = Stats(feat_time, scale_time, invert_time, fsc_time, sw.ElapsedMilliseconds)

                new EntropyModel2(input.dag, regions, ih, fsc, distance_f, indivisibles, times)
