namespace ExceLint
    open System
    open System.Collections.Generic
    open System.Collections.Immutable
    open Utils
    open CommonTypes
    open CommonFunctions
    open Distances
    open MathNet.Numerics.Distributions

    module EntropyModelBuilder2 =
        open FastDependenceAnalysis

        type ImmDistanceFMaker = ROInvertedHistogram -> ImmDistanceF

        [<Struct>]
        type Stats(feat_ms: int64, scale_ms: int64, invert_ms: int64, fsc_ms: int64, infer_ms: int64) =
            member self.FeatureTimeMS = feat_ms
            member self.ScaleTimeMS = scale_ms
            member self.InvertTimeMS = invert_ms
            member self.FastSheetCounterMS = fsc_ms
            member self.InferTimeMS = infer_ms

        let fix_probability(fixes: (ImmutableHashSet<AST.Address>*ImmutableHashSet<AST.Address>)[])(ih: ROInvertedHistogram) =
            // compute fix frequency
            let fix_freq = new Dict<Countable*Countable,int>()

            for (s,t) in fixes do
                let (_,_,sco) = ih.[Seq.head s]
                let (_,_,tco) = ih.[Seq.head t]
                let sco_res = sco.ToCVectorResultant
                let tco_res = tco.ToCVectorResultant
                let fix = sco_res, tco_res

                // count fixes
                if not (fix_freq.ContainsKey fix) then
                    fix_freq.Add(fix, 1)
                else
                    fix_freq.[fix] <- fix_freq.[fix] + 1

            fix_freq

        let prob_at_least_height_n(n: int)(dist: int[][])(trials: int) : double =
            let max_heights: int[] = Array.zeroCreate (trials + 1)
            for d in dist do
                let max_height = Array.max d
                max_heights.[max_height] <- max_heights.[max_height] + 1

            // compute CDF
            let mutable prop = 0
            let mutable total = 0
            for height = 0 to trials do
                total <- total + max_heights.[height]
                if height >= n then
                    prop <- prop + max_heights.[height]

            let prob = double prop / double total

            prob

        // computes kinda-sorta-EMD for a simulated "fix" of the source, s
        let FixDistance(s: ImmutableHashSet<AST.Address>)(t: ImmutableHashSet<AST.Address>)(ih: ROInvertedHistogram) : double =
            let (_,_,rep_co) = ih.[(Seq.head t)]

            // we know that the distance between t and t is 0, so we
            // just need to know the distance between s and s'
            // because t' = s' union t
            s |> Seq.sumBy (fun a ->
                     let (_,_,co) = ih.[a]
                     let fixed_co = co.UpdateResultant rep_co
                     co.EuclideanDistance fixed_co
                 )

        let ScoreForCell(addr: AST.Address)(ih: ROInvertedHistogram) : Countable =
            let (_,_,v) = ih.[addr]
            v

        let AddressIsFormulaValued(a: AST.Address)(ih: ROInvertedHistogram)(graph: Graph) : bool =
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

        let AddressIsNumericValued(a: AST.Address)(ih: ROInvertedHistogram)(graph: Graph) : bool =
            if graph.Values.ContainsKey a then
                let s = graph.Values.[a]
                let mutable d: double = 0.0
                let b = Double.TryParse(s, &d)
                let nf = not (AddressIsFormulaValued a ih graph)
                b && nf
            else
                false

        let AddressIsWhitespaceValued(a: AST.Address)(ih: ROInvertedHistogram)(graph: Graph) : bool =
            if graph.Values.ContainsKey a then
                let s = graph.Values.[a]
                let b = String.IsNullOrWhiteSpace(s)
                let nf = not (AddressIsFormulaValued a ih graph)
                b && nf
            else
                true

        let AddressIsStringValued(a: AST.Address)(ih: ROInvertedHistogram)(graph: Graph) : bool =
            let nn = not (AddressIsNumericValued a ih graph)
            let nf = not (AddressIsFormulaValued a ih graph)
            let nws = not (AddressIsWhitespaceValued a ih graph)
            nn && nf && nws

        let ClusterIsFormulaValued(c: ImmutableHashSet<AST.Address>)(ih: ROInvertedHistogram)(graph: Graph) : bool =
            c |> Seq.forall (fun addr -> AddressIsFormulaValued addr ih graph)

        let ClusterDirectionVector(c1: ImmutableHashSet<AST.Address>) : Countable =
            let (lt,rb) = Utils.BoundingRegion c1 0
            Vector(double (rb.X - lt.X), double (rb.Y - lt.Y), 0.0)

        let IsComputationChain(s: ImmutableHashSet<AST.Address>)(t: ImmutableHashSet<AST.Address>)(ih: ROInvertedHistogram) : bool =
            if s.Count > 1 then
                false
            else
                // get countable for first element
                let (_,_,sco) = ih.[Seq.head s]
                let sloc = sco.ToLocationVector

                // sort target by location-only distance from source
                let ts = 
                    t
                    |> Seq.sortBy (fun a -> 
                           let (_,_,aco) = ih.[a]
                           let aloc = aco.ToLocationVector
                           sloc.EuclideanDistance aloc
                       )
                    |> Seq.toList
                let furthest = ts |> List.rev |> List.head
                let ts' = ts |> List.rev |> List.tail |> List.rev
                let ts'' = (Seq.head s) :: ts'
                // list must be in order from furthest to closest
                let xs = List.rev ts''

                let rec isChain(xs: AST.Address list)(x: AST.Address) : bool =
                    match xs with
                    | x' :: xs' ->
                        // get relative vector for x
                        let (_,_,xco) = ih.[x]
                        let xrloc = xco.RelativeToLocationVector

                        // does it point at x'?
                        // yes, recurse; no, fail here
                        let (_,_,x'co) = ih.[x']
                        let x'coloc = x'co.ToLocationVector
                        if xrloc = x'coloc then
                            isChain xs' x'
                        else
                            false
                    | [] -> true

                isChain xs furthest

        // does s *exclusively* reference t?  If so, it's an aggregate.
        // off-by-one references and aggregates with extra constant
        // will return false;
        let IsAggregation(s: ImmutableHashSet<AST.Address>)(t: ImmutableHashSet<AST.Address>)(ih: ROInvertedHistogram) : bool =
            if s.Count > 1 then
                false
            else
                // get countable for first element
                let (_,_,sco) = ih.[Seq.head s]
                let sloc = sco.ToLocationVector

                // gin up fingerprint for a reference that points to all of t
                let vs =
                    t
                    |> Seq.map (fun a ->
                         let (_,_,tco) = ih.[a]
                         let tloc = tco.ToLocationVector

                         // compute address offset relative to sloc
                         let rel = tloc.Sub sloc

                         rel
                       )

                // make new source resultant
                let res =
                    vs
                    |> Seq.fold (fun (accv: Countable)(v: Countable) -> accv.Add v) ((Seq.head vs).Zero)

                // update resultant
                let sco' = sco.UpdateResultant res

                // if they're the same, return true
                sco = sco'

        type EntropyModel2(alpha: double, graph: Graph, regions: ImmutableClustering[], ih: ROInvertedHistogram, fsc: FastSheetCounter, d: ImmDistanceFMaker, indivisibles: ImmutableClustering[], stats: Stats) =
            // save the reverse lookup for later use
            let revLookups = Array.map (fun r -> ReverseClusterLookup r) regions

            member self.InvertedHistogram : ROInvertedHistogram = ih

            member self.Clustering(z: int) : ImmutableClustering =
                regions.[z]
                |> (fun s -> CommonTypes.makeImmutableGenericClustering s)

            member self.ZForWorksheet(sheet: string) : int = fsc.ZForWorksheet sheet

            member self.DependenceGraph = graph

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

                // split source clusters if necessary
                let source_clusters'' =
                    source_clusters'
                    |> Seq.map (fun cs ->
                           let isRect = FasterBinaryMinEntropyTree.ImmClusterIsRectangular cs
                           if isRect then
                               [| cs |]
                           else
                               let z = fsc'.ZForWorksheet ((Seq.head cs).WorksheetName)

                               // get bounding box for cluster
                               let (lt,rb) = Utils.BoundingRegion cs 0
                               let lt_xy = (lt.X, lt.Y)
                               let rb_xy = (rb.X, rb.Y)

                               // run tree analysis again on subtree
                               let tree = FasterBinaryMinEntropyTree.DecomposeAt fsc' z ih' lt_xy rb_xy

                               // get clusters
                               let regs = FasterBinaryMinEntropyTree.Regions tree
                               let clusters = regs |> Array.map (fun leaf -> leaf.Cells ih fsc)
                               clusters
                       )
                    |> Seq.concat
                    // make sure that new clusters do not contain any of the removed sources
                    |> Seq.filter (fun cs ->
                           (cs.Intersect source).Count = 0
                       )

                // remove sources
                let a = source_clusters |> Seq.fold (fun (acc: ImmutableClustering)(sc: ImmutableHashSet<AST.Address>) -> acc.Remove sc) regions.[z]

                // add modified sources
                let b = source_clusters'' |> Seq.fold (fun (acc: ImmutableClustering)(sc: ImmutableHashSet<AST.Address>) -> acc.Add sc) a

                // remove target
                let c = b |> Seq.filter (fun cl -> (cl.Intersect target).Count = 0) |> fun xs -> CommonTypes.makeImmutableGenericClustering xs

                // add merged
                let cs' = c.Add merged

                // re-coalesce the sheet that changed
                let region' = FasterBinaryMinEntropyTree.Coalesce cs' ih' indivisibles'.[z]

                // update region in question
                let regions' = Array.copy regions
                regions'.[z] <- region'

                new EntropyModel2(alpha, graph, regions', ih', fsc', d, indivisibles', stats)
                
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

            member self.AllWorkbooksNormalizedEntropy : double = 
                let cs = fsc.WorksheetIndices |> Seq.map (fun z -> self.Clustering z) 
                let wb_cluster = cs |> Seq.reduce (fun c1 c2 -> c1.Union c2)
                BinaryMinEntropyTree.NormalizedClusteringEntropy wb_cluster

            member self.NumSingletonFormulas : int = 
                let cs = fsc.WorksheetIndices |> Seq.map (fun z -> self.Clustering z) 
                let wb_cluster = cs |> Seq.reduce (fun c1 c2 -> c1.Union c2)
                let singleton_num =
                    wb_cluster |>
                    Seq.map (fun c -> c |> Seq.toArray) |>
                    Seq.fold (fun acc c -> if c.Length = 1 && graph.isFormula(c.[0]) then acc + 1 else acc) 0
                singleton_num

            member self.NumSingletonsNonWhitespace : int = 
                let cs = fsc.WorksheetIndices |> Seq.map (fun z -> self.Clustering z) 
                let wb_cluster =
                    cs
                    |> Seq.reduce (fun c1 c2 -> c1.Union c2)
                    |> Seq.map (fun c -> Seq.toArray c) 
                let singleton_num =
                    wb_cluster
                    |> Seq.filter (fun c -> not (AddressIsWhitespaceValued (c.[0]) ih graph))
                    |> Seq.fold (fun acc c -> if c.Length = 1 then acc + 1 else acc) 0
                singleton_num

            member self.NumSingletons : int = 
                let cs = fsc.WorksheetIndices |> Seq.map (fun z -> self.Clustering z) 
                let wb_cluster = cs |> Seq.reduce (fun c1 c2 -> c1.Union c2)
                let singleton_num =
                    wb_cluster |>
                    Seq.fold (fun acc c -> if c.Count = 1 then acc + 1 else acc) 0
                singleton_num

            member self.NumClusters : int =
                let cs = fsc.WorksheetIndices |> Seq.map (fun z -> self.Clustering z) 
                let wb_cluster = cs |> Seq.reduce (fun c1 c2 -> c1.Union c2)
                wb_cluster.Count

            member self.NumClustersNonWhitespace : int =
                let cs = fsc.WorksheetIndices |> Seq.map (fun z -> self.Clustering z) 
                let wb_cluster = cs |> Seq.reduce (fun c1 c2 -> c1.Union c2) |> Seq.map (fun c -> Seq.toArray c) 
                let wb_cluster_non_ws =
                    wb_cluster
                    |> Seq.filter (fun c -> not (AddressIsWhitespaceValued (c.[0]) ih graph))
                Seq.length wb_cluster_non_ws

            member self.ClusterSizes : int list =
                let cs = fsc.WorksheetIndices |> Seq.map (fun z -> self.Clustering z) 
                let wb_cluster = cs |> Seq.reduce (fun c1 c2 -> c1.Union c2)
                wb_cluster |> Seq.map (fun c -> c.Count) |> Seq.toList 

            member self.ClusterSizesNonWhitespace : int list =
                let cs = fsc.WorksheetIndices |> Seq.map (fun z -> self.Clustering z) 
                let wb_cluster = cs |> Seq.reduce (fun c1 c2 -> c1.Union c2) |> Seq.map (fun c -> Seq.toArray c) 
                wb_cluster
                |> Seq.filter (fun c -> not (AddressIsWhitespaceValued (c.[0]) ih graph))
                |> Seq.map (fun c -> c.Length) |> Seq.toList 

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
                let rng = new Random()
                let wsname = fsc.WorksheetForZ z
                let fcount = graph.getAllFormulaAddrs() |> Array.filter (fun addr -> addr.WorksheetName = wsname) |> Array.length
                let thresh = 0.5

                // get models & prune potential fixes
                let fixes =
                    self.PrevailingDirectionAdjacencies z true
                    |> Array.map (fun (target, addr) ->
                           // get z for s and t worksheet
                           let z = fsc.ZForWorksheet addr.WorksheetName

                           // ensure that source and target are on the same sheet
                           assert (z = fsc.ZForWorksheet (Seq.head target).WorksheetName)

                           let source_class = revLookups.[z].[addr]
                           // Fix equivalence class?
                           // Formulas and strings must all be fixed as a cluster together;
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

                // sort so that small fixes are favored when deduping converse fixes
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
                    // filtering of targets and sources must happen here because
                    // we may have swapped them above.
                    // ALL DOMAIN KNOWLEDGE GOES HERE

                    // no duplicates
                    |> Array.distinctBy (fun (s,_,t) -> s,t)
                    // no self fixes
                    |> Array.filter (fun (s,_,t) -> s <> t)
                    // modification: 2018-04-14: all sources must be formulas
                    |> Array.filter (fun (s,_,_) -> ClusterIsFormulaValued s ih graph)
                    // all targets must be formulas
                    |> Array.filter (fun (_,_,t) -> ClusterIsFormulaValued t ih graph)
                    // no whitespace sources, for now
                    |> Array.filter (fun (s,_,_) -> s |> Seq.forall (fun a -> not (AddressIsWhitespaceValued a ih graph)))
                    // no string sources, for now
                    |> Array.filter (fun (s,_,_) -> s |> Seq.forall (fun a -> not (AddressIsStringValued a ih graph)))
                    // no single-cell targets
                    |> Array.filter (fun (_,_,t) -> t.Count > 1)
                    // computation chains do not need fixing
                    |> Array.filter (fun (s,_,t) -> not (IsComputationChain s t ih))
                    // nor do aggregates
                    |> Array.filter (fun (s,_,t) -> not (IsAggregation s t ih))

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

                // compute fix frequency & gin up null hypothesis
                let fix_freq = fix_probability (fixes''' |> Array.map (fun (s,_,t) -> s,t)) ih
                let uniform_prior = Array.init (fix_freq.Count) (fun _ -> 1.0 / double (fix_freq.Count))

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

                // count again
                let chosen_fix_freq = fix_probability (dps' |> Array.map (fun (s,t,_) -> s,t )) ih

                let probable_fixes =
                    dps'
                    // don't propose fixes where the # source fingerprints / # target fingerprints is common;
                    // let's say more than 95% for now
                    |> Array.map (fun (s,t,dp) ->
                           let (_,_,sco) = ih.[Seq.head s]
                           let (_,_,tco) = ih.[Seq.head t]
                           let sco_res = sco.ToCVectorResultant
                           let tco_res = tco.ToCVectorResultant
                           let fix = sco_res, tco_res

                           // get the count for this fix
                           let count = chosen_fix_freq.[fix]
                           let trials = (self.Cutoff + 1)

                           // hypothesis test: is this fix overrepresented given uniform distribution of bugs over targets?
                           let X = Multinomial.Samples(rng, uniform_prior, trials) |> Seq.take 10000 |> Seq.toArray
                           let p = prob_at_least_height_n count X trials
                           
                           s, t, dp, p
                       )
                    // keep if we can't rule out the null hypothesis
                    |> Array.filter (fun (_,_,_,p) -> p > 0.05)

                // produce one model for each adjacency
                let models = 
                    probable_fixes
                    |> Array.Parallel.map (fun (source, target, wdotproduct, prob) ->
                            // produce a new model for each adjacency
                            let numodel = self.MergeCluster source target

                            // compute entropy difference
                            let entropyDelta = self.EntropyDiff z numodel

                            // compute fix distance
                            let dist = FixDistance source target ih

                            // compute fix frequency
                            let (_,_,sco) = ih.[Seq.head source]
                            let (_,_,tco) = ih.[Seq.head target]
                            let sco_res = sco.ToCVectorResultant
                            let tco_res = tco.ToCVectorResultant
                            let fix = sco_res, tco_res
                            let SiT = double fix_freq.[fix]  // | S intersect T |
                            let freq = SiT / double (fix_freq |> Seq.sumBy (fun kvp -> kvp.Value))

                            // eliminate all zero or negative-scored fixes
                            let pf = ProposedFix(source, target, entropyDelta, wdotproduct, dist, freq, prob)
                            if pf.Score <= 0.0 then
                                None
                            else
                                Some (pf)
                        )
                    |> Array.choose id
                    // no duplicates (happens for whole-cluster merges)
                    |> Array.distinctBy (fun fix -> fix.Source, fix.Target)

                // sort from highest score to lowest
                let ranking =
                    models
                    |> Array.sortByDescending (fun f -> f.Score)
                    // no duplicate sources (this keeps the one sorted first)
                    |> Array.distinctBy (fun f -> f.Source)

                ranking

            member self.Scores : ScoreTable = ROInvertedHistogramToScoreTable ih

            member self.ScoreTimeMs = stats.FeatureTimeMS + stats.ScaleTimeMS + stats.InvertTimeMS + stats.FastSheetCounterMS + stats.InferTimeMS

            member self.NumWorksheets = fsc.NumWorksheets

            // compute the cutoff based on a percentage of the number of formulas,
            // by default PCT_TO_FLAG %
            member self.Cutoff : int =
                // this is a workbook-wide cutoff
                let num_formulas = graph.getAllFormulaAddrs().Length
                int (Math.Floor((double num_formulas) * alpha))

            static member InitialSetup(z: int)(ih: ROInvertedHistogram)(fsc: FastSheetCounter)(indivisibles: ImmutableClustering[]) : ImmutableClustering =
                // get the initial tree
                let clusters = FasterBinaryMinEntropyTree.Decompose fsc z ih
                
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
                        let m = EntropyModel2.Initialize input
                        let numSheets = m.NumWorksheets
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
                                ranking = ranking;
                                fixes = fixes;
                                score_time = m.ScoreTimeMs;
                                ranking_time = fixtime + combtime + rtime;
                                sig_threshold_idx = 0;
                                cutoff_idx = m.Cutoff;
                                weights = EntropyModel2.Weights fixeses;    // this just returns entropy delta for now
                                clustering = CommonFunctions.ToMutableClustering (EntropyModel2.RankingToClusters fixes);
                                escapehatch = Some (m :> obj);
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

                // get all NLFRs for every formula cell
                let _runf = fun () -> runEnabledFeatures cells input.dag input.config input.progress
                let (ns: ScoreTable,feat_time: int64) = PerfUtils.runMillis _runf ()

                // make HistoBin lookup by address
                let _runhisto = fun () -> invertedHistogram ns input.dag input.config
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
                let indivisibles = EntropyModel2.InitIndivisibles fsc.NumWorksheets

                // init all sheets
                let sw = System.Diagnostics.Stopwatch.StartNew()
                let regions = [| 0 .. (fsc.NumWorksheets - 1) |] |> Array.Parallel.map (fun z -> EntropyModel2.InitialSetup z ih fsc indivisibles)
                sw.Stop()
//                assert (FasterBinaryMinEntropyTree.SheetAnalysesAreDistinct regions)

                // collate stats
                let times = Stats(feat_time, 0L, invert_time, fsc_time, sw.ElapsedMilliseconds)

                new EntropyModel2(input.alpha, input.dag, regions, ih, fsc, distance_f, indivisibles, times)
