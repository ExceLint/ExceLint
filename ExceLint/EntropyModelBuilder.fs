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
            let cells = ih.Keys |> Seq.toArray

            // do region inference
            let tree = BinaryMinEntropyTree.Infer ih
            let regions = BinaryMinEntropyTree.Clustering tree ih indivisibles

            // save the reverse lookup for later use
            let revLookup = ReverseClusterLookup regions

            member self.NumCells : int = cells.Length
            member self.InvertedHistogram : ROInvertedHistogram = ih
            member self.Clustering : ImmutableClustering = regions
            member self.Tree : BinaryMinEntropyTree = tree
            member self.Ranking : Ranking =
                failwith "coming soon to a theatre near you"
            member self.AnalysisBase = cells
            member self.Analysis = failwith "not yet"
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

            member self.MostLikelyAnomaly: AST.Address =
                // for each cluster

                // get all adjacent cells

                // produce a merge for each cell

                // compute entropy for each new model

                // if a cell is chosen twice
                // because it is adjacent to multiple clusters,
                // select the best fix

                // rank by smallest entropy

                // return best to user

                failwith "not done"

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