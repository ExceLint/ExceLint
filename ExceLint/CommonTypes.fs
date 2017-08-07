namespace ExceLint
    module CommonTypes =
        open System.Collections.Generic
        open System.Collections.Immutable
        open System.Collections
        open System
        open Utils

        let private nop = Depends.Progress.NOPProgress()

        type Weight = double
        type Feature = string
        type Count = int
        type Distribution = Dict<Feature,Dict<Scope.SelectID,Dict<Countable,Set<AST.Address>>>>
        type ScoreTable = Dict<Feature,(AST.Address*Countable)[]> // feature -> (address, countable)
        type FlatScoreTable = Dict<Feature*AST.Address,Countable>
        type ConditioningSetSizeTable = Dict<Scope.Selector,Dict<AST.Address,Count>>
        type HistoBin = Feature*Scope.SelectID*Countable
        type ROInvertedHistogram = ImmutableDictionary<AST.Address,HistoBin>
        type InvertedHistogram = System.Collections.Generic.Dictionary<AST.Address,HistoBin>
        type FreqTable = Dict<HistoBin,Count>
        type ClusterTable = Dict<HistoBin,AST.Address list>
        type Weights = IDictionary<AST.Address,Weight>
        type Ranking = KeyValuePair<AST.Address,double>[]
        type HypothesizedFixes = Dict<AST.Address,Dict<Feature,Countable>>
        type Causes = Dict<AST.Address,(HistoBin*Count*Weight)[]>
        type GenericClustering<'p> = HashSet<HashSet<'p>>
        type ImmutableGenericClustering<'p> = ImmutableHashSet<ImmutableHashSet<'p>>
        type Clustering = GenericClustering<AST.Address>
        type ImmutableClustering = ImmutableGenericClustering<AST.Address>
        type ChangeSet = {
            mutants: KeyValuePair<AST.Address,string>[];
            scores: ScoreTable;
            freqtable: FreqTable;
        }

        type NoFormulasException(msg: string) =
                inherit Exception(msg)
        type HBDistance = HistoBin -> HistoBin -> ClusterTable -> FlatScoreTable -> double
        type Edge(pair: HashSet<AST.Address>*HashSet<AST.Address>) = 
            member self.tupled = fst pair, snd pair
            override self.Equals(o: obj) =
                let (x,y) = self.tupled
                let (x',y') = (o :?> Edge).tupled
                x = x' && y = y'
            override self.ToString() =
                let (x,y) = self.tupled
                x.ToString() + " -> " + y.ToString()

        [<Struct>]
        type ProposedFix(source: ImmutableHashSet<AST.Address>, target: ImmutableHashSet<AST.Address>, entropyDelta: double, weighted_dp: double, distance: double, fix_freq: double, ff_score: double) =
            member self.Source = source
            member self.Target = target
            member self.EntropyDelta = entropyDelta
            member self.Distance = distance
            member self.WeightedDotProduct = weighted_dp
            member self.TargetSize = double (Seq.length target)
            member self.Score = - self.TargetSize / (self.EntropyDelta * self.Distance)
            member self.FixFrequency = fix_freq
            member self.FixFrequencyScore = ff_score

        type DistanceF = HashSet<AST.Address> -> HashSet<AST.Address> -> double
        type ImmDistanceF = ImmutableHashSet<AST.Address> -> ImmutableHashSet<AST.Address> -> double
        type Distances = Dict<Edge,double>
        type ClusterStep = {
                source: Set<AST.Address>;
                target: Set<AST.Address>;
                distance: double;
                f: double;
                within_cluster_sum_squares: double;
                between_cluster_sum_squares: double;
                total_sum_squares: double;
                num_clusters: int;
                in_critical_region: bool;
             }
        type MinDistComparer(d: DistanceF) =
            interface IComparer<Edge> with
                member self.Compare(x: Edge, y: Edge) =
                    let (xs,xt) = x.tupled
                    let (ys,yt) = y.tupled
                    let distx = d xs xt
                    let disty = d ys yt

                    if distx < disty then
                        -1
                    else if distx = disty then
                        // sortedset discards elements with the same
                        // sort order even if they are not defined as equal;
                        // this behavior is very bad for ExceLint.
                        // return zero only iff x = y, otherwise
                        // sort deterministically but arbitrarily
                        if x.Equals(y) then
                            0
                        else
                            x.GetHashCode().CompareTo(y.GetHashCode())
                    else
                        1

        type ClusterIDs = Dict<HashSet<AST.Address>,int>

        exception AnalysisCancelled

        type Input = {
            app: Microsoft.Office.Interop.Excel.Application;
            config: FeatureConf;
            dag: Depends.DAG;
            alpha: double;
            progress: Depends.Progress;
        }

        let SimpleInput(app: Microsoft.Office.Interop.Excel.Application)(config: FeatureConf)(dag: Depends.DAG) : Input =
            { app = app; config = config; dag = dag; alpha = 0.05; progress = nop; }

        type HistoAnalysis = {
            scores: ScoreTable;
            ftable: FreqTable;
            csstable: ConditioningSetSizeTable;
            ranking: Ranking;
            causes: Causes;
            fixes: HypothesizedFixes option;
            score_time: int64;
            ftable_time: int64;
            csstable_time: int64;
            ranking_time: int64;
            causes_time: int64;
            sig_threshold_idx: int;
            cutoff_idx: int;
            weights: Weights;
        }

        type COFAnalysis = {
            scores: ScoreTable;
            ranking: Ranking;
            fixes: Dictionary<AST.Address,HashSet<AST.Address>>;
            fixes_time: int64;
            score_time: int64;
            ranking_time: int64;
            sig_threshold_idx: int;
            cutoff_idx: int;
            weights: Weights;
        }

        type ClusterAnalysis = {
            numcells: int;
            numformulas: int;
            scores: ScoreTable;
            ranking: Ranking;
            score_time: int64;
            ranking_time: int64;
            sig_threshold_idx: int;
            cutoff_idx: int;
            weights: Weights;
            clustering: Clustering;
            fixes: ProposedFix[];
        }

        type Analysis =
        | Histogram of HistoAnalysis
        | COF of COFAnalysis
        | Cluster of ClusterAnalysis

        type AnalysisOutcome =
        | Success of Analysis
        | Cancellation
        | CantRun of string

        type Pipe = Input -> Analysis -> AnalysisOutcome

        type PipeStart = Input -> AnalysisOutcome

        let comb (fn2: Pipe)(fn1: PipeStart) : PipeStart =
            fun (input: Input) ->
                match (fn1 input) with
                | Success(analysis) -> fn2 input analysis
                | Cancellation -> Cancellation

        let (+>) (fn1: PipeStart)(fn2: Pipe) : PipeStart = comb fn2 fn1

        let ToImmutableClustering(cs: Clustering) : ImmutableClustering =
            let hs_of_ic = cs |> Seq.map (fun c -> c.ToImmutableHashSet())
            hs_of_ic.ToImmutableHashSet()

        let makeImmutableGenericClustering<'p>(cs: seq<ImmutableHashSet<'p>>) : ImmutableGenericClustering<'p> =
            cs.ToImmutableHashSet()