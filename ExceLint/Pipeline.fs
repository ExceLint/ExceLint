namespace ExceLint
    module Pipeline =
        open System.Collections.Generic
        open System.Collections
        open System
        open Utils

        type Weight = double
        type Feature = string
        type Count = int
        type Distribution = Dict<Feature,Dict<Scope.SelectID,Dict<Countable,Set<AST.Address>>>>
        type ScoreTable = Dict<Feature,(AST.Address*Countable)[]> // feature -> (address, countable)
        type FlatScoreTable = Dict<Feature*AST.Address,Countable>
        type ConditioningSetSizeTable = Dict<Scope.Selector,Dict<AST.Address,Count>>
        type HistoBin = Feature*Scope.SelectID*Countable
        type InvertedHistogram = System.Collections.ObjectModel.ReadOnlyDictionary<AST.Address,HistoBin>
        type FreqTable = Dict<HistoBin,Count>
        type ClusterTable = Dict<HistoBin,AST.Address list>
        type Weights = IDictionary<AST.Address,Weight>
        type Ranking = KeyValuePair<AST.Address,double>[]
        type HypothesizedFixes = Dict<AST.Address,Dict<Feature,Countable>>
        type Causes = Dict<AST.Address,(HistoBin*Count*Weight)[]>
        type Clustering = HashSet<HashSet<AST.Address>>
        type ChangeSet = {
            mutants: KeyValuePair<AST.Address,string>[];
            scores: ScoreTable;
            freqtable: FreqTable;
            selcache: Scope.SelectorCache;
            sidcache: Scope.SelectIDCache;
        }

        exception AnalysisCancelled

        type Input = {
            app: Microsoft.Office.Interop.Excel.Application;
            config: FeatureConf;
            dag: Depends.DAG;
            alpha: double;
            progress: Depends.Progress;
        }

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
            scores: ScoreTable;
            ranking: Ranking;
            score_time: int64;
            ranking_time: int64;
            sig_threshold_idx: int;
            cutoff_idx: int;
            weights: Weights;
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