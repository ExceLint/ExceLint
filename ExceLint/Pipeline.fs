namespace ExceLint
    module Pipeline =
        open System.Collections.Generic
        open System.Collections
        open System
        open Utils

        type Weight = double
        type Hash = double
        type Feature = string
        type Count = int
        type Distribution = Dict<Feature,Dict<Hash,Set<AST.Address>>>
        type ScoreTable = Dict<string,(AST.Address*Hash)[]>
        type FlatScoreTable = Dict<string*AST.Address,Hash>
        type ConditioningSetSizeTable = Dict<Scope.Selector,Dict<AST.Address,Count>>
        type HistoBin = Feature*Scope.SelectID*Hash
        type FreqTable = Dict<HistoBin,Count>
        type Weights = IDictionary<AST.Address,Weight>
        type Ranking = KeyValuePair<AST.Address,double>[]
        type Causes = Dict<AST.Address,(HistoBin*Count*Weight)[]>
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

        type Analysis = {
            scores: ScoreTable;
            ftable: FreqTable;
            csstable: ConditioningSetSizeTable;
            ranking: Ranking;
            causes: Causes;
            score_time: int64;
            ftable_time: int64;
            csstable_time: int64;
            ranking_time: int64;
            causes_time: int64;
            sig_threshold_idx: int;
            cutoff_idx: int;
            weights: Weights;
        }

        type AnalysisOutcome =
        | Success of Analysis
        | Cancellation

        type Pipe = Input -> Analysis -> AnalysisOutcome

        type PipeStart = Input -> AnalysisOutcome

        let comb (fn2: Pipe)(fn1: PipeStart) : PipeStart =
            fun (input: Input) ->
                match (fn1 input) with
                | Success(analysis) -> fn2 input analysis
                | Cancellation -> Cancellation

        let (+>) (fn1: PipeStart)(fn2: Pipe) : PipeStart = comb fn2 fn1