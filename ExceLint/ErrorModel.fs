namespace ExceLint
    open System.Collections.Generic
    open System.Collections
    open System
    open Utils
    open ConfUtils
    open Pipeline

    type ErrorModel(input: Input, analysis: Analysis, config: FeatureConf) =
        member self.ScoreTimeInMilliseconds : int64 = analysis.score_time

        member self.FrequencyTable : Pipeline.FreqTable = analysis.ftable

        member self.FrequencyTableTimeInMilliseconds : int64 = analysis.ftable_time

        member self.RankingTimeInMilliseconds : int64 = analysis.ranking_time

        member self.ConditioningSetSizeTimeInMilliseconds: int64 = analysis.csstable_time

        member self.CausesTimeInMilliseconds: int64 = analysis.causes_time

        member self.NumScoreEntries : int = Array.fold (fun acc (pairs: (AST.Address*double)[]) ->
                                                acc + pairs.Length
                                            ) 0 (analysis.scores.Values |> Seq.toArray)

        member self.NumFreqEntries : int = analysis.ftable.Count

        member self.NumRankedEntries : int = analysis.ranking.Length

        member self.Scopes : Scope.Selector[] = config.EnabledScopes

        member self.Features : string[] = config.EnabledFeatures

        member self.causeOf(addr: AST.Address) : KeyValuePair<HistoBin,Tuple<int,double>>[] =
            Array.map (fun cause ->
                let (bin,count,beta) = cause
                let tup = new Tuple<int,double>(count,beta)
                new KeyValuePair<HistoBin,Tuple<int,double>>(bin,tup)
            ) analysis.causes.[addr]

        member self.weightOf(addr: AST.Address) : double = analysis.weights.[addr]

        member self.rankByFeatureSum() : Ranking =
            if ErrorModel.rankingIsSane analysis.ranking input.dag (input.config.IsEnabled "AnalyzeOnlyFormulas") then
                analysis.ranking
            else
                failwith "ERROR: Formula-only analysis returns non-formulas."

        // note that this cutoff is INCLUSIVE
        member self.Cutoff : int = analysis.cutoff

        member self.inspectSelectorFor(addr: AST.Address, sel: Scope.Selector, dag: Depends.DAG) : KeyValuePair<AST.Address,(string*double)[]>[] =
            let selcache = Scope.SelectorCache()
            let sID = sel.id addr dag selcache

            let d = new Dict<AST.Address,(string*double) list>()

            Seq.iter (fun (kvp: KeyValuePair<string,(AST.Address*double)[]>) ->
                let fname: string = kvp.Key
                let scores: (AST.Address*double)[] = kvp.Value

                let valid_scores =
                    Array.choose (fun (addr2,score) ->
                        if sel.id addr2 dag selcache = sID then
                            Some (addr2,score)
                        else
                            None
                    ) scores

                Array.iter (fun (addr2,score) ->
                    if d.ContainsKey addr2 then
                        d.[addr2] <- (fname,score) :: d.[addr2]
                    else
                        d.Add(addr2, [(fname,score)])
                ) valid_scores
            ) analysis.scores

            let debug = Seq.map (fun (kvp: KeyValuePair<AST.Address,(string*double) list>) ->
                                    let addr2: AST.Address = kvp.Key
                                    let scores: (string*double)[] = kvp.Value |> List.toArray

                                    new KeyValuePair<AST.Address,(string*double)[]>(addr2, scores)
                                    ) d

            debug |> Seq.toArray

        static member private rankingIsSane(r: Ranking)(dag: Depends.DAG)(formulasOnly: bool) : bool =
            if formulasOnly then
                Array.forall (fun (kvp: KeyValuePair<AST.Address,double>) -> dag.isFormula kvp.Key) r
            else
                true

        static member prettyHistoBinDesc(hb: HistoBin) : string =
            let (feature_name,selector,fscore) = hb
            "(" +
            "feature name: " + feature_name + ", " +
            "selector: " + Scope.Selector.ToPretty selector + ", " +
            "feature value: " + fscore.ToString() +
            ")"