namespace ExceLint
    open System.Collections.Generic
    open System.Collections
    open System
    open Utils
    open ConfUtils
    open CommonTypes

    type ErrorModel(input: Input, analysis: Analysis, config: FeatureConf) =
        let r = match analysis with
                      | Histogram h -> h.ranking
                      | COF c -> c.ranking
                      | Cluster c -> c.ranking

        let rankByAddress = Array.mapi (fun rank (pair: KeyValuePair<AST.Address,double>) ->
                                let addr = pair.Key
                                addr, rank
                            ) r |> adict

        member self.AllCells : HashSet<AST.Address> = new HashSet<AST.Address>(input.dag.allCells())

        member self.DependenceGraph : Depends.DAG = input.dag

        member self.ScoreTimeInMilliseconds : int64 =
            match analysis with 
            | Histogram h -> h.score_time
            | COF c -> c.score_time
            | Cluster c -> c.score_time

        member self.FrequencyTable : CommonTypes.FreqTable =
            match analysis with 
            | Histogram h -> h.ftable
            | COF c -> failwith "Not valid for COF analysis."
            | Cluster c -> failwith "Not valid for cluster analysis."

        member self.IsCOF : bool =
            match analysis with
            | COF _ -> true
            | _ -> false

        member self.FrequencyTableTimeInMilliseconds : int64 =
            match analysis with 
            | Histogram h -> h.ftable_time
            | COF c -> 0L
            | Cluster c -> 0L

        member self.RankingTimeInMilliseconds : int64 =
            match analysis with 
            | Histogram h -> h.ranking_time
            | COF c -> c.ranking_time
            | Cluster c -> c.ranking_time

        member self.ConditioningSetSizeTimeInMilliseconds: int64 =
            match analysis with 
            | Histogram h -> h.csstable_time
            | COF c -> 0L
            | Cluster c -> 0L

        member self.CausesTimeInMilliseconds: int64 =
            match analysis with 
            | Histogram h -> h.causes_time
            | COF c -> c.fixes_time
            | Cluster c -> 0L

        member self.NumScoreEntries : int =
            let scores = match analysis with 
                         | Histogram h -> h.scores
                         | COF c -> c.scores
                         | Cluster c -> c.scores

            Array.fold (fun acc (pairs: (AST.Address*Countable)[]) ->
                acc + pairs.Length
            ) 0 (scores.Values |> Seq.toArray)

        member self.NumFreqEntries : int =
            match analysis with 
            | Histogram h -> h.ftable.Count
            | COF c -> 0
            | Cluster c -> 0

        member self.NumRankedEntries : int =
            match analysis with 
            | Histogram h -> h.ranking.Length
            | COF c -> c.ranking.Length
            | Cluster c -> c.ranking.Length

        member self.Scopes : Scope.Selector[] = config.EnabledScopes

        member self.Features : string[] = config.EnabledFeatures

        member self.Fixes : HypothesizedFixes option =
            match analysis with 
            | Histogram h -> h.fixes
            | COF c -> failwith "Not valid for COF analysis."
            | Cluster c -> failwith "Not valid for cluster analysis."

        member self.COFFixes : Dictionary<AST.Address,HashSet<AST.Address>> =
            match analysis with
            | COF c -> c.fixes
            | _ -> failwith "Not valid for non-COF analysis."

        member self.Scores : ScoreTable =
            match analysis with 
            | Histogram h -> h.scores
            | COF c -> c.scores
            | Cluster c -> c.scores

        member self.causeOf(addr: AST.Address) : KeyValuePair<HistoBin,Tuple<int,double>>[] =
            let causes = match analysis with 
                         | Histogram h -> h.causes
                         | _ -> failwith "Not valid for COF analysis."

            Array.map (fun cause ->
                let (bin,count,beta) = cause
                let tup = new Tuple<int,double>(count,beta)
                new KeyValuePair<HistoBin,Tuple<int,double>>(bin,tup)
            ) causes.[addr]

        member self.isAnomalous(addr: AST.Address) : bool =
            if not (rankByAddress.ContainsKey addr) then
                false
            else
                rankByAddress.[addr] <= self.Cutoff

        member self.weightOf(addr: AST.Address) : double =
            match analysis with 
            | Histogram h -> h.weights.[addr]
            | COF c -> c.weights.[addr]
            | Cluster c -> c.weights.[addr]

        member self.ranking() : Ranking =
            if ErrorModel.rankingIsSane r input.dag (input.config.IsEnabled "AnalyzeOnlyFormulas") then
                r
            else
                failwith "ERROR: Formula-only analysis returns non-formulas."

        // note that this cutoff is INCLUSIVE
        member self.Cutoff : int =
            match analysis with 
            | Histogram h -> h.cutoff_idx
            | COF c -> c.cutoff_idx
            | Cluster c -> c.cutoff_idx

        member self.Distribution : Distribution =
            match analysis with 
            | Histogram h -> ErrorModel.toDistribution h.causes
            | _ -> failwith "Not valid for non-histogram analysis."

        member self.inspectSelectorFor(addr: AST.Address, sel: Scope.Selector, dag: Depends.DAG) : KeyValuePair<AST.Address,(string*Countable)[]>[] =
            let selcache = Scope.SelectorCache()
            let sID = sel.id addr dag selcache

            let d = new Dict<AST.Address,(string*Countable) list>()

            let scores = match analysis with
                         | Histogram h -> h.scores
                         | COF c -> c.scores
                         | Cluster c -> c.scores

            Seq.iter (fun (kvp: KeyValuePair<string,(AST.Address*Countable)[]>) ->
                let fname: string = kvp.Key
                let scores: (AST.Address*Countable)[] = kvp.Value

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
            ) scores

            let debug = Seq.map (fun (kvp: KeyValuePair<AST.Address,(string*Countable) list>) ->
                                    let addr2: AST.Address = kvp.Key
                                    let scores: (string*Countable)[] = kvp.Value |> List.toArray

                                    new KeyValuePair<AST.Address,(string*Countable)[]>(addr2, scores)
                                    ) d

            debug |> Seq.toArray

        static member toDistribution(causes: Causes) : Distribution =
            let d = new Distribution()

            for addr_entry in causes do
                let addr = addr_entry.Key
                for bin_entry in addr_entry.Value do
                    let (hb,_,_) = bin_entry
                    let (feat,scope,hash) = hb

                    // init outer dict
                    if not (d.ContainsKey(feat)) then
                        d.Add(feat, new Dict<Scope.SelectID,Dict<Countable,Set<AST.Address>>>())

                    // init inner dict
                    if not (d.[feat].ContainsKey(scope)) then
                        d.[feat].Add(scope, new Dict<Countable,Set<AST.Address>>())

                    // init address set
                    if not (d.[feat].[scope].ContainsKey(hash)) then
                        d.[feat].[scope].Add(hash, set [])

                    // add address
                    d.[feat].[scope].[hash] <- d.[feat].[scope].[hash].Add(addr)

            d

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