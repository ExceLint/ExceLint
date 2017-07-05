namespace ExceLint
    open System
    open System.Collections.Generic
    open Utils
    open CommonTypes
    open CommonFunctions

    module SpectralModelBuilder =
        let private conditioningSetWeight(addr: AST.Address)(sel: Scope.Selector)(csstable: ConditioningSetSizeTable)(config: FeatureConf) : double =
            if config.IsEnabled "WeightByConditioningSetSize" then
                1.0 / double (csstable.[sel].[addr])
            else
                1.0

        let private getCause(addr: AST.Address)(cells: AST.Address[])(sel: Scope.Selector)(ftable: FreqTable)(scores: FlatScoreTable)(csstable: ConditioningSetSizeTable)(config: FeatureConf)(dag: Depends.DAG) : (HistoBin*int*double)[] =
            Array.map (fun fname -> 
                // get selector ID
                let sID = sel.id addr dag
                // get feature score
                let fscore = scores.[fname,addr]
                // get score count
                let count = ftable.[(fname,sID,fscore)]
                // get weight coefficient (beta)
                let beta = conditioningSetWeight addr sel csstable config
                (fname,sID,fscore),count,beta
            ) (config.EnabledFeatures)

        let private causes(cells: AST.Address[])(ftable: FreqTable)(scores: ScoreTable)(csstable: ConditioningSetSizeTable)(config: FeatureConf)(progress: Depends.Progress)(dag: Depends.DAG) : Causes =
            let fscores = makeFlatScoreTable scores

            // get histogram bin heights for every given cell
            // and for every enabled scope and feature
            let carr =
                Array.map (fun addr ->
                    let causes = Array.map (fun sel ->
                                    if progress.IsCancelled() then
                                        raise AnalysisCancelled

                                    getCause addr cells sel ftable fscores csstable config dag

                                    ) (config.EnabledScopes) |> Array.concat
                    (addr, causes)
                ) cells

            carr |> adict

        let private sizeOfConditioningSet(addr: AST.Address)(sel: Scope.Selector)(sidcache: Scope.SelectIDCache)(dag: Depends.DAG) : int =
            // get selector ID
            let sID = sel.id addr dag

            // get number of cells with matching sID
            if sidcache.ContainsKey sID then
                sidcache.[sID].Count
            else
                failwith ("sID cache is missing sID " + sID.ToString())

        let private buildFrequencyTable(scoretable: ScoreTable)(progress: Depends.Progress)(dag: Depends.DAG)(config: FeatureConf): FreqTable*Scope.SelectIDCache =
            // as a side-effect, maintain a selectID cache to make
            // conditioning set size lookup fast
            let s = new Scope.SelectIDCache()

            let d = new Dict<HistoBin,int>()
            Array.iter (fun fname ->
                Array.iter (fun (sel: Scope.Selector) ->
                    Array.iter (fun (addr: AST.Address, score: Countable) ->
                        if progress.IsCancelled() then
                            raise AnalysisCancelled

                        // fetch SelectID for this selector and address
                        let sID = sel.id addr dag

                        // update SelectIDCache if necessary
                        if not (s.ContainsKey(sID)) then
                            s.Add(sID, set [addr])
                        else if not (s.[sID].Contains(addr)) then
                            s.[sID] <- s.[sID].Add(addr)

                        if d.ContainsKey (fname,sID,score) then
                            let freq = d.[(fname,sID,score)]
                            d.[(fname,sID,score)] <- freq + 1
                        else
                            d.Add((fname,sID,score), 1)
                        progress.IncrementCounter()
                    ) (scoretable.[fname])
                ) (config.EnabledScopes)
            ) (config.EnabledFeatures)
            d,s

        let private buildCSSTable(cells: AST.Address[])(progress: Depends.Progress)(dag: Depends.DAG)(sidcache: Scope.SelectIDCache)(config: FeatureConf): ConditioningSetSizeTable =
            let d = new ConditioningSetSizeTable()
            Array.iter (fun (sel: Scope.Selector) ->
                Array.iter (fun cell ->
                    if progress.IsCancelled() then
                        raise AnalysisCancelled

                    progress.IncrementCounter()

                    // size of cell's conditioning set when conditioned on sel
                    let n = sizeOfConditioningSet cell sel sidcache dag

                    // initialize nested storage, if necessary
                    if not (d.ContainsKey sel) then
                        d.Add(sel, new Dict<AST.Address,int>())

                    // add to dictionary
                    d.[sel].Add(cell, n)
                ) cells
            ) (config.EnabledScopes)
            d

        let private justHashes(P: Distribution)(feature: Feature)(scope: Scope.SelectID) : Set<Countable> =
            let mutable s = set[]
            for hash_row in P.[feature].[scope] do
                let hash = hash_row.Key
                let addr = hash_row.Value |> Set.toList |> List.head
                let other_addresses = Seq.map (fun (pair: KeyValuePair<Countable,Set<AST.Address>>) -> pair.Value |> Set.toList) (P.[feature].[scope]) |> Seq.toList |> List.concat |> List.distinct
                let allsame = List.forall (fun (other_addr: AST.Address) -> other_addr.A1Worksheet() = addr.A1Worksheet()) other_addresses
                assert allsame
                s <- s.Add(hash)
            s

        let private justScopes(P: Distribution)(feature: Feature) : Set<Scope.SelectID> =
            let mutable s = set[]
            for hash_row in P.[feature] do
                s <- s.Add(hash_row.Key)
            s

        // the mean
        let private binCentroid(cells: Countable[]) : Countable =
            assert (cells.Length > 0)
            let n = double (cells.Length)

            cells
            |> Array.fold (fun (acc: Countable)(c: Countable) -> acc.Add c) (cells.[0].Zero)
            |> (fun c -> c.ScalarDivide n)

        let euclideanDistance(cell1: Countable)(cell2: Countable) : double =
            // note that it is possible for the two cells to be the same;
            // like when the cell happens to be the exact centroid
            Math.Sqrt((cell1.Negate.Add cell2).VectorMultiply(cell1.Negate.Add cell2))

        let private sameSheet(P: Distribution)(feature: Feature)(scope: Scope.SelectID)(other_hash: Countable)(anom_hash: Countable) =
            let acells = P.[feature].[scope].[anom_hash]
            let ocells = P.[feature].[scope].[other_hash]
            let bothcells = Set.union acells ocells |> Set.toList
            let (allsame,_) = List.fold (fun (is_same: bool, ws_opt: string option)(addr: AST.Address) ->
                                    match ws_opt with
                                    | Some(ws) ->
                                        let wssame = ws = addr.A1Worksheet()
                                        (is_same && wssame, Some ws)
                                    | None -> (is_same, Some(addr.A1Worksheet()))
                                ) (true,None) bothcells
            allsame

        let private earthMoversDistance(P: Distribution)(feature: AST.Address -> Countable)(fname: Feature)(scope: Scope.SelectID)(other_hash: Countable)(anom_hash: Countable): double =
            assert (other_hash <> anom_hash)
            assert (sameSheet P fname scope other_hash anom_hash)

            // get every cell in the named anomalous bin in the sheets conditional table only
            let dirt = P.[fname].[scope].[anom_hash] |> Set.toArray |> Array.map (fun a -> feature a)

            // get every cell in the named other bin(s) across all conditional tables
            let other_dirt = P.[fname].[scope].[other_hash] |> Set.toArray |> Array.map (fun a -> feature a)

            // compute the centroid of the other_hash
            let centroid = binCentroid other_dirt

            // compute work required to move dirt
            // amount * distance
            let dist = Array.sumBy (fun coord -> euclideanDistance coord centroid) dirt

            dist

        let private EMDsbyFeature(feature: AST.Address -> Countable)(fname: Feature)(causes: Causes) : Dict<Countable,Countable*double> =
            // the initial distribution
            let P = ErrorModel.toDistribution causes

            // get all sheet scopes 
            let scopes = justScopes P fname
            assert (Set.forall (fun (scope: Scope.SelectID) -> scope.IsKind = Scope.SameSheet) scopes)

            // find the distance to move every cell with a given hash to the nearest hash of a larger bin
            List.map (fun scope -> 
                // get set of feature hashes from distribution
                let hashes = justHashes P fname scope

                List.map (fun (a: Countable, hs: Countable[]) ->
                    // do not consider smaller bins
                    let a_count = P.[fname].[scope].[a].Count
                    let bigger = Array.filter (fun h -> P.[fname].[scope].[h].Count > a_count) hs

                    if bigger.Length > 0 then
                        let f = (fun h -> earthMoversDistance P feature fname scope h a) 
                        let min_hash = argmin f bigger
                        let min_distance = f min_hash
                        assert (sameSheet P fname scope a min_hash)
                        a, (min_hash, min_distance)
                    else
                        // If there is no closest hash, then we say that
                        // the closest hash is itself with a distance
                        // of positive infinity in order to put it at the
                        // end of the ranking.  Seems a bit like a hack.
                        a, (a, Double.PositiveInfinity)
                ) (cartesianProductByX hashes hashes)
            ) (scopes |> Set.toList)
            |> List.concat
            |> adict

        let private rankByEMD(cells: AST.Address[])(input: Input)(causes: Causes) : Ranking*HypothesizedFixes option =
            let emds = Array.map (fun fname ->
                            // get feature function
                            let feature = (fun addr -> (input.config.FeatureByName fname) addr input.dag)
                            // return (fname, distance) pair
                            fname, EMDsbyFeature feature fname causes
                        ) (input.config.EnabledFeatures) |> adict

            let ranking' = Array.map (fun (cell: AST.Address) ->
                                let sum = Array.sumBy (fun fname ->
                                                // get feature function
                                                let feature = input.config.FeatureByName fname
                                                // compute the hash of the cell using feature
                                                let hash = feature cell input.dag
                                                // find the EMD to the closest hash
                                                let (min_hash,min_dist) = emds.[fname].[hash]
                                                min_dist
                                            ) (input.config.EnabledFeatures)
                                cell, sum
                            ) cells
                            |> Array.sortBy (fun (cell,sum) -> sum)
                            |> Array.map (fun (cell,sum) -> new KeyValuePair<AST.Address,double>(cell,sum))

            let hf = new HypothesizedFixes()
            Array.iter (fun (cell: AST.Address) ->
                Array.iter (fun fname ->
                    // get feature function
                    let feature = input.config.FeatureByName fname
                    // compute the hash of the cell using feature
                    let hash = feature cell input.dag
                    // find the closest hash
                    let (min_hash,_) = emds.[fname].[hash]

                    // add entry for cell, if necessary
                    if not (hf.ContainsKey(cell)) then
                        hf.Add(cell, new Dict<Feature,Countable>())

                    // add closest hash
                    hf.[cell].Add(fname, min_hash)

                ) (input.config.EnabledFeatures)
            ) cells

            (ranking', Some hf)

        // sum the count of the appropriate feature bin of every feature
        // for the given address
        let private sumFeatureCounts(addr: AST.Address)(sel: Scope.Selector)(ftable: FreqTable)(scores: FlatScoreTable)(config: FeatureConf)(dag: Depends.DAG) : int =
            Array.sumBy (fun fname -> 
                // get selector ID
                let sID = sel.id addr dag
                // get feature score
                let fscore = scores.[fname,addr]
                // get score count
                ftable.[(fname,sID,fscore)]
            ) (config.EnabledFeatures)

        // for every cell and for every requested conditional,
        // find the bin height for the cell, then sum all
        // of these bin heights to produce a total ranking score
        // THIS IS WHERE ALL THE ACTION HAPPENS, FOLKS
        let private totalHistoSums(cells: AST.Address[])(ftable: FreqTable)(scores: ScoreTable)(csstable: ConditioningSetSizeTable)(config: FeatureConf)(progress: Depends.Progress)(dag: Depends.DAG) : Ranking*HypothesizedFixes option =
            let fscores = makeFlatScoreTable scores

            // get sums for every given cell
            // and for every enabled scope
            let addrSums: (AST.Address*double)[] =
                Array.map (fun addr ->
                    if progress.IsCancelled() then
                            raise AnalysisCancelled

                    let sum = Array.sumBy (fun sel ->
                                    // compute conditioning set weight
                                    let beta_sel_i = conditioningSetWeight addr sel csstable config

                                    // this is the conditional count
                                    let count = sumFeatureCounts addr sel ftable fscores config dag

                                    // weigh
                                    beta_sel_i * (double count)
                                ) (config.EnabledScopes)
                    addr, sum
                ) cells

            // return KeyValuePairs
            let ranking = Array.map (fun (addr,sum) -> new KeyValuePair<AST.Address,double>(addr,double sum)) addrSums

            ranking, None

        let private shouldNotHaveZeros(r: Ranking) : bool =
            Array.TrueForAll (r, fun (kvp: KeyValuePair<AST.Address,double>) -> kvp.Value > 0.0)

        let private transpose(mat: 'a[][]) : 'a[][] =
            // assumes that all subarrays are the same length
            Array.map (fun i ->
                Array.map (fun j ->
                    mat.[j].[i]
                ) [| 0..mat.Length - 1 |]
            ) [| 0..(mat.[0]).Length - 1 |]

        let private genChanges(cell: AST.Address)(refs: AST.Address[])(dag: Depends.DAG)(config: FeatureConf)(progress: Depends.Progress)(app: Microsoft.Office.Interop.Excel.Application) : ChangeSet[] =
            // for each referencing formula, systematically generate all ref variants
            let fs' = Array.mapi (fun i f ->
                        // get AST
                        let ast = dag.getASTofFormulaAt f

                        let mutator = ASTMutator.mutateExpr ast cell

                        let cabs_rabs = mutator AST.AddressMode.Absolute AST.AddressMode.Absolute
                        let cabs_rrel = mutator AST.AddressMode.Absolute AST.AddressMode.Relative
                        let crel_rabs = mutator AST.AddressMode.Relative AST.AddressMode.Absolute
                        let crel_rrel = mutator AST.AddressMode.Relative AST.AddressMode.Relative

                        [|(f, cabs_rabs); (f, cabs_rrel); (f, crel_rabs); (f, crel_rrel); |]
                        ) refs

            // make the first index the mode, the second index the formula
            let fsT = transpose fs'

            Array.map (fun (addrs_exprs: (AST.Address*AST.Expression)[]) ->
                // generate formulas for each expr AST
                let mutants = Array.map (fun (addr, expr: AST.Expression) ->
                                new KeyValuePair<AST.Address,string>(addr, expr.WellFormedFormula)
                                ) addrs_exprs

                // get new DAG
                let dag' = dag.CopyWithUpdatedFormulas(mutants, app, true, progress)

                // get the set of buckets
                let mutBuckets = runEnabledFeatures (
                                    Array.map (fun (kvp: KeyValuePair<AST.Address,string>) ->
                                        kvp.Key
                                    ) mutants
                                    ) dag' config nop

                // compute frequency tables
                let (mutFtable,mutSIDCache) = buildFrequencyTable mutBuckets nop dag' config
                    
                { mutants = mutants; scores = mutBuckets; freqtable = mutFtable; }
            ) fsT

        let private countBuckets(ftable: FreqTable) : int =
            // get total number of non-zero buckets in the entire table
            Seq.filter (fun (elem: KeyValuePair<HistoBin,int>) -> elem.Value > 0) ftable
            |> Seq.length

        let private chooseLikelyAddressMode(input: AST.Address)(refs: AST.Address[])(dag: Depends.DAG)(config: FeatureConf)(progress: Depends.Progress)(app: Microsoft.Office.Interop.Excel.Application) : ChangeSet =
            // generate all variants for the formulas that refer to this cell
            let css = genChanges input refs dag config progress app

            // count the buckets for the default
            let ref_fs = Array.map (fun (ref: AST.Address) ->
                            new KeyValuePair<AST.Address,string>(ref,dag.getFormulaAtAddress(ref))
                            ) refs
            let def_buckets = runEnabledFeatures (
                                    Array.map (fun (kvp: KeyValuePair<AST.Address,string>) ->
                                        kvp.Key
                                    ) ref_fs
                                    ) dag config nop
            let (def_freq, def_sidcache) = buildFrequencyTable def_buckets nop dag config
            let def_count = countBuckets def_freq

            // find the variants that minimize the bucket count
            let mutant_counts = Array.map (fun mutant ->
                                    countBuckets mutant.freqtable
                                ) css

            let mode_idx = argindexmin (fun mutant ->
                                // count histogram buckets
                                countBuckets mutant.freqtable
                            ) css

            if mutant_counts.[mode_idx] < def_count then
                css.[mode_idx]
            else
                { mutants = ref_fs; scores = def_buckets; freqtable = def_freq; }

        let private mutateDAG(cs: ChangeSet)(dag: Depends.DAG)(app: Microsoft.Office.Interop.Excel.Application)(p: Depends.Progress) : Depends.DAG =
            dag.CopyWithUpdatedFormulas(cs.mutants, app, true, p)

        let inferAddressModes(input: Input)(analysis: Analysis) : AnalysisOutcome =
            match analysis with
            | Histogram h ->
                if input.config.IsEnabled "InferAddressModes" then
                    try
                        let cells = analysisBase input.config input.dag

                        // convert ranking into map
                        let rankmap = h.ranking
                                        |> Array.map (fun (pair: KeyValuePair<AST.Address,double>) -> (pair.Key,pair.Value))
                                        |> toDict

                        let rank_positions = h.ranking
                                                |> Array.mapi (fun i (pair: KeyValuePair<AST.Address,double>) -> (pair.Key, i))
                                                |> toDict

                        // get all the formulas that ref each cell
                        let refss = Array.map (fun i -> i, input.dag.getFormulasThatRefCell i) cells |> toDict

                        // rank inputs by their impact on the ranking
                        let crank = Array.map (fun input_addr ->
                                        let anomalous_formulas = Array.filter (
                                                                    fun formula_addr ->
                                                                        let pos = try
                                                                                    rank_positions.[formula_addr]
                                                                                    with
                                                                                    | e ->
                                                                                        let is_formula = input.dag.isFormula formula_addr
                                                                                        h.cutoff_idx + 1
                                                                        pos <= h.cutoff_idx
                                                                    ) (refss.[input_addr])

                                        if anomalous_formulas.Length > 0 then
                                            let sum = Array.sumBy (fun formula ->
                                                            rankmap.[formula]
                                                        ) anomalous_formulas
                                            let average_score = sum / double (anomalous_formulas.Length)
                                            Some(input_addr, average_score)
                                        else
                                            None
                                    ) cells |>
                                    Array.choose id |>
                                    Array.sortBy (fun (input,score) -> score) |>
                                    Array.map (fun (input,score) -> input)

                        // for each input cell, try changing all refs to either abs or rel;
                        // if anomalousness drops, keep new interpretation
                        let dag' = Array.fold (fun accdag i ->
                                        // get referring formulas
                                        let refs = refss.[i]

                                        if refs.Length <> 0 then
                                            // run inference
                                            let cs = chooseLikelyAddressMode i refs accdag input.config input.progress input.app

                                            // update DAG
                                            mutateDAG cs accdag input.app input.progress
                                        else
                                            accdag
                                    ) input.dag crank

                        // score
                        let scores = runEnabledFeatures cells dag' input.config input.progress

                        // count freqs
                        let (freqs,sidcache) = buildFrequencyTable scores input.progress dag' input.config

                        // compute conditioning set size
                        let csstable = buildCSSTable cells input.progress dag' sidcache input.config

                        // rerank
                        let (ranking,_) = totalHistoSums cells freqs scores csstable input.config input.progress input.dag

                        // get causes
                        let causes = causes cells freqs scores csstable input.config input.progress input.dag

                        Success(Histogram { h with scores = scores; ftable = freqs; ranking = ranking; causes = causes; })
                    with
                    | AnalysisCancelled -> Cancellation
                    | e ->
                        // for breakpoint-friendliness
                        raise e
                else
                    Success(Histogram h)
            | COF c ->
                // do nothing for now
                Success(COF c)
            | Cluster c ->
                // do nothing for now
                Success(Cluster c)

        let runSpectralModel(input: Input) : AnalysisOutcome =
                try
                    let cells = (analysisBase input.config input.dag)

                    let _runf = fun () -> runEnabledFeatures cells input.dag input.config input.progress

                    // get scores for each feature: featurename -> (address, score)[]
                    let (scores: ScoreTable,score_time: int64) = PerfUtils.runMillis _runf ()

                    // build frequency table: (featurename, selector, score) -> freq
                    let _freqf = fun () -> buildFrequencyTable scores input.progress input.dag input.config
                    let (ftable,sidcache),ftable_time = PerfUtils.runMillis _freqf ()

                    // build conditioning set size table
                    let _cssf = fun () -> buildCSSTable cells input.progress input.dag sidcache input.config
                    let csstable,csstable_time = PerfUtils.runMillis _cssf ()

                    // save causes
                    let _causef = fun () -> causes cells ftable scores csstable input.config input.progress input.dag
                    let causes,causes_time = PerfUtils.runMillis _causef ()

                    // rank
                    let _rankf = fun () ->
                                    if input.config.IsEnabledSpectralRanking then
                                        // note that zero scores are OK here
                                        rankByEMD cells input causes
                                    else
                                        let (r,hfo) = totalHistoSums cells ftable scores csstable input.config input.progress input.dag
                                        assert shouldNotHaveZeros r
                                        r, hfo
                    let (ranking,fixes),ranking_time = PerfUtils.runMillis _rankf ()

                    Success(
                        Histogram(
                          {
                            scores = scores;
                            ftable = ftable;
                            csstable = csstable;
                            ranking = ranking;
                            causes = causes;
                            fixes = fixes;
                            score_time = score_time;
                            ftable_time = ftable_time;
                            csstable_time = csstable_time;
                            ranking_time = ranking_time;
                            causes_time = causes_time;
                            sig_threshold_idx = 0;
                            cutoff_idx = 0;
                            weights = new Dictionary<AST.Address,double>();
                          }
                        )
                    )
                with
                | AnalysisCancelled -> Cancellation

