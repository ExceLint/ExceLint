namespace ExceLint
    open System.Collections.Generic
    open System
    open FastDependenceAnalysis

    module ConfUtils =
        type RunnerMap = Map<string, AST.Address -> Graph -> Countable>

    type DistanceMetric =
    | NearestNeighbor
    | EarthMover
    | MeanCentroid

    // a C#-friendly configuration object that is also pure/fluent
    type FeatureConf private (userConf: Map<string,Capability>, limitToSheet: string option, alpha: double) =
        let nop(cell: AST.Address)(dag: Graph) : Countable = Countable.Num 0.0

        // this function adds a feature and its capabilities to the configuration map,
        // and returns a new FeatureConf
        let capabilityConstructorHelper(paramz: string*Capability)(self: FeatureConf)(on: bool)(config: Map<string,Capability>) : FeatureConf =
            let (name,cap) = paramz
            if on then
                FeatureConf(
                    config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner }),
                    limitToSheet,
                    alpha
                )
            else
                if config.ContainsKey(name) then
                    FeatureConf(
                        config.Remove(name),
                        limitToSheet,
                        alpha
                    )
                else
                    self

        let _base = BaseFeature.run

        let _capabilities : Map<string,Capability> =
            [   
                Vector.ShallowInputVectorMixedFullCVectorResultantOSI.capability;
            ] |> Map.ofList

        let _config = Map.fold (fun (acc: Map<string,Capability>)(fname: string)(cap: Capability) ->
                        let cap' : Capability =
                            {   enabled = cap.enabled;
                                kind = cap.kind;
                                runner = if cap.enabled then cap.runner else nop
                            }
                        Map.add fname cap acc
                        ) _capabilities userConf

        let _features : ConfUtils.RunnerMap = Map.map (fun (fname: string)(cap: Capability) -> cap.runner) _config

        new() = FeatureConf(Map.empty, None, 0.05)

        // set significance threshold
        member self.setThresh(alphanew: double) : FeatureConf =
            new FeatureConf(userConf, limitToSheet, alphanew)

        // fluent constructors
        member self.enableShallowInputVectorMixedFullCVectorResultantOSI(on: bool) : FeatureConf =
            capabilityConstructorHelper Vector.ShallowInputVectorMixedFullCVectorResultantOSI.capability self on _config
        member self.analyzeRelativeToAllCells(on: bool) : FeatureConf =
            let name = "ScopeAllCells"
            let cap : Capability = { enabled = true; kind = ConfigKind.Scope; runner = nop}
            capabilityConstructorHelper (name,cap) self on _config
        member self.analyzeRelativeToColumns(on: bool) : FeatureConf =
            let name = "ScopeColumns"
            let cap : Capability = { enabled = true; kind = ConfigKind.Scope; runner = nop}
            capabilityConstructorHelper (name,cap) self on _config

        member self.analyzeRelativeToRows(on: bool) : FeatureConf =
            let name = "ScopeRows"
            let cap : Capability = { enabled = true; kind = ConfigKind.Scope; runner = nop}
            capabilityConstructorHelper (name,cap) self on _config
        member self.analyzeRelativeToLevels(on: bool) : FeatureConf =
            let name = "ScopeLevels"
            let cap : Capability = { enabled = true; kind = ConfigKind.Scope; runner = nop}
            capabilityConstructorHelper (name,cap) self on _config
        member self.analyzeRelativeToSheet(on: bool) : FeatureConf =
            let name = "ScopeSheets"
            let cap : Capability = { enabled = true; kind = ConfigKind.Scope; runner = nop}
            capabilityConstructorHelper (name,cap) self on _config
        member self.inferAddressModes(on: bool) : FeatureConf =
            let name = "InferAddressModes"
            let cap : Capability = { enabled = true; kind = ConfigKind.Misc; runner = nop }
            capabilityConstructorHelper (name,cap) self on _config
        member self.analyzeOnlyFormulas(on: bool) : FeatureConf =
            let name = "AnalyzeOnlyFormulas"
            let cap : Capability = { enabled = true; kind = ConfigKind.Misc; runner = nop }
            capabilityConstructorHelper (name,cap) self on _config
        member self.weightByIntrinsicAnomalousness(on: bool) : FeatureConf =
            let name = "WeightByIntrinsicAnomalousness"
            let cap : Capability = { enabled = true; kind = ConfigKind.Misc; runner = nop }
            capabilityConstructorHelper (name,cap) self on _config
        member self.weightByConditioningSetSize(on: bool) : FeatureConf =
            let name = "WeightByConditioningSetSize"
            let cap : Capability = { enabled = true; kind = ConfigKind.Misc; runner = nop }
            capabilityConstructorHelper (name,cap) self on _config
        member self.spectralRanking(on: bool) : FeatureConf =
            let name = "SpectralRanking"
            let cap : Capability = { enabled = true; kind = ConfigKind.Misc; runner = nop }
            capabilityConstructorHelper (name,cap) self on _config
        member self.enableDistanceNearestNeighbor(on: bool) : FeatureConf =
            let name = "NearestNeighborDistance"
            let cap : Capability = { enabled = true; kind = ConfigKind.Misc; runner = nop }
            capabilityConstructorHelper (name,cap) self on _config
        member self.enableDistanceMeanCentroid(on: bool) : FeatureConf =
            let name = "MeanCentroidDistance"
            let cap : Capability = { enabled = true; kind = ConfigKind.Misc; runner = nop }
            capabilityConstructorHelper (name,cap) self on _config
        member self.enableDistanceEarthMover(on: bool) : FeatureConf =
            let name = "EarthMoverDistance"
            let cap : Capability = { enabled = true; kind = ConfigKind.Misc; runner = nop }
            capabilityConstructorHelper (name,cap) self on _config
        member self.enableOldClusteringAlgorithm(on: bool) : FeatureConf =
            let name = "oldclusteralgo"
            let cap : Capability = { enabled = true; kind = ConfigKind.Misc; runner = nop }
            capabilityConstructorHelper (name,cap) self on _config
        member self.limitAnalysisToSheet(wsname: string) : FeatureConf =
            let name = "limitToCurrentSheet"
            let cap : Capability = { enabled = true; kind = ConfigKind.Misc; runner = nop }
            FeatureConf(
                _config.Add(name, cap),
                Some wsname,
                alpha
            )
        member self.enableDebugMode(on: bool) : FeatureConf =
            let name = "DebugMode"
            let cap : Capability = { enabled = true; kind = ConfigKind.Misc; runner = nop }
            capabilityConstructorHelper (name,cap) self on _config

        // getters
        member self.Threshold = alpha
        member self.FeatureByName
            with get(name) = _features.[name]
        member self.EnabledFeatures
            with get() : string[] = 
                _config |>
                    Map.toArray |>
                    Array.choose (fun (fname,cap) ->
                                    if cap.enabled && cap.kind = ConfigKind.Feature then
                                        Some fname
                                    else None)
        member self.EnabledScopes
            with get() : Scope.Selector[] =
                _config |>
                    Map.toArray |>
                    Array.choose (fun (confname,cap) ->
                                    if cap.enabled && cap.kind = ConfigKind.Scope then
                                        match confname with
                                        | "ScopeAllCells" -> Some Scope.AllCells
                                        | "ScopeColumns" -> Some Scope.SameColumn
                                        | "ScopeRows" -> Some Scope.SameRow
                                        | "ScopeLevels" -> Some Scope.SameLevel
                                        | "ScopeSheets" -> Some Scope.SameSheet
                                        | _ -> failwith "Unknown scope selector."
                                    else None)


        member self.IsEnabled(name: string) : bool =
            _config.ContainsKey name && _config.[name].enabled

        member self.IsEnabledOptCondAllCells : bool = _config.ContainsKey "ScopeAllCells" && _config.["ScopeAllCells"].enabled
        member self.IsEnabledOptCondRows : bool = _config.ContainsKey "ScopeRows" && _config.["ScopeRows"].enabled
        member self.IsEnabledOptCondCols : bool = _config.ContainsKey "ScopeColumns" && _config.["ScopeColumns"].enabled
        member self.IsEnabledOptCondLevels : bool = _config.ContainsKey "ScopeLevels" && _config.["ScopeLevels"].enabled
        member self.IsEnabledOptCondSheets : bool = _config.ContainsKey "ScopeSheets" && _config.["ScopeSheets"].enabled
        member self.IsEnabledOptAddrmodeInference : bool = _config.ContainsKey "InferAddressModes" && _config.["InferAddressModes"].enabled
        member self.IsEnabledOptWeightIntrinsicAnomalousness : bool = _config.ContainsKey "WeightByIntrinsicAnomalousness" && _config.["WeightByIntrinsicAnomalousness"].enabled
        member self.IsEnabledOptWeightConditioningSetSize : bool = _config.ContainsKey "WeightByConditioningSetSize" && _config.["WeightByConditioningSetSize"].enabled
        member self.IsEnabledSpectralRanking : bool = _config.ContainsKey "SpectralRanking" && _config.["SpectralRanking"].enabled
        member self.IsEnabledAnalyzeOnlyFormulas : bool = _config.ContainsKey "AnalyzeOnlyFormulas" && _config.["AnalyzeOnlyFormulas"].enabled
        member self.IsEnabledAnalyzeAllCells : bool = not (_config.ContainsKey "AnalyzeOnlyFormulas") || not (_config.["AnalyzeOnlyFormulas"].enabled)
        member self.Cluster : bool =
            let (name0,_) = Vector.ShallowInputVectorMixedFullCVectorResultantOSI.capability
            (_config.ContainsKey name0 && _config.[name0].enabled)
        member self.DistanceMetric =
            if _config.ContainsKey "MeanCentroidDistance" then
                DistanceMetric.MeanCentroid
            else if _config.ContainsKey "NearestNeighborDistance" then
                DistanceMetric.NearestNeighbor
            else if _config.ContainsKey "EarthMoverDistance" then
                DistanceMetric.EarthMover
            else
                DistanceMetric.EarthMover // default if nothing is specified
        member self.IsLimitedToSheet : string option = limitToSheet
        member self.DebugMode : bool = _config.ContainsKey "DebugMode" && _config.["DebugMode"].enabled

        // make sure that config option combinations make sense;
        // returns a 'corrected' config
        member self.validate : FeatureConf =
            let config = if self.IsEnabledSpectralRanking then
                            self.analyzeRelativeToAllCells(false)
                                .analyzeRelativeToRows(false)
                                .analyzeRelativeToColumns(false)
                                .analyzeRelativeToLevels(false)
                                .analyzeRelativeToSheet(true)
                         else if self.Cluster then
                            self.analyzeRelativeToAllCells(false)
                                .analyzeRelativeToRows(false)
                                .analyzeRelativeToColumns(false)
                                .analyzeRelativeToLevels(false)
                                .analyzeRelativeToSheet(true)
                         else
                            self

            // if the user did not explicitly ask to analyze
            // all cells, the default is to analyze only formulas
            let config' = if not config.IsEnabledAnalyzeAllCells then
                              config.analyzeOnlyFormulas(true)
                          else
                              config

            config'

        member self.rawConf = _config

        /// Returns the (set of changed options, set of removed options, set of added options)
        member self.diff(otherconf: FeatureConf) : Set<string>*Set<string>*Set<string> =
            let my_keys = _config |> Map.toSeq |> Seq.map fst |> Set.ofSeq
            let your_keys = otherconf.rawConf |> Map.toSeq |> Seq.map fst |> Set.ofSeq
            let all_keys = Set.union my_keys your_keys

            // return
            // 1. keys changed by you
            // 2. keys no longer present in you
            // 3. keys introduced in you
            let changed = Set.filter (fun (k: string) -> _config.[k].enabled <> otherconf.IsEnabled(k)) (Set.intersect my_keys your_keys)
            let added = Set.difference your_keys my_keys
            let removed = Set.difference my_keys your_keys
            (changed, removed, added)

        static member simpleConf(m: Map<string,Capability>) : Map<string,bool> =
            Map.map (fun (k: string)(v: Capability) -> v.enabled) m

        override self.Equals(obj: Object) : bool =
            let other_fc = obj :?> FeatureConf
            (FeatureConf.simpleConf _config) = (FeatureConf.simpleConf other_fc.rawConf)
