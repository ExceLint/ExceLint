namespace ExceLint
    open System.Collections.Generic
    open System.Collections
    open System

    module ConfUtils =
        type RunnerMap = Map<string, AST.Address -> Depends.DAG -> double>

    // a C#-friendly configuration object that is also pure/fluent
    type FeatureConf private (userConf: Map<string,Feature.Capability>) =
        let _base = Feature.BaseFeature.run 

        let _capabilities : Map<string,Feature.Capability> =
            [   Degree.InDegree.capability;
                Degree.OutDegree.capability;
                Degree.CombinedDegree.capability;
                Vector.DeepInputVectorRelativeL2NormSum.capability;
                Vector.DeepOutputVectorRelativeL2NormSum.capability;
                Vector.DeepInputVectorAbsoluteL2NormSum.capability;
                Vector.DeepOutputVectorAbsoluteL2NormSum.capability;
                Vector.DeepInputVectorMixedL2NormSum.capability;
                Vector.DeepOutputVectorMixedL2NormSum.capability
                Vector.ShallowInputVectorRelativeL2NormSum.capability;
                Vector.ShallowOutputVectorRelativeL2NormSum.capability;
                Vector.ShallowInputVectorAbsoluteL2NormSum.capability;
                Vector.ShallowOutputVectorAbsoluteL2NormSum.capability;
                Vector.ShallowInputVectorMixedL2NormSum.capability;
                Vector.ShallowOutputVectorMixedL2NormSum.capability;
                Proximity.Above.capability;
                Proximity.Below.capability;
                Proximity.Left.capability;
                Proximity.Right.capability
            ] |> Map.ofList

        let nop(cell: AST.Address)(dag: Depends.DAG) : double = 0.0

        let _config = Map.fold (fun (acc: Map<string,Feature.Capability>)(fname: string)(cap: Feature.Capability) ->
                        let cap' : Feature.Capability =
                            {   enabled = cap.enabled;
                                kind = cap.kind;
                                runner = if cap.enabled then cap.runner else nop
                            }
                        Map.add fname cap acc
                        ) _capabilities userConf

        let _features : ConfUtils.RunnerMap = Map.map (fun (fname: string)(cap: Feature.Capability) -> cap.runner) _config

        new() = FeatureConf(Map.empty)

        // fluent constructors
        member self.enableInDegree(on: bool) : FeatureConf =
            let (name,cap) = Degree.InDegree.capability
            if on then
                FeatureConf(
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.enableOutDegree(on: bool) : FeatureConf =
            let (name,cap) = Degree.OutDegree.capability
            if on then
                FeatureConf(
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.enableCombinedDegree(on: bool) : FeatureConf =
            let (name,cap) = Degree.CombinedDegree.capability
            if on then
                FeatureConf(
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.enableDeepInputVectorRelativeL2NormSum(on: bool) : FeatureConf =
            let (name,cap) = Vector.DeepInputVectorRelativeL2NormSum.capability
            if on then
                FeatureConf(
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.enableDeepOutputVectorRelativeL2NormSum(on: bool) : FeatureConf =
            let (name,cap) = Vector.DeepOutputVectorRelativeL2NormSum.capability
            if on then
                FeatureConf(
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.enableDeepInputVectorAbsoluteL2NormSum(on: bool) : FeatureConf =
            let (name,cap) = Vector.DeepInputVectorAbsoluteL2NormSum.capability
            FeatureConf(
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableDeepOutputVectorAbsoluteL2NormSum(on: bool) : FeatureConf =
            let (name,cap) = Vector.DeepOutputVectorAbsoluteL2NormSum.capability
            if on then
                FeatureConf(
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.enableDeepInputVectorMixedL2NormSum(on: bool) : FeatureConf =
            let (name,cap) = Vector.DeepInputVectorMixedL2NormSum.capability
            if on then
                FeatureConf(
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.enableDeepOutputVectorMixedL2NormSum(on: bool) : FeatureConf =
            let (name,cap) = Vector.DeepOutputVectorMixedL2NormSum.capability
            if on then
                FeatureConf(
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.enableShallowInputVectorRelativeL2NormSum(on: bool) : FeatureConf =
            let (name,cap) = Vector.ShallowInputVectorRelativeL2NormSum.capability
            FeatureConf(
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableShallowOutputVectorRelativeL2NormSum(on: bool) : FeatureConf =
            let (name,cap) = Vector.ShallowOutputVectorRelativeL2NormSum.capability
            if on then
                FeatureConf(
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.enableShallowInputVectorAbsoluteL2NormSum(on: bool) : FeatureConf =
            let (name,cap) = Vector.ShallowInputVectorAbsoluteL2NormSum.capability
            if on then
                FeatureConf(
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.enableShallowOutputVectorAbsoluteL2NormSum(on: bool) : FeatureConf =
            let (name,cap) = Vector.ShallowOutputVectorAbsoluteL2NormSum.capability
            if on then
                FeatureConf(
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.enableShallowInputVectorMixedL2NormSum(on: bool) : FeatureConf =
            let (name,cap) = Vector.ShallowInputVectorMixedL2NormSum.capability
            if on then
                FeatureConf(
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.enableShallowOutputVectorMixedL2NormSum(on: bool) : FeatureConf =
            let (name,cap) = Vector.ShallowOutputVectorMixedL2NormSum.capability
            if on then
                FeatureConf(
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.enableProximityAbove(on: bool) : FeatureConf =
            let (name,cap) = Proximity.Above.capability
            if on then
                FeatureConf(
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.enableProximityBelow(on: bool) : FeatureConf =
            let (name,cap) = Proximity.Below.capability
            if on then
                FeatureConf(
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.enableProximityLeft(on: bool) : FeatureConf =
            let (name,cap) = Proximity.Left.capability
            if on then
                FeatureConf(
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.enableProximityRight(on: bool) : FeatureConf =
            let (name,cap) = Proximity.Right.capability
            if on then
                FeatureConf(
                    _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.analyzeRelativeToAllCells(on: bool) : FeatureConf =
            let name = "ScopeAllCells"
            let cap : Feature.Capability = { enabled = true; kind = Feature.ConfigKind.Scope; runner = nop}
            if on then
                FeatureConf(
                    _config.Add(name, cap)
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.analyzeRelativeToColumns(on: bool) : FeatureConf =
            let name = "ScopeColumns"
            let cap : Feature.Capability = { enabled = true; kind = Feature.ConfigKind.Scope; runner = nop}
            if on then
                FeatureConf(
                    _config.Add(name, cap)
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.analyzeRelativeToRows(on: bool) : FeatureConf =
            let name = "ScopeRows"
            let cap : Feature.Capability = { enabled = true; kind = Feature.ConfigKind.Scope; runner = nop}
            if on then
                FeatureConf(
                    _config.Add(name, cap)
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.analyzeRelativeToLevels(on: bool) : FeatureConf =
            let name = "ScopeLevels"
            let cap : Feature.Capability = { enabled = true; kind = Feature.ConfigKind.Scope; runner = nop}
            if on then
                FeatureConf(
                    _config.Add(name, cap)
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.analyzeRelativeToSheet(on: bool) : FeatureConf =
            let name = "ScopeSheets"
            let cap : Feature.Capability = { enabled = true; kind = Feature.ConfigKind.Scope; runner = nop}
            if on then
                FeatureConf(
                    _config.Add(name, cap)
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.inferAddressModes(on: bool) : FeatureConf =
            let name = "InferAddressModes"
            let cap : Feature.Capability = { enabled = true; kind = Feature.ConfigKind.Misc; runner = nop }
            if on then
                FeatureConf(
                    _config.Add(name, cap)
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.analyzeOnlyFormulas(on: bool) : FeatureConf =
            let name = "AnalyzeOnlyFormulas"
            let cap : Feature.Capability = { enabled = true; kind = Feature.ConfigKind.Misc; runner = nop }
            if on then
                FeatureConf(
                    _config.Add(name, cap)
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.weightByIntrinsicAnomalousness(on: bool) : FeatureConf =
            let name = "WeightByIntrinsicAnomalousness"
            let cap : Feature.Capability = { enabled = true; kind = Feature.ConfigKind.Misc; runner = nop }
            if on then
                FeatureConf(
                    _config.Add(name, cap)
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.weightByConditioningSetSize(on: bool) : FeatureConf =
            let name = "WeightByConditioningSetSize"
            let cap : Feature.Capability = { enabled = true; kind = Feature.ConfigKind.Misc; runner = nop }
            if on then
                FeatureConf(
                    _config.Add(name, cap)
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self
        member self.spectralRanking(on: bool) : FeatureConf =
            let name = "SpectralRanking"
            let cap : Feature.Capability = { enabled = true; kind = Feature.ConfigKind.Misc; runner = nop }
            if on then
                FeatureConf(
                    _config.Add(name, cap)
                )
            else
                if _config.ContainsKey(name) then
                    FeatureConf(
                        _config.Remove(name)
                    )
                else
                    self

        // getters
        member self.FeatureByName
            with get(name) = _features.[name]
        member self.EnabledFeatures
            with get() : string[] = 
                _config |>
                    Map.toArray |>
                    Array.choose (fun (fname,cap) ->
                                    if cap.enabled && cap.kind = Feature.ConfigKind.Feature then
                                        Some fname
                                    else None)
        member self.EnabledScopes
            with get() : Scope.Selector[] =
                _config |>
                    Map.toArray |>
                    Array.choose (fun (confname,cap) ->
                                    if cap.enabled && cap.kind = Feature.ConfigKind.Scope then
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
        member self.IsEnabledAnalyzeAllCells : bool = _config.ContainsKey "AnalyzeOnlyFormulas" && not (_config.["AnalyzeOnlyFormulas"].enabled)

        // make sure that config option combinations make sense;
        // returns a 'corrected' config
        member self.validate : FeatureConf =
            let config = if self.IsEnabledSpectralRanking then
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

            // for now, our feature is always the mixed shallow L2 norm
            config'.enableShallowInputVectorMixedL2NormSum(true)

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

        static member simpleConf(m: Map<string,Feature.Capability>) : Map<string,bool> =
            Map.map (fun (k: string)(v: Feature.Capability) -> v.enabled) m

        override self.Equals(obj: Object) : bool =
            let other_fc = obj :?> FeatureConf
            (FeatureConf.simpleConf _config) = (FeatureConf.simpleConf other_fc.rawConf)
