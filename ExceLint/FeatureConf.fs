namespace ExceLint
    open System.Collections.Generic
    open System.Collections
    open System

    module ConfUtils =
        // we're using C# Dictionary instead of F# map
        // for debugging (it's inspectable) and speed purposes
        type Dict<'a,'b> = Dictionary<'a,'b>
        let adict(a: seq<('a*'b)>) = new Dict<'a,'b>(a |> dict)

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
        member self.enableInDegree() : FeatureConf =
            FeatureConf(
                let (name,cap) = Degree.InDegree.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableOutDegree() : FeatureConf =
            FeatureConf(
                let (name,cap) = Degree.OutDegree.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableCombinedDegree() : FeatureConf =
            FeatureConf(
                let (name,cap) = Degree.CombinedDegree.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableDeepInputVectorRelativeL2NormSum() : FeatureConf =
            FeatureConf(
                let (name,cap) = Vector.DeepInputVectorRelativeL2NormSum.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableDeepOutputVectorRelativeL2NormSum() : FeatureConf =
            FeatureConf(
                let (name,cap) = Vector.DeepOutputVectorRelativeL2NormSum.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableDeepInputVectorAbsoluteL2NormSum() : FeatureConf =
            FeatureConf(
                let (name,cap) = Vector.DeepInputVectorAbsoluteL2NormSum.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableDeepOutputVectorAbsoluteL2NormSum() : FeatureConf =
            FeatureConf(
                let (name,cap) = Vector.DeepOutputVectorAbsoluteL2NormSum.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableDeepInputVectorMixedL2NormSum() : FeatureConf =
            FeatureConf(
                let (name,cap) = Vector.DeepInputVectorMixedL2NormSum.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableDeepOutputVectorMixedL2NormSum() : FeatureConf =
            FeatureConf(
                let (name,cap) = Vector.DeepOutputVectorMixedL2NormSum.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableShallowInputVectorRelativeL2NormSum() : FeatureConf =
            FeatureConf(
                let (name,cap) = Vector.ShallowInputVectorRelativeL2NormSum.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableShallowOutputVectorRelativeL2NormSum() : FeatureConf =
            FeatureConf(
                let (name,cap) = Vector.ShallowOutputVectorRelativeL2NormSum.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableShallowInputVectorAbsoluteL2NormSum() : FeatureConf =
            FeatureConf(
                let (name,cap) = Vector.ShallowInputVectorAbsoluteL2NormSum.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableShallowOutputVectorAbsoluteL2NormSum() : FeatureConf =
            FeatureConf(
                let (name,cap) = Vector.ShallowOutputVectorAbsoluteL2NormSum.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableShallowInputVectorMixedL2NormSum() : FeatureConf =
            FeatureConf(
                let (name,cap) = Vector.ShallowInputVectorMixedL2NormSum.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableShallowOutputVectorMixedL2NormSum() : FeatureConf =
            FeatureConf(
                let (name,cap) = Vector.ShallowOutputVectorMixedL2NormSum.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableProximityAbove() : FeatureConf =
            FeatureConf(
                let (name,cap) = Proximity.Above.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableProximityBelow() : FeatureConf =
            FeatureConf(
                let (name,cap) = Proximity.Below.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableProximityLeft() : FeatureConf =
            FeatureConf(
                let (name,cap) = Proximity.Left.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.enableProximityRight() : FeatureConf =
            FeatureConf(
                let (name,cap) = Proximity.Right.capability
                _config.Add(name, { enabled = true; kind = cap.kind; runner = cap.runner })
            )
        member self.analyzeRelativeToAllCells() : FeatureConf =
            FeatureConf(
                let name = "ScopeAllCells"
                let cap : Feature.Capability = { enabled = true; kind = Feature.ConfigKind.Scope; runner = nop}
                _config.Add(name, cap)
            )
        member self.analyzeRelativeToColumns() : FeatureConf =
            FeatureConf(
                let name = "ScopeColumns"
                let cap : Feature.Capability = { enabled = true; kind = Feature.ConfigKind.Scope; runner = nop}
                _config.Add(name, cap)
            )
        member self.analyzeRelativeToRows() : FeatureConf =
            FeatureConf(
                let name = "ScopeRows"
                let cap : Feature.Capability = { enabled = true; kind = Feature.ConfigKind.Scope; runner = nop}
                _config.Add(name, cap)
            )
        member self.inferAddressModes() : FeatureConf =
            FeatureConf(
                let name = "InferAddressModes"
                let cap : Feature.Capability = { enabled = true; kind = Feature.ConfigKind.Misc; runner = nop }
                _config.Add(name, cap)
            )

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
                                        | _ -> failwith "Unknown scope selector."
                                    else None)

        member self.IsEnabled(name: string) : bool =
            _config.ContainsKey name && _config.[name].enabled