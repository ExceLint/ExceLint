namespace ExceLint
    open Utils
    open CommonTypes

    module CommonFunctions =
            // _analysis_base specifies which cells should be ranked:
            // 1. allCells means all cells in the spreadsheet
            // 2. onlyFormulas means only formulas
            let analysisBase(config: FeatureConf)(d: Depends.DAG) : AST.Address[] =
                let cs = if config.IsEnabled("AnalyzeOnlyFormulas") then
                            d.getAllFormulaAddrs()
                         else
                            d.allCells()
                let cs' = match config.IsLimitedToSheet with
                          | Some(wsname) -> cs |> Array.filter (fun addr -> addr.A1Worksheet() = wsname)
                          | None -> cs 
                cs'

            let runEnabledFeatures(cells: AST.Address[])(dag: Depends.DAG)(config: FeatureConf)(progress: Depends.Progress) =
                config.EnabledFeatures |>
                Array.map (fun fname ->
                    // get feature lambda
                    let feature = config.FeatureByName fname

                    let fvals =
                        Array.map (fun cell ->
                            if progress.IsCancelled() then
                                raise AnalysisCancelled

                            progress.IncrementCounter()
                            cell, feature cell dag
                        ) cells
                    
                    fname, fvals
                ) |> adict

            let invertedHistogram(scoretable: ScoreTable)(selcache: Scope.SelectorCache)(dag: Depends.DAG)(config: FeatureConf) : InvertedHistogram =
                assert (config.EnabledScopes.Length = 1 && config.EnabledFeatures.Length = 1)

                let d = new Dict<AST.Address,HistoBin>()

                Array.iter (fun fname ->
                    Array.iter (fun (sel: Scope.Selector) ->
                        Array.iter (fun (addr: AST.Address, score: Countable) ->
                            // fetch SelectID for this selector and address
                            let sID = sel.id addr dag selcache

                            // get binname
                            let binname = fname,sID,score

                            d.Add(addr,binname)
                        ) (scoretable.[fname])
                    ) (config.EnabledScopes)
                ) (config.EnabledFeatures)

                new InvertedHistogram(d)

            let centroid(c: seq<AST.Address>)(ih: InvertedHistogram) : Countable =
                c
                |> Seq.map (fun a ->
                    let (_,_,c) = ih.[a]    // get histobin for address
                    c                       // get countable from bin
                   )
                |> Countable.Mean               // get mean

            let cartesianProductByX(xset: Set<'a>)(yset: Set<'a>) : ('a*'a[]) list =
                // cartesian product, grouped by the first element,
                // excluding the element itself
                Set.map (fun x -> x, (Set.difference yset (Set.ofList [x])) |> Set.toArray) xset |> Set.toList

            let cartesianProduct(xs: seq<'a>)(ys: seq<'b>) : seq<'a*'b> =
                xs |> Seq.collect (fun x -> ys |> Seq.map (fun y -> x, y))

            let ToCountable(a: AST.Address)(ih: InvertedHistogram) : Countable =
                let (_,_,v) = ih.[a]
                v  