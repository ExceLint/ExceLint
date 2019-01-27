namespace ExceLint
    open CommonTypes
    open EntropyModelBuilder2
    open FastDependenceAnalysis
        module ModelBuilder =
            type App = Microsoft.Office.Interop.Excel.Application

            let initEntropyModel2(app: App)(config: FeatureConf)(g: Graph)(progress: Progress) : EntropyModel2 =
                let config' = config.validate
                let input : Input = { app = app; config = config'; dag = g; alpha = config.Threshold; progress = progress; }
                EntropyModel2.Initialize input
            
            // for old references
            let initEntropyModel(app: App)(config: FeatureConf)(g: Graph)(progress: Progress) : EntropyModel2 =
                initEntropyModel2 app config g progress

            let analyze(app: App)(config: FeatureConf)(g: Graph)(alpha: double)(progress: Progress) =
                let config' = config.validate

                let input : Input = { app = app; config = config'; dag = g; alpha = alpha; progress = progress; }

                if g = null then
                    None
                else
                    // handle errors
                    match EntropyModel2.runClusterModel input with
                    | Success(analysis) -> Some (ErrorModel(input, analysis, config'))
                    | Cancellation -> None
                    | CantRun msg -> None