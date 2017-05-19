namespace ExceLint
    open System.Collections.Generic
    open System.Collections
    open System
    open Utils
    open ConfUtils
    open CommonTypes
    open CommonFunctions
    open ClusterModelBuilder
    open COFModelBuilder
    open SpectralModelBuilder

        module ModelBuilder =
            let VisualizeLSH(app: Microsoft.Office.Interop.Excel.Application)(config: FeatureConf)(dag: Depends.DAG)(alpha: double)(progress: Depends.Progress) =
                let config' = config.validate
                let input : Input = { app = app; config = config'; dag = dag; alpha = alpha; progress = progress; }
                LSHViz(input)

            let initStepClusterModel(app: Microsoft.Office.Interop.Excel.Application)(config: FeatureConf)(dag: Depends.DAG)(alpha: double)(progress: Depends.Progress) : ClusterModel =
                let config' = config.validate
                let input : Input = { app = app; config = config'; dag = dag; alpha = alpha; progress = progress; }
                ClusterModel input

            let analyze(app: Microsoft.Office.Interop.Excel.Application)(config: FeatureConf)(dag: Depends.DAG)(alpha: double)(progress: Depends.Progress) =
                let config' = config.validate

                let input : Input = { app = app; config = config'; dag = dag; alpha = alpha; progress = progress; }

                if dag.IsCancelled() then
                    None
                elif input.config.IsCOF then
                    let pipeline = runCOFModel              // produce initial (unsorted) ranking
                                    +> weights              // compute weights
                                    +> reweightRanking      // modify ranking scores
                                    +> canonicalSort        // sort
                                    +> cutoffIndex          // compute initial cutoff index
                                    +> kneeIndexOpt         // optionally compute knee index
                                    +> inferAddressModes    // remove anomaly candidates
                                    +> canonicalSort
                                    +> kneeIndexOpt

                    match pipeline input with
                    | Success(analysis) -> Some (ErrorModel(input, analysis, config'))
                    | Cancellation -> None
                elif input.config.Cluster then
                    let pipeline = runClusterModel

                    match pipeline input with
                    | Success(analysis) -> Some (ErrorModel(input, analysis, config'))
                    | Cancellation -> None
                    | CantRun msg -> raise (NoFormulasException msg)
                else
                    let pipeline = runSpectralModel                 // produce initial (unsorted) ranking
                                    +> weights              // compute weights
                                    +> reweightRanking      // modify ranking scores
                                    +> canonicalSort        // sort
                                    +> cutoffIndex          // compute initial cutoff index
                                    +> kneeIndexOpt         // optionally compute knee index
                                    +> inferAddressModes    // remove anomaly candidates
                                    +> canonicalSort
                                    +> kneeIndexOpt

                    match pipeline input with
                    | Success(analysis) -> Some (ErrorModel(input, analysis, config'))
                    | Cancellation -> None