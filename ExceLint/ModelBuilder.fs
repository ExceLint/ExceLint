namespace ExceLint
    open System.Collections.Generic
    open System.Collections
    open System
    open Utils
    open ConfUtils
    open CommonTypes
    open CommonFunctions
    open ClusterModelBuilder
    open EntropyModelBuilder
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

            let initEntropyModel(app: Microsoft.Office.Interop.Excel.Application)(config: FeatureConf)(dag: Depends.DAG)(progress: Depends.Progress) : EntropyModel =
                let config' = config.validate
                let input : Input = { app = app; config = config'; dag = dag; alpha = 0.00; progress = progress; }
                EntropyModel.Initialize input

            let analyze(app: Microsoft.Office.Interop.Excel.Application)(config: FeatureConf)(dag: Depends.DAG)(alpha: double)(progress: Depends.Progress) =
                let config' = config.validate

                let input : Input = { app = app; config = config'; dag = dag; alpha = alpha; progress = progress; }

                if dag = null then
                    None
                else
                    let pipe =
                        // COF clustering
                        if input.config.IsCOF then
                            let pipeline = runCOFModel              // produce initial (unsorted) ranking
                                            +> weights              // compute weights
                                            +> reweightRanking      // modify ranking scores
                                            +> canonicalSort        // sort
                                            +> cutoffIndex          // compute initial cutoff index
                                            +> kneeIndexOpt         // optionally compute knee index
                                            +> inferAddressModes    // remove anomaly candidates
                                            +> canonicalSort
                                            +> kneeIndexOpt
                            pipeline input 
                        // LSH-NN clustering
                        elif input.config.LSHNNCluster then
                            runClusterModel input
                        // NN clustering
                        elif input.config.OldCluster then
                            OldClusterModel.runClusterModel input
                        // entropy clustering
                        elif input.config.Cluster then
                            EntropyModel.runClusterModel input
                        else
                        // spectral clustering
                            let pipeline = runSpectralModel         // produce initial (unsorted) ranking
                                            +> weights              // compute weights
                                            +> reweightRanking      // modify ranking scores
                                            +> canonicalSort        // sort
                                            +> cutoffIndex          // compute initial cutoff index
                                            +> kneeIndexOpt         // optionally compute knee index
                                            +> inferAddressModes    // remove anomaly candidates
                                            +> canonicalSort
                                            +> kneeIndexOpt
                            pipeline input 
                    // handle errors
                    match pipe with
                    | Success(analysis) -> Some (ErrorModel(input, analysis, config'))
                    | Cancellation -> None
                    | CantRun msg -> raise (NoFormulasException msg)