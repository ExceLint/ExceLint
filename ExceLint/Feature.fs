namespace Feature
    open Depends
    open System

    type ConfigKind =
    | Feature
    | Scope
    | Misc
    
    type ICountable<'T> =
       abstract member Add: 'T -> 'T
       abstract member MeanFoldDefault: 'T
       abstract member Negate: 'T
       abstract member ScalarDivide: double -> 'T
       abstract member VectorMultiply: 'T -> double
       abstract member Sqrt: 'T
       abstract member Abs: 'T

    type Capability<'T> = { enabled : bool; kind: ConfigKind; runner: AST.Address -> Depends.DAG -> ICountable<'T>; }

    type FeatureLambda<'T> = AST.Address -> DAG -> ICountable<'T>

    type BaseFeature<'T>() =
        static member run (cell: AST.Address) (dag: DAG): ICountable<'T> = failwith "Feature must provide run method."
        static member capability : string*Capability<'T> = failwith "Feature must provide capability."