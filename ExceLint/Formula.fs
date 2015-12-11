module Formula

    open Feature
    open Depends
    open Parcel

    type MaxNestingDepth() = 
        inherit BaseFeature()

        static member private getMaxNestingDepth (fAddr: AST.Address) (dag: DAG) =
            let rec findMaxPath (formulaAST: AST.Expression) (depth: int) =
                match formulaAST with
                | ReferenceExpr _ -> depth 
                | BinOpExpr _ e1 e2 -> 1 + max (findMaxPath e1 depth) (findMaxPath e2 depth)
                | UnaryOpExpr _ e -> 1 + findMaxPath e
                | ParensExpr e -> 1 + findMaxPath e
            in findMaxPath (parseFormulaAtAddress fAddr dag) 0
           
        static member run cell dag =
            MaxNestingDepth.getMaxNestingDepth cell dag