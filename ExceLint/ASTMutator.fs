module ASTMutator

    let rec mutateExpr(e: AST.Expression)(ref: AST.Address)(newCMode: AST.AddressMode)(newRMode: AST.AddressMode) : AST.Expression =
        match e with
        | AST.ReferenceExpr(re) -> AST.ReferenceExpr(mutateReference re ref newCMode newRMode)
        | AST.BinOpExpr(op,e1,e2) -> AST.BinOpExpr(op, mutateExpr e1 ref newCMode newRMode, mutateExpr e2 ref newCMode newRMode)
        | AST.UnaryOpExpr(op,e1) -> AST.UnaryOpExpr(op, mutateExpr e1 ref newCMode newRMode)
        | AST.ParensExpr(e1) -> mutateExpr e1 ref newCMode newRMode

    and mutateReference(re: AST.Reference)(ref: AST.Address)(newCMode: AST.AddressMode)(newRMode: AST.AddressMode) : AST.Reference =
        match re with
        | :? AST.ReferenceRange as rr -> rr :> AST.Reference
        | :? AST.ReferenceAddress as ra -> mutateAddress ra ref newCMode newRMode :> AST.Reference
        | :? AST.ReferenceFunction as rf -> rf :> AST.Reference
        | :? AST.ReferenceConstant as rc -> rc :> AST.Reference
        | :? AST.ReferenceString as rs -> rs :> AST.Reference
        | :? AST.ReferenceNamed as rn -> rn :> AST.Reference

    and mutateAddress(ra: AST.ReferenceAddress)(ref: AST.Address)(newCMode: AST.AddressMode)(newRMode: AST.AddressMode) : AST.ReferenceAddress =
        if ra.Address = ref then
            let env = AST.Env(ra.Path, ra.WorkbookName, ra.WorksheetName)
            let addr' = AST.Address(ref.Row, ref.Col, newRMode, newCMode, env)
            AST.ReferenceAddress(env, addr')
        else
            ra