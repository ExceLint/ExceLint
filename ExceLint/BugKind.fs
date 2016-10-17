namespace ExceLint
    type BugKind =
    | NotABug
    | ReferenceBug
    | ReferenceBugInverse
    | ConstantWhereFormulaExpected
        static member ToKind(kindstr: string) =
            match kindstr with
            | "no" -> NotABug
            | "ref" -> ReferenceBug
            | "refi" -> ReferenceBugInverse
            | "cwfe" -> ConstantWhereFormulaExpected
            | _ -> failwith ("Unknown bug type '" + kindstr + "'")
        override self.ToString() : string =
            match self with
            | NotABug -> "Not a bug"
            | ReferenceBug -> "Reference bug"
            | ReferenceBugInverse -> "Reference bug (inverse)"
            | ConstantWhereFormulaExpected -> "Constant where formula expected"
        member self.ToLog() : string =
            match self with
            | NotABug -> "no"
            | ReferenceBug -> "ref"
            | ReferenceBugInverse -> "refi"
            | ConstantWhereFormulaExpected -> "cwfe"
        static member AllKinds : BugKind[] =
            Microsoft.FSharp.Reflection.FSharpType.GetUnionCases typeof<BugKind>
            |> Array.map (fun uci -> BugKind.ToKind uci.Name)
        static member DefaultKind : BugKind = NotABug