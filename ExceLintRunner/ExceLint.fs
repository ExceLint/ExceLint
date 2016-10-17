module ExceLint

    open System
    open System.Collections.Generic

    type Address = string
    type Worksheet = string
    type Workbook = string
    type Path = string

    type BugKinds =
    | NotABug
    | ReferenceBug
    | ConstantWhereFormulaExpected
        static member ToKind(kindstr: string) =
            match kindstr with
            | "no" -> NotABug
            | "ref" -> ReferenceBug
            | "cwfe" -> ConstantWhereFormulaExpected
            | _ -> failwith "Unknown bug type '" + kindstr + "'"

    let ExceLintGTToAddress(addrstr: Address)(worksheetname: Worksheet)(workbookname: Workbook)(path: Path) : AST.Address =
        // we force the mode to absolute because
        // that's how Depends reads them
        AST.Address.FromA1StringForceMode(
            addrstr.ToUpper(),
            AST.AddressMode.Absolute,
            AST.AddressMode.Absolute,
            worksheetname,
            (if workbookname.EndsWith(".xls") then workbookname else workbookname + ".xls"),
            IO.Path.GetFullPath(path)   // ensure absolute path
        )

    type GroundTruth(gtpath: string) =
        let raw = CSV.OurGroundTruth.Load(gtpath)

        let k = new Dictionary<AST.Address,BugKinds>()

        do
            Seq.iter (fun (row: CSV.OurGroundTruth.Row) ->
                let addr = ExceLintGTToAddress row.Addr row.Worksheet row.Workbook row.Path
                k.Add(addr, BugKinds.ToKind(row.Bug_kind))
            ) raw.Rows

        member self.IsABug(addr: AST.Address) = if k.ContainsKey addr then k.[addr] <> BugKinds.NotABug else false
        member self.TrueRefBugsByWorkbook(workbookname: Workbook) =
            new HashSet<AST.Address>(
                Seq.filter (fun (pair: KeyValuePair<AST.Address,BugKinds>) ->
                    pair.Key.WorkbookName = workbookname
                ) k
                |> Seq.map (fun (pair: KeyValuePair<AST.Address,BugKinds>) -> pair.Key)
            )
