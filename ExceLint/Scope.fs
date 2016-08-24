namespace ExceLint

module Scope =
    open System.Collections.Generic
    open Utils

    type Path = string

    [<CustomEquality; CustomComparison>]
    type XYPath = {
        x: int option;
        y: int option;
        fullpath: string option;
    }
    with
        override self.Equals(obj: obj) =
            let other = obj :?> XYPath
            self.x = other.x &&
            self.y = other.y &&
            self.fullpath = other.fullpath
        override self.GetHashCode() : int =
            let xhc = match self.x with
                        | Some(i) -> i.GetHashCode()
                        | None -> 0

            let yhc = match self.y with
                        | Some(i) -> i.GetHashCode()
                        | None -> 0

            let phc = match self.fullpath with
                        | Some(s) -> s.GetHashCode()
                        | None -> 0
            xhc||| yhc ||| phc
        interface System.IComparable with
            member self.CompareTo obj =
                let other = obj :?> XYPath
                if self.x < other.x then
                    -1
                else if self.x = other.x then
                    0
                else
                    1

    type Level = {
        fn_x: int;
        fn_y: int;
        fn_fullpath: string;
        level: int
    }
    with
        override self.ToString() =
            "{ fn_x: " + self.fn_x.ToString() + "; fn_y: " + self.fn_y.ToString() + "; fn_fullpath: " + self.fn_fullpath + "; level: " + self.level.ToString() + " }"

    [<CustomEquality; CustomComparison>]
    type SelectID =
    | AllID
    | ColumnID of XYPath
    | RowID of XYPath
    | LevelID of Set<Level>
    | SheetID of Path
        member self.IsKind : Selector =
            match self with
            | AllID -> Selector.AllCells
            | ColumnID _ -> Selector.SameColumn
            | RowID _ -> Selector.SameRow
            | LevelID _ -> Selector.SameLevel
            
        override self.Equals(obj: obj) : bool =
            let other = obj :?> SelectID

            match self with
            | AllID -> true
            | ColumnID(xyp1) ->
                match other with
                | ColumnID(xyp2) -> xyp1 = xyp2
                | _ -> false
            | RowID(xyp1) ->
                match other with
                | RowID(xyp2) -> xyp1 = xyp2
                | _ -> false
            | LevelID(lv1) ->
                match other with
                | LevelID(lv2) -> not (Set.isEmpty (Set.intersect lv1 lv2))
                | _ -> false
            | SheetID(p1) ->
                match other with
                | SheetID(p2) -> p1 = p2
                | _ -> false
        override self.GetHashCode() : int =
            match self with
            | AllID -> 0
            | ColumnID(xyp) -> xyp.GetHashCode()
            | RowID(xyp) -> xyp.GetHashCode()
            | LevelID(lv) -> lv.GetHashCode()
        interface System.IComparable with
            member self.CompareTo(obj: obj) : int =
                let other = obj :?> SelectID
                match self,other with
                | AllID,AllID -> 0
                | ColumnID(xyp1),ColumnID(xyp2) -> 
                    (xyp1 :> System.IComparable).CompareTo(xyp2)
                | RowID(xyp1),RowID(xyp2) ->
                    (xyp1 :> System.IComparable).CompareTo(xyp2)
                | LevelID(lv1),LevelID(lv2) ->
                    if lv1.Count < lv2.Count then
                        -1
                    else if lv1.Count = lv2.Count then
                        0
                    else
                        1
                | _,_ -> failwith "incomparable"

    and Selector =
    | AllCells
    | SameColumn
    | SameRow
    | SameLevel
    | SameSheet
        // the selector ID is a hash value that says how to
        // compute conditional distributions.  E.g., if addr1
        // and addr2 have the same SameColumn ID, then they are
        // in the same column.
        member self.id(addr: AST.Address)(dag: Depends.DAG)(cache: SelectorCache) : SelectID =
            cache.fetchOrStore addr dag self
        static member ToPretty(id: SelectID) : string =
            match id with
            | AllID(_) -> "All Cells"
            | ColumnID(xyp) -> "Column " + AST.Address.IntToColChars(xyp.x.Value)
            | RowID(xyp) -> "Row " + xyp.y.Value.ToString()
            | LevelID(levels) ->
                if levels.Count = 0 then
                    ""
                else
                    let lstrs = Set.map (fun l -> l.ToString()) levels |> Set.toList
                    "Levels [" + (List.reduce (fun (acc: string)(l: string) -> acc + ", " + l.ToString()) lstrs) + "]"
            | SheetID(p) -> "Sheet: " + p
        static member Kinds = [| AllCells; SameColumn; SameRow; SameLevel |]
        override self.ToString() : string =
            match self with
            | AllCells -> "All Cells"
            | SameColumn -> "Same Column"
            | SameRow -> "Same Row"
            | SameLevel -> "Same Level"
            | SameSheet -> "Same Sheet"

    and SelectorCache() =
        let levelsOf(input: AST.Address)(dag: Depends.DAG) : Set<Level> =
            let terminal_formulas = new HashSet<AST.Address>(dag.terminalFormulaNodes false)
            let unfilt_refdepts = dag.AllRefDistancesFromInput input
            let refdepths = Seq.filter (fun (kvp: KeyValuePair<AST.Address,HashSet<int>>) ->
                                terminal_formulas.Contains kvp.Key
                            ) unfilt_refdepts

            Seq.map (fun (kvp: KeyValuePair<AST.Address,HashSet<int>>) ->
                let faddr = kvp.Key;
                let distances = kvp.Value
                Seq.map (fun (d: int) ->
                    { fn_x = faddr.X; fn_y = faddr.Y; fn_fullpath = faddr.Path + ":" + faddr.WorkbookName + ":" + faddr.WorksheetName; level = d }
                ) distances
            ) refdepths |> Seq.concat |> Set.ofSeq

        let _cache = new Dict<Selector,Dict<AST.Address,SelectID>>()

        let fullpath(addr: AST.Address) : Path =
            addr.Path + ":" + addr.WorkbookName + ":" + addr.WorksheetName

        member self.fetchOrStore(addr: AST.Address)(dag: Depends.DAG)(sel: Selector) : SelectID =
            if not (_cache.ContainsKey(sel)) then
                _cache.Add(sel, new Dict<AST.Address,SelectID>())

            if not (_cache.[sel].ContainsKey(addr)) then
                let sID = match sel with
                            | AllCells -> AllID
                            | SameColumn -> ColumnID { x = Some addr.X; y = None; fullpath = Some (fullpath addr) }
                            | SameRow -> RowID { x = None; y = Some addr.Y; fullpath = Some (fullpath addr)}
                            | SameLevel -> LevelID (levelsOf addr dag)
                            | SameSheet -> SheetID (fullpath addr)
                _cache.[sel].Add(addr, sID)

                sID
            else
                _cache.[sel].[addr]

    and SelectIDCache = Dict<SelectID,Set<AST.Address>>

