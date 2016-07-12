module Scope
//    type SelectID = (int option*int option*string option)

    type XYPath = {
        x: int option;
        y: int option;
        fullpath: string option;
    }

    type SelectID =
    | AllID of XYPath
    | ColumnID of XYPath
    | RowID of XYPath
        
    type Selector =
    | AllCells
    | SameColumn
    | SameRow
        // the selector ID is a hash value that says how to
        // compute conditional distributions.  E.g., if addr1
        // and addr2 have the same SameColumn ID, then they are
        // in the same column.
        member self.id(addr: AST.Address) : SelectID =
            match self with
            | AllCells -> AllID ({ x = None; y = None; fullpath = None })
            | SameColumn -> ColumnID { x = Some addr.X; y = None; fullpath = Some (addr.Path + ":" + addr.WorkbookName + ":" + addr.WorksheetName) }
            | SameRow -> RowID { x = None; y = Some addr.Y; fullpath = Some (addr.Path + ":" + addr.WorkbookName + ":" + addr.WorksheetName)}
        static member ToPretty(id: SelectID) : string =
            match id with
            | AllID(_) -> "AllCells"
            | ColumnID(xyp) -> "Column " + xyp.x.ToString()
            | RowID(xyp) -> "Row " + xyp.y.ToString()
        static member Kinds = [| AllCells; SameColumn; SameRow |]
