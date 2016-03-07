module Scope
//    type SelectID(col: int option, row: int option, path: string option) =
//        member self.ToString() =
//            match col with
//            | Some(c) -> "Column " + c.ToString() + " in " + path.Value
//            | None ->
//                match row with
//                | Some(r) -> "Row " + r.ToString() + " in " + path.Value
//                | None ->
//                    match path with
//                    | Some(p) -> "All cells in " + p
//                    | None -> "All cells"

    type SelectID = (int option*int option*string option)
    
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
            | AllCells -> None, None, None
            | SameColumn -> Some addr.X, None, Some (addr.Path + ":" + addr.WorkbookName + ":" + addr.WorksheetName)
            | SameRow -> None, Some addr.Y, Some (addr.Path + ":" + addr.WorkbookName + ":" + addr.WorksheetName)
        static member Kinds = [| AllCells; SameColumn; SameRow |]