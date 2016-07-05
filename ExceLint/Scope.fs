module Scope
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
        static member ToPretty(id: SelectID) : string =
            match id with
            | None,None,None -> "AllCells"
            | Some(x),None,Some(path) -> "Column " + x.ToString()
            | None,Some(y),Some(path) -> "Row " + y.ToString()
            | _ -> failwith "Unknown selector"
        static member Kinds = [| AllCells; SameColumn; SameRow |]
