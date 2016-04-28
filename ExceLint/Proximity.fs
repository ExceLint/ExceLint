module Proximity

open Feature
open Depends
open ExceLint.Vector


type Outcome =
| Same = 1
| NotSame = 0
| Boundary = -1

let cellAbove(cell: AST.Address) : AST.Address option =
    if cell.Row - 1 < 1 then
        None
    else
        Some(AST.Address.fromR1C1withMode(cell.Row - 1, cell.Col, AST.AddressMode.Absolute, AST.AddressMode.Absolute, cell.WorksheetName, cell.WorkbookName, cell.Path))

let cellBelow(cell: AST.Address) : AST.Address option =
    // there is no bottom boundary
    Some(AST.Address.fromR1C1withMode(cell.Row + 1, cell.Col, AST.AddressMode.Absolute, AST.AddressMode.Absolute, cell.WorksheetName, cell.WorkbookName, cell.Path))

let cellLeft(cell: AST.Address) : AST.Address option =
    if cell.Col - 1 < 1 then
        None
    else
        Some(AST.Address.fromR1C1withMode(cell.Row, cell.Col - 1, AST.AddressMode.Absolute, AST.AddressMode.Absolute, cell.WorksheetName, cell.WorkbookName, cell.Path))

let cellRight(cell: AST.Address) : AST.Address option =
    // there is no right boundary
    Some(AST.Address.fromR1C1withMode(cell.Row, cell.Col + 1, AST.AddressMode.Absolute, AST.AddressMode.Absolute, cell.WorksheetName, cell.WorkbookName, cell.Path))

let cellProximal(cell: AST.Address)(dag: DAG)(selector: AST.Address -> AST.Address option) : double =
    match selector cell with
        | Some(cell') -> 
            let thisHash = ShallowInputVectorMixedL2NormSum.run cell dag
            let proxHash = ShallowInputVectorMixedL2NormSum.run cell' dag
            if thisHash = proxHash then
                double(Outcome.Same)
            else
                double(Outcome.NotSame)
        | None ->
            double(Outcome.Boundary)

type Above() =
    inherit BaseFeature()
    static member run(cell: AST.Address)(dag: DAG) : double =
        cellProximal cell dag cellAbove
    static member capability : string*Capability =
        (typeof<Above>.Name,
            { enabled = false; kind = ConfigKind.Feature; runner = Above.run } )

type Below() =
    inherit BaseFeature()
    static member run(cell: AST.Address)(dag: DAG) : double =
        cellProximal cell dag cellBelow
    static member capability : string*Capability =
        (typeof<Below>.Name,
            { enabled = false; kind = ConfigKind.Feature; runner = Below.run } )

type Left() =
    inherit BaseFeature()
    static member run(cell: AST.Address)(dag: DAG) : double =
        cellProximal cell dag cellLeft
    static member capability : string*Capability =
        (typeof<Left>.Name,
            { enabled = false; kind = ConfigKind.Feature; runner = Left.run } )

type Right() =
    inherit BaseFeature()
    static member run(cell: AST.Address)(dag: DAG) : double =
        cellProximal cell dag cellRight
    static member capability : string*Capability =
        (typeof<Right>.Name,
            { enabled = false; kind = ConfigKind.Feature; runner = Right.run } )