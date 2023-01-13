module Spreadsheet

open System
open Utils
open Parser
open System.IO

type CellValue = Expr of string | Str of string | IntValue of int | DateValue of DateTime | DoubleValue of double | Empty with
    member this.forDisplay =
        match this with
            | Expr(s) -> s
            | Str(s) -> sprintf "\"%s\"" s
            | IntValue(i) -> i.ToString()
            | DateValue(d) -> 
                if (d.TimeOfDay.Milliseconds) = 0 then d.ToShortDateString()
                else d.ToString()
            | DoubleValue(d) -> sprintf "%g" d
            | Empty -> ""

type RValue = 
    | NoRValue
    | StringRValue of string
    | DoubleRValue of Double
    | DateRValue of DateTime
    | BoolRValue of bool
    | ErrorRValue of string
    | CellRangeRValue of string
    member this.forDisplay =
        match this with
        | NoRValue -> ""
        | StringRValue(s) -> s
        | DoubleRValue(d) -> sprintf "%g" d
        | DateRValue(d) -> 
            if (d.TimeOfDay.Milliseconds) = 0 then d.ToShortDateString()
            else d.ToString()
        | BoolRValue(b) ->
            if b then "True"
            else "False"
        | ErrorRValue(s) -> "!" + s
        | CellRangeRValue(s) -> s


type RowName = string
type ColName = string 
type CellName = { Row: RowName; Col: ColName }

type Col = { Name: string }

type Cell = { Contents: CellValue; mutable Value: RValue; Column: Col; ExtraColumns: Col list }

type Row = { Name: string; Cells: Cell list }

type Sheet = { Id: int; Name: string; Rows: Row list; Cols: Col list }

type CellRange = 
| SingleCell of CellName
| RectangularRange of StartCell:CellName * EndCell:CellName
| SingleCol of ColName

type RowColPair = { Row: Row; Col: Col }

type NormalisedRectangularRange = { TopLeft: RowColPair; BottomRight: RowColPair }

let addColumn sheet col  =
    { sheet with Cols = sheet.Cols @ [col] }

let addRow sheet row  =
    { sheet with Rows = sheet.Rows @ [row] }

let findColumn sheet (name:string) =
    let n = name.ToUpper()
    sheet.Cols |> List.tryFind( fun(x) -> x.Name = n )

let lookupColumn sheet (name:string) =
    let n = name.ToUpper()
    sheet.Cols |> List.find( fun(x) -> x.Name = n )

let findCellInRow row (name:string) =
    let n = name.ToUpper()
    row.Cells |> List.tryFind( fun(x) -> x.Column.Name = n )

let findCellInRowByCol row (col:Col) =
    row.Cells |> List.tryFind( fun(x) -> x.Column = col )

let findRow sheet name =
    sheet.Rows |> List.tryFind( fun(x) -> x.Name = name )

let lookupRow sheet name =
    sheet.Rows |> List.find( fun(x) -> x.Name = name )

let findCell sheet row (col:string) =
    let r = findRow sheet row
    match r with
        | Some foundRow ->
            let cell = findCellInRow foundRow col
            cell
        | None ->
            None



type NextParsedThing<'a> =
    | NothingGiven
    | NotThing of remaining:string
    | ParsedThing of item:'a * remaining:string
    | ParsedThingError of err:string * remaining:string

let snipWord str =
    let str1 = skipWhite(str)
    if String.length str1 = 0 then NothingGiven
    else
        let mutable i = 0
        let mutable cont = true
        while i < String.length str1 && cont do
            match str1.[i] with
                | t when Char.IsLetterOrDigit(t) -> i <- i + 1; cont <- true
                | _ -> cont <- false
        if i = 0 then NotThing(str1)
        else ParsedThing(str1.[0..i-1], str1.[i..])

type CellAtomName =
| ColOnly of colName:string * remaining:string
| ColRow of colName:string * rowName:string * remaining:string

let parseCellLocation str = 
    let l = String.length str
    let mutable p = 0
    while p < l && Char.IsLetter(str.[p]) do
        p <- p + 1
    let colName = str.[0..p-1]
    if (p >= l) || not ( Char.IsDigit str.[p] ) then ColOnly(colName, str.[p..])
    else
        let pd = p
        while p < l && Char.IsDigit(str.[p]) do
            p <- p + 1
        ColRow( colName, str.[pd..p-1 ], str.[p..])

type ValidCellReference(rowcol: RowColPair, cell: Cell option ) =
    member this.RowCol = rowcol
    member this.Cell = cell

type CellValidation =
    | RowDoesNotExist of rowname:RowName
    | ColDoesNotExist of colname:ColName
    | CellRefValid of validcell: ValidCellReference
    member this.display =
        match this with
        | RowDoesNotExist(r) -> sprintf "Row %s does not exist" r
        | ColDoesNotExist(c) -> sprintf "Column %s does not exist" c
        | CellRefValid(validcell) -> sprintf "%s%s" validcell.RowCol.Col.Name validcell.RowCol.Row.Name

let validateCellReference sheet (cr:CellName) =
    match findColumn sheet cr.Col with
    | None -> ColDoesNotExist(cr.Col)
    | Some col ->
        match findRow sheet cr.Row with
        | None -> RowDoesNotExist(cr.Row)
        | Some row ->
            CellRefValid( ValidCellReference( { Row = row; Col = col }, findCellInRowByCol row col ) )

let isCellRefValid sheet (cr:CellName) = 
    match validateCellReference sheet cr with
    | ColDoesNotExist( colname) -> false
    | RowDoesNotExist(r) -> false
    | CellRefValid(vcr) -> true

let normaliseRectangularRange sheet (sr:CellName) (er:CellName) =
    let srr = lookupRow sheet sr.Row
    let err = lookupRow sheet er.Row
    // get rows in order
    let (start_row, end_row) = if sr.Row = er.Row then (srr, err )
                               elif (sheet.Rows |> List.findIndex( fun f -> f = srr )) < (sheet.Rows |> List.findIndex( fun f -> f = err )) then (srr, err )
                               else (err, srr )
    let src = lookupColumn sheet sr.Col
    let erc = lookupColumn sheet er.Col
    let (start_col, end_col ) = if sr.Col = er.Col then (src, erc)
                                elif (sheet.Cols |> List.findIndex( fun f -> f = src )) < (sheet.Cols |> List.findIndex( fun f -> f = erc )) then (src, erc )
                                else (erc, src)
    { TopLeft = { Row = start_row; Col = start_col }; BottomRight = { Row = end_row; Col = end_col } }



let parseCellRange (str:string) =
    if not (Char.IsLetter str.[0] ) then NotThing( str)
    else
        let firstPart = parseCellLocation str
        match firstPart with
        | ColOnly(rs,rem) ->
            ParsedThing( SingleCol(rs) , rem)
        | ColRow(cs, rs, rem ) ->
            if String.length rem = 0 || not (rem.[0] = ':') then ParsedThing( SingleCell( { Row = rs; Col = cs } ), rem )
            else
                let str1 = rem.[1..] // past the :
                if String.length str1 = 0 || not (Char.IsLetter str.[0] )  then ParsedThingError("Incomplete range", str1 )
                else
                    let secondPart = parseCellLocation str1
                    match secondPart with
                    | ColOnly(re, rem1) ->
                        ParsedThingError("Incomplete range: cell followed by col", rem1 )
                    | ColRow(ce, re, rem1 ) ->
                        ParsedThing( RectangularRange( { Row = rs; Col = cs }, { Row = re; Col = ce } ), rem1 )

type GatheredCells = 
    | GatheredError of err:CellValidation
    | CellList of list:ValidCellReference list * range:NormalisedRectangularRange

let rec calcThingsBetween acc start finish list = 
    if start = finish then [start] // optimisation
    else 
        match list with
        | h::rest -> if h = start then calcThingsBetween [h] start finish rest
                     elif h = finish then acc @ [ h ] // the result
                     elif acc.IsEmpty then calcThingsBetween acc start finish rest // not got to start yet
                     else calcThingsBetween (acc @ [ h ]) start finish rest // in the middle
        | [] -> acc // I don't think we should get here unless first list is empty (which it can't)

// from a range produce a list of cells
let gatherCells sheet range =
    match range with
    | SingleCell(cr) ->
        let v = validateCellReference sheet cr
        match v with
        | CellRefValid(vcr) -> CellList( [vcr], { TopLeft=vcr.RowCol; BottomRight=vcr.RowCol })
        | _ ->  GatheredError( v )
    | SingleCol(colr) ->        
        match findColumn sheet colr with
        | None -> GatheredError (ColDoesNotExist colr)
        | Some col -> 
            CellList( sheet.Rows |> List.map ( fun(x) -> ValidCellReference( { Row =  x; Col = col } , findCellInRowByCol x col) ), { TopLeft={ Row=sheet.Rows.Head; Col=col}; BottomRight = {Row=sheet.Rows.[sheet.Rows.Length-1]; Col=col} } )
    | RectangularRange(sr,er) ->
        if not (isCellRefValid sheet sr) then GatheredError ( validateCellReference sheet sr )
        elif not (isCellRefValid sheet er) then GatheredError ( validateCellReference sheet er )
        else
            let nr = normaliseRectangularRange sheet sr er

            let rows = calcThingsBetween [] nr.TopLeft.Row nr.BottomRight.Row sheet.Rows
            let cols = calcThingsBetween [] nr.TopLeft.Col nr.BottomRight.Col sheet.Cols
            let rec calcCells rows cols acc =
                match rows with 
                | r::rest -> 
                    let rec calcCellsForRow row cols cellacc =
                        match cols with 
                        | c::rest ->
                            calcCellsForRow row rest (cellacc @  [ ValidCellReference( { Row = row; Col = c } , findCellInRowByCol row c ) ])
                        | [] -> cellacc
                    let acc_this = calcCellsForRow r cols acc
                    calcCells rest cols acc_this
                | [] -> acc
            CellList(calcCells rows cols [], nr )

let replaceRow (rows:Row list) (row:Row) =
    rows |> List.map( fun r -> if r.Name = row.Name then row else r )

let cellMerge sheet row startCol endCol =
    let all_cols_affected = calcThingsBetween [] startCol endCol sheet.Cols
    match all_cols_affected with
    | _::other_cols ->
        // if there is a cell for startCol then the contents of that becomes the new cell
        // if there isn't a cell for startCol then the new cell is empty
        let c = findCellInRowByCol row startCol
        let new_cell = 
            match c with
            | Some cell -> 
                // it may have already been merged. The new list will overwrite the old one as it must be bigger
                { cell with ExtraColumns = other_cols }
            | None -> 
                { Contents=Empty; Value=NoRValue; Column=startCol; ExtraColumns = other_cols }
        // we need to delete any cells in the row that are in range
        let remaining_cells = row.Cells |> List.filter( fun x -> not ( List.contains x.Column all_cols_affected ) )
        let new_rows = replaceRow sheet.Rows { row with Cells = remaining_cells @ [new_cell] }
        { sheet with Rows = new_rows }
    | [] ->             
        sheet // probably an error


let doAddCol sheet cmd =
    let c = snipWord cmd
    match c with
        | NothingGiven | ParsedThingError(_,_) -> 
            printfn "Expected col name"
            sheet
        | NotThing(rest) -> 
            printfn "Do not understand, expected col name"
            sheet
        | ParsedThing(cw, rest) -> 
            if findColumn sheet cw = None then
                let newSheet = addColumn sheet { Name = cw.ToUpper() }
                newSheet
            else
                printfn "Column exists %A" cw
                sheet

let doAddRow sheet cmd =
    let c = snipWord cmd
    match c with
        | NothingGiven | ParsedThingError(_,_) -> 
            printfn "Expected row name"
            sheet
        | NotThing(rest) -> 
            printfn "Do not understand, expected row name"
            sheet
        | ParsedThing(cw, rest) -> 
            if findRow sheet cw = None then
                addRow sheet { Name = cw; Cells = [] }
            else
                printfn "Row exists %A" cw
                sheet



let setCell sheet (vcr:ValidCellReference) value =
    let new_cell = { Column = vcr.RowCol.Col; Value = NoRValue; Contents = value; ExtraColumns = [] }
    let row = lookupRow sheet vcr.RowCol.Row.Name
    let new_cells = 
        match vcr.Cell with
        | None -> row.Cells @ [new_cell]
        | Some old_cell -> row.Cells |> List.map( fun c -> if c = old_cell then new_cell else c  ) // remove cell if there
    let new_row = { row with Cells = new_cells }
    let new_rows = replaceRow  sheet.Rows new_row
    { sheet with Rows = new_rows }

let printCells sheet cells normRange =
    let cols = calcThingsBetween [] normRange.TopLeft.Col normRange.BottomRight.Col sheet.Cols
    printf "%5s" ""    //space right
    for c in cols do
        printf "%10s" c.Name
    printf "\n"
    let rec cellprinter lastrow (cells : ValidCellReference list ) =
        match cells with
        | c::rest -> 
                        if c.RowCol.Row <> lastrow then printf "\n%-5s" c.RowCol.Row.Name
                        else printf " | "
                        match c.Cell with
                        | None -> printf "%10s" ""
                        | Some cc -> printf "%10s" cc.Contents.forDisplay
                        cellprinter c.RowCol.Row rest
        | [] -> printf "\n"
    cellprinter { Name = ""; Cells=[] } cells



type SingleCellFindResult = 
    | RowNotValid of string
    | RowNotGiven
    | ColNotValid of string
    | NoCell 
    | CellAtLocation of Cell

let findSingleCellFromRef sheet cellref = 
    let cl = parseCellLocation cellref
    match cl with
    | ColRow(cr,r,_) ->
        let row = findRow sheet r
        match row with
        | Some(fr) -> 
            let cell = findCellInRow fr cr
            match cell with
            | Some(c) -> CellAtLocation(c)
            | None -> 
                // is there a column or is it just empty
                let col = findColumn sheet cr
                match col with
                | Some(_) -> NoCell
                | None -> ColNotValid(cr)
        | None -> RowNotValid(r)
    | ColOnly(_) -> RowNotGiven

let rec doCell sheet cmd : Sheet =
    let range = parseCellRange cmd
    match range with
    | NothingGiven -> printfn "Nothing given"; sheet
    | NotThing(rest) -> printfn "Not a cell reference"; sheet
    | ParsedThingError(e,rest) ->
        printfn "Error: %A" e; sheet
    | ParsedThing(cellRange, rest) ->
        let gathered = gatherCells sheet cellRange
        match gathered with
        | GatheredError(e) -> Console.WriteLine e.display 
        | CellList(cells, nr) ->
           printCells sheet cells nr

        let anyleft = skipWhite(rest)
        match anyleft with
        | "" -> sheet
        | _ -> doCell sheet anyleft

let getCellValue (t:string) =
    match t with
    | _ when t.Length=0 -> Empty
    | x when t.[0]='=' -> Expr(x)
    | x when t.[0]='\"' -> Str(x.Substring(1,x.Length-2))
    //| x when fst (Int32.TryParse(t)) -> IntValue (Int32.Parse(x))
    | _ when isInt t -> IntValue((int)t)
    | _ when isDouble t -> DoubleValue((double)t)
    | _ -> Str(t)

type SetResult = 
    | SetResultError of string
    | SetResultOk of Sheet

let rec doSetCell sheet cellref (text:string) : SetResult =
    let range = parseCellRange cellref
    match range with
    | NothingGiven -> SetResultError("Nothing given")
    | NotThing(rest) -> SetResultError( "Not a cell reference")
    | ParsedThingError(e,rest) ->
        SetResultError(sprintf "Error: %A" e)
    | ParsedThing(cellRange, rest) ->
        let t = text.Trim()
        let gcells = gatherCells sheet cellRange
        match gcells with
        | GatheredError(err) -> SetResultError(err.display)
        | CellList(cells, nr) ->
            let mutable s = sheet
            let cv = getCellValue t
            for c in cells do
                s <- setCell s c cv
            SetResultOk(s)


let rec doSet sheet cmd : Sheet =
    let range = parseCellRange cmd
    match range with
    | NothingGiven -> printfn "Nothing given"; sheet
    | NotThing(rest) -> printfn "Not a cell reference"; sheet
    | ParsedThingError(e,rest) ->
        printfn "Error: %A" e; sheet
    | ParsedThing(cellRange, rest) ->
        let t = rest.Trim()
        let gcells = gatherCells sheet cellRange
        match gcells with
        | GatheredError(err) -> Console.WriteLine err.display; sheet
        | CellList(cells, nr) ->
            let mutable s = sheet
            let cv = getCellValue t
            for c in cells do
                s <- setCell s c cv
            s

let doShow sheet =
    printf "%5s" ""    //space right
    for c in sheet.Cols do
        printf "%16s" c.Name
    printfn ""
    for r in sheet.Rows do
        printf "%-5s" r.Name
        for c in sheet.Cols do
            printf "|"
            match findCellInRowByCol r c with
            | None -> printf "%15s" ""
            | Some cell ->
                match cell.Contents with
                | Expr(e) -> printf "%15s" (cell.Contents.forDisplay + "(" + cell.Value.forDisplay + ")")
                | _ -> printf "%15s" cell.Contents.forDisplay
        printfn ""

let doSave sheet filename =
    let s = seq {
        yield sprintf "Id=%d" sheet.Id
        yield sprintf "Name=%s" sheet.Name
        for col in sheet.Cols do
            yield sprintf "Col=%s" col.Name 
        for r in sheet.Rows do
            yield sprintf "Row=%s" r.Name
            for c in r.Cells do
                let rec extraColumns (cols: Col list) = 
                    match cols with
                    | head::tail -> head.Name + "," + (extraColumns tail)
                    | [] -> ""
                let extraCol = extraColumns c.ExtraColumns
                let extraColText = if extraCol.Length > 1 then sprintf "{%s}" (extraCol.Substring(0, extraCol.Length-1)) else ""
                                     
                yield sprintf "|%s|=%s'%s" c.Column.Name extraColText c.Contents.forDisplay
            yield "EndRow"
    }
    File.WriteAllLines( filename, s)
    sheet

let doLoad filename =
    let lines = File.ReadAllLines(filename)
    let mutable s = { Id = -1; Name = ""; Rows = []; Cols = [] }
    let mutable current_row = { Name = ""; Cells = [] }
    for l in lines do
        match l with
        | Regex "Id=(.+)" [ id ] -> s <- { s with Id = (int)id  }
        | Regex "Name=(.+)" [ name ] -> s <- { s with Name = name }
        | Regex "Col=(.+)" [ col ] -> s <- { s with Cols = s.Cols @ [ { Name=col }] }
        | Regex "Row=(.+)" [ row ] -> current_row <- { Name = row; Cells = [] }
        | "EndRow" -> s <- { s with Rows = s.Rows @ [current_row ] }
        | Regex "\|(.+)\|=(.*)'(.+)" [col; extra; cell ] -> 
            let c = findColumn s col
            match c with
            | Some(col) -> 
                let mutable ec = []
                if extra.Length > 0 then 
                    let extras = extra.Substring(1,extra.Length-2).Split(',') // get rid of {}
                    for e in extras do
                        match findColumn s e with
                        | Some(extracol) ->
                            ec <- ec @ [extracol]
                        | None -> failwith "Unknown extra column"
                else ()
                current_row <- { current_row with Cells = current_row.Cells @ [ { Contents = getCellValue cell; Value=NoRValue; Column=col; ExtraColumns = ec } ] } // new cell
            | None -> failwith "Unknown column"
        | _ -> ()
    s
    


