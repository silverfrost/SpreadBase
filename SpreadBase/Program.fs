// Learn more about F# at http://fsharp.org

open Saturn
open Giraffe
open Giraffe.Core
open Giraffe.ResponseWriters
open Giraffe.GiraffeViewEngine
open FSharp.Control.Tasks.V2


open System
open Utils
open Spreadsheet
open Parser
open Eval
open Microsoft.AspNetCore.Http

let doCommand sheet cmd =
    let c = snipWord cmd
    match c with
        | NothingGiven ->
            sheet
        | NotThing(rest) -> 
            printfn "Do not understand, expected word"
            sheet
        | ParsedThing(cw, rest) -> 
            let restw = skipWhite(rest)
            let more = String.length restw > 0
            match cw with
                | "addcol" ->
                    if more then doAddCol sheet restw
                    else sheet
                | "addrow" ->
                    if more then doAddRow sheet restw
                    else sheet
                | "show" ->
                    calculateRValues sheet
                    doShow sheet
                    sheet
                | "cell" ->
                    match more with
                    | false -> sheet
                    | true -> doCell sheet restw
                | "set" ->
                    match more with
                    | false -> sheet
                    | true -> doSet sheet restw
                | "save" ->
                    if more then doSave sheet restw
                    else printfn "no filename"; sheet
                | "load" ->
                    if more then doLoad restw
                    else printfn "no filename"; sheet
                | _ -> 
                    printfn "Unknown command %A" cw
                    sheet
        | ParsedThingError (_,_) -> sheet // not really possible

let testSheet id = 
    let cola = { Name = "A" }
    let colb = { Name = "B" }
    let c = { Contents = Str "Hello"; Column = cola; Value = NoRValue; ExtraColumns = [] }
    let c1 = { Contents = Str "World"; Column = colb; Value = NoRValue; ExtraColumns = [] }
    let r = { Name = "1"; Cells = [c; c1] }
    let mutable s = { Id = id; Name = "sheet1"; Rows = [r]; Cols = [cola; colb ]}
    for cn in ['C'..'H'] do
        s <- doCommand s (sprintf "addcol %c" cn)
    for rn in [2..8] do
        s <- doCommand s (sprintf "addrow %d" rn)
    let setup = [ "set b1 76"; "set b2 42"; "set b3 =b1+b2"; "set b4 =b3+1"; "set a2 =sum(b1:b4,3)"; "set c1 =b4" ; "set d4 56" ; "set e4 67" ]
    setup |> List.map( fun x -> s <- doCommand s x) |> ignore
    s <- cellMerge s (lookupRow s "4" ) (lookupColumn s "d") (lookupColumn s "e")
    calculateRValues s
    s

//[<EntryPoint>]
let main1 argv =
    //testParser()
    let cols = List.empty
    let cola = { Name = "A" }
    let colb = { Name = "B" }
    let c = { Contents = Str "Hello"; Column = cola; Value = NoRValue; ExtraColumns = [] }
    let c1 = { Contents = Str "World"; Column = colb; Value = NoRValue; ExtraColumns = [] }
    let r = { Name = "1"; Cells = [c; c1] }
    let mutable s = { Id = -1; Name = "sheet1"; Rows = [r]; Cols = [cola; colb ]}
    for cn in ['C'..'H'] do
        s <- doCommand s (sprintf "addcol %c" cn)
    for rn in [2..8] do
        s <- doCommand s (sprintf "addrow %d" rn)
    let setup = [ "set b1 76"; "set b2 42"; "set b3 =b1+b2"; "set b4 =b3+1"; "set a2 =sum(b1:b4,3)"; "set c1 =b4" ]
    setup |> List.map( fun x -> s <- doCommand s x) |> ignore
    calculateRValues s
    testEval s

    let mutable running = true
    while running do
        Console.Write ">"
        let cmd = Console.ReadLine()
        if cmd="q" then running <- false
        else s <- doCommand s cmd

    0 // return an integer exit code

let sheetFile id =
    sprintf "%d.sb" id

let loadSheet id =
    let fn = sheetFile id
    if System.IO.File.Exists(fn) then 
        let s = doLoad fn
        calculateRValues s
        s
    else testSheet id

let saveSheet sheet id  =
    doSave sheet (sheetFile id) |> ignore

let renderSheet sheet =
    table [_class "sheet" ] [
        tr [] [
            th [] [] // row
            for c in sheet.Cols do
                th [ _class "colheader"; _id (sprintf "c%s" c.Name) ] [str c.Name ] 
        ]
        for r in sheet.Rows do
            tr [ _class "sr" ] [ 
                td [ _class "rowname"; _id (sprintf "r%s" r.Name)  ] [str r.Name]
                let rec renderCells  acc (cols: Col list) (row:Row) =
                    match cols with
                    | c::colrest -> 
                        let id = sprintf "c%s%s" c.Name row.Name
                        let onclick = _onclick (sprintf "cell_click(\"%s\")" id)
                        let cell_html, remaining_cols =
                            match findCellInRowByCol row c with
                            | None -> td [ _class "cell"; _id id; onclick ] [],colrest
                            | Some cell -> 
                                let attrs, colsleft = if List.isEmpty cell.ExtraColumns then [],colrest
                                                      else [ _colspan (sprintf "%d" (cell.ExtraColumns.Length+1)) ], colrest.[cell.ExtraColumns.Length..]
                                let content, moreattrs, moreclass = 
                                    match cell.Value with
                                    | NoRValue -> [], [], ""
                                    | StringRValue(s) -> [ str s ], [], ""
                                    | DoubleRValue(d) -> [ str (sprintf "%g" d)], [ ], "numcell"
                                    | ErrorRValue(s) -> [ str ("!" + s)], [ ], "errcell"
                                    | _ -> [str cell.Value.forDisplay ], [], ""
                                let cls = "cell" + if moreclass.Length=0 then "" else " " + moreclass
                                td ([ _class cls ] @ attrs @ moreattrs @ [_id id] @ [onclick]) content, colsleft
                        renderCells (acc @ [cell_html]) remaining_cols row
                    | [] -> 
                        acc
                let cells = renderCells [] sheet.Cols r
                for c in cells do
                    c
            ]
    ]

let renderSheetAsDataGrid sheet =
    script [] [

        rawText "\nvar grid = canvasDatagrid(); document.body.appendChild(grid); grid.data = [\n"
        for r in sheet.Rows do
            rawText "{"
            for c in sheet.Cols do
                match findCellInRowByCol r c with
                | None -> rawText (sprintf "%s: ''," c.Name)
                | Some cell ->
                    match cell.Value with
                    | _ -> rawText (sprintf "%s: '%s'," c.Name cell.Value.forDisplay)
            rawText "},\n"
        rawText "];"
        for c in 1..sheet.Cols.Length do
            rawText (sprintf "grid.setColumnWidth(%d,%d);" (c-1) 150 )
    ]


let index sheet =
  div [] [
      h2 [] [ str "Sheet" ]
      label [] [str "f(x)" ]
      input [ _id "root"; _value (sprintf "/api/%d" sheet.Id); _hidden ]
      input [ _id "formula-bar" ]
      label [ _id "cell-name" ] [str ""]
      label [ _id "message" ] [str ""]
      //renderSheet sheet
      renderSheetAsDataGrid sheet
  ]

type GetCellResponse = { CellValid: Boolean; Txt: string; RefedCells: string list; ErrorMessage: string  }

let rec cellListAsStringList acc (cells : ValidCellReference list) = 
    match cells with
    | h::rest -> cellListAsStringList (acc @ [ sprintf "%s%s" h.RowCol.Col.Name h.RowCol.Row.Name ]) rest
    | [] -> acc

let cellText (id:int) =
    fun (next : HttpFunc) (ctx : HttpContext) ->
        task {
            let cell = ctx.GetQueryStringValue "cell"
            let r  = match cell with
                     | Ok t ->
                        let cr = t.[1..] // first is a c
                        let sheet = loadSheet id
                        let fr = findSingleCellFromRef sheet cr
                        match fr with
                        | CellAtLocation c ->
                            match c.Contents with
                            | Expr(s) ->
                                try
                                    let node = parseExpression s.[1..]
                                    let cells_refed = calculateCellsReferenced sheet node []
                                    let cells_refed_sl = cellListAsStringList [] cells_refed
                                    { CellValid = true; Txt = c.Contents.forDisplay; RefedCells= cells_refed_sl; ErrorMessage="" }
                                with
                                | Failure(m) -> 
                                    { CellValid = true; Txt = c.Contents.forDisplay; RefedCells= []; ErrorMessage=m }
                                | _ -> 
                                    { CellValid = true; Txt = c.Contents.forDisplay; RefedCells= []; ErrorMessage="huh" }
                            | _ ->
                                { CellValid = true; Txt = c.Contents.forDisplay; RefedCells= []; ErrorMessage="" }
                        | NoCell ->
                            { CellValid = true; Txt = ""; RefedCells= []; ErrorMessage="" }
                        | _ ->
                            { CellValid = false; Txt = "error"; RefedCells= []; ErrorMessage="error" }
                     | Error _ ->
                        { CellValid = false; Txt = "error"; RefedCells= []; ErrorMessage="error" }
            return! json r next ctx
        }

type SetCellResponse = { Success: Boolean; Message: string }

let setCellText (id:int)  =
    fun (next : HttpFunc) (ctx : HttpContext) ->
        task {
            let cell = ctx.GetQueryStringValue "cell"
            let text = ctx.GetQueryStringValue "text"
            let txt = match text with
                        | Ok t -> t
                        | Error e -> ""
            let r  = match cell with
                     | Ok t ->
                        let cr = t.[1..] // first is a c
                        let sheet = loadSheet id
                        if t = "" then { Success = false; Message = "No selection "}
                        else
                            match doSetCell sheet t txt with
                            | SetResultError e -> { Success = false; Message = e}
                            | SetResultOk s -> 
                                saveSheet s id
                                { Success = true; Message = "Done" }
                     | Error e ->  
                        { Success = false; Message = e}
            return! json r next ctx
        }


let fullPage dom = 
    html [] [
        head [] [
            link [ _rel "stylesheet"; _href "https://cdn.jsdelivr.net/npm/bootstrap@4.6.0/dist/css/bootstrap.min.css" ]
            link [ _rel "stylesheet"; _href "/app.css" ]
            script [_src "https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"] []
            script [_src "https://unpkg.com/canvas-datagrid"] []
        ]
        body [] [
            dom
            script [_src "/app.js"] []
        ]
    ]

let drawSheet (id:int) =
    htmlView (fullPage (index (loadSheet id)) )

let spreadRouter = router {
    getf "/%i" drawSheet
    getf "/api/%i/getCellText" cellText
    getf "/api/%i/setCellText" setCellText
}

let app = application {
    use_router spreadRouter
    use_static "static"
    url("http://192.168.1.105:5000/")
    url("https://localhost:5001/")
}


let ts = testSheet

run app

