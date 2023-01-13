module Eval

open System
open Utils
open Spreadsheet
open Parser



let equalise (nodes:RValue list) =
    match (nodes:RValue list) with
    | StringRValue(s)::StringRValue(s1):: [] -> nodes.Head, nodes.[1]
    | DoubleRValue(s)::DoubleRValue(s1):: [] -> nodes.Head, nodes.[1]
    | DateRValue(s)::DateRValue(s1):: [] -> nodes.Head, nodes.[1]
    | BoolRValue(s)::BoolRValue(s1):: [] -> nodes.Head, nodes.[1]
    | StringRValue(s)::NoRValue::[] -> nodes.Head, StringRValue("")
    | NoRValue::StringRValue(s)::[] -> StringRValue(""), nodes.Head
    | NoRValue::DoubleRValue(s)::[] -> failwith "Empty"
    | DoubleRValue(s)::NoRValue::[] -> failwith "Empty"
    | DoubleRValue(s)::StringRValue(s1):: [] ->
        let flag,d = Double.TryParse s1
        if flag then nodes.Head, DoubleRValue(d)
        else failwith "Cannot mix data types"
    | StringRValue(s)::DoubleRValue(s1):: [] ->
        let flag,d = Double.TryParse s
        if flag then DoubleRValue(d), nodes.[1]
        else failwith "Cannot mix data types"
    | _ -> failwithf "Cannot equalise %A" nodes


let doPlus (nodes:RValue list) =
    let (left, right) = equalise nodes
    match left,right with
    | DoubleRValue(l), DoubleRValue(r) -> DoubleRValue(l+r)
    | StringRValue(l), StringRValue(r) -> StringRValue(l+r)
    | _ -> failwithf "Unexpected %A " (left,right)

let doAmpersand (nodes:RValue list) =
    StringRValue(nodes.[0].forDisplay + nodes.[1].forDisplay)

let doMinus (nodes:RValue list) =
    let (left, right) = equalise nodes
    match left,right with
    | DoubleRValue(l), DoubleRValue(r) -> DoubleRValue(l-r)
    | StringRValue(l), StringRValue(r) -> failwith "Cannot subtract strings"
    | _ -> failwithf "Unexpected %A " (left,right)

let doTimes (nodes:RValue list) =
    let (left, right) = equalise nodes
    match left,right with
    | DoubleRValue(l), DoubleRValue(r) -> DoubleRValue(l*r)
    | StringRValue(l), StringRValue(r) -> failwith "Cannot multiply strings"
    | _ -> failwithf "Unexpected %A " (left,right)

let doDivide (nodes:RValue list) =
    let (left, right) = equalise nodes
    match left,right with
    | DoubleRValue(l), DoubleRValue(r) -> if r = 0.0 then failwith "Division by zero"
                                          else DoubleRValue(l/r)
    | StringRValue(l), StringRValue(r) -> failwith "Cannot divide strings"
    | _ -> failwithf "Unexpected %A " (left,right)

let doComparison (nodes:RValue list) fn fn1 fn2 fn3 =
    let (left, right) = equalise nodes
    match left,right with
    | DoubleRValue(l), DoubleRValue(r) -> BoolRValue(fn l r)
    | StringRValue(l), StringRValue(r) -> BoolRValue(fn1 l r)
    | DateRValue(l), DateRValue(r) -> BoolRValue(fn2 l r)
    | BoolRValue(l), BoolRValue(r) -> BoolRValue(fn3 l r)
    | _ -> failwithf "Unexpected %A " (left,right)



type CellRangeType =
    | SingleCellRValue of RValue
    | RectangleCells of GatheredCells

let evaluateSingleCell sheet cellref =
    let c = findSingleCellFromRef sheet cellref
    match c with
    | NoCell -> NoRValue
    | CellAtLocation(cell) -> cell.Value
    | ColNotValid(c) -> failwithf "Column not valid: %s" c
    | RowNotValid(r) -> failwithf "Row not valid: %s" r
    | RowNotGiven -> failwith "Row not given"

let evaluateCellRectangle sheet cellref =
    let r = parseCellRange cellref
    match r with
    | NothingGiven -> failwith "Nothing given"
    | NotThing(rest) -> failwith "Not a cell reference"
    | ParsedThingError(e,rest) -> failwithf "Error %s" e
    | ParsedThing(cellRange, rest) ->
        let gathered = gatherCells sheet cellRange
        match gathered with
        | GatheredError(e) -> failwithf "Error %s" e.display
        | CellList(cells, nr) -> cells, nr
           
let collectCellRange sheet range =
    let vcr, nr = evaluateCellRectangle sheet range
    vcr |> List.map ( fun x -> match x.Cell with | Some(cell) -> cell.Value | _ -> NoRValue )

let getDouble (values:RValue list) =
    match values.Head with
    | DoubleRValue(d) -> d
    | _ -> failwith "FAIL"

let doSin value =
    Math.Sin value

let doASin value =
    if value > 1.0 || value < -1.0 then failwith "asin domain -1..1"
    else Math.Asin value

let doSinh value =
    Math.Sinh value

let doASinh value =
    Math.Asinh value

let doCos value =
    Math.Cos value

let doACos value =
    if value > 1.0 || value < -1.0 then failwith "acos domain -1..1"
    else Math.Acos value

let doCosh value =
    Math.Cosh value

let doACosh value =
    Math.Acosh value

let doTan value =
    Math.Tan value

let doATan value =
    Math.Atan value

let doTanh value =
    Math.Tanh value

let doATanh value =
    Math.Atanh value

let doExp value =
    Math.Exp value

let doLn value =
    if value <= 0.0 then failwith "Ln domain +ve"
    else Math.Log value

let doLog10 value =
    if value <= 0.0 then failwith "Log10 domain +ve"
    else Math.Log10 value

let doPower sheet (values: RValue list) evalNode =
    match values.[0], values.[1] with
    | DoubleRValue(num), DoubleRValue(pow) -> DoubleRValue( Math.Pow (num, pow ))
    | _ -> failwith "FAIL!"

let doInt (value:float) =
    Math.Floor value 

let doCeiling (value:float) =
    Math.Ceiling value 

let doTrunc (value:float) =
    Math.Truncate value 

let doRound sheet (values: RValue list) evalNode =
    match values.[0], values.[1] with
    | DoubleRValue(num), DoubleRValue(digits) -> DoubleRValue( Math.Round (num, (int)digits ))
    | _ -> failwith "FAIL!"

let doAverage (values:float list) =
    if values.IsEmpty then failwith "Division by zero"
    else (List.sum values) / (float)values.Length

let doCount (values:float list) =
    (float)values.Length

let doCounta (values:RValue list) =
    DoubleRValue( (float)values.Length )

let doPi() = DoubleRValue( Math.PI )

let doE() = DoubleRValue( Math.E )

let doSum values =
    List.sum values

let doEomonth (values : RValue list) =
    assert (values.Length = 2)
    match values.[0], values.[1] with
    | DateRValue( d ), DoubleRValue(n) 
        -> 
            let som = DateTime( d.Year, d.Month, 1) // start of month
            let mon = som.AddMonths( (int)( Math.Truncate (n+1.0) ))
            let eolm = mon.AddDays(-1.0) // end of last month
            DateRValue( eolm)
    | _,_ -> failwith "err"

let doSomonth (values : RValue list) =
    assert (values.Length = 2)
    match values.[0], values.[1] with
    | DateRValue( d ), DoubleRValue(n) 
        -> 
            let som = DateTime( d.Year, d.Month, 1) // start of month
            let mon = som.AddMonths( (int)( Math.Truncate (n) ))
            DateRValue( mon)
    | _,_ -> failwith "err"

let doAstime (values : RValue list) =
    assert (values.Length = 1)
    match values.[0] with
    | DoubleRValue( d ) 
        -> 
            let ticks = d * 3600.0 * 10000000.0
            let ts = TimeSpan( (int64)ticks)
            StringRValue( ts.ToString("hh\:mm\:ss") )
    | _ -> failwith "err"
 
type OperatorWithCriteria = { Tok: Token; Expr: TreeNode }

let rvalueToTreeNode r =
    match r with
    | DoubleRValue d -> { tok = NumberConst(d); children = [] }
    | StringRValue s -> { tok = StringConst(s); children = [] }
    | DateRValue d -> { tok = DateConst(d); children = [] }
    | BoolRValue d -> { tok = BoolConst(d); children = [] }
    | CellRangeRValue crf -> { tok = CellRange(crf); children = [] }
    | ErrorRValue e -> { tok = StringConst("!" + e); children = [] }
    | NoRValue -> { tok = StringConst(""); children = [] }
    
let rec isCritTrue sheet (rangeCriteria : (RValue*OperatorWithCriteria) list ) evalNode =
    match rangeCriteria with
    | (v1,owc)::tail ->
        let node = { tok = owc.Tok; children = [rvalueToTreeNode (v1)] @ [owc.Expr] }
        let res = evalNode sheet node false
        match res with
        | BoolRValue(b) -> 
            if b then isCritTrue sheet tail evalNode
            else false
        | _ -> false
    | [] -> true

// cellToSum = list of values we are processing, here so we know the size
// rangeCriteriaValues = list containing alternate Range and criteria. In sumif(V,r1,c1,r2,c2). r1 is a range and c1 a criteria expression
// acc is a list of tuples, one for each range/criteria paid. The first in the tuple is the list of rvalues that make up the range. The second is a treenode for the criteria
let rec gatherRangeCriteria sheet (cellsToSum : RValue list) rangeCriteriaValues acc = 
    match rangeCriteriaValues with
    | CellRangeRValue(cr)::crit::tail ->
        let cellsForCriteria = collectCellRange sheet cr
        if cellsForCriteria.Length <> cellsToSum.Length then failwith "Ranges must be same size"
        else 
            match crit with
            | CellRangeRValue _ -> failwith "Criteria cannot be a range" 
            | StringRValue(scrit) ->
                // strings can have an operator at the start
                let tok,len = if scrit.StartsWith('=') then Plus,1
                              elif scrit.StartsWith("<>") then NotEqual,2
                              elif scrit.StartsWith("<=") then LessEqual,2
                              elif scrit.StartsWith(">=") then GreaterEqual,2
                              elif scrit.StartsWith(">") then Greater,1
                              elif scrit.StartsWith("<") then Less,1
                              else Equal,0
                let n = parseExpression scrit.[len..]
                gatherRangeCriteria sheet cellsToSum tail (acc @ [cellsForCriteria,{ Tok = tok; Expr = n }] )
                
            | _ -> 
                gatherRangeCriteria sheet cellsToSum tail (acc @ [cellsForCriteria,{ Tok = Equal; Expr = rvalueToTreeNode crit }] )
    | _ :: crit :: tail -> failwith "Expected range for criteria"
    | [] -> acc
    | _ -> failwith "Expected range and criteria"

// applies the function fn in turn to each value in a range if the criteria are met
// fn = function that takes two doubles and returns one. (+) is such a function and ends up summing the rnage
// cells = cells to be passed to fn
// rangeCrit = to be included in the calculation a set of criteria must be met. rangeCrit is that list of criteria.
//             each item in the list is a tuple containing a set of cell values and a criteria to match with them.
//             Say you have sumifs(b1:b4, c1:c4, ">50", d1:d4, "<200")
//             rangeCrit is a list with the first item being a list of RValues c1:c4 and the tree >50
//                                          second items is a  list of RValues d1:d4 and the tree <200
//             This function takes the first item from each of the list and sees if the criteria is matched (i.e. is c1 >50 && d1<200). If it is it applies the function
// accumulator = result of applying fn successively
// evalNode = function capable of evaluating trees
let rec applyRangedCriteria fn sheet cells rangeCrit accumulator evalNode =
    match cells with
    | cv::rest ->
        // from the range/criteria we need to take the next in the ranges
        let rec nextRangeCrit ranges individualCellsCrit remainingCellsCrit = 
            match ranges with
            | ((head::tail),crit)::rest -> nextRangeCrit rest (individualCellsCrit @ [head,crit]) (remainingCellsCrit @ [tail,crit])
            | [] -> individualCellsCrit, remainingCellsCrit
        let currentRangeCriteria, remainingRangeCriteria = nextRangeCrit rangeCrit [] []
        let amount,apply = 
            match cv with
            | DoubleRValue d -> d,true
            | StringRValue s ->
                let flag,d = Double.TryParse s
                if flag then d,true
                else 0.0,false
            | _ -> 0.0,false
        let accum = 
            if apply && (isCritTrue sheet currentRangeCriteria evalNode) then fn accumulator amount
            else accumulator // do not apply this one
        applyRangedCriteria fn sheet rest remainingRangeCriteria accum evalNode
    | [] -> 
        accumulator

let doIfs sheet (values : RValue list) evalNode fn accumulator =
    // arg1 must be the range of cells we are processing
    // arg2 is a range, same dimensions as arg1, which will be used as a criteria for arg3
    // arg3 is a criteria. cells in arg2 are compared against the criteria and if true the corresponding cells in arg1 are summed
    // range/criteria pairs can fill arg4/5, arg6/7 etc. Again the range has to match the same dimensions as arg1
    // example: sumifs(A1:A5, B1:B5, "<today()", B1:B5, ">C1")
    match values with
    | CellRangeRValue(vr)::restValues ->
        let cellsToSum = collectCellRange sheet vr
        let rangeCriteria = gatherRangeCriteria sheet cellsToSum restValues []
        // we now have a list in which each element is a list*criteria
        let total = applyRangedCriteria fn sheet cellsToSum rangeCriteria accumulator evalNode
        DoubleRValue(total)

    | _ -> failwith "Expected range as 1st argument"


let doSumifs sheet (values : RValue list) evalNode = 
    doIfs sheet values evalNode (+) 0.0


let doMaxifs sheet (values : RValue list) evalNode = 
    let maxfn a b = 
        if a > b then a
        else b
    doIfs sheet values evalNode maxfn 0.0

let doMinifs sheet (values : RValue list) evalNode = 
    let minfn a b = 
        if a < b then a
        else b
    doIfs sheet values evalNode minfn Double.MaxValue

type DoubleFunction = { FnPtr: Double -> Double }

type DoubleListFunction = { FnPtr: Double list -> Double; AllowCellRange:bool  }

type ListFunction = { FnPtr: RValue list -> RValue }

type NoArgFunction = { FnPtr: unit -> RValue }

type SpecificArgFunction = { FnPtr: RValue list -> RValue; Args: string }

type GeneralFunction = { FnPtr: Sheet -> RValue list -> ( Sheet -> TreeNode -> bool -> RValue ) -> RValue; MinArgs: int; MaxArgs: int; AllNumeric:bool; AllowCellRange:bool }


type NamedFunction =
    | NamedDoubleFunction of DoubleFunction
    | NamedDoubleListFunction of DoubleListFunction
    | NamedListFunction of ListFunction
    | NamedGeneralFunction of GeneralFunction
    | NamedNoArgFunction of NoArgFunction
    | NamedSpecificArgFunction of SpecificArgFunction

let namedFunctions  = [ 
    ( "sin", NamedDoubleFunction( { FnPtr = doSin } ) )
    ( "asin", NamedDoubleFunction( { FnPtr = doASin } ) )
    ( "sinh", NamedDoubleFunction( { FnPtr = doSinh } ) )
    ( "asinh", NamedDoubleFunction( { FnPtr = doASinh } ) )
    ( "cos", NamedDoubleFunction( { FnPtr = doCos } ) )
    ( "acos", NamedDoubleFunction( { FnPtr = doACos } ) )
    ( "cosh", NamedDoubleFunction( { FnPtr = doCosh } ) )
    ( "acosh", NamedDoubleFunction( { FnPtr = doACosh } ) )
    ( "tan", NamedDoubleFunction( { FnPtr = doTan } ) )
    ( "atan", NamedDoubleFunction( { FnPtr = doATan } ) )
    ( "tanh", NamedDoubleFunction( { FnPtr = doTanh } ) )
    ( "atanh", NamedDoubleFunction( { FnPtr = doATanh } ) )
    ( "exp", NamedDoubleFunction( { FnPtr = doExp } ) )
    ( "ln", NamedDoubleFunction( { FnPtr = doLn } ) )
    ( "log10", NamedDoubleFunction( { FnPtr = doLog10 } ) )
    ( "power", NamedGeneralFunction( { FnPtr = doPower; MinArgs = 2; MaxArgs = 2; AllNumeric = true; AllowCellRange = false } ) )
    ( "pi", NamedNoArgFunction( { FnPtr = doPi } ) )
    ( "e", NamedNoArgFunction( { FnPtr = doE } ) )
    ( "sum", NamedDoubleListFunction( { FnPtr = doSum; AllowCellRange = true }))
    ( "int", NamedDoubleFunction( { FnPtr = doInt } ) )
    ( "ceiling", NamedDoubleFunction( { FnPtr = doCeiling } ) )
    ( "trunc", NamedDoubleFunction( { FnPtr = doTrunc } ) )  
    ( "round", NamedGeneralFunction( { FnPtr = doRound; MinArgs = 2; MaxArgs = 2; AllNumeric = true; AllowCellRange = false } ) )
    ( "average", NamedDoubleListFunction( { FnPtr = doAverage; AllowCellRange = true } ))
    ( "count", NamedDoubleListFunction( { FnPtr = doCount; AllowCellRange = true } ))
    ( "counta", NamedListFunction( { FnPtr = doCounta } ))
    ( "today", NamedNoArgFunction( { FnPtr = fun () -> DateRValue( DateTime.Today ) } ))
    ( "now", NamedNoArgFunction( { FnPtr = fun () -> DateRValue( DateTime.Now ) } ))
    ( "eomonth", NamedSpecificArgFunction( { FnPtr = doEomonth; Args = "TD" } ))
    ( "somonth", NamedSpecificArgFunction( { FnPtr = doSomonth; Args = "TD" } ))
    ( "sumifs", NamedGeneralFunction( { FnPtr = doSumifs; MinArgs = 3; MaxArgs = 31; AllNumeric = false; AllowCellRange = true } ))
    ( "maxifs", NamedGeneralFunction( { FnPtr = doMaxifs; MinArgs = 3; MaxArgs = 31; AllNumeric = false; AllowCellRange = true } ))
    ( "minifs", NamedGeneralFunction( { FnPtr = doMinifs; MinArgs = 3; MaxArgs = 31; AllNumeric = false; AllowCellRange = true } ))
    ( "astime", NamedSpecificArgFunction( { FnPtr = doAstime; Args = "D" } ))
    //    { Name="sin"; FnPtr=doSin; MinArgs=1; MaxArgs=1; AllNumeric = true }
]

let findNamedFunction x = namedFunctions |> List.tryFind (fun f -> fst f = x )

let doesNamedFunctionAllowCellRanges x =
    let fnptr = findNamedFunction x
    match fnptr with
    | None -> false
    | Some(f) -> 
        match snd f with
        | NamedGeneralFunction nf -> nf.AllowCellRange
        | NamedDoubleListFunction nf -> nf.AllowCellRange
        | NamedListFunction(_) -> true
        | _ -> false

let areAllNumeric values =
    match values |> List.tryFind (fun v -> match v with | DoubleRValue(_) -> false | _ -> true ) with
    | Some(_) -> false
    | None -> true

let areAllNumericOrRange values =
    match values |> List.tryFind (fun v -> match v with | DoubleRValue(_) | CellRangeRValue(_) -> false | _ -> true ) with
    | Some(_) -> false
    | None -> true



let rec filterDoubleValues sheet acc values = 
    match values with
    | h::rest ->
        match h with
        | DoubleRValue(d) ->
            filterDoubleValues sheet (acc @ [d]) rest
        | StringRValue(s) -> // can we convert to a number
            let flag,d = Double.TryParse s
            if flag then filterDoubleValues sheet (acc @ [d]) rest
            else filterDoubleValues sheet acc rest
        | CellRangeRValue (r) ->
            // calculate the range
            let rangeValues = collectCellRange sheet r
            let subacc = filterDoubleValues sheet [] rangeValues
            filterDoubleValues sheet (acc @ subacc) rest

        | _ -> // ignore
            filterDoubleValues sheet acc rest
    | [] ->
        acc

let rec filterNotEmptyValues sheet acc values = 
    match values with
    | h::rest ->
        match h with
        | NoRValue -> filterNotEmptyValues sheet acc rest
        | CellRangeRValue (r) ->
            // calculate the range
            let rangeValues = collectCellRange sheet r
            let subacc = filterNotEmptyValues sheet [] rangeValues
            filterNotEmptyValues sheet (acc @ subacc) rest

        | _ -> filterNotEmptyValues sheet (acc @ [h]) rest
    | [] ->
        acc

let evaluateFunction sheet (fn:String) (values:RValue list) evalNode =
    let fnptr = findNamedFunction (fn.ToLower())
    match fnptr with
    | None -> failwithf "Unknown function %s" fn
    | Some(f) -> 
        match snd f with
        | NamedGeneralFunction nf -> 
            if nf.MinArgs > values.Length then failwithf "%s requires at least %d arguments" fn nf.MinArgs
            elif nf.MaxArgs < values.Length then failwithf "%s requires can only have a maximum of %d arguments" fn nf.MinArgs
            else
                if nf.AllNumeric && (( nf.AllowCellRange && not (areAllNumericOrRange values)) || (not nf.AllowCellRange && not (areAllNumeric values))) then failwithf "Expected numeric arguments in %s" fn 
                else nf.FnPtr sheet values evalNode
        | NamedDoubleFunction df ->
            if values.Length <> 1 then failwithf "%s requires one argument" fn
            else 
                match values.[0] with
                | DoubleRValue(d) -> DoubleRValue( df.FnPtr d )
                | _ -> failwithf "%s requires a number" fn
        | NamedDoubleListFunction df ->
            let doubleValues = filterDoubleValues sheet [] values
            DoubleRValue(df.FnPtr doubleValues)
        | NamedListFunction nl ->
            let allvalues = filterNotEmptyValues sheet [] values
            nl.FnPtr allvalues
        | NamedSpecificArgFunction sf ->
            if values.Length <> sf.Args.Length then failwithf "%s expects %d arguments" fn sf.Args.Length
            else
                // arg count is right so convert each one to the correct type or fail
                let rec convertArg (inargs: RValue list) (argtype:string) outargs : RValue list =
                    match inargs with
                    | h::rest ->
                        let recurse = convertArg rest argtype.[1..]
                        let argchar = argtype.[0] 
                        match h with
                        | DoubleRValue(d) ->
                            match argchar with
                            | 'D' -> recurse (outargs @ [h]) // want double, got double
                            | 'S' -> recurse (outargs @ [ StringRValue( h.forDisplay )]) // want string, got double
                            | 'T' -> failwith "Expected date" // want date, got double
                            | 'R' -> failwith "Expected cell range" // want range, got double
                            | _ -> failwith "Error"
                        | StringRValue(s) ->
                            match argchar with
                            | 'D' -> 
                                let valid, d = Double.TryParse s
                                if valid then recurse (outargs @ [ DoubleRValue(d)] ) // want double, got string
                                else failwith "Expected a number"
                            | 'S' -> recurse (outargs @ [ h ]) // want string, got string
                            | 'T' -> 
                                let valid, d = DateTime.TryParse s
                                if valid then recurse (outargs @ [ DateRValue(d)] ) // want date, got string
                                else failwith  "Expected a date"
                            | 'R' -> failwith "Expected cell range" // want range, got string
                            | _ -> failwith "Error"
                        | DateRValue(d) ->
                            match argchar with
                            | 'D' -> failwith "Expected number" // want double, got date
                            | 'S' -> recurse (outargs @ [ StringRValue( h.forDisplay )]) // want string, got date
                            | 'T' -> recurse (outargs @ [h])  // want date, got date
                            | 'R' -> failwith "Expected cell range" // want range, got double
                            | _ -> failwith "Error"
                        | CellRangeRValue(r) ->
                            match argchar with
                            | 'D' -> failwith "Expected range" // want double, got date
                            | 'S' -> failwith "Expected range" // want string, got date
                            | 'T' -> failwith "Expected range"  // want date, got date
                            | 'R' -> recurse (outargs @ [h])  
                            | _ -> failwith "Error"
                        | _ -> failwith "Error"

                    | [] -> outargs
                let args = convertArg values sf.Args []
                sf.FnPtr args


        | NamedNoArgFunction nf ->
            if values.Length <> 0 then failwithf "%s does not have any one arguments" fn
            else nf.FnPtr()


let rec evaluateNode (sheet: Sheet) (root:TreeNode) allowCellRange =
    let child_allow_cell_range = 
        match root.tok with
        | FunctionCall(fn) -> doesNamedFunctionAllowCellRanges fn
        | _ -> false
    let child_rvalues = root.children |> List.map( fun x -> evaluateNode sheet x child_allow_cell_range )
    match root.tok with
    | NumberConst(d) -> DoubleRValue(d)
    | StringConst(s) -> StringRValue(s)
    | CellReference(cr) -> evaluateSingleCell sheet cr
    | CellRange(cr) -> if allowCellRange then CellRangeRValue(cr) else failwith "Unexpected cell range"
    | FunctionCall(fn) -> evaluateFunction sheet fn child_rvalues evaluateNode
    | Plus -> doPlus child_rvalues
    | Minus -> doMinus child_rvalues
    | Times -> doTimes child_rvalues
    | Divide -> doDivide child_rvalues
    | Equal -> doComparison child_rvalues (=) (=) (=) (=)
    | NotEqual -> doComparison child_rvalues (<>) (<>) (<>) (<>)
    | Greater -> doComparison child_rvalues (>) (>) (>) (>)
    | GreaterEqual -> doComparison child_rvalues (>=) (>=) (>=) (>=)
    | Less -> doComparison child_rvalues (<) (<) (<) (<)
    | LessEqual -> doComparison child_rvalues (<=) (<=) (<=) (<=)
    | Ampersand -> doAmpersand child_rvalues
    | Ident(s) -> failwithf "Unknown: %s" s
   // | DateValue(d) -> DateValue(d)
    | _ -> failwithf "Unexpected node %A" root


let evaluate (root:TreeNode) (sheet: Sheet) =
    try 
        evaluateNode sheet root false
    with
        Failure ex -> ErrorRValue ex

let eval s sheet =
    let n = parseExpression s
   // pTree n
    let e = evaluate n sheet
    printfn "%s = %s" s e.forDisplay

let testEval sheet =
    eval "123" sheet
    eval "1+2+3" sheet
    eval "1+2/3" sheet
    eval "1+2/0" sheet
    eval "A1 + B1" sheet
    eval "sin(1.57)*sin(1.57) + cos(1.57)*cos(1.57)" sheet
    eval "pi()" sheet
    eval "tan(pi()/2)" sheet
    eval "asin(1)" sheet
    eval "asin(2)" sheet
    eval "power(8,3)" sheet

type CellCalcState = { Cell:Cell; Node: TreeNode; CellRefs: ValidCellReference list }

let rec calculateCellsReferenced sheet node cells =
    let rec processChildren children sheet cells =
        match children with
        | h::tail -> 
            let new_cells = calculateCellsReferenced sheet h cells
            processChildren tail sheet new_cells
        | [] -> cells
    let cells_from_children = processChildren node.children sheet cells
    match node.tok with
    | CellReference cr | CellRange cr
        ->
            let pcr = parseCellRange cr
            match pcr with
            | NothingGiven(_) -> failwith "Not a cell reference"
            | NotThing(_) -> failwith "Not a cell reference"
            | ParsedThingError(e,_) -> failwith e
            | ParsedThing(pt,_) -> 
                let gathered = gatherCells sheet pt
                match gathered with
                | GatheredError(e) -> failwith e.display
                | CellList(reflist,norm_range) -> cells_from_children @ reflist
    | _ -> cells_from_children

let calculateRValues sheet =
    let mutable nonconst_cells = []
    for r in sheet.Rows do
        for c in r.Cells do
            match c.Contents with
            | Empty -> c.Value <- NoRValue
            | Str(s) ->
                let isdate,d  = DateTime.TryParse(s)
                c.Value <- if isdate then DateRValue(d)
                           else StringRValue(s)
            | IntValue(i) -> c.Value <- DoubleRValue(double i)
            | DoubleValue(d) -> c.Value <- DoubleRValue(d)
            | Expr(s) -> 
                        try 
                            let node = parseExpression s.[1..]
                            let cells_refed = calculateCellsReferenced sheet node []
                            nonconst_cells <- nonconst_cells @ [ { Cell=c; Node=node; CellRefs = cells_refed }]
                        with
                        | Failure m -> 
                            c.Value <- ErrorRValue(m)  // some sort of error
            | DateValue(d) -> c.Value <- NoRValue

    // we now have a list of expr cells. We need to order them so that they can 
    // be evaluated in an order that allows for the fact that they reference each other
   // let process_cell sheet cell =
    let is_cell_in_list (list: ValidCellReference list) (cell:Cell) =
        let r =  list |> List.tryFind( fun x -> match x.Cell with
                                                | Some xc -> xc = cell 
                                                | None -> false // empty cell
                                     )
        match r with
        | Some _ -> true
        | None -> false

    while nonconst_cells.Length > 0 do
        let mutable new_nonconst_list = []
        for c in nonconst_cells do
            let is_present = nonconst_cells |> List.tryFind( fun x -> is_cell_in_list c.CellRefs x.Cell )
            match is_present with
            | Some _ -> new_nonconst_list <- new_nonconst_list @ [c]
            | None -> // no other references so evaluate it
                c.Cell.Value <- evaluate c.Node sheet
        nonconst_cells <- if new_nonconst_list.Length = nonconst_cells.Length then
                              // no cells done, so they must be circular
                              nonconst_cells |> List.map( fun x -> x.Cell.Value <- ErrorRValue("Circular")) |> ignore
                              []
                          else
                              new_nonconst_list
