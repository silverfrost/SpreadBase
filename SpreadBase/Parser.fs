module Parser

open Utils
open System

type Token = 
    | NoToken
    | Ident of string
    | FunctionCall of string
    | NumberConst of double
    | DateConst of DateTime
    | BoolConst of Boolean
    | LeftBracket
    | RightBracket
    | Plus
    | Minus
    | UnaryMinus
    | Times
    | Divide
    | Comma
    | Equal
    | Less
    | LessEqual
    | Greater
    | GreaterEqual
    | NotEqual
    | Ampersand
    | CellReference of string
    | CellRange of string
    | StringConst of string
    | NotRecognised of string
    | ErrorToken of string
    | Backstop

type Operator = { tok:Token; name: string; prec: int; infix: bool; absorb: bool }

type TreeNode = { tok: Token; children: TreeNode list }

type OpStackItem = { node: TreeNode; prec: Operator }


let operator_prec = [ 
    { tok = Plus; name = "+"; prec = 50; infix=true; absorb=false };
    { tok = Minus; name ="-"; prec = 50; infix=true; absorb=false };
    { tok = Times; name ="*"; prec = 60; infix=true; absorb=false };
    { tok = Divide; name ="/"; prec = 70; infix=true; absorb=false };
    { tok = Less; name ="<"; prec = 30; infix=true; absorb=false };
    { tok = LessEqual; name ="<="; prec = 30; infix=true; absorb=false };
    { tok = Greater; name =">"; prec = 30; infix=true; absorb=false };
    { tok = GreaterEqual; name =">="; prec = 30; infix=true; absorb=false };
    { tok = Equal; name ="="; prec = 30; infix=true; absorb=false };
    { tok = NotEqual; name ="<>"; prec = 30; infix=true; absorb=false };
    { tok = UnaryMinus; name ="-ve"; prec = 90; infix=false; absorb=false};
    { tok = Ampersand; name = "&"; prec = 50; infix=true; absorb=false };
    { tok = LeftBracket; name ="("; prec = 10; infix=false; absorb=true }
]

let findOperator t = operator_prec |> List.find( fun x -> x.tok = t )

let nextToken str =
    let s = skipWhite str
    if s.Length = 0 then NoToken,s
    else 
        let ch = s.[0]
        let mutable p = 1
        let l = s.Length
        let strleft () = s.[p..]
        let strleftp () = s.[p+1..]
        match ch with
        | x when Char.IsLetter ch -> 
                while p<l && (Char.IsLetterOrDigit s.[p] || s.[p] = ':') do
                    p <- p + 1
                if (p<l && s.[p]='(') then
                    FunctionCall(s.[0..p-1]), strleftp()
                else
                    let ident = s.[0..p-1]
                    if ident.IndexOf ':' >= 0 then CellRange(ident), strleft()
                    elif ident.Length=2 && Char.IsDigit s.[1] then CellReference(ident), strleft()
                    else Ident(ident), strleft()

        | x when Char.IsDigit ch ->
                while p<l && (Char.IsDigit s.[p] || s.[p] = '.'  || s.[p] = 'e') do
                    p <- p + 1
                let ident = s.[0..p-1]
                let flag,d = Double.TryParse ident
                if flag then NumberConst(d), strleft()
                else NotRecognised(ident), strleft()
        | '\"' -> 
                while p<l && s.[p]<>'\"' do
                    p <- p + 1
                let str = if p = 1 then ""
                          else s.[1..p-1]
                if (p<l) then StringConst(str), strleftp()
                else  StringConst(str), ""
        | '(' -> LeftBracket, strleft()
        | ')' -> RightBracket, strleft()
        | '+' -> Plus, strleft()
        | '-' -> Minus, strleft()
        | '*' -> Times, strleft()        
        | '/' -> Divide, strleft()
        | '&' -> Ampersand, strleft()
        | ',' -> Comma, strleft()
        | '=' -> Equal, strleft()
        | '<' -> if p<l && s.[p] = '=' then LessEqual, strleftp()
                 elif p<l && s.[p] = '>' then NotEqual, strleftp()
                 else Less, strleft()
        | '>' -> if p<l && s.[p] = '=' then GreaterEqual, strleftp()
                 else Greater, strleft()

        | _ -> NotRecognised(s.[p..p]), strleft()

let parseExpression e =

    let rec stack_op op_stack val_stack new_op_item =
        let tn::rest_op_stack = op_stack
        if new_op_item.prec.prec = tn.prec.prec && tn.prec.absorb then
            rest_op_stack, val_stack // absorb both
        elif new_op_item.prec.prec >= tn.prec.prec then
            [new_op_item] @ op_stack, val_stack // just push it 
        else
            // we now have an operator of lower precedence wanting to go on a higher
            if tn.prec.infix then
                match val_stack with
                | a::b::rest_val_stack ->
                        let new_value_node = { tn.node with children = [b] @ [a] } // weld the two top items on the value stack into one
                        stack_op rest_op_stack ([new_value_node] @ rest_val_stack) new_op_item
                | _ -> failwith "Expected value"
            else
                // just a prefix operator
                let a::rest_val_stack = val_stack
                let new_value_node = { tn.node with children = [a]  } 
                stack_op rest_op_stack ([new_value_node] @ rest_val_stack) new_op_item

    let back_stop = { node = { tok = Backstop; children = [] }; prec = { tok = Backstop; name = "bs"; prec = 0; infix = false; absorb = false }}

    let tidy_up_opstack op_stack val_stack =
        let (new_op_stack, new_val_stack) = stack_op op_stack val_stack back_stop
        // we should now have two nodes on the op_stack and one on the value stack
        if new_op_stack.Length<>2 || new_val_stack.Length<>1 then failwith "Syntax error"
        else new_val_stack.[0]

    let rec parseNextToken bras terminators op_stack val_stack opnext expr_text =

        let nt = nextToken expr_text
        match nt with
        | NoToken,_ -> op_stack, val_stack, "" // end of expression
        | NotRecognised(id),_ -> 
                failwithf "Unrecognised: %s" id

        | NumberConst(d), expr_text_left -> 
                if opnext then
                    failwithf "Syntax error: %f" d // number when operator expected
                else
                    let new_val_stack = [ { tok = NumberConst(d); children = [] } ] @ val_stack
                    parseNextToken bras terminators op_stack new_val_stack true expr_text_left

        | StringConst(s), expr_text_left  
        | Ident(s), expr_text_left 
        | CellReference(s), expr_text_left | CellRange(s), expr_text_left
            -> 
                if opnext then
                    failwithf "Syntax error: %s" s // string when operator expected
                else
                    let new_val_stack = [ { tok = fst nt; children = [] } ] @ val_stack
                    parseNextToken bras terminators op_stack new_val_stack true expr_text_left


        | LeftBracket, expr_text_left -> 
                if opnext then failwith "Did not expect (" // ( when operator was expected
                else
                    let new_op_stack = [{ node = { tok = LeftBracket; children = [] }; prec = findOperator LeftBracket }] @ op_stack // push (
                    parseNextToken ([LeftBracket] @ bras) terminators new_op_stack val_stack false expr_text_left
        | RightBracket, expr_text_left ->
            if not opnext then failwith "Did not expect )" // ) when value was expected (or postfix operator)
            else
                match bras with
                | [] -> 
                    match terminators |> List.tryFind( fun x -> x = RightBracket ) with
                    | Some x -> op_stack, val_stack, expr_text // right bracket for end of function, pass back with text at )
                    | None -> failwith ") without (" // no brackets
                | LeftBracket::morebras -> 
                    // )  so unstack to the (
                    let (new_op_stack, new_val_stack) = stack_op op_stack val_stack { node = { tok = LeftBracket; children = [] }; prec = findOperator LeftBracket }
                    parseNextToken morebras terminators new_op_stack new_val_stack true expr_text_left
                | _ -> failwith "Expected )" // not bracket on top of bracket stack
        | Comma, expr_text_left ->
            match terminators |> List.tryFind( fun x -> x = Comma ) with
            | Some x -> op_stack, val_stack, expr_text // comma, pass back with text at )
            | None -> failwith "Unexpected comma"
            

        | FunctionCall(fn), expr_text_left ->
            // we need to collect the arguments. expr_text_left has already absorbed the ( so we are at the first arg

            let rec parseFunctionCall nodes (expr_text:string) =
                match nextToken expr_text with
                | NoToken,_ -> failwith "Too many (" // no closing )
                | NotRecognised(id),_ -> failwithf "Unrecognised: %s" id
                | RightBracket,rest -> // no arguments in call
                    [], rest
                | _ -> // regular start of argument
                    let (arg_op_stack, arg_val_stack, expr_text_to_term) = parseNextToken [] ([Comma] @ [RightBracket]) [back_stop] [] false expr_text
                    let arg = tidy_up_opstack arg_op_stack arg_val_stack
                    let new_nodes = nodes @ [arg]
                    match nextToken expr_text_to_term with
                    | NoToken,_ -> parseFunctionCall new_nodes expr_text_to_term // let rec take the strain
                    | NotRecognised(id),_ -> parseFunctionCall new_nodes expr_text_to_term // let rec take the strain
                    | RightBracket, text_after_bracket -> new_nodes, text_after_bracket // we return the text left
                    | Comma, text_after_comma -> parseFunctionCall new_nodes text_after_comma
                    | x,_ -> failwithf "Expected , or ), got %A" x
            let (nodes, text_after_call ) = parseFunctionCall [] expr_text_left
            let fncall = { tok = FunctionCall(fn); children = nodes } // node to contain function call
            parseNextToken bras terminators op_stack ([fncall] @ val_stack) true text_after_call

        | t,expr_text_left -> 
                let tt = 
                    if t = Minus && not opnext then
                        UnaryMinus
                    else
                        t
                let op = findOperator tt
                if op.infix <> opnext then
                    failwithf "Syntax error: %s" op.name // prefix operator in postfix position
                else
                    let (new_op_stack, new_val_stack) = stack_op op_stack val_stack { node = { tok = tt; children = [] }; prec = op }
                    parseNextToken bras terminators new_op_stack new_val_stack false expr_text_left

    
    let (op_stack, val_stack, text_left) = parseNextToken [] [] [back_stop] [] false e
    tidy_up_opstack op_stack val_stack

let rec printTree node indent =
    printfn "%s %A" indent node.tok
    for n in node.children do
        printTree n (indent + "    ") |> ignore

let pTree node = 
    printTree node ""

let testParser () =
    //try
    let t = [ "123"; "7*6"; "SUM(A1:B6)"; "-6+GG(8)"; "((6*8+9/5+sum(log(sin(a*b)+8))))"; "3<4" ]
    //t |> List.map ( fun x -> parseExpression x ) |> List.map ( fun x -> pTree x )
    let parseandprint x = 
        printfn "%s" x 
        parseExpression >> pTree

    //t |> List.map ( fun x -> pTree (parseExpression x) )
    t |> List.map ( fun x -> (parseandprint x) x)
   // printTree (parseExpression "1*(2+5)*sum(67,34+2,add(6))" ) ""
    //with
    //| Failure msg -> printfn "Error: %s" msg
    let mutable more = true
    let mutable s = "12.5 and bill(A5:C6) and A5 is go =<=<>>"
    while more do
        more <- 
            match nextToken s with
                |    NoToken,_ -> false
                |    NotRecognised(id),_ -> printfn "tok = %A" id; false
                |    t,l -> printfn "tok = %A" t; s <- l; true
    true
    
