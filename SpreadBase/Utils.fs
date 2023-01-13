
module Utils

open System
open System.Text.RegularExpressions

let isInt (s:string) = 
    let mutable m:Int64 = 0L
    let (b,m) = Int64.TryParse s
    b

let isDouble (s:string) = 
    let mutable m:Double = 0.0
    let (b,m) = Double.TryParse s
    b

let rec skipWhite str:string =
    if String.length str = 0 then str
    else if str.[0] = ' ' then skipWhite str.[1..]
    else str

let (|Regex|_|) pattern input =
    let m = Regex.Match(input, pattern)
    if m.Success then Some(List.tail [ for g in m.Groups -> g.Value ])
    else None