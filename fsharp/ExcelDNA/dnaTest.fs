#light
open ExcelDna.Integration

[<ExcelFunction(Category="FSharp Functions", Description="funky(x,y) = xxxy")>]
let funky x y = x*x*x*y

[<ExcelFunction(Category="FSharp Functions", Description="n! = 1*2*...*n")>]
let rec factorial n = 
    if n<=1
    then 1
    else n * (factorial (n-1))