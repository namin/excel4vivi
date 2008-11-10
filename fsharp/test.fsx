#light

#r "Microsoft.Office.Interop.Excel"

open Microsoft.Office.Interop.Excel
open System

// Open an Excel application.
let excel =
    let app = new ApplicationClass()
    app.Visible <- true
    app

// Workaround for old format or invalid type library error 
// http://support.microsoft.com/kb/320369
Threading.Thread.CurrentThread.CurrentCulture <- new System.Globalization.CultureInfo("en-US")

// Add a new workbook.
let workbook = excel.Workbooks.Add()

// Reference the first sheet.    
let sheet = workbook.Worksheets.Item(1) :?> _Worksheet

/// Sets the cell in given row and column to the given value.
let set row col value =
    let cell = sheet.Cells.Item(row,col) :?> Range
    cell.Value2 <- value

/// Fills the worksheet with a calculation table.
let table f n =
    for i in {1..n} do
        set i 1 i
        set i 2 (f i)

// Lists the square of numbers from 1 to 12.
table (fun x -> x*x) 12