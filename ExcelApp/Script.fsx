// This file is a script that can be executed with the F# Interactive.  
// It can be used to explore and test the library project.
// Note that script files will not be part of the project build.
#r "office.dll"
#r "Microsoft.Office.Interop.Excel.dll"

#load "ExcelApp.fs"

open System
open System.IO
let a = Excel.App(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),"Excel","recnik.xlsm"))
printfn "%s" (a.getString("A5"))
a.Close