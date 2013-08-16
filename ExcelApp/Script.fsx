// This file is a script that can be executed with the F# Interactive.  
// It can be used to explore and test the library project.
// Note that script files will not be part of the project build.
#r "office.dll"
#r "Microsoft.Office.Interop.Excel.dll"

#load "ExcelApp.fs"

open System
open System.IO

let myDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
let fileName = Path.Combine(myDocuments, "Visual Studio 2010", "Projects", "VocabularyGame", "dictionary.xlsm")
let a = Excel.App(fileName)
printfn "%s" (a.getString("A5"))
a.Close