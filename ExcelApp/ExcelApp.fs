namespace Excel

open Microsoft.Office.Interop.Excel

type App(filePath) =
    let app = ApplicationClass()
    let releaseObject(o) =
        try
            try System.Runtime.InteropServices.Marshal.ReleaseComObject(o) |> ignore
            with ex -> printfn "Unable to release the Object %s" (ex.ToString())
        finally
            System.GC.Collect()
    let workbook = app.Workbooks.Open filePath
    let worksheet = workbook.Worksheets.get_Item(1) :?> _Worksheet
    member this.xlApp = app
    member this.xlWorkBook = workbook
    member this.xlWorkSheet = worksheet
    member this.Close() =
        workbook.Close false
        app.Quit |> ignore
        releaseObject worksheet
        releaseObject workbook
        releaseObject app
    member this.get (cell:string) = this.xlWorkSheet.get_Range(cell).Value2
    member this.getByte (cell:string) = this.get(cell) :?> byte
    member this.getDouble (cell:string) = this.get(cell) :?> double
    member this.getInt (cell:string) = this.get(cell) :?> int
    member this.getString (cell:string) = this.get(cell) :?> string