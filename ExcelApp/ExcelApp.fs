namespace Excel

open Microsoft.Office.Interop.Excel
open System

type App(filePath, visible) =
    let app = ApplicationClass(Visible = visible)
    let releaseObject(o) =
        try
            try System.Runtime.InteropServices.Marshal.ReleaseComObject(o) |> ignore
            with ex -> printfn "Unable to release the Object %s" (ex.ToString())
        finally
            System.GC.Collect()
    let workbook = app.Workbooks.Open filePath
    let worksheet = workbook.Worksheets.get_Item(1) :?> _Worksheet
    new(filePath) = App(filePath, false)

    member this.xlApp = app
    member this.xlWorkBook = workbook
    member this.xlWorkSheet = worksheet
    member this.MainWindowHandle = IntPtr(app.Hwnd)

    member this.get (cell:string) = this.getRange(cell).Value2
    member this.getByte (cell:string) = this.get(cell) :?> byte
    member this.getDouble (cell:string) = this.get(cell) :?> double
    member this.getInt (cell:string) = this.get(cell) :?> int
    member this.getRange (cell:string) : Range = this.xlWorkSheet.get_Range(cell)
    member this.getString (cell:string) = this.get(cell) :?> string
    member this.selectRange (cell:string) = this.getRange(cell).Select() |> ignore

    member this.Close() =
        workbook.Close false
        app.Quit |> ignore
        releaseObject worksheet
        releaseObject workbook
        releaseObject app