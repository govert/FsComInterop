module MyFunctions

open ExcelDna.Integration
open Microsoft.Office.Interop.Excel
open System.Diagnostics

[<ExcelFunction(Description="My first .NET function")>]
let HelloDna name =
    "Hello " + name

[<ExcelFunction(Description = "Test Excel Interop")>]
let TestComVersion () =
    let application = ExcelDnaUtil.Application :?> Microsoft.Office.Interop.Excel.Application
    sprintf "%A" application.Version
    
let private SetStatusHelloImpl () = 
      let application = ExcelDnaUtil.Application :?> Application
      application.StatusBar <- "Hello from the command!"


[<ExcelCommand(MenuName="F# COM Test", MenuText="Set Status Hello", ShortCut="^F")>]
let SetStatusHello ()  =
    Debug.Print "Hello from SetStatus"
    SetStatusHelloImpl ()

[<ExcelCommand(MenuName="F# COM Test", MenuText="Clear Status", ShortCut="%F")>]
let ClearStatus () =
      let application = ExcelDnaUtil.Application :?> Application
      application.StatusBar <- "Hello from the command!"

    