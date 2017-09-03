'REFERENCES:
'System.Core
'Microsoft.Office.Interop.Excel
'Microsoft.Office.Interop.PowerPoint
'Microsoft.Office.Interop.Word
'office
'UIAutomationClient
'UIAutomationTypes
'
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports System.Windows.Automation
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Diagnostics

Module ToPDFMain
    Sub Main(args() As String)
        On Error GoTo HandleError
        If args.Length < 3 Then
            Console.WriteLine("Usage: ToPDF.exe <appType (PP,XL,WD)> <source> <destination> <options>")
            Console.WriteLine("See documentation for further details.")
            Exit Sub
        End If

        args = args.ToArray()
        Dim app As String = args(0)
        Dim source As String = args(1)
        Dim destination As String = args(2)
        Dim arr As Array = (args.ToArray())

        Select Case app
            Case "PP"
                Dim ppex As New PowerPointExporter
                ppex.Export(source, destination)
            Case "XL"
                Dim xlex As New ExcelExporter
                xlex.Export(source, destination)

            Case "WD"
                Dim wdex As New WordExporter
                wdex.Export(source, destination)
            Case Else
                Console.WriteLine("Usage: ToPDF.exe <appType (PP,XL,WD)> <source> <destination> <options>")
                Console.WriteLine("See documentation for further details.")
        End Select
        Exit Sub

HandleError:
        Dim e As Object = Err.GetException()
        If (TypeOf e Is System.NullReferenceException) Then
            Console.WriteLine("ERR: Invalid path specified.")
        ElseIf TypeOf e Is System.Runtime.InteropServices.COMException Then
            Console.WriteLine("COMException: " & e.Message)
        Else
            Console.WriteLine(Err.Description & "[" & Err.HelpContext & "@" & Err.Erl & "]")
        End If
    End Sub
End Module
Class PowerPointExporter
    Sub Export(path As String, newPath As String)
        Dim application As New PowerPoint.Application
        application.DisplayAlerts = False
        Dim presentation As PowerPoint.Presentation = application.Presentations.Open(path, WithWindow:=False) 'open presentation and prevent it from being visible (2nd param)
        Dim helper As New PowerPointHelper(application.HWND)
        presentation.SaveAs(newPath, PowerPoint.PpSaveAsFileType.ppSaveAsPDF)
        'Hide this window::
        '    Publishing...
        '    ahk_class CMsoProgressBarWindow
        '    ahk_exe POWERPNT.EXE
        presentation.Close()
        application.Quit()
    End Sub
End Class
Class ExcelExporter
    Sub Export(path As String, newPath As String)
        Dim application As New Excel.Application
        application.Visible = False
        Dim wb As Excel.Workbook = application.Workbooks.Open(path)
        For Each ws As Excel.Worksheet In wb.Sheets
            With ws.PageSetup
                .FitToPagesTall = 1
                .FitToPagesWide = 1
            End With
        Next
        wb.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, newPath)
        application.Quit()
    End Sub
End Class
Class WordExporter
    Sub Export(path As String, newPath As String)
        Dim application As New Word.Application
        application.Visible = False
        Dim doc As Word.Document = application.Documents.Open(path)
        doc.SaveAs2(newPath, Word.WdSaveFormat.wdFormatPDF)
        doc.Close()
        application.Quit()
    End Sub
End Class

Class PowerPointHelper
    Private Declare Auto Function ShowWindow Lib "user32" (ByVal hwnd As IntPtr, nCmdShow As Integer) As Boolean
    Private Declare Function GetWindowModuleFileName Lib "user32.dll" (hwnd As IntPtr, lpszFileName As StringBuilder, cchFileNameMax As UInteger) As UInteger
    Private Declare Function GetParent Lib "user32.dll" (hwnd As IntPtr) As IntPtr

    Sub New(hwnd As UInteger)
        Dim eventHandler As New AutomationEventHandler(AddressOf OnWindowOpen)
        Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, AutomationElement.RootElement, TreeScope.Descendants, eventHandler)
    End Sub

    Public Sub OnWindowOpen(src As Object, e As AutomationEventArgs)
        Dim SourceElement As AutomationElement

        'Try to cast automation element, if stop.
        Try
            SourceElement = DirectCast(src, AutomationElement)
        Catch err As ElementNotAvailableException
            Return
        End Try

        'If window has just opened and criteria is met, hide window.
        If e.EventId Is WindowPattern.WindowOpenedEvent Then
            With SourceElement.Current
                If .Name = "Publishing..." And .ClassName = "CMsoProgressBarWindow" Then
                    Dim hwnd As IntPtr = SourceElement.Current.NativeWindowHandle
                    If Process.GetProcessById(.ProcessId).MainModule.ModuleName = "POWERPNT.EXE" Then
                        ShowWindow(hwnd, 0)
                    End If
                End If
            End With
        End If
    End Sub
End Class