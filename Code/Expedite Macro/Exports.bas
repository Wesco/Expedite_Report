Attribute VB_Name = "Exports"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : ExportSheets
' Date : 9/25/2013
' Desc : Save report to the disk
'---------------------------------------------------------------------------------------
Sub ExportSheets()
    Dim PrevDispAlert As Boolean
    Dim SaveDialog As Object
    Dim FilePath As String
    Dim FileName As String
    Dim FileExt As String
    Dim NameLen As String
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim s As Worksheet
    Dim SaveFile As Variant
    Dim i As Long

    Set SaveDialog = Application.FileDialog(msoFileDialogSaveAs)

    PrevDispAlert = Application.DisplayAlerts
    FilePath = Environ("userprofile") + "\Desktop\"
    FileName = "Expedite Report " & Format(Date, "yyyy-mm-dd")
    FileExt = ".xlsx"

    Sheets(Array("Expedite Report", "0-14 Days", "15-30 Days", "31+ Days")).Copy

    For Each s In ActiveWorkbook.Sheets
        s.Select
        TotalCols = s.UsedRange.Columns.Count
        s.UsedRange.AutoFilter
        s.Range(Cells(1, 1), Cells(1, TotalCols)).HorizontalAlignment = xlCenter
        s.UsedRange.Columns.AutoFit
    Next

    Sheets(1).Select
    Application.DisplayAlerts = True

    On Error GoTo Save_Err
    With SaveDialog
        .InitialFileName = FilePath & FileName & FileExt
        SaveFile = .Show
    End With
    
    If SaveFile = -1 Then
        ActiveWorkbook.SaveAs FileName:=FilePath & FileName & FileExt, FileFormat:=xlWorkbook
    ElseIf SaveFile = 0 Then
        MsgBox "Expedite report not saved."
    End If
    On Error GoTo 0

    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlert
    Exit Sub

Save_Err:
    MsgBox "An error occurred while trying to save."
    Resume
End Sub
