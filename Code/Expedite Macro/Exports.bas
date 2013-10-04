Attribute VB_Name = "Exports"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : ExportSheets
' Date : 9/25/2013
' Desc : Save report to the network
'---------------------------------------------------------------------------------------
Sub ExportSheets()
    Dim FilePath As String
    Dim FileName As String
    Dim FileExt As String
    Dim NameLen As String
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim s As Worksheet
    Dim i As Long

    FilePath = "\\br3615gaps\gaps\Expedite Report\" & Format(Date, "yyyy") & "\" & Format(Date, "mmmm") & "\"
    FileName = "Expedite Report " & Format(Date, "yyyy-mm-dd")
    FileExt = ".xlsx"

    If Not FolderExists(FilePath) Then
        RecMkDir FilePath
    End If

    NameLen = Len(FileName)
    For i = 1 To 50
        If FileExists(FilePath & FileName & FileExt) Then
            FileName = Left(FileName, NameLen) & " (" & i & ")"
        End If
    Next

    Sheets(Array("0-14 Days", "15-30 Days", "31+ Days")).Copy

    For Each s In ActiveWorkbook.Sheets
        s.Select
        TotalCols = s.UsedRange.Columns.Count
        s.UsedRange.AutoFilter
        s.Range(Cells(1, 1), Cells(1, TotalCols)).HorizontalAlignment = xlCenter
        s.UsedRange.Columns.AutoFit
    Next

    Sheets(1).Select
    ActiveWorkbook.SaveAs FilePath & FileName & FileExt, xlOpenXMLWorkbook
    ActiveWorkbook.Close

    Email "JAbercrombie@wescodist.com", Subject:="Expedite Report", Body:="""" & FilePath & FileName & FileExt & """"
End Sub
