Attribute VB_Name = "Imports"
Option Explicit

Sub ImportExpediteReport(Destination As Range)
    Dim PrevDispAlert As Boolean
    Dim FilePath As String
    Dim FileName As String
    Dim Found As Boolean
    Dim i As Integer
    Dim dt As Date

    PrevDispAlert = Application.DisplayAlerts
    Application.DisplayAlerts = False

    For i = 0 To 60
        dt = Date - i
        FilePath = "\\br3615gaps\gaps\Expedite Report\" & Format(dt, "yyyy") & "\" & Format(Date, "mmmm") & "\"
        FileName = "Expedite Report " & Format(dt, "yyyy-mm-dd") & ".xlsx"
        If FileExists(FilePath & FileName) Then
            Found = True
            Workbooks.Open FilePath & FileName
            Sheets("Expedite Report").Select
            ActiveSheet.UsedRange.Copy Destination:=Destination
            ActiveWorkbook.Close
            Exit For
        End If
    Next

    Application.DisplayAlerts = PrevDispAlert

    If Found = False Then
        Err.Raise Errors.FILE_NOT_FOUND, "ImportExpediteReport", "Expedite report not found"
    End If
End Sub
