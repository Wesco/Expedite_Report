Attribute VB_Name = "Program"
Option Explicit

Sub CreateReport()
    On Error GoTo Failed_Import
    UserImportFile Sheets("Expedite Report").Range("A1"), True
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    RemoveColumns       'Remove columns not needed on the report
    RemoveDuplicates    'Remove duplicate lines
    RemoveBuyerCodes    'Remove buyer codes that do not need to be reviewed
    RemoveSODS          'Remove all SO and DS POs
    RemoveLTZ           'Remove all items that have been completely received or over received
    CalculateAge        'Calculate PO Age
    SortAZ              'Sort oldest to newest, add a filter column
    FilterAndSplit      'Filter the data by age and split into three sheets
    ExportSheets        'Export sheets to a new workbook and save to the network
    Clean               'Remove all data from the macro workbook

    Sheets("Macro").Select
    Range("C7").Select
    MsgBox "Complete!"

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

Failed_Import:
    MsgBox ERR.Description, vbOKOnly, ERR.Source
End Sub

'---------------------------------------------------------------------------------------
' Proc : Clean
' Date : 9/25/2013
' Desc : Removes all data from the macro workbook
'---------------------------------------------------------------------------------------
Sub Clean()
    Dim PrevDispAlert As Boolean
    Dim PrevScrnUpdat As Boolean
    Dim PrevWkbk As Workbook
    Dim s As Worksheet

    PrevDispAlert = Application.DisplayAlerts
    PrevScrnUpdat = Application.ScreenUpdating
    Set PrevWkbk = ActiveWorkbook

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    ThisWorkbook.Activate

    'Remove filters and delete cells
    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.AutoFilterMode = False
            s.Cells.Delete
            s.Cells(1, 1).Select
        End If
    Next

    Sheets("Macro").Select
    Range("C7").Select

    PrevWkbk.Activate
    Application.DisplayAlerts = PrevDispAlert
    Application.ScreenUpdating = PrevScrnUpdat
End Sub
