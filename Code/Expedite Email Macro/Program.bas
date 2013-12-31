Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    ImportSupplierContacts Sheets("Contacts").Range("A1")
    ImportExpediteReport Sheets("Expedite Report").Range("A1")
    SendReports

End Sub

Private Sub SendReports()
    Dim SupAddress As String
    Dim SupColData As Variant
    Dim SupplierList As Variant
    Dim SupplierCol As Integer
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim i As Long

    Sheets("Expedite Report").Select
    TotalCols = ActiveSheet.UsedRange.Columns.Count + 1
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    SupplierCol = FindColumn("Supplier#")
    SupAddress = Left(Columns(SupplierCol).Address(False, False), InStr(1, Columns(SupplierCol).Address(False, False), ":") - 1) & "2"

    'Lookup Contacts
    Cells(1, TotalCols).Value = "Contact"
    With Range(Cells(2, TotalCols), Cells(TotalRows, TotalCols))
        .Formula = "=IFERROR(VLOOKUP(" & SupAddress & ", Contacts!A:B,2,FALSE),"""")"
        .Value = .Value
    End With

    'Applying sort
    With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Cells(1, SupplierCol), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SetRange ActiveSheet.UsedRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Get the list of supplier
    SupColData = Range(Cells(1, 1), Cells(TotalRows, TotalCols)).Value
    ActiveSheet.UsedRange.RemoveDuplicates SupplierCol, xlYes
    SupplierList = Range(Cells(2, SupplierCol), Cells(ActiveSheet.UsedRange.Rows.Count, SupplierCol)).Value
    Range(Cells(1, 1), Cells(TotalRows, TotalCols)).Value = SupColData

    'Filter expedite report by supplier
    For i = 1 To UBound(SupplierList, 1)
        ActiveSheet.UsedRange.AutoFilter Field:=SupplierCol, Criteria1:=SupplierList(i, 1)
        ActiveSheet.UsedRange.Copy Destination:=Sheets("Report").Range("A1")
        Sheets("Report").Select
        If Cells(2, SupplierCol).Value <> "" Then
            'Email Report
        End If
    Next
End Sub

Sub Clean()
    Dim PrevDispAlert As Boolean
    Dim s As Worksheet

    PrevDispAlert = Application.DisplayAlerts
    Application.DisplayAlerts = False

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.AutoFilterMode = False
            s.Cells.Delete
            s.Cells(1, 1).Select
        End If
    Next

    Application.DisplayAlerts = PrevDispAlert
    Sheets("Macro").Select
    Range("C7").Select
End Sub
