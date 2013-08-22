Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    On Error GoTo Failed_Import
    UserImportFile Sheets("Expedite Report").Range("A1"), False
    On Error GoTo 0

    RemoveColumns
    CalculateAge
    SortAZ
    FilterAndSplit

    Exit Sub

Failed_Import:
    MsgBox ERR.Description, vbOKOnly, ERR.Source
End Sub

Sub RemoveColumns()
    Dim i As Integer

    Sheets("Expedite Report").Select

    For i = ActiveSheet.UsedRange.Columns.Count To 1 Step -1
        If Cells(1, i).Value <> "BR" And _
           Cells(1, i).Value <> "BC" And _
           Cells(1, i).Value <> "po no" And _
           Cells(1, i).Value <> "line No" And _
           Cells(1, i).Value <> "SO Sim" And _
           Cells(1, i).Value <> "SO Item" And _
           Cells(1, i).Value <> "Supplier#" And _
           Cells(1, i).Value <> "Sim" And _
           Cells(1, i).Value <> "Item" And _
           Cells(1, i).Value <> "Desc" And _
           Cells(1, i).Value <> "Ord Tot" And _
           Cells(1, i).Value <> "Open Qty" And _
           Cells(1, i).Value <> "Line Date Promissed" And _
           Cells(1, i).Value <> "PO Date" And _
           Cells(1, i).Value <> "supplier name" Then
            Columns(i).Delete
        End If
    Next
End Sub

Sub CalculateAge()
    Dim TotalCols As Integer
    Dim colPODate As Integer
    Dim TotalRows As Long
    Dim PODtAddr As String
    Dim LnDtAddr As String


    Sheets("Expedite Report").Select
    colPODate = FindColumn("PO Date")
    PODtAddr = Cells(2, colPODate).Address(False, False)
    LnDtAddr = Cells(2, FindColumn("Line Date Promissed")).Address(False, False)
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column + 1

    With Range(Cells(2, colPODate - 1), Cells(TotalRows, colPODate))
        .Value = .Value
        .NumberFormat = "m/d/yyyy;@"
    End With

    Cells(1, TotalCols).Value = "PO Age"
    With Range(Cells(2, TotalCols), Cells(TotalRows, TotalCols))
        .Formula = "=IF(TODAY()-IF(" & LnDtAddr & "=""""," & PODtAddr & "," & _
                   LnDtAddr & ")<0,0,TODAY()-IF(" & LnDtAddr & "=""""," & PODtAddr & "," & LnDtAddr & "))"
        .NumberFormat = "@"
    End With
End Sub

Sub SortAZ()
    Dim TotalRows As Long
    Dim TotalCols As Integer


    Sheets("Expedite Report").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    Sheets("Expedite Report").Sort.SortFields.Clear
    Sheets("Expedite Report").Sort.SortFields.Add Key:=Range("O1"), _
                                                  SortOn:=xlSortOnValues, _
                                                  Order:=xlDescending, _
                                                  DataOption:=xlSortNormal
    With Sheets("Expedite Report").Sort
        .SetRange Range(Cells(2, 1), Cells(TotalRows, TotalCols))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    TotalCols = ActiveSheet.UsedRange.Columns.Count + 1

    Cells(1, TotalCols).Value = "Filter"
    Range(Cells(2, TotalCols), Cells(TotalRows, TotalCols)).Formula = "=IF(AND(O2>=15, O2<=30),""15-30"",IF(O2>30,""31+"",""0-15""))"
End Sub


Sub FilterAndSplit()
    Dim TotalRows As Long

    Sheets("Expedite Report").Select

    ActiveSheet.UsedRange.AutoFilter 16, "31+", xlAnd
    ActiveSheet.UsedRange.Copy Destination:=Sheets("31+ Days").Range("A1")

    ActiveSheet.UsedRange.AutoFilter 16, "15-30", xlAnd
    ActiveSheet.UsedRange.Copy Destination:=Sheets("15-30 Days").Range("A1")

    ActiveSheet.UsedRange.AutoFilter 16, "0-15", xlAnd
    ActiveSheet.UsedRange.Copy Destination:=Sheets("0-15 Days").Range("A1")

    Sheets("31+ Days").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Columns("P:P").Delete
    Range("O2:O" & TotalRows).Value = Range("O2:O" & TotalRows).Value

    Sheets("15-30 Days").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Columns("P:P").Delete
    Range("O2:O" & TotalRows).Value = Range("O2:O" & TotalRows).Value

    Sheets("0-15 Days").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Columns("P:P").Delete
    Range("O2:O" & TotalRows).Value = Range("O2:O" & TotalRows).Value

End Sub

Sub Clean()
    Dim s As Worksheet

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.AutoFilterMode = False
            s.Cells.Delete
        End If
    Next
End Sub












































