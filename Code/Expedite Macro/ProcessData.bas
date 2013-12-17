Attribute VB_Name = "ProcessData"
Option Explicit

Sub CalculateAge()
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim PODtAddr As String
    Dim colLnDt As Integer
    Dim colPODate As Integer

    Sheets("Expedite Report").Select
    colPODate = FindColumn("PO Date")
    colLnDt = FindColumn("Line Promise Date")
    PODtAddr = Cells(2, colPODate).Address(False, False)
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column + 1

    With Range(Cells(2, colPODate), Cells(TotalRows, colPODate))
        .Value = .Value
        .NumberFormat = "m/d/yyyy;@"
    End With

    With Range(Cells(2, colLnDt), Cells(TotalRows, colLnDt))
        .Value = .Value
        .NumberFormat = "m/d/yyyy;@"
    End With

    Cells(1, TotalCols).Value = "PO Age"
    With Range(Cells(2, TotalCols), Cells(TotalRows, TotalCols))
        .Formula = "=TODAY()-" & PODtAddr
        .NumberFormat = "@"
    End With
End Sub

Sub SortAZ()
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim AgeAddr As String


    Sheets("Expedite Report").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count
    AgeAddr = Cells(2, FindColumn("PO Age")).Address(False, False)

    Sheets("Expedite Report").Sort.SortFields.Clear
    Sheets("Expedite Report").Sort.SortFields.Add Key:=Cells(1, FindColumn("PO Age")), _
                                                  SortOn:=xlSortOnValues, _
                                                  Order:=xlDescending, _
                                                  DataOption:=xlSortNormal
    With Sheets("Expedite Report").Sort
        .SetRange Range(Cells(1, 1), Cells(TotalRows, TotalCols))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    TotalCols = ActiveSheet.UsedRange.Columns.Count + 1

    Cells(1, TotalCols).Value = "Filter"
    Range(Cells(2, TotalCols), Cells(TotalRows, TotalCols)).Formula = _
    "=IF(AND(" & AgeAddr & ">=15, " & AgeAddr & "<=30),""15-30"",IF(" & AgeAddr & ">30,""31+"",""0-15""))"
End Sub

Sub FilterAndSplit()
    Dim TotalRows As Long
    Dim FilterCol As Long

    Sheets("Expedite Report").Select
    FilterCol = FindColumn("Filter")

    ActiveSheet.UsedRange.AutoFilter FilterCol, "31+", xlAnd
    ActiveSheet.UsedRange.Copy Destination:=Sheets("31+ Days").Range("A1")

    ActiveSheet.UsedRange.AutoFilter FilterCol, "15-30", xlAnd
    ActiveSheet.UsedRange.Copy Destination:=Sheets("15-30 Days").Range("A1")

    ActiveSheet.UsedRange.AutoFilter FilterCol, "0-15", xlAnd
    ActiveSheet.UsedRange.Copy Destination:=Sheets("0-14 Days").Range("A1")

    Sheets("31+ Days").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Columns(FilterCol).Delete
    Range(Cells(2, FilterCol), Cells(TotalRows, FilterCol)).Value = _
    Range(Cells(2, FilterCol), Cells(TotalRows, FilterCol)).Value

    Sheets("15-30 Days").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Columns(FilterCol).Delete
    Range(Cells(2, FilterCol), Cells(TotalRows, FilterCol)).Value = _
    Range(Cells(2, FilterCol), Cells(TotalRows, FilterCol)).Value

    Sheets("0-14 Days").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Columns(FilterCol).Delete
    Range(Cells(2, FilterCol), Cells(TotalRows, FilterCol)).Value = _
    Range(Cells(2, FilterCol), Cells(TotalRows, FilterCol)).Value

    Sheets("Expedite Report").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Columns(FilterCol).Delete
    Range(Cells(2, FilterCol), Cells(TotalRows, FilterCol)).Value = _
    Range(Cells(2, FilterCol), Cells(TotalRows, FilterCol)).Value
End Sub
