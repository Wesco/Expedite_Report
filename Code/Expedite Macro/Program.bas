Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    On Error GoTo Failed_Import
    UserImportFile Sheets("Expedite Report").Range("A1"), False
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    RemoveColumns
    RemoveBuyerCodes
    RemoveSODS
    RemoveLTZ
    CalculateAge
    SortAZ
    FilterAndSplit
    ExportSheets
    Clean
    Sheets("Macro").Select
    Range("C7").Select
    MsgBox "Complete!"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

Failed_Import:
    MsgBox ERR.Description, vbOKOnly, ERR.Source
End Sub

Sub RemoveColumns()
    Dim i As Integer

    Sheets("Expedite Report").Select

    For i = ActiveSheet.UsedRange.Columns.Count To 1 Step -1
        If Cells(1, i).Value <> "BR" And _
           Cells(1, i).Value <> "WBC" And _
           Cells(1, i).Value <> "PO No" And _
           Cells(1, i).Value <> "Line No" And _
           Cells(1, i).Value <> "SO Sim" And _
           Cells(1, i).Value <> "SO Item" And _
           Cells(1, i).Value <> "Supplier#" And _
           Cells(1, i).Value <> "Sim" And _
           Cells(1, i).Value <> "Item" And _
           Cells(1, i).Value <> "Desc" And _
           Cells(1, i).Value <> "Ord Tot" And _
           Cells(1, i).Value <> "Open Qty" And _
           Cells(1, i).Value <> "Line Date Requested" And _
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
    LnDtAddr = Cells(2, FindColumn("Line Date Requested")).Address(False, False)
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column + 1

    With Range(Cells(2, colPODate - 1), Cells(TotalRows, colPODate))
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
End Sub

Sub RemoveBuyerCodes()
    Dim BCAddr As Integer
    Dim TotalRows As Long
    Dim BuyerCodes As Variant
    Dim CurrCell As String
    Dim i As Long

    Sheets("Expedite Report").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    BCAddr = FindColumn("WBC")

    For i = TotalRows To 2 Step -1
        CurrCell = Trim(Cells(i, 1).Value & Cells(i, BCAddr).Value)
        If Cells(i, 1).Value <> "3605" Then
            If CurrCell <> "3615CC" And CurrCell <> "3615CQ" And CurrCell <> "3615CS" And CurrCell <> "3615DK" And _
               CurrCell <> "3615EA" And CurrCell <> "3615EB" And CurrCell <> "3615IK" And CurrCell <> "3615JS" And _
               CurrCell <> "3615T1" And CurrCell <> "3615T2" And CurrCell <> "3615T3" And CurrCell <> "3615T4" And _
               CurrCell <> "3615W1" And CurrCell <> "3615W2" And CurrCell <> "3615W3" And CurrCell <> "3615W4" And _
               CurrCell <> "3615H1" And CurrCell <> "3615H2" And CurrCell <> "3615H3" And CurrCell <> "3615H4" And _
               CurrCell <> "3615F1" And CurrCell <> "3615F2" And CurrCell <> "3615F3" And CurrCell <> "3615F4" And _
               CurrCell <> "3615LT" And CurrCell <> "3615PP" And CurrCell <> "3615MC" And CurrCell <> "3615MH" And _
               CurrCell <> "3615ML" And CurrCell <> "3615MS" And CurrCell <> "3615SK" And CurrCell <> "3615SW" And _
               CurrCell <> "3615ST" And CurrCell <> "3615VB" And CurrCell <> "3615VK" And CurrCell <> "3615DR" And _
               CurrCell <> "3615LP" And CurrCell <> "3615CK" And CurrCell <> "3625BR" And CurrCell <> "3625BS" And _
               CurrCell <> "3625EF" And CurrCell <> "3625EK" And CurrCell <> "3625EW" And CurrCell <> "3625LT" And _
               CurrCell <> "3625RC" And CurrCell <> "3625TB" And CurrCell <> "3625UT" And CurrCell <> "3625WD" And _
               CurrCell <> "3625WH" And CurrCell <> "3625WJ" Then
                Rows(i).Delete
            End If
        End If
    Next
End Sub

Sub RemoveSODS()
    Dim TotalCols As Integer
    Dim ColHeaders As Variant
    Dim SOSimCol As Integer

    Sheets("Expedite Report").Select
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    SOSimCol = FindColumn("SO Sim")
    ColHeaders = Range(Cells(1, 1), Cells(1, TotalCols))

    ActiveSheet.UsedRange.AutoFilter SOSimCol, "=*DS*", xlAnd
    ActiveSheet.UsedRange.Delete

    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, TotalCols)).Value = ColHeaders

    ActiveSheet.UsedRange.AutoFilter SOSimCol, "=*SO*", xlAnd
    ActiveSheet.UsedRange.Delete

    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, TotalCols)).Value = ColHeaders

    Columns(SOSimCol).Delete
    Columns(FindColumn("SO Item")).Delete
End Sub

Sub RemoveLTZ()
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim OpenQtyCol As Integer
    Dim ColHeaders As Variant

    Sheets("Expedite Report").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count
    ColHeaders = Range(Cells(1, 1), Cells(1, TotalCols))
    OpenQtyCol = FindColumn("Open Qty")

    With Range(Cells(2, OpenQtyCol), Cells(TotalRows, OpenQtyCol))
        .Value = .Value
    End With

    ActiveSheet.UsedRange.AutoFilter OpenQtyCol, "<=0", xlAnd
    ActiveSheet.UsedRange.Cells.Delete
    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, TotalCols)).Value = ColHeaders
End Sub

Sub ExportSheets()
    Dim FilePath As String
    Dim FileName As String
    Dim FileExt As String
    Dim NameLen As String
    Dim i As Long

    FilePath = "\\br3615gaps\gaps\Expedite Report\"
    FileName = "Expedite Report " & Format(Date, "yyyy-mm-dd")
    FileExt = ".xlsx"

    NameLen = Len(FileName)
    For i = 1 To 50
        If FileExists(FilePath & FileName & FileExt) Then
            FileName = Left(FileName, NameLen) & " (" & i & ")"
        End If
    Next

    Sheets(Array("0-14 Days", "15-30 Days", "31+ Days")).Copy
    ActiveWorkbook.SaveAs FilePath & FileName & FileExt, xlOpenXMLWorkbook
    ActiveWorkbook.Close

    Email "JAbercrombie@wescodist.com", Subject:="Expedite Report", Body:="""" & FilePath & FileName & FileExt & """"
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












































