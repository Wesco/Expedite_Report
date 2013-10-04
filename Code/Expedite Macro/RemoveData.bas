Attribute VB_Name = "RemoveData"
Option Explicit

Sub RemoveColumns()
    Dim i As Integer

    Sheets("Expedite Report").Select

    'Remove all columns that are not listed
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
           Cells(1, i).Value <> "Line Promise Date" And _
           Cells(1, i).Value <> "PO Date" And _
           Cells(1, i).Value <> "Rcd Tot" And _
           Cells(1, i).Value <> "supplier name" Then
            Columns(i).Delete
        End If
    Next
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

Sub RemoveDuplicates()
    Dim POAddr As String
    Dim LNAddr As String
    Dim TotalRows As Long

    Sheets("Expedite Report").Select
    POAddr = Cells(2, FindColumn("PO No")).Address(False, False)
    LNAddr = Cells(2, FindColumn("Line No")).Address(False, False)
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    Columns(1).Insert
    Range("A1").Value = "UID"
    Range("A2:A" & TotalRows).Formula = "=" & POAddr & "&" & LNAddr
    ActiveSheet.UsedRange.RemoveDuplicates Columns:=1, Header:=xlYes
    Columns(1).Delete
End Sub
