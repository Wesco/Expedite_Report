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
    Dim BCArry As Variant
    Dim TotalRows As Long
    Dim BuyerCodes As Variant
    Dim FoundBC As Boolean
    Dim CurrCell As String
    Dim i As Long
    Dim j As Integer

    Sheets("Expedite Report").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    BCAddr = FindColumn("WBC")
    BCArry = Array("3615CC", "3615CS", "3615EB", "3615F1", "3615F2", "3615F3", "3615F4", "3615H1", "3615H2", "3615H3", "3615H4", _
                   "3615JS", "3615LP", "3615LT", "3615MC", "3615MH", "3615ML", "3615MS", "3615PP", "3615ST", "3615T1", "3615T2", _
                   "3615T3", "3615T4", "3615W1", "3615W2", "3615W3", "3615W4", "3615SW", "3615EA", "3615DK", "3615CQ", "3615IK", _
                   "3615SK", "3615VB", "3615VK", "3615UT", "3615DR", "3615MY", "3615IR", "3615CK", "3615CP", "3615ZP", "3615ZK", _
                   "3625BR", "3625OR", "3625BS", "3625OR", "3625EF", "3625OR", "3625EK", "3625OR", "3625EW", "3625OR", "3625LT", _
                   "3625OR", "3625RC", "3625OR", "3625TB", "3625OR", "3625UT", "3625OR", "3625WD", "3625OR", "3625WH", "3625OR", _
                   "3625WJ", "3625OR", "3625JC")


    For i = TotalRows To 2 Step -1
        CurrCell = Trim(Cells(i, 1).Value & Cells(i, BCAddr).Value)
        FoundBC = False

        If Cells(i, 1).Value <> "3605" Then
            For j = 0 To UBound(BCArry)
                If CurrCell = BCArry(j) Then
                    FoundBC = True
                    Exit For
                End If
            Next
            
            If FoundBC = False Then
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
    Columns(1).Insert
    Range("A1").Value = "UID"
    POAddr = Cells(2, FindColumn("PO No")).Address(False, False)
    LNAddr = Cells(2, FindColumn("Line No")).Address(False, False)
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    Range("A2:A" & TotalRows).Formula = "=" & POAddr & "&" & LNAddr
    ActiveSheet.UsedRange.RemoveDuplicates Columns:=1, Header:=xlYes
    Columns(1).Delete
End Sub
