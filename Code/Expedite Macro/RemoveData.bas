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
