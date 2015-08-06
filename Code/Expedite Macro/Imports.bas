Attribute VB_Name = "Imports"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : ImportExpedite
' Date : 2/19/2014
' Desc : Imports the most recently created expedite report and sorts it
'---------------------------------------------------------------------------------------
Sub ImportExpedite()
    Dim PrevDispAlerts As Boolean
    Dim ColPromiseDate As Integer
    Dim ColSupplierNum As Integer
    Dim ColBranchNum As Integer
    Dim ColPOAge As Integer
    Dim FilePath As String
    Dim FileName As String
    Dim FileExt As String
    Dim TotalRows As Long
    Dim i As Long

    PrevDispAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    For i = 0 To 30
        FilePath = "\\br3615gaps\gaps\Expedite Report\" & Format(Date - i, "yyyy") & "\" & Format(Date - i, "mmmm") & "\"
        FileName = "Expedite Report " & Format(Date - i, "yyyy-mm-dd")
        FileExt = ".xlsx"

        If FileExists(FilePath & FileName & FileExt) Then
            Workbooks.Open FilePath & FileName & FileExt
            Sheets("Expedite Report").Select

            'Remove any previous filters
            ActiveSheet.AutoFilterMode = False

            TotalRows = ActiveSheet.UsedRange.Rows.Count

            'Find the columns to sort and filter by
            ColPromiseDate = FindColumn("Line Promise Date")
            ColSupplierNum = FindColumn("Supplier#")
            ColBranchNum = FindColumn("BR")
            ColPOAge = FindColumn("PO Age")

            If ColPromiseDate > 0 And ColSupplierNum > 0 And ColBranchNum > 0 And ColPOAge > 0 Then
                'ActiveSheet.UsedRange.AutoFilter Field:=ColPromiseDate, Criteria1:="="
                ActiveSheet.UsedRange.AutoFilter Field:=ColPromiseDate, Criteria1:="<" & Format(Date, "m/d/yyyy"), Operator:=xlOr, Criteria2:="="
                ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Expedite Report").Range("A1")
                ActiveWorkbook.Close

                Sheets("Expedite Report").Select

                With ActiveSheet.Sort
                    .SortFields.Clear

                    'Sort by supplier number
                    .SortFields.Add Key:=Range(Cells(2, ColSupplierNum), Cells(TotalRows, ColSupplierNum)), _
                                    SortOn:=xlSortOnValues, _
                                    Order:=xlAscending, DataOption:=xlSortNormal

                    'Sort by Branch Number
                    .SortFields.Add Key:=Range(Cells(2, ColBranchNum), Cells(TotalRows, ColBranchNum)), _
                                    SortOn:=xlSortOnValues, _
                                    Order:=xlAscending, _
                                    DataOption:=xlSortTextAsNumbers

                    'Sort by PO Age
                    .SortFields.Add Key:=Range(Cells(2, ColPOAge), Cells(TotalRows, ColPOAge)), _
                                    SortOn:=xlSortOnValues, _
                                    Order:=xlDescending, _
                                    DataOption:=xlSortNormal
                    .SetRange ActiveSheet.UsedRange
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                Exit For
            Else
                MsgBox "Failed to find 'Line Promise Date'"
                Err.Raise CustErr.COLNOTFOUND, "ImportExpedite", "Failed to find a column."
            End If
        End If
    Next

    Application.DisplayAlerts = PrevDispAlerts
End Sub

'---------------------------------------------------------------------------------------
' Proc : ImportContacts
' Date : 2/19/2014
' Desc : Imports the supplier contact master
'---------------------------------------------------------------------------------------
Sub ImportContacts()
    Dim PrevDispAlerts As Boolean
    Dim FilePath As String
    Dim FileName As String
    Dim TotalRows As Long

    PrevDispAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    FilePath = "\\br3615gaps\gaps\Contacts\"
    FileName = "Supplier Contact Master.xlsx"

    Workbooks.Open FilePath & FileName
    ActiveSheet.AutoFilterMode = False
    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Contact Master").Range("A1")
    ActiveWorkbook.Close
    
    'Store suppliers as text
    ThisWorkbook.Activate
    Sheets("Contact Master").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Columns("A:A").Insert Shift:=xlToRight
    Range("A1:A" & TotalRows).Formula = "=""=""&""""""""&B1&"""""""""
    Range("A1:A" & TotalRows).Value = Range("A1:A" & TotalRows).Value
    Columns("B:B").Delete

    Application.DisplayAlerts = PrevDispAlerts
End Sub
