Attribute VB_Name = "Program"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : CreateReport
' Date : 2/19/2014
' Desc : Imports an expedite report from WPS, then formats and saves it to the network
'---------------------------------------------------------------------------------------
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
    MsgBox Err.Description, vbOKOnly, Err.Source
End Sub

'---------------------------------------------------------------------------------------
' Proc : ImportReport
' Date : 2/19/2014
' Desc : Imports the most recent expedite report and gets
'---------------------------------------------------------------------------------------
Sub ImportReport()
    ImportExpedite
    ImportContacts
End Sub

'---------------------------------------------------------------------------------------
' Proc : SendEmail
' Date : 2/19/2014
' Desc : Emails items to their respective suppliers
'---------------------------------------------------------------------------------------
Sub SendEmail()
    'Email variables
    Dim Body As String
    Dim Subject As String
    Dim SuppName As String
    Dim Contact As String
    Dim PONumber As String
    Dim Created As String
    Dim Branch As String

    'Loop conditionals
    Dim PrevCell As String
    Dim CurrCell As String
    Dim NextCell As String
    Dim StartRow As Long
    Dim EndRow As Long
    Dim TotalRows As Long

    'Loop counters
    Dim i As Long
    Dim j As Long

    Sheets("PO Conf").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Branch = Sheets("473").Range("A2")

    For i = 2 To TotalRows
        PrevCell = Cells(i - 1, 3).Value
        CurrCell = Cells(i, 3).Value
        NextCell = Cells(i + 1, 3).Value

        If CurrCell <> PrevCell And CurrCell <> NextCell Then
            'Only one PO for this supplier
            PONumber = Cells(i, 1).Value
            Created = Format(Cells(i, 2).Value, "mmm dd, yyyy")
            SuppName = Cells(i, 4).Value

            Contact = Cells(i, 5).Value
            Subject = "Please send an estimated ship date for PO# " & Branch & "-" & PONumber
            Body = "<tr>" & _
                   "<td>" & Branch & "-" & PONumber & "</td>" & _
                   "<td>" & Created & "</td>" & _
                   "<td>" & SuppName & "</td>" & _
                   "</tr>"

            If Contact <> "" Then
                Email Contact, Subject:=Subject, Body:=EmailHeader & Body & EmailFooter
            End If

            'Reset email body
            Body = ""
        ElseIf CurrCell = NextCell And CurrCell <> PrevCell Then
            'First cell for this supplier
            StartRow = i
        ElseIf CurrCell <> NextCell And CurrCell = PrevCell Then
            'Last cell for this supplier
            EndRow = i

            'Add all rows to the email in a table
            For j = StartRow To EndRow
                PONumber = Cells(j, 1).Value
                Created = Format(Cells(j, 2).Value, "mmm dd, yyyy")
                SuppName = Cells(j, 4).Value

                Body = Body & "<tr>" & _
                       "<td>" & Branch & "-" & PONumber & "</td>" & _
                       "<td>" & Created & "</td>" & _
                       "<td>" & SuppName & "</td>" & _
                       "</tr>"
            Next
            Subject = "Please send estimated ship dates"
            Contact = Cells(i, 5).Value

            If Contact <> "" Then
                Email Contact, Subject:=Subject, Body:=EmailHeader & Body & EmailFooter
            End If

            'Reset email body
            Body = ""
        End If
    Next

    MsgBox "Complete!"
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
