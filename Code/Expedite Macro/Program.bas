Attribute VB_Name = "Program"
Option Explicit
Public Const RepositoryName As String = "Expedite_Report"
Public Const VersionNumber As String = "1.1.2"

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
    LookupEmails
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
    Dim ItemNum As String
    Dim SimNum As String
    Dim PromDt As String

    'Loop conditionals
    Dim PrevCell As String
    Dim CurrCell As String
    Dim NextCell As String
    Dim StartRow As Long
    Dim EndRow As Long
    Dim TotalRows As Long

    'Column numbers
    Dim ColPromDt As Integer
    Dim ColContact As Integer
    Dim ColSupNum As Integer
    Dim ColPONum As Integer
    Dim ColBrNum As Integer
    Dim ColName As Integer
    Dim ColDate As Integer
    Dim ColItem As Integer
    Dim ColSim As Integer

    'Loop counters
    Dim i As Long
    Dim j As Long

    Sheets("Expedite Report").Select

    TotalRows = ActiveSheet.UsedRange.Rows.Count
    ColPromDt = FindColumn("Line Promise Date")
    ColPONum = FindColumn("PO No")
    ColBrNum = FindColumn("BR")
    ColName = FindColumn("supplier name")
    ColDate = FindColumn("PO Date")
    ColSupNum = FindColumn("Supplier#")
    ColContact = FindColumn("Email")
    ColItem = FindColumn("Item")
    ColSim = FindColumn("Sim")

    If ColPONum = 0 Or ColBrNum = 0 Or ColName = 0 Or ColDate = 0 Or ColPromDt = 0 Or _
       ColSupNum = 0 Or ColContact = 0 Or ColItem = 0 Or ColSim = 0 Then
        Err.Raise CustErr.COLNOTFOUND, "SendEmail", "A column could not be found."
    End If

    For i = 2 To TotalRows
        PrevCell = Cells(i - 1, ColSupNum).Value
        CurrCell = Cells(i, ColSupNum).Value
        NextCell = Cells(i + 1, ColSupNum).Value

        If CurrCell <> PrevCell And CurrCell <> NextCell Then
            'Only one PO for this supplier
            Branch = Cells(i, ColBrNum).Value
            PONumber = Cells(i, ColPONum).Value
            Created = Format(Cells(i, ColDate).Value, "mmm dd, yyyy")
            SuppName = Cells(i, ColName).Value
            SimNum = Cells(i, ColSim).Value
            ItemNum = Cells(i, ColItem).Value
            PromDt = Format(Cells(i, ColPromDt).Value, "mmm dd, yyyy")

            Contact = Cells(i, ColContact).Value
            Subject = "Please send an estimated ship date for PO# " & Branch & "-" & PONumber
            Body = "<tr>" & _
                   "<td>" & Branch & "-" & PONumber & "</td>" & _
                   "<td>" & Created & "</td>" & _
                   "<td>" & PromDt & "</td>" & _
                   "<td>" & SimNum & ItemNum & "</td>" & _
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
                Branch = Cells(j, ColBrNum).Value
                PONumber = Cells(j, ColPONum).Value
                Created = Format(Cells(j, ColDate).Value, "mmm dd, yyyy")
                SuppName = Cells(j, ColName).Value
                SimNum = Cells(j, ColSim).Value
                ItemNum = Cells(j, ColItem).Value
                PromDt = Format(Cells(j, ColPromDt).Value, "mmm dd, yyyy")

                Body = Body & "<tr>" & _
                       "<td>" & Branch & "-" & PONumber & "</td>" & _
                       "<td>" & Created & "</td>" & _
                       "<td>" & PromDt & "</td>" & _
                       "<td>" & SimNum & ItemNum & "</td>" & _
                       "<td>" & SuppName & "</td>" & _
                       "</tr>"
            Next
            Subject = "Please send estimated ship dates"
            Contact = Cells(i, ColContact).Value

            If Contact <> "" Then
                Email Contact, Subject:=Subject, Body:=EmailHeader & Body & EmailFooter
            End If

            'Reset email body
            Body = ""
        End If
    Next

    MsgBox "Complete!"
End Sub

Private Function EmailHeader()
    EmailHeader = "<html>" & _
                  "<style>" & _
                  "table{border:1px solid black; border-collapse:collapse;}" & _
                  "table,th,td{border:1px solid black;}" & _
                  "td{padding:5px; text-align:center;}" & _
                  "th{padding:5px;}" & _
                  "</style>" & _
                  "Dear Supplier," & _
                  "<br>" & _
                  "<br>" & _
                  "Please review the list of orders below and confirm that they have been received and provide an estimated ship date. " & _
                  "<br>" & _
                  "If you are receiving this for a second time, we may not have received an estimated shipping date in your original response or the promise date has passed." & _
                  "<br>" & _
                  "<br>" & _
                  "<br>" & _
                  "<table>" & _
                  "<th>PO</th><th>CREATED</th><th>PROMISED</th><th>SIM</th><th>SUPPLIER</th>"
End Function

Private Function EmailFooter()
    EmailFooter = "</table>" & _
                  "<br>" & _
                  "<br>" & _
                  "Thanks in advance for your help!<br>" & _
                  "<br>" & _
                  "<span style='font-size:8.0pt;font-family:""Arial"",""sans-serif""'>" & _
                  Environ("username") & "@wesco.com" & " | office: 704-393-6636 | fax: 704-393-6645<br>" & _
                  "<b>WESCO Distribution<br>" & _
                  "5521 Lakeview Road, Suite W, Charlotte, NC 28269</b>" & _
                  "</span>" & _
                  "</body></html>"
End Function

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
