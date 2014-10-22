Attribute VB_Name = "AHF_Mail"
Option Explicit

'Pauses for x# of milliseconds
'Used for email function to prevent
'all emails from being sent at once
'Example: "Sleep 1500" will pause for 1.5 seconds
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Enum Mail_Type
    OTLK
    SMTP
End Enum

'---------------------------------------------------------------------------------------
' Proc : Exists
' Date : 3/18/2014
' Desc : Checks to see if a file exists and has read access
'---------------------------------------------------------------------------------------
Private Function Exists(ByVal FilePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    'Remove trailing backslash
    If InStr(Len(FilePath), FilePath, "\") > 0 Then
        FilePath = Left(FilePath, Len(FilePath) - 1)
    End If

    'Check to see if the file exists and has read access
    On Error GoTo File_Error
    If fso.FileExists(FilePath) Then
        fso.OpenTextFile(FilePath, 1).Read 0
        Exists = True
    Else
        Exists = False
    End If
    On Error GoTo 0

    Exit Function

File_Error:
    Exists = False
End Function

'---------------------------------------------------------------------------------------
' Proc  : Sub Email
' Date  : 10/11/2012
' Desc  : Sends an email using Outlook
' Ex    : Email SendTo:=email@example.com, Subject:="example email", Body:="Email Body", SleepTime:=1000
'---------------------------------------------------------------------------------------
Sub Email(SendTo As String, Optional CC As String, Optional BCC As String, Optional Subject As String, Optional Body As String, _
          Optional Attachment As Variant, Optional SleepTime As Long = 0, Optional MailType As Mail_Type = Mail_Type.OTLK)

    If MailType = OTLK Then
        OTLK_Mail SendTo, CC, BCC, Subject, Body, Attachment
    ElseIf MailType = SMTP Then
        SMTP_Mail SendTo, CC, BCC, Subject, Body, Attachment
    End If

    'Wait if a sleep time was specified
    If SleepTime > 0 Then
        Sleep SleepTime
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : OTLK_Mail
' Date : 10/21/2014
' Desc : Send email using outlook
'---------------------------------------------------------------------------------------
Private Sub OTLK_Mail(SendTo As String, Optional CC As String, Optional BCC As String, Optional Subject As String, Optional Body As String, Optional Attachment As Variant)
    Dim Mail_Object As Variant    'Outlook application object
    Dim Mail_Single As Variant    'Email object
    Dim Att As Variant            'Attachment string if array is passed

    Set Mail_Object = CreateObject("Outlook.Application")
    Set Mail_Single = Mail_Object.CreateItem(0)

    'Add attachments
    Select Case TypeName(Attachment)
        Case "Variant()"
            For Each Att In Attachment
                If Att <> Empty Then
                    If Exists(Att) = True Then
                        Mail_Single.attachments.Add Att
                    End If
                End If
            Next
        Case "String"
            If Attachment <> Empty Then
                If Exists(Attachment) = True Then
                    Mail_Single.attachments.Add Attachment
                End If
            End If
    End Select

    'Setup email
    With Mail_Single
        .Subject = Subject
        .To = SendTo
        .CC = CC
        .BCC = BCC
        .HTMLbody = Body
    End With

    On Error GoTo SEND_FAILED
    Mail_Single.Send
    On Error GoTo 0
    Exit Sub

SEND_FAILED:
    MsgBox "Mail to '" & Mail_Single.To & "' could not be sent."
    Mail_Single.Delete
End Sub

'---------------------------------------------------------------------------------------
' Proc : SMTP_Mail
' Date : 10/21/2014
' Desc : Send email using SMTP
'---------------------------------------------------------------------------------------
Private Sub SMTP_Mail(SendTo As String, Optional CC As String, Optional BCC As String, Optional Subject As String, Optional Body As String, Optional Attachment As Variant)
    Const cdoSendUsingPort As Integer = 2   'Send the message using the network (SMTP over the network).
    Const cdoNTLM As Integer = 2            'NTLM Auth
    Const cdoSchema As String = "http://schemas.microsoft.com/cdo/configuration/"
    Dim objMessage As Object
    Dim Att As Variant

    Set objMessage = CreateObject("CDO.Message")

    With objMessage
        .Subject = Subject
        .From = Environ("username") & "@wesco.com"
        .To = SendTo
        .CC = CC
        .BCC = BCC
        .HTMLbody = Body
    End With

    'Add attachments
    Select Case TypeName(Attachment)
        Case "Variant()"
            For Each Att In Attachment
                If Att <> Empty Then
                    If Exists(Att) = True Then
                        objMessage.AddAttachment Att
                    End If
                End If
            Next
        Case "String"
            If Attachment <> Empty Then
                If Exists(Attachment) = True Then
                    objMessage.AddAttachment Attachment
                End If
            End If
    End Select

    With objMessage.Configuration.Fields
        .Item(cdoSchema & "sendusing") = cdoSendUsingPort
        .Item(cdoSchema & "smtpserver") = "email.wescodist.com"
        .Item(cdoSchema & "smtpauthenticate") = cdoNTLM
        .Item(cdoSchema & "smtpserverport") = 25
        .Item(cdoSchema & "smtpconnectiontimeout") = 15
        .Update
    End With

    On Error GoTo SEND_FAILED
    objMessage.Send
    On Error GoTo 0
    Exit Sub

SEND_FAILED:
    MsgBox "Mail to '" & objMessage.To & "' could not be sent."
    Set objMessage = Nothing
End Sub
