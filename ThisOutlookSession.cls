VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub PatchRequest_Main()

Dim WorkItemID As Double
Dim LoopExit As Boolean

LoopExit = False
'-->>>







'-->>>
UserForm_InputData.Show
If DataValidation Then
    MsgBox "Last value of arrWorkItems(" & UBound(arrWorkItems) & "): " & arrWorkItems(UBound(arrWorkItems)) & vbCrLf & _
            "Your seperator is: " & strSeperander
Else
    Exit Sub
End If

Exit Sub
'GetSMTPAddressForRecipients (Outlook.ActiveWindow)


strPatchID = InputBox("Please New Patch ID :", strInputTitle)

Do Until LoopExit
CurrentID = InputBox("Enter your first work item number for " & strPatchID & vbCrLf & "" _
             & "If this your last input, use" & """" & "." & """" & " at the end", strInputTitle)
    If CurrentID = "." Or CurrentID = "," Or Len(CurrentID) = 0 Then
       LoopExit = True
    ElseIf Right(CurrentID, 1) = "." Or Right(CurrentID, 1) = "," Then
       If Len(CurrentID) > 1 Then
            strWorkItems = strWorkItems & CurrentID
            LoopExit = True
       End If
    Else
        strWorkItems = strWorkItems & CurrentID & "."
    End If
Loop

If Len(strWorkItems) = 0 Then
    MsgBox "Process terminated. No work items were provided", vbOKOnly, strInputTitle
    Exit Sub
End If
    
'MsgBox("Enter Patch ID", vbYesNo, "Patch Outlook")
GetSMTPAddressForRecipients (Outlook.ActiveWindow)



End Sub



Function GetSMTPAddressForRecipients(mail As Outlook.MailItem)
    Dim recips As Outlook.Recipients
    Dim recip As Outlook.Recipient
    Dim pa As Outlook.PropertyAccessor
    
    Const PR_SMTP_ADDRESS As String = _
        "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
    Set recips = mail.Recipients
    'recip.Type values: To=1, CC=2, BCC = 3
    For Each recip In recips
        Set pa = recip.PropertyAccessor
           MsgBox recip.Name & "; " & recip.Type & "; SMTP=" _
           & pa.GetProperty(PR_SMTP_ADDRESS)
    Next
    Set OutMail = Nothing
End Function

