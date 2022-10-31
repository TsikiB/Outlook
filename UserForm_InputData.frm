VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_InputData 
   Caption         =   "Au10tix Patch - Data Input"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9540.001
   OleObjectBlob   =   "UserForm_InputData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_InputData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub UserForm_initialize()
    DataValidation = False
    CommandButton_Submit.SetFocus
End Sub

Private Sub CommandButton_Cancel_Click()
        DataValidation = False
        Unload Me
End Sub

Sub CommandButton_Submit_Click()

'On Error Resume Next
strPatchID = TextBox_PatchID
strWorkItems = TextBox_Content

For i = 1 To Len(strWorkItems)
    If IsNumeric(Mid(strWorkItems, i, 1)) Then
    Else
        strSeperander = Mid(strWorkItems, i, 1)
        Exit For
    End If
Next

For Each i In Split(TextBox_Content, strSeperander)
    If IsEmpty(arrWorkItems) Then
        arrWorkItems = Array(i)
    Else
        ReDim Preserve arrWorkItems(0 To UBound(arrWorkItems) + 1) As Variant
        arrWorkItems(UBound(arrWorkItems)) = i
    End If
Next

Select Case True
    Case OpBtn_Web.Value = True
            strDevTeam = mailAddressWeb
    Case OpBtn_BOS.Value = True
            strDevTeam = mailAddressBOS
    Case OpBtn_Infra.Value = True
            strDevTeam = mmailAddressInfra
    Case OpBtn_DataService.Value = True
            strDevTeam = mailAddressDataService
    Case OpBtn_Analytics.Value = True
            strDevTeam = mailAddressAnalytics
    Case Else
            strDevTeam = "Release.Management@Au10tix.com"
End Select
       
DataValidation = True
Unload Me

End Sub


