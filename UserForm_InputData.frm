VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_InputData 
   Caption         =   "Au10tix Patch - Data Input"
   ClientHeight    =   4500
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
  CommandButton_Submit.SetFocus
End Sub

Private Sub CommandButton_Cancel_Click()
        Unload Me
End Sub

Private Sub CommandButton_RunUpdate_Click()

If DataValidation Then
End If



Public Function DataValidation()





End Sub
