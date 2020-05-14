VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCP 
   Caption         =   "Mar's Album Data Conversion"
   ClientHeight    =   3120
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   6108
   OleObjectBlob   =   "frmCP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Call Run
End Sub

Private Sub OptionButton2_Click()

End Sub

Private Sub UserForm_Initialize()
ctlOrigOrder = True
End Sub
