VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelection 
   Caption         =   "UserForm1"
   ClientHeight    =   2040
   ClientLeft      =   80
   ClientTop       =   300
   ClientWidth     =   8430.001
   OleObjectBlob   =   "frmSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLoad_Click()

    popUpFormResponseIndex = Me.cmbProfileList.ListIndex
    
    Unload Me
End Sub
