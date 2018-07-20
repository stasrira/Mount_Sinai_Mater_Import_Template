VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelection 
   ClientHeight    =   2730
   ClientLeft      =   -45
   ClientTop       =   -150
   ClientWidth     =   8310.001
   OleObjectBlob   =   "frmSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbProfileList_Change()
    Dim prof_details As clsFieldSettingProfile
    
    Set prof_details = dictProfiles(Me.cmbProfileList.ListIndex)
    'Debug.Print prof_details.Description, prof_details.Owner, prof_details.Name, Format(prof_details.Created, "Short Date")
    
    'lblProfDetails.Caption = prof_details.Name & vbCrLf & prof_details.Description & vbCrLf & "Last updated on " & Format(prof_details.Created, "Short Date")
    txtDesc.Text = prof_details.Name & vbCrLf & prof_details.Description & vbCrLf & "Last updated on " & Format(prof_details.Created, "Short Date")
    txtDesc.Locked = True
    txtDesc.MultiLine = True
    
    txtCurProfile.Text = "Last loaded profile: " & GetConfigValue("FieldSetting_LastLoadedProfile")
    txtCurProfile.Locked = True
    txtCurProfile.MultiLine = True
    
End Sub

Private Sub cmbProfileList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then 'if enter was pressed on the dropdown control, load the form
        cmdLoad_Click
    ElseIf KeyCode = 27 And Shift = 0 Then 'if Escape was pressed, close the form
        cmdCancel_Click
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancel_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 And Shift = 0 Then 'if Escape was pressed, close the form
        cmdCancel_Click
    End If
End Sub

Private Sub cmdLoad_Click()
    popUpFormResponseIndex = Me.cmbProfileList.ListIndex 'pass selected index to the global variable
    
    Unload Me
End Sub

Private Sub cmdLoad_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 And Shift = 0 Then 'if Escape was pressed, close the form
        cmdCancel_Click
    End If
End Sub

