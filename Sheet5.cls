VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    'in case of any changes on the Field Settings sheet, clean all values out of dictFieldSettings dictionary
    dictFieldSettings.RemoveAll
    
    If Target.Column = 1 Then
        'if the first column was modified, list of fields should be re-sync betweent the FieldSetting sheet and RawData and Validated sheets
        SyncFieldsAccrossSheets
    End If
    
    'Apply dropdown settings to appropriate cells on RawData
    ApplyDropdownSettingsToCells
    
End Sub


