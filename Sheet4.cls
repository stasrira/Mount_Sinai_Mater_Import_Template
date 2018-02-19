VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    dictFieldSettings.RemoveAll
'    MsgBox "The last modification performed on the ""Dictionary"" sheet caused re-setting of the ""FieldSetting"" values (default values might got affected)." & vbCrLf & vbCrLf & _
'        "Consider re-running validation of the ""RawData"" sheet to make sure that all validation results are current", vbInformation, "Modification of ""Dictionary"" Sheet"
End Sub
