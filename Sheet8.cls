VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    'apply position changes only if first column of the sheet is updated
    If Target.Column = 1 Then
        'last 2 parameters sets shift of the "row" and "column" position cell (on the sheet) relative to the current cell
        GetPositionValues Target, 1, 2
    End If
End Sub

Private Sub Worksheet_Deactivate()
    RequestCopyScansToDict "MT_Vessel ID", Me.Name, "A1"
End Sub
