VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    'last 2 parameters sets shift of the "row" and "column" position cell (on the sheet) relative to the current cell
    UpdatePositionValues Me.UsedRange, 1, 1, 2

    AssignFormulaToCheckVesselIdMatch Me.UsedRange, 1 'assign formula to corresponding cells of column containing flags for matching to Vessel ID field on Rawdata
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    'do not proceed if whole row or column was updated
    If Target.EntireRow.Address = Target.Address Or Target.EntireColumn.Address = Target.Address Then
        Exit Sub
    End If
    
    
    'apply position changes only if first column of the sheet is updated
    If Target.Cells(1).Column = 1 Then
        'last 2 parameters sets shift of the "row" and "column" position cell (on the sheet) relative to the current cell
        UpdatePositionValues Target, 1, 1, 2

        AssignFormulaToCheckVesselIdMatch Target, 1 'assign formula to corresponding cells of column containing flags for matching to Vessel ID field on Rawdata
    End If
    
End Sub

Private Sub Worksheet_Deactivate()
    RequestCopyScansToDict "MT_Vessel ID", Me.Name, "A1"
End Sub
