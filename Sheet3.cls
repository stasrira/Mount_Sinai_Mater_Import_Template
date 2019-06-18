VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    're-evaluate formula matching barcodes existing on a sheet to MT_Vessel IDs on RawData or Validated sheets

    AssignFormulaToCheckVesselIdMatch Me.UsedRange, 2
    
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    
    On Error GoTo err1
    
    Dim rCol As Range, i As Integer
    
    'do not proceed if whole row or column was delete/updated
    If Target.EntireRow.Address = Target.Address Or Target.EntireColumn.Address = Target.Address Then
        Exit Sub
    End If
    
    'apply the logic below only for first 4 columns (the ones being copied from the flatbed scanner file); exit sub in other cases.
    If Target.Column > 4 Then
        Exit Sub
    End If
    
    AssignFormulaToCheckVesselIdMatch Target, 2 'assign formula to corresponding cells of column containing flags for matching to Vessel ID field on Rawdata

    'ask user if records below insertion/update have to be deleted
    If Target.Cells.Count > 1 Then
        Dim iResponse As Integer, rStart As Long, rEnd As Long
        
        iResponse = MsgBox("Do you want to proceed with deleting any values located under the area being updated with new values?", _
                        vbOKCancel, "Pasting Flatbed Scanner Output")
                        
        If iResponse = vbOK Then
            With Target.Worksheet
                rStart = Target.Offset(Target.Rows.Count).Cells(1).Row
                rEnd = Target.Offset(.Rows.Count - rStart).End(xlUp).Cells(1).Row + 1
                'clear un-needed values
                Application.ScreenUpdating = False
                Application.EnableEvents = False
                
                Range("A" & CStr(rStart) & ":" & "D" & CStr(rEnd)).Clear 'clear records located under the entered ones
                
                Application.ScreenUpdating = True
                Application.EnableEvents = True
            End With
        End If
    End If
    
    Exit Sub
    
err1:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
 
End Sub

Private Sub Worksheet_Deactivate()
    'The following will handle copying data entered on the Flatbedscanner sheet to the Dictionary sheet to populate values for Vessel ID and Box ID dropdowns
    
    RequestCopyScansToDict "MT_Box Label", Me.Name, "D1"
    RequestCopyScansToDict "MT_Vessel ID", Me.Name, "B1"
    
    ApplyDropdownSettingsToCells 'apply any dropdown value changes to the dropdown fields
    
End Sub





