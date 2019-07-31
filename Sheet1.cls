VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Private bVoidDropdownChangeEvent As Boolean

Private Sub Worksheet_Activate()
    If bFieldHeadersWereSynced Then
        NotifyUserAboutFieldSyncChanges
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rCell As Range
    Dim cellRow As Integer
    Dim bStopValidation As Boolean
    
'    Debug.Print "RawData Change event " & Target.Address & " ==> Time: " & Now()
'    Debug.Print "bVoidAutomatedValidation = " & bVoidAutomatedValidation
    
    If Not bVoidAutomatedValidation Then 'proceed only if global flag (bVoidAutomatedValidation) to cancel auto validation is not set to True
    
        If Target.EntireRow.Address = Target.Address Or Target.EntireColumn.Address = Target.Address Then
            'Entire row was changed (added or deleted). Do not proceed with validation. Optional: notify user to run whole page validation instead.
            bStopValidation = True
        End If
        
        If Not bStopValidation Then
            'Proceed with validation for individual cells or islands of cells
            For Each rCell In Target.Cells
'                Debug.Print "RawData Change event - validation of " & rCell.Address & " ==> Time: " & Now()
                Validate_Cell_Value rCell
            Next
        End If
    
    End If
       
End Sub






