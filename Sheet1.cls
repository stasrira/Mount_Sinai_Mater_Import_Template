VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private bVoidDropdownChangeEvent As Boolean

Private Sub Worksheet_Activate()
    If bFieldHeadersWereSynced Then
        NotifyUserAboutFieldSyncChanges
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rCell As Range
    Dim cellRow As Integer
    Dim bStopValidation As Boolean
    
    'MsgBox "Wks Change event " & Target.Address
    
    If Not bVoidAutomatedValidation Then 'proceed only if global flag (bVoidAutomatedValidation) to cancel auto validation is not set to True
    
        If Target.EntireRow.Address = Target.Address Or Target.EntireColumn.Address = Target.Address Then
            'Entire row was changed (added or deleted). Do not proceed with validation. Optional: notify user to run whole page validation instead.
            bStopValidation = True
            
'            MsgBox "Entire row/column was modified. Validation won't be performed for such modification." & vbCrLf & vbCrLf & _
'                    "If needed, the validation still can be applied through 'Add-In/MenuCommands/MSSM Menu/Validate RawData Sheet' menu.", vbInformation

        End If
        
        If Not bStopValidation Then
        
            'for testing only ============= Start ===========
'            Dim tTime As Date
'            Debug.Print Target.Address
'            tTime = Now()
'            Debug.Print tTime
            'for testing only ============= End ===========
            
            'Proceed with validation for individual cells or islands of cells
            For Each rCell In Target.Cells
                Validate_Cell_Value rCell
            Next
            
            'for testing only ============= Start ===========
'            'Debug.Print obValidationResult.ValidatedCellProperties.CellAddress & " - " & obValidationResult.ValidatedCellProperties.CellFieldName
'            Debug.Print Now()
'            Debug.Print DateDiff("s", tTime, Now())
'            Debug.Print Target.Address
'            Debug.Print "----------------------------------------------"
            'for testing only ============= End ===========
            
        End If
    
    End If
       
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

'MsgBox "Wks Selection Change event START " & Target.Address

    Dim xCombox As OLEObject
    Dim xStr As String
    Dim xWs As Worksheet
    Set xWs = Application.ActiveSheet
    
    On Error Resume Next
    
    'Hide and unlink the dropdown from previous cell
    Set xCombox = xWs.OLEObjects("TempCombo")
    With xCombox
        .ListFillRange = ""
        .LinkedCell = ""
        .Visible = False
    End With
    
    If bVoidDropDownFunctionality Then Exit Sub 'exit sub if bVoidDropDownFunctionality flag was set to True (through the custom menu).

    bVoidDropdownChangeEvent = True 'set this flag On to avoid triggering the cell's "Change" event during linking combo box to the associated cell
    
    If Target.Validation.Type = 3 Then
        Target.Validation.InCellDropdown = False
        Cancel = True
        xStr = Target.Validation.Formula1
        xStr = Right(xStr, Len(xStr) - 1)
        If xStr = "" Then Exit Sub
        With xCombox
            .Visible = True
            .Left = Target.Left
            .Top = Target.Top
            .Width = Target.Width + 5
            .Height = Target.Height + 5
            .ListFillRange = xStr
            .LinkedCell = Target.Address
        End With
        xCombox.Activate
        Me.TempCombo.DropDown
    End If

bVoidDropdownChangeEvent = False

'MsgBox "Wks Selection Change event END " & Target.Address

End Sub
        
Private Sub TempCombo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Worksheet_Change Application.ActiveCell
    'MsgBox "Dropdown KeyDown Event " & Application.ActiveCell.Address
    
    Select Case KeyCode
        Case 9
            Application.ActiveCell.Offset(0, 1).Activate
        Case 13
            Application.ActiveCell.Offset(1, 0).Activate
    End Select
    
End Sub

Private Sub TempCombo_Change()
    If bVoidDropdownChangeEvent Then Exit Sub
    
    'MsgBox "Dropdown Change Event " & Application.ActiveCell.Address
    
    Worksheet_Change Application.ActiveCell
    'Validate_Cell_Value Application.ActiveCell

End Sub





