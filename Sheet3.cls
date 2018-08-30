VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Change(ByVal Target As Range)
    
    On Error GoTo Err1
    
    Dim rCell As Range
    
    'apply the logic below only for first 4 columns (the ones being copied from the flatbed scanner file); exit sub in other cases.
    If Target.Column > 4 Then
        Exit Sub
    End If
    
    Application.EnableEvents = False
    For Each rCell In Target.Cells
        'If IsNumeric(rCell.Value) Then
        If IsNumeric(rCell.Value) And Not IsEmpty(rCell.Value) Then 'IsEmpty is used to avoid adding "'" to an empty cells that are considered Numeric by VBA
            rCell.Value = "'" & rCell.Value
        End If
    Next
    
    If Target.Cells.Count > 1 Then
        Dim iResponse As Integer, rStart As Long, rEnd As Long
    
        iResponse = MsgBox("Do you want to proceed with deleting any values located under the area being updated with new values?", _
                        vbOKCancel, "Pasting Flatbed Scanner Output")
                        
        If iResponse = vbOK Then
            With Target.Worksheet
                rStart = Target.Offset(Target.Rows.Count).Cells(1).Row
                rEnd = Target.Offset(.Rows.Count - rStart).Cells(1).Row
                Range("A" & CStr(rStart) & ":" & "D" & CStr(rEnd)).Clear
            End With
        End If
    End If
    
Err1:
    Application.EnableEvents = True
 
End Sub

Private Sub Worksheet_Deactivate()
    'The following will handle copying data entered on the Flatbedscanner sheet to the Dictionary sheet to populate values for Vessel ID and Box ID dropdowns
    
    Dim oFieldSettings As clsFieldSettings
    Dim sourceRangeAddress As String
    
    'set parameters for MT_Box Label field.
    Set oFieldSettings = GetFieldSettingsInstance(Nothing, False, "MT_Box Label")
    If oFieldSettings.DataAvailable And Len(Trim(oFieldSettings.FieldDropDownValueLookupRange)) Then
        sourceRangeAddress = "D1" 'first cell of the column storring the list of Box IDs on the FlatbedScans tab
        CopyValuesToDictionarySheet sourceRangeAddress, oFieldSettings.FieldDropDownValueLookupRange, True
    End If
    'CopyBoxIDsToDictionarySheet oFieldSettings.FieldDropDownValueLookupRange
    'CopyBoxIDsToDictionarySheet "BA3" ' - for test only
    
    Set oFieldSettings = GetFieldSettingsInstance(Nothing, False, "MT_Vessel ID")
    If oFieldSettings.DataAvailable And Len(Trim(oFieldSettings.FieldDropDownValueLookupRange)) Then
        sourceRangeAddress = "B1" 'first cell of the column storring the list of Tube IDs on the FlatbedScans tab
        CopyValuesToDictionarySheet sourceRangeAddress, oFieldSettings.FieldDropDownValueLookupRange, True
    End If
    
    ApplyDropdownSettingsToCells 'apply any dropdown value changes to the dropdown fields
    
End Sub

'copy values (i.e. Box ID ) entered on the Flatbedscanner sheet to the Dictionary sheet to populate values for corresponede dropdown (i.e. MT_Box ID)
Private Sub CopyValuesToDictionarySheet(sourceRandgeAddress As String, targetRangeAddress As String, Optional CopyUniqueValuesOnly As Boolean = False)

    With Worksheets(cFlatbedScansWorksheetName)
        
        Dim targetRange As Range, sourceRange As Range, tRange As Range
        Dim targetRangeStart As Range, targetRangeEnd As Range
        
        Set sourceRange = .Range(sourceRandgeAddress & ":" & .Range(sourceRandgeAddress).Offset(.Rows.Count - .Range(sourceRandgeAddress).Row).End(xlUp).Address) 'this source range will include all cells (in this column) located below the given cell
        
        With Worksheets(cDictionayWorksheetName)
            'clear target range
            Set targetRangeStart = .Range(targetRangeAddress)
            Set targetRangeEnd = .Range(targetRangeAddress).Offset(.Rows.Count - .Range(targetRangeAddress).Row).End(xlUp)
            If targetRangeEnd.Row < targetRangeStart.Row Then
                Set targetRangeEnd = targetRangeStart
            End If
            Set targetRange = .Range(targetRangeStart.Address & ":" & targetRangeEnd.Address)
            'Set targetRange = .Range(targetRangeAddress & ":" & .Range(targetRangeAddress).Offset(.Rows.Count - .Range(targetRangeAddress).Row).End(xlUp).Address)
            targetRange.Clear
            targetRange.Offset(0, 2).Clear
            
            'reset target range to point to the first cell only, this is required to copy all available cells from the source range
            Set targetRange = .Range(targetRangeAddress) 'this range points the first cell in column that will hold copied values (i.e. "BA3")
        End With
        
        If sourceRange.Cells.Count > 1 Then 'proceed with copying data only if the source has some data (beside the header cell)
            'copy unique list of Box IDs from Flatbed scanner sheet to Dictionary. Because Box IDs repeats for muptiple Barcodes, only unique values will be copied to dictionary
            sourceRange.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=targetRange, Unique:=CopyUniqueValuesOnly
            'copy the same information (as in previous row) to the 2nd column shifted to the right on 2 cells
            sourceRange.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=targetRange.Offset(0, 2), Unique:=CopyUniqueValuesOnly
            '.Range("D1:" & .Range("D1").End(xlDown).Address).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets(cDictionayWorksheetName).Range("BC3"), Unique:=True
            
            'sort all copied values (in both columns)
            Set tRange = Worksheets(cDictionayWorksheetName).Range(targetRange.Address & ":" & targetRange.Offset(0, 2).End(xlDown).Address)
            tRange.Sort key1:=tRange, Header:=xlYes
            
            'delete header titles (Excel always copies headers) from the target range
            targetRange.Offset(0, 2).Delete xlShiftUp
            targetRange.Delete xlShiftUp 'delete the first cell (that holds title of the source column) from the target range
            
             With Worksheets(cDictionayWorksheetName)
                'this will re-evaluate range of the actually copied values and set number format for each cell to "Text"
                'clear target range
                        'Old code - TO DELETE - Set targetRange = .Range(targetRangeAddress & ":" & .Range(targetRangeAddress).Offset(.Rows.Count - .Range(targetRangeAddress).Row).End(xlUp).Address)
                'After deleting cells above, Start and End Ranges have to be re-defined
                Set targetRangeStart = .Range(targetRangeAddress)
                Set targetRangeEnd = .Range(targetRangeAddress).Offset(.Rows.Count - .Range(targetRangeAddress).Row).End(xlUp)
                Set targetRange = .Range(targetRangeStart.Address & ":" & targetRangeEnd.Address)
                '.Rows.Count - .Row
                targetRange.NumberFormat = "@" 'this will set number format to Text
                targetRange.Offset(0, 2).NumberFormat = "@" 'this will set number format to Text
                
                Dim rBlanks1 As Range, rBlanks2 As Range
                'this will delete blank cells in the target range. This might be needed if flatbed scanner returns some empty values for absent tubes
                If targetRange.Cells.Count > 1 Then
                    On Error GoTo Err1 'this should catch cases when no empty cells are found - excel generates an error in such situation
                    
                    Set rBlanks1 = targetRange.Offset(0, 2).SpecialCells(xlCellTypeBlanks)
                    Set rBlanks2 = targetRange.SpecialCells(xlCellTypeBlanks)
                    rBlanks1.Delete xlShiftUp
                    rBlanks2.Delete xlShiftUp
                    
Err1:
                    On Error GoTo 0 'resume default error handling
                    
                End If
            End With
        End If
        
    End With
End Sub


