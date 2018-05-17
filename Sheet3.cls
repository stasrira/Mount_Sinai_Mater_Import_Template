VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Deactivate()
    'The following will handle copying data entered on the Flatbedscanner sheet to the Dictionary sheet to populate values for Vessel ID and Box ID dropdowns
    
    Dim oFieldSettings As clsFieldSettings
    Dim sourceRangeAddress As String
    
    'set parameters for MT_Box Label field.
    Set oFieldSettings = GetFieldSettingsInstance(Nothing, False, "MT_Box Label")
    sourceRangeAddress = "D1" 'first cell of the column storring the list of Box IDs on the FlatbedScans tab
    CopyValuesToDictionarySheet sourceRangeAddress, oFieldSettings.FieldDropDownValueLookupRange, True
    'CopyBoxIDsToDictionarySheet oFieldSettings.FieldDropDownValueLookupRange
    'CopyBoxIDsToDictionarySheet "BA3" ' - for test only
    
    Set oFieldSettings = GetFieldSettingsInstance(Nothing, False, "MT_Vessel ID")
    sourceRangeAddress = "B1" 'first cell of the column storring the list of Tube IDs on the FlatbedScans tab
    CopyValuesToDictionarySheet sourceRangeAddress, oFieldSettings.FieldDropDownValueLookupRange, True
    
    ApplyDropdownSettingsToCells 'apply any dropdown value changes to the dropdown fields
    
'    With Worksheets(cFlatbedScansWorksheetName)
'
'        Dim targetRange As Range, sourceRange As Range, tRange As Range
'
'        Set sourceRange = .Range("D1:" & .Range("D1").End(xlDown).Address) 'this source range will include all cells (in this column) located below the given cell
'        Set targetRange = Worksheets(cDictionayWorksheetName).Range("BA3") 'this range points the first cell in column that will hold copied values
'
'        targetRange.Clear
'        targetRange.Offset(0, 2).Clear
'
'        'copy unique list of Box IDs from Flatbed scanner sheet to Dictionary. Because Box IDs repeats for muptiple Barcodes, only unique values will be copied to dictionary
'        sourceRange.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=targetRange, Unique:=True
'        'copy the same information (as in previous row) to the 2nd column shifted to the right on 2 cells
'        sourceRange.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=targetRange.Offset(0, 2), Unique:=True
'        '.Range("D1:" & .Range("D1").End(xlDown).Address).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets(cDictionayWorksheetName).Range("BC3"), Unique:=True
'
'        'sort all copied values (in both columns)
'        Set tRange = Worksheets(cDictionayWorksheetName).Range(targetRange.Address & ":" & targetRange.Offset(0, 2).End(xlDown).Address)
'        tRange.Sort key1:=tRange, Header:=xlYes
'
'        'delete header titles (Excel always copies headers) from the target range
'        targetRange.Offset(0, 2).Delete xlShiftUp
'        targetRange.Delete xlShiftUp 'delete the first cell (that holds title of the source column) from the target range
'
'    End With
End Sub

'copy Box ID values entered on the Flatbedscanner sheet to the Dictionary sheet to populate values for MT_Box ID dropdown
'Private Sub CopyBoxIDsToDictionarySheet(targetRangeAddress As String)
'
'    With Worksheets(cFlatbedScansWorksheetName)
'
'        Dim targetRange As Range, sourceRange As Range, tRange As Range
'
'        Set sourceRange = .Range("D1:" & .Range("D1").End(xlDown).Address) 'this source range will include all cells (in this column) located below the given cell
'        Set targetRange = Worksheets(cDictionayWorksheetName).Range(targetRangeAddress) 'this range points the first cell in column that will hold copied values (i.e. "BA3")
'
'        targetRange.Clear
'        targetRange.Offset(0, 2).Clear
'
'        'copy unique list of Box IDs from Flatbed scanner sheet to Dictionary. Because Box IDs repeats for muptiple Barcodes, only unique values will be copied to dictionary
'        sourceRange.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=targetRange, Unique:=True
'        'copy the same information (as in previous row) to the 2nd column shifted to the right on 2 cells
'        sourceRange.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=targetRange.Offset(0, 2), Unique:=True
'        '.Range("D1:" & .Range("D1").End(xlDown).Address).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets(cDictionayWorksheetName).Range("BC3"), Unique:=True
'
'        'sort all copied values (in both columns)
'        Set tRange = Worksheets(cDictionayWorksheetName).Range(targetRange.Address & ":" & targetRange.Offset(0, 2).End(xlDown).Address)
'        tRange.Sort key1:=tRange, Header:=xlYes
'
'        'delete header titles (Excel always copies headers) from the target range
'        targetRange.Offset(0, 2).Delete xlShiftUp
'        targetRange.Delete xlShiftUp 'delete the first cell (that holds title of the source column) from the target range
'
'    End With
'End Sub



'copy values (i.e. Box ID ) entered on the Flatbedscanner sheet to the Dictionary sheet to populate values for corresponede dropdown (i.e. MT_Box ID)
Private Sub CopyValuesToDictionarySheet(sourceRandgeAddress As String, targetRangeAddress As String, Optional CopyUniqueValuesOnly As Boolean = False)

    With Worksheets(cFlatbedScansWorksheetName)
        
        Dim targetRange As Range, sourceRange As Range, tRange As Range
        
        Set sourceRange = .Range(sourceRandgeAddress & ":" & .Range(sourceRandgeAddress).End(xlDown).Address) 'this source range will include all cells (in this column) located below the given cell
        Set targetRange = Worksheets(cDictionayWorksheetName).Range(targetRangeAddress) 'this range points the first cell in column that will hold copied values (i.e. "BA3")
        
        targetRange.Clear
        targetRange.Offset(0, 2).Clear
        
        'copy unique list of Box IDs from Flatbed scanner sheet to Dictionary. Because Box IDs repeats for muptiple Barcodes, only unique values will be copied to dictionary
        sourceRange.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=targetRange, Unique:=True
        'copy the same information (as in previous row) to the 2nd column shifted to the right on 2 cells
        sourceRange.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=targetRange.Offset(0, 2), Unique:=CopyUniqueValuesOnly
        '.Range("D1:" & .Range("D1").End(xlDown).Address).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets(cDictionayWorksheetName).Range("BC3"), Unique:=True
        
        'sort all copied values (in both columns)
        Set tRange = Worksheets(cDictionayWorksheetName).Range(targetRange.Address & ":" & targetRange.Offset(0, 2).End(xlDown).Address)
        tRange.Sort key1:=tRange, Header:=xlYes
        
        'delete header titles (Excel always copies headers) from the target range
        targetRange.Offset(0, 2).Delete xlShiftUp
        targetRange.Delete xlShiftUp 'delete the first cell (that holds title of the source column) from the target range
        
    End With
End Sub
