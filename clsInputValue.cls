VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInputValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'TODO Replace with property calls
Private FieldValue As String
Private FieldValueValidated As String
'Private FieldAddress As String
Private fieldName As String
'Private FieldDefaultValue As String
'Private FieldDefaultBool As Boolean
'Private FieldErrorMessage As String
'Private FieldRequiredBool As Boolean

'Private Const cRawDataWorksheetName = "RawData"
'Private Const cValidatedWorksheetName = "Validated"

Private oFieldSettings As clsFieldSettings
Private cellProperties As clsCellProperties

Public ValidationErrors As clsFieldValidationErrors
Public ValidationResult As clsValidationResult

Public Function InitializeValues(Target As Range) As Boolean
    
    'Initiate cell Object to get cell's address properties
    Set cellProperties = New clsCellProperties
    cellProperties.InitializeValues Target.Address
    
    'Exit the function if no field name is assigned to the column where the field is located
    If Len(Trim(cellProperties.CellFieldName)) = 0 Then
        InitializeValues = False
        Exit Function
    End If
    
    'Find name of the field being updated
    'OLD --> FieldName = Application.Worksheets("Validated").Range(Left(Target.Address, InStrRev(Target.Address, "$")) + "1").value
    'FieldName = Application.Worksheets(cRawDataWorksheetName).Range(CellProperties.CellCollumn + "1").value
    FieldValue = Target.value
    'FieldAddress = Target.Address
    
    Set ValidationErrors = New clsFieldValidationErrors
    ValidationErrors.ValidatedCellAddress = cellProperties.CellAddress
    
    'get the Field Settings for the field being updated
    Set oFieldSettings = GetFieldSettingsInstance(cellProperties, True, cellProperties.CellFieldName)
    
    InitializeValues = True
End Function

'Validate field value
Public Sub ValidateFieldValue(ByRef obValidatedResult As clsValidationResult) 'As clsValidationResult
    Dim iFieldValueValidated As String
    'object that will be returned as an output of this function
    'Dim obValidatedResult As New clsValidationResult
    
    If oFieldSettings.DataAvailable Then 'verify that Field Setting information was retrieved.
    
        obValidatedResult.ValidationStatus = Undefined '-1
                
        'validate if Default value is provided and assign it FieldValue variable
        If FieldValue = "" And oFieldSettings.FieldDefaultBool Then
            FieldValue = oFieldSettings.FieldDefaultValue
            'iFieldValueValidated = oFieldSettings.FieldDefaultValue
            obValidatedResult.ValidationStatus = DefaultAssigned
        End If
            
        If FieldValue = "" Then 'if no value was provided and no default value was set
            If oFieldSettings.FieldRequiredBool And iFieldValueValidated = "" Then
                'if field is required and default value is empty, record a validation error
                ValidationErrors.Add "No value was provided to the required field and no default value is available.", "Input value validation module", RequiredFieldEmpty
                obValidatedResult.ValidationStatus = ValidationError
                'for test only --> ValidationErrors.Add "No value was provided to the required field and no default value is available #2 - TEST.", "Input value validation module", UnexpectedValue
            Else
                'pass the value as is
                iFieldValueValidated = FieldValue
                'obValidatedResult.ValidationStatus = ValidationPassed
                obValidatedResult.ValidationStatus = IIf(obValidatedResult.ValidationStatus > 0, obValidatedResult.ValidationStatus, ValidationPassed)
            End If
        End If
        
        If FieldValue <> "" Then
            'if no validation error was detected
            'Check if the field type is dropdown. Verify the entered value if it is a dropdown
            If oFieldSettings.FieldDropDownBool Then
                'Verify provided value against the expected values; it will convert a value based on the provided map;
                'It will return error message if provided value does not match to values from the expected list
                Dim obDropdownVerify As New clsDropdownValue
                obDropdownVerify.InitializeValues oFieldSettings.FieldDropDownValueLookupRange, oFieldSettings.FieldErrorMessage, ValidationErrors
                iFieldValueValidated = obDropdownVerify.VaildateValue(FieldValue)
                Set obDropdownVerify = Nothing
                obValidatedResult.ValidationStatus = IIf(ValidationErrors.ErrorCount > 0, ValidationError, MapConversionApplied)
            ElseIf oFieldSettings.FieldDateType Then
                If IsDate(FieldValue) Then
                    iFieldValueValidated = FieldValue
                    'obValidatedResult.ValidationStatus = ValidationPassed
                    obValidatedResult.ValidationStatus = IIf(obValidatedResult.ValidationStatus > 0, obValidatedResult.ValidationStatus, ValidationPassed)
                Else
                    obValidatedResult.ValidationStatus = ValidationError
                    'todo assign error status
                    ValidationErrors.Add "Incorrect date value was provided.", "Input value validation module", IncorrectDate
                End If
            ElseIf oFieldSettings.FieldNumericOnlyType Then
                If IsNumeric(FieldValue) Then
                    iFieldValueValidated = FieldValue
                    'obValidatedResult.ValidationStatus = ValidationPassed
                    obValidatedResult.ValidationStatus = IIf(obValidatedResult.ValidationStatus > 0, obValidatedResult.ValidationStatus, ValidationPassed)
                Else
                    obValidatedResult.ValidationStatus = ValidationError
                    'todo assign error status
                    ValidationErrors.Add "Not numeric value was provided.", "Input value validation module", NotNumericValue
                End If
            Else
                'pass the value as is
                iFieldValueValidated = FieldValue
                obValidatedResult.ValidationStatus = IIf(obValidatedResult.ValidationStatus > 0, obValidatedResult.ValidationStatus, ValidationPassed)
                'If obValidatedResult.ValidationStatus <> DefaultAssigned Then
                '    obValidatedResult.ValidationStatus = ValidationPassed
                'End If
            End If
        End If
        
        'Check if this field is a trigger for calculations
        If oFieldSettings.FieldCalcTrigger Then
            ProcessCalculationTrigger oFieldSettings.fieldName, obValidatedResult, oFieldSettings.FieldCalcOverwriteExistingVal
        End If
        
    End If
    
    'assign validated value to the class' variable
    FieldValueValidated = iFieldValueValidated
    
    'Assign values for various properties of the ValidateResults object
    obValidatedResult.ValidatedValue = iFieldValueValidated
    obValidatedResult.InitialValue = FieldValue
    Set obValidatedResult.ValidatedCellProperties = cellProperties
    Set obValidatedResult.ValidationErrors = ValidationErrors
    
    'assing conditional formating properties to CellProperties object based on the list of presented errors in ValidationErrors class
    Dim iTotalErrorNum As ValidationErrorStatus
    iTotalErrorNum = ValidationErrors.TotalValidationErrorStatus

    With obValidatedResult.ValidatedCellProperties
    
        Select Case iTotalErrorNum
            Case ValidationErrorStatus.NoErrors
                'if no errors, conditional formating might be applied if some data transformation was applied
                Select Case obValidatedResult.ValidationStatus
                    Case ValidationOutcomeStatus.DefaultAssigned
                        .CellConditionalInteriorColor = BackgroundColors.Green
                        .CellConditionalFontColor = FontColors.Black
                    Case ValidationOutcomeStatus.MapConversionApplied
                        .CellConditionalInteriorColor = BackgroundColors.Green
                        .CellConditionalFontColor = FontColors.DarkGreen
                    Case ValidationOutcomeStatus.CalculatedValueApplied
                        .CellConditionalInteriorColor = BackgroundColors.Green
                        .CellConditionalFontColor = FontColors.DarkYellow
                    Case ValidationOutcomeStatus.ValidationPassed
                        .CellConditionalInteriorColor = BackgroundColors.NoColor
                        .CellConditionalFontColor = FontColors.Black
                    Case Else
                        .CellConditionalInteriorColor = BackgroundColors.Blue
                        .CellConditionalFontColor = FontColors.Black
                End Select
            Case ValidationErrorStatus.RequiredFieldEmpty
                .CellConditionalInteriorColor = BackgroundColors.Red
                .CellConditionalFontColor = FontColors.white
            Case ValidationErrorStatus.UnexpectedValue, ValidationErrorStatus.IncorrectDate
                .CellConditionalInteriorColor = BackgroundColors.Orange
                .CellConditionalFontColor = FontColors.Black
            Case ValidationErrorStatus.CombinationOfErrors
                .CellConditionalInteriorColor = BackgroundColors.Yellow
                .CellConditionalFontColor = FontColors.DarkYellow
            Case Else
                .CellConditionalInteriorColor = BackgroundColors.Blue
                .CellConditionalFontColor = FontColors.DarkRed
        End Select
        
    End With
    
    'Set ValidateFieldValue = obValidatedResult
    
    'Set obValidatedResult = Nothing
    
End Sub

Public Sub UpdateValidatedCell(Optional CleanOutValue As Boolean = False)
    Dim strValue As String
    
    'if CleanOutValue is set to True, it will clear out the current cell
    If CleanOutValue Then
        strValue = ""
        ApplyConditionalFormating True
    Else
        strValue = FieldValueValidated
        ApplyConditionalFormating False
    End If
    
    Application.Worksheets(cValidatedWorksheetName).Range(cellProperties.CellAddress).value = strValue
End Sub

Private Sub ApplyConditionalFormating(Optional CleanOutValue As Boolean = False)
    'this function will apply conditional formating for validated cells
    If CleanOutValue Then 'this will clear out any existing conditional formating on a cell
        ApplyFormatingToCell Worksheets(cRawDataWorksheetName).Range(cellProperties.CellAddress), BackgroundColors.NoColor, FontColors.Black
        ApplyFormatingToCell Worksheets(cValidatedWorksheetName).Range(cellProperties.CellAddress), BackgroundColors.NoColor, FontColors.Black
    Else 'this will apply a new conditional formating to the cell
        ApplyFormatingToCell Worksheets(cRawDataWorksheetName).Range(cellProperties.CellAddress), cellProperties.CellConditionalInteriorColor, cellProperties.CellConditionalFontColor
        ApplyFormatingToCell Worksheets(cValidatedWorksheetName).Range(cellProperties.CellAddress), cellProperties.CellConditionalInteriorColor, cellProperties.CellConditionalFontColor
    End If
End Sub

'moved to mdlGeneric
'Private Sub ApplyFormatingToCell(ByRef vRange As Range, bgrColor As BackgroundColors, fntColor As FontColors)
'    vRange.Interior.Color = bgrColor
'    vRange.Font.Color = fntColor
'    vRange.BorderAround xlContinuous, xlHairline, xlColorIndexAutomatic
''    vRange.Borders.LineStyle = xlContinuous 'xlNone
''    vRange.Borders.Weight = xlHairline
'End Sub

Private Sub ProcessCalculationTrigger(fieldName As String, ByRef obValidatedResult As clsValidationResult, CalcOverwriteExistingVal As Boolean)
    Select Case fieldName
        Case "MT_Source Aliquot Label"
            ParseSourcAliquotLabel CalcOverwriteExistingVal
        Case "MT_Vessel ID"
            If Get_MiscSettingValue(fieldName, "scanner") = "FlatbedScans" Then
                GetScannedPositions_flatbedscans CalcOverwriteExistingVal
            Else
                GetScannedPositions_handledscans CalcOverwriteExistingVal
            End If
        Case "MT_Source Specimen"
            Set_Aliquot_Prep_State_Field_ForCodedValues CalcOverwriteExistingVal
        Case "Sample Weight Reported", "Sample Volume Reported"
            Set_Initial_Current_Values CalcOverwriteExistingVal
        Case Else
            ValidationErrors.Add "Field """ & fieldName & """ marked as ""Calculation Trigger"", however no custom calculation function was allocated for it." & vbCrLf & _
                "To resolve this error clear ""Calculation Trigger"" setting of this field or contact technical support to resolve the issue.", _
                "Process Calculation Trigger module", FieldCalculationError
            obValidatedResult.ValidationStatus = ValidationError
    End Select
End Sub

Private Sub Set_Aliquot_Prep_State_Field_ForCodedValues(CalcOverwriteExistingVal As Boolean)
    Dim curValue As String, curAddress As String
    
    If IsNumeric(Trim(FieldValue)) Then
         On Error Resume Next
         
        'Update Aliquot_Prep_State fields with value from the source specimen field
        With Worksheets(cRawDataWorksheetName)
            curAddress = FindFieldAddress("MT_Aliquot_Preparation_State", cellProperties.cellRow)
            If Len(Trim(curAddress)) > 0 Then
                With .Range(curAddress)
                    curValue = .value
                    .value = IIf(Not CalcOverwriteExistingVal And Len(Trim(curValue)) > 0, curValue, Trim(FieldValue))
                End With
            End If
            
        End With
    End If
End Sub

Private Sub Set_Initial_Current_Values(CalcOverwriteExistingVal As Boolean)
    Dim curValue As String, curAddress As String
    Dim curAddrWeight As String, curAddrVolume As String
    Dim curValueWeight As String, curValueVolume As String
    Dim finalOutValue As String

    With Worksheets(cRawDataWorksheetName)
    
        'get weight & volume columm's addresses
        curAddrWeight = FindFieldAddress("Sample Weight Reported", cellProperties.cellRow)
        curAddrVolume = FindFieldAddress("Sample Volume Reported", cellProperties.cellRow)
        
        'get weight value
        If Len(Trim(curAddrWeight)) > 0 Then
            With .Range(curAddrWeight)
                curValueWeight = .value
            End With
        End If
        
        'get volume value
        If Len(Trim(curAddrVolume)) > 0 Then
            With .Range(curAddrVolume)
                curValueVolume = .value
            End With
        End If
        
        'Weight and Volume value should be mutually exclusive.
        'OK case - If only one if filled with data, pass that data to the target fields.
        'Error case - If these 2 fields are filled with some values or both are blank, report an error - fill with "Can't Resolve" and highlight as Unexpected value.
        
        finalOutValue = "Can't Resolve" 'default value
        finalOutValue = IIf((Len(Trim(curValueWeight)) > 0 And Len(Trim(curValueVolume)) = 0), curValueWeight, finalOutValue) 'assign Weight value
        finalOutValue = IIf((Len(Trim(curValueWeight)) = 0 And Len(Trim(curValueVolume)) > 0), curValueVolume, finalOutValue) 'assign Volume value
           
        'update "Initial Amount"
        curAddress = FindFieldAddress("Initial Amount", cellProperties.cellRow)
        If Len(Trim(curAddress)) > 0 Then
            With .Range(curAddress)
                curValue = .value
                .value = IIf(Not CalcOverwriteExistingVal And Len(Trim(curValue)) > 0, curValue, Trim(finalOutValue))
            End With
        End If
        
        'update "current value"
        curAddress = FindFieldAddress("Current Amount", cellProperties.cellRow)
        If Len(Trim(curAddress)) > 0 Then
            With .Range(curAddress)
                curValue = .value
                .value = IIf(Not CalcOverwriteExistingVal And Len(Trim(curValue)) > 0, curValue, Trim(finalOutValue))
            End With
        End If
        
        
    End With
End Sub

Private Sub ParseSourcAliquotLabel(CalcOverwriteExistingVal As Boolean)
'TODO - Parsing and field mapping is hardcoded in this function. Eventually this should be replaced with a generic approach for calculated fields.

'Parsing logic is below:
'- Biospecimen ID (BID) (sSubjectId)– First 5 digits – All collections for an
'acute test will have the same BID for a particpant
'- Timepoint (sTimepoint)– Two digits following BID – Current model includes
'timepoints 1-11 for blood, 1-4 for muscle, and 1-2 for adipose
'Collection
'- Sample Type (sSourceSpecimen) – two digits following Timepoint – Each sample
'type (EDTA, serum, adipose, muscle, etc) is assigned a sample
'type code (00-99) that is included in the ID for every collection
'tube and aliquot
'- Unique Cryovial – Two digits following sample type – Each
'aliquot created of a given sample type has available aliquots 01-
'99

    'verify that the length of the field's value is correct and = 11
    If Len(FieldValue) = 11 Then
        
        Dim sSampleId As String
        Dim sSubjectId As String
        Dim sTimepoint As String
        Dim sSourceSpecimen As String
        Dim curValue As String, curAddress As String
        
        'Parse field's values
        sSampleId = Left(FieldValue, Len(FieldValue) - 2) 'whole value without 2 last characters
        sSubjectId = Left(FieldValue, 5)
        sSourceSpecimen = Mid(FieldValue, 8, 2)
        sTimepoint = Mid(FieldValue, 6, 2)
        
        On Error Resume Next
        'Update associated fields
        With Worksheets(cRawDataWorksheetName)
            curAddress = FindFieldAddress("MT_Sample ID", cellProperties.cellRow)
            If Len(Trim(curAddress)) > 0 Then
                With .Range(curAddress)
                    curValue = .value
                    .value = IIf(Not CalcOverwriteExistingVal And Len(Trim(curValue)) > 0, curValue, sSampleId)
                End With
            End If
            
            curAddress = FindFieldAddress("MT_SubjectID", cellProperties.cellRow)
            If Len(Trim(curAddress)) > 0 Then
                With .Range(curAddress)
                    curValue = .value
                    .value = IIf(Not CalcOverwriteExistingVal And Len(Trim(curValue)) > 0, curValue, sSubjectId)
                End With
            End If
            
            curAddress = FindFieldAddress("MT_Source Specimen", cellProperties.cellRow)
            If Len(Trim(curAddress)) > 0 Then
                With .Range(curAddress)
                    curValue = .value
                    .value = IIf(Not CalcOverwriteExistingVal And Len(Trim(curValue)) > 0, curValue, sSourceSpecimen)
                End With
            End If
            
            curAddress = FindFieldAddress("MT_Study Timepoint", cellProperties.cellRow)
            If Len(Trim(curAddress)) > 0 Then
                With .Range(curAddress)
                    curValue = .value
                    .value = IIf(Not CalcOverwriteExistingVal And Len(Trim(curValue)) > 0, curValue, sTimepoint)
                End With
            End If
            
        End With
        On Error GoTo 0
    Else
    'calculation procedure cannot continue
    End If
    
End Sub

Private Sub GetScannedPositions_flatbedscans(CalcOverwriteExistingVal As Boolean)
    Dim rCell As Range
    Dim scanStatus As String
    Dim bPos As clsBoxPosition
    Dim curValue As String, curAddress As String
    Dim scanBoxID As String
        
    With Worksheets(cFlatbedScansWorksheetName)
        If Len(FieldValue) > 0 Then
            'TODO: avoid using hardcoded addresses as one below
            Set rCell = .Range("$B:$B").Find(FieldValue, , LookIn:=xlValues) 'search for a particular Vessel ID value
            
            If Not rCell Is Nothing Then 'if the Vessel ID was found, proceed here
                scanStatus = rCell.Offset(0, 1).value
                
                If scanStatus = "OK" Then
                    'parse the location value into coordinates: row/column
                    Set bPos = New clsBoxPosition
                    bPos.SplitScannedLocationToPositions (rCell.Offset(0, -1).value)
                    'get Box ID value reported by flatbedscanner
                    scanBoxID = rCell.Offset(0, 2).value
                    
                    'On Error Resume Next
                    'Update associated fields
                    With Worksheets(cRawDataWorksheetName)
                        'update Row(p4) field value on the Rawdata sheet
                        curAddress = FindFieldAddress("Row_p4", cellProperties.cellRow)
                        If Len(Trim(curAddress)) > 0 Then
                            With .Range(curAddress)
                                curValue = .value
                                .value = IIf(Not CalcOverwriteExistingVal And Len(Trim(curValue)) > 0, curValue, bPos.Row_p4)
                            End With
                        End If
                    
'                        .Range(FindFieldAddress("Row_p4", cellProperties.cellRow)).Value = _
'                            IIf(Not CalcOverwriteExistingVal And Len(Trim(Range(FindFieldAddress("Row_p4", cellProperties.cellRow)).Value)) > 0, _
'                                Range(FindFieldAddress("Row_p4", cellProperties.cellRow)).Value, _
'                                bPos.Row_p4)
                                
                        'update Column(p5) field value on the Rawdata sheet
                        curAddress = FindFieldAddress("Column_p5", cellProperties.cellRow)
                        If Len(Trim(curAddress)) > 0 Then
                            With .Range(curAddress)
                                curValue = .value
                                .value = IIf(Not CalcOverwriteExistingVal And Len(Trim(curValue)) > 0, curValue, bPos.Column_p5)
                            End With
                        End If
'                        .Range(FindFieldAddress("Column_p5", cellProperties.cellRow)).Value = _
'                            IIf(Not CalcOverwriteExistingVal And Len(Trim(Range(FindFieldAddress("Column_p5", cellProperties.cellRow)).Value)) > 0, _
'                                Range(FindFieldAddress("Column_p5", cellProperties.cellRow)).Value, _
'                                bPos.Column_p5)
                         'Old code
'                        .Range(FindFieldAddress("Row_p4", cellProperties.cellRow)).Value = bPos.Row_p4
'                        .Range(FindFieldAddress("Column_p5", cellProperties.cellRow)).Value = bPos.Column_p5
                        
                        'update MT_Box Label field value on the Rawdata sheet
                        curAddress = FindFieldAddress("MT_Box Label", cellProperties.cellRow)
                        If Len(Trim(curAddress)) > 0 Then
                            With .Range(curAddress)
                                curValue = .value
                                .value = IIf(Not CalcOverwriteExistingVal And Len(Trim(curValue)) > 0, curValue, scanBoxID)
                            End With
                        End If

                        
                    End With
                    'On Error GoTo 0
                    
                    Set bPos = Nothing
                End If
            End If
        End If
    End With
End Sub

Private Sub GetScannedPositions_handledscans(CalcOverwriteExistingVal As Boolean)
    Dim rCell As Range
    Dim curValue As String, curAddress As String
    Dim sRow As String, sCol As String
        
    With Worksheets(cHandledScansWorksheetName)
        If Len(FieldValue) > 0 Then
            'TODO: avoid using hardcoded addresses as one below
            Set rCell = .Range("$A:$A").Find(FieldValue, , LookIn:=xlValues) 'search for a particular Vessel ID value
            
            If Not rCell Is Nothing Then 'if the Vessel ID was found, proceed here
                sRow = rCell.Offset(0, 1).value
                sCol = rCell.Offset(0, 2).value
                
                'On Error Resume Next
                'Update associated fields
                With Worksheets(cRawDataWorksheetName)
                    'update Row(p4) field value on the Rawdata sheet
                    curAddress = FindFieldAddress("Row_p4", cellProperties.cellRow)
                    If Len(Trim(curAddress)) > 0 Then
                        With .Range(curAddress)
                            curValue = .value
                            .value = IIf(Not CalcOverwriteExistingVal And Len(Trim(curValue)) > 0, curValue, sRow)
                        End With
                    End If
                
                    'update Column(p5) field value on the Rawdata sheet
                    curAddress = FindFieldAddress("Column_p5", cellProperties.cellRow)
                    If Len(Trim(curAddress)) > 0 Then
                        With .Range(curAddress)
                            curValue = .value
                            .value = IIf(Not CalcOverwriteExistingVal And Len(Trim(curValue)) > 0, curValue, sCol)
                        End With
                    End If
                    
                End With
                'On Error GoTo 0
                
            End If
        End If
    End With
End Sub

'==========This function was moved to mdlGeneric module - START ======================
''This function will return the location of the value of the requested field. It will use the cellProperties object to identify the correct row to be used
'Private Function FindFieldAddress(fieldName As String, cellRow As String, Optional sWorksheetName As String = "RawData") As String
'    'fieldName example: {MT_Project}
'
'    fieldName = Replace(Replace(fieldName, "{", ""), "}", "") 'remove figure brackets from the name
'
'    'This will find address of the cell to be evaluated
'    With Worksheets(sWorksheetName)
'        Dim fnr As Range, fcol As String
'
'        'Find the supplied column name on the Validated sheet
'        '   Used Range method and Cells object will retrun the cell of the actually used last column on the sheet
'        Set fnr = .Range("A1:" & Cells(1, .UsedRange.Columns.Count).Address).Find(fieldName, LookIn:=xlValues)
'        If Not fnr Is Nothing Then
'            'find the column associated with the given field name
'            'fcol = Replace(Left(fnr.Address, InStrRev(fnr.Address, "$")), "$", "")
'            'identify an address of the cell that should be used for evaluation
'            'FindFieldAddress = fcol & cellRow 'cValidatedWorksheetName & "!" &
'
'            FindFieldAddress = Cells(cellRow, fnr.Column).Address
'        Else
'            FindFieldAddress = ""
'        End If
'    End With
'
'End Function
'==========This function was moved to mdlGeneric module - END ======================

'Public Sub CleanObjectReferences()
'    'clean global object references
'    Set oFieldSettings = Nothing
'    Set cellProperties = Nothing
'    Set ValidationErrors = Nothing
'
'End Sub

Private Sub Class_Terminate()
    'clean global object references
    Set oFieldSettings = Nothing
    Set cellProperties = Nothing
    Set ValidationErrors = Nothing
End Sub
