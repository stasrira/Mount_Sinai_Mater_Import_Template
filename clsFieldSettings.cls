VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFieldSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public fieldName As String
Public FieldDefaultValue As String
Public FieldDefaultBool As Boolean
Public FieldErrorMessage As String
Public FieldRequiredBool As Boolean
Public FieldDropDownValueLookupRange As String
Public FieldDropDownBool As Boolean
Public DataAvailable As Boolean
Public FieldCalculated As Boolean
Public FieldCalcTrigger As Boolean
Public FieldCalcOverwriteExistingVal As Boolean
Public FieldDateType As Boolean
Public FieldExcludeFromExport As Boolean 'if set to true, the column corresponding to this field won't be included into the export file
'Public FieldColumnAddressOnRawDataSheet As String 'will hold the address of the column on the RawData sheet corresponding to the current field
Public FieldNumericOnlyType As Boolean
Public FieldMiscSettings As String

Private FieldSettingsFirstCell As String 'sets to "A2" during initialization of the class
Private fieldRowNum As String 'Will hold address of the row containing the field on the FieldSettings sheet. This variable will be used from different Subs/Functions

'Private Const cFieldsRange = "A2:A100" ' part of old code

'TODO this assingment should idealy be going from a table structure on the Dictionary sheet
Const cAddrDropdownErrorMessage = "$B"
Const cAddrDefault = "$C"
Const cAddrRequired = "$D"
Const cAddrDropDown = "$E"
Const cAddrDropDownValueLookupRange = "$F"
Const cAddrCalcTrigger = "$G"
Const cAddrCalcOverwriteExistingVal = "$H"
Const cAddrCalculated = "$I"
Const cAddrDateField = "$J"
Const cAddrExcludeFromExport = "$K"
Const cAddrNumericOnly = "$L"
Const cAddrMiscSettings = "$M"


Public Sub InitializeValues(cellProperties As clsCellProperties)  '(vFieldName As String)

'for test only
'Debug.Print cellProperties.CellFieldName

    'Assign Field Value
    fieldName = cellProperties.CellFieldName 'vFieldName
    
    
    With Worksheets(cSettingsWorksheetName)
        Dim fnr As Range
        Dim iRows As Integer
        'Dim fn_row As String 'Will hold address of the row containing the field
        
        iRows = .UsedRange.Rows.Count 'number of actually used rows
        
        'identify range of actually used cells on the given spreadsheet and apply Find function
        Set fnr = .Range(FieldSettingsFirstCell & ":" & Cells(iRows, 1).Address).Find(fieldName, LookIn:=xlValues)
        'Old code --> Set fnr = .Range(cFieldsRange).Find(fieldName, LookIn:=xlValues)
        If Not fnr Is Nothing Then
            DataAvailable = True 'update a flag to reflect that object can be safely used
            
            'fieldRowNum = Right(fnr.Address, Len(fnr.Address) - InStrRev(fnr.Address, "$") + 1) 'it will contain value presenting row of the record, i.e. $10
            fieldRowNum = "$" & fnr.Row 'it will contain value presenting row of the record, i.e. $10
            
            'Populate values for all variables
            
            'read Default as boolean field value
            'FieldRequiredBool = .Range(cAddrRequired + fieldRowNum).value
            FieldDefaultBool = Len(.Range(cAddrDefault + fieldRowNum).value) > 0
            
            'read Default field value
            'Default value can be hardcoded for regular fields or retrieved through hardcoded VLookup function for dropdowns on FieldSettings sheet (the current version)
            FieldDefaultValue = .Range(cAddrDefault + fieldRowNum).value
'TODO - retrieve default for dropdown fields from the range programmatically - it will be a more generic way of getting it
            
            'read DropDown field value
            FieldDropDownBool = .Range(cAddrDropDown + fieldRowNum).value
            'read error message value for no match of the dropdown field
            FieldErrorMessage = .Range(cAddrDropdownErrorMessage + fieldRowNum).value
            'read Dropdown lookup range (on the Dictionary sheet)
            FieldDropDownValueLookupRange = .Range(cAddrDropDownValueLookupRange + fieldRowNum).value
            'read Date field value
            FieldDateType = .Range(cAddrDateField + fieldRowNum).value
            'read exclude from export value
            FieldExcludeFromExport = .Range(cAddrExcludeFromExport + fieldRowNum).value
            'read Numeric Only value
            FieldNumericOnlyType = .Range(cAddrNumericOnly + fieldRowNum).value
            'read Misc_Settings value
            FieldMiscSettings = .Range(cAddrMiscSettings + fieldRowNum).value
            
            UpdateVolatileSettings cellProperties
            
            'read Required field value
            'Call function that will evaluate all volatile values - values that depends on values of other fileds
            'FieldRequiredBool = ValidateBoolValue(.Range(cAddrRequired + fieldRowNum).value, cellProperties)
            'UpdateVolatileSettings
            
            'read Calculated field value
            'FieldCalculated = .Range(cAddrCalculated + fieldRowNum).value
            'FieldCalculated = ValidateBoolValue(.Range(cAddrCalculated + fieldRowNum).value, cellProperties)
            
            'read Calculation Trigger field value
            'FieldCalcTrigger = .Range(cAddrCalcTrigger + fieldRowNum).value
            'FieldCalcTrigger = ValidateBoolValue(.Range(cAddrCalcTrigger + fieldRowNum).value, cellProperties)
            
        Else
            'TODO - Generate error message that the field name was not found
            DataAvailable = False
        End If
    End With
End Sub

'This function will calculate volatile values for properties that depends on values of other fields
'This function is public since the FieldSettings objects stored in the dictFieldSettings dictionary will be required to re-evaluate volatile properties before each use.
Public Sub UpdateVolatileSettings(cellProperties As clsCellProperties) ', Optional ColumnToValidate As String = cAddrRequired
    If DataAvailable Then 'run the below code only if FieldSettings instance has data available
        With Worksheets(cSettingsWorksheetName)
            'read Required field value
            FieldRequiredBool = ValidateBoolValue(.Range(cAddrRequired + fieldRowNum).value, cellProperties)
            'read Calculated field value
            FieldCalculated = ValidateBoolValue(.Range(cAddrCalculated + fieldRowNum).value, cellProperties)
            'read Calculation Trigger field value
            FieldCalcTrigger = ValidateBoolValue(.Range(cAddrCalcTrigger + fieldRowNum).value, cellProperties)
            'read Calculation Overwrites Existing Values field value
            FieldCalcOverwriteExistingVal = ValidateBoolValue(.Range(cAddrCalcOverwriteExistingVal + fieldRowNum).value, cellProperties)
        End With
    End If
End Sub

Private Function ValidateBoolValue(vFormulaToValidate As String, cellProperties As clsCellProperties)
'IF ({MT_Project} = "MoTrPAC", 1, 0)

    If Len(vFormulaToValidate) > 0 Then
        'if passed value is not empty, try to validate it
        
        Dim regex As Object, str As String
        Dim matches As Object, match As Object
        
        'intiate regex object and pass there the search pattern
        Set regex = CreateObject("VBScript.RegExp")
        With regex
          .Pattern = "\{\w*\}" 'search for patterns of filed names inside fo the curved brackets, i.e. {MT_Project}
          .Global = True
        End With
             
        str = vFormulaToValidate
        'for test only:
        'str = "IF ({MT_Project} = ""MoTrPAC"" and {abc}, 1, 0)"
         
        If regex.Test(str) Then
            'if pattern was found, the found field names are needed to be filled with the actual field values from Validated sheet
            Set matches = regex.Execute(str)
                 
            For Each match In matches
              'Debug.Print match.value '{MT_Project}
              str = Replace(str, match.value, FindFieldValue(match.value, cellProperties))
              'str = Replace(str, match.value, FindFieldAddress(match.value, cellProperties.cellRow))
            Next match
        End If
        
        str = Replace(str, " ", "")
        
        'Evaluate value of the formula
        With Worksheets(cValidatedWorksheetName)
            If Not IsError(.Evaluate(str)) Then
                ValidateBoolValue = .Evaluate(str)
            Else
                ValidateBoolValue = False
            End If
            
        End With
    Else
        ValidateBoolValue = False
    End If
    
End Function

'This function will return the actual value of the requested field. It will use the cellProperties object to identify the correct row to be used
Private Function FindFieldValue(fieldName As String, cellProperties As clsCellProperties) As String
    'fieldName example: {MT_Project}
    'The following is using HLOOKUP function to find needed value - i.e. HLOOKUP("MT_Project", A1:Z100,4, FALSE)
    On Error Resume Next 'this is required to avoid run time errors, if HLookup cannot find a needed value
    
    fieldName = Replace(Replace(fieldName, "{", ""), "}", "") 'remove figure brackets from the name
    
    With Worksheets(cValidatedWorksheetName)
        FindFieldValue = """" & Application.WorksheetFunction.HLookup(fieldName, .Range("A1:Z" & cellProperties.cellRow), cellProperties.cellRow, False) & """"
    End With

End Function


Private Sub Class_Initialize()
    FieldSettingsFirstCell = cFieldSettings_FirstFieldCell '"A2"
End Sub
