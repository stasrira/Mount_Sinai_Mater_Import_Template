Attribute VB_Name = "mdlGeneric"
Option Explicit

Public Const cHelpTitle = "Sample Entry Validation Tool"
Public Const cHelpVersion = "1.017"
Public Const cHelpDescription = "Questions and technical support: email to stasrirak.ms@gmail.com"

Public Const cRawDataWorksheetName = "RawData"
Public Const cValidatedWorksheetName = "Validated"
Public Const cSettingsWorksheetName = "FieldSettings"
Public Const cDictionayWorksheetName = "Dictionary"
Public Const cFlatbedScansWorksheetName = "FlatbedScans"
Public Const cHandledScansWorksheetName = "HandledScans"
Public Const cConfigWorksheetName = "Configuration"


Public Const cCustomMenuName = "&MSSM MENU"
Public Const cCustomMenu_SubMenuSettings = "Settings"
Public Const cCustomMenu_SetDropdowonFunc = "Set Dropdown Functionality "
Public Const cCustomMenu_SetDropdowonFunc_ShortCut = "    CTRL+SHIFT+D"
Public Const cCustomMenu_SetValidationFunc = "Set Automatic Validation "
Public Const cCustomMenu_SetValidationFunc_ShortCut = "          CTRL+SHIFT+V"

Public Const cFieldSettings_FirstFieldCell = "A2"
Public Const cRawData_FirstColumnCell = "A1"
Public Const cValidated_FirstColumnCell = "A1"
Public Const cConfig_FirstFieldCell = "A2"

Public dictValidationResults As New Dictionary
Public dictFieldSettings As New Dictionary

Public bVoidAutomatedValidation As Boolean
Public bVoidDropDownFunctionality As Boolean
Public bFieldHeadersWereSynced As Boolean
Public bSetCtrlVPasteAsValues As Boolean

Public Enum ValidationErrorStatus
    NoErrors = 1
    RequiredFieldEmpty = 2
    UnexpectedValue = 3
    CombinationOfErrors = 4
    IncorrectDate = 5
    FieldCalculationError = 6
    NotNumericValue = 7
End Enum

Public Enum ValidationOutcomeStatus
    ValidationPassed = 0
    DefaultAssigned = 1
    MapConversionApplied = 2
    ValidationError = 3
    CalculatedValueApplied = 4
End Enum

Public Enum BackgroundColors
    white = 16777215 'RGB(255, 255, 255)
    Green = 13561798 'RGB(198, 239, 206)
    Orange = 8696052 'RGB(244, 176, 132)
    LightRed = 13551615 'RGB(255, 199, 206)
    Red = 255 'RGB(255, 0, 0)
    Blue = 15189684 'RGB(180, 198, 231)
    Yellow = 10284031 'RGB(255, 235, 156)
    NoColor = -4142 'xlNone 'No Color (default Excel filling)
End Enum

Public Enum FontColors
    white = 16777215 'RGB(255, 255, 255)
    DarkGreen = 24832 'RGB(0, 97, 0)
    DarkRed = 393372 'RGB(156, 0, 6)
    DarkYellow = 22428 'RGB(156, 87, 0)
    Black = 0 'RGB(0, 0, 0)
End Enum

Private Type ValidationReportMsg
    ValidationMessage As String
    MsgBoxStyle As VbMsgBoxStyle
End Type

Private CalcState As Long
Private EventState As Boolean
Private PageBreakState As Boolean

Sub OptimizeCode_Begin()

    Application.ScreenUpdating = False
    
'    EventState = Application.EnableEvents
'    Application.EnableEvents = False
    
    CalcState = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    PageBreakState = ActiveSheet.DisplayPageBreaks
    ActiveSheet.DisplayPageBreaks = False

End Sub

Sub OptimizeCode_End()

    ActiveSheet.DisplayPageBreaks = PageBreakState
    Application.Calculation = CalcState
    'Application.EnableEvents = EventState
    Application.ScreenUpdating = True

End Sub

Public Sub Validate_Cell_Value(ByVal Target As Range)
'    Application.Worksheets("Validated").Range(Target.Address).value = Target.value
'    Dim FieldName As String
'    FieldName = Application.Worksheets("Validated").Range(Left(Target.Address, InStrRev(Target.Address, "$")) + "1").value

    'Avoid validating the 1st row (the column headers row)
    If Target.Row = 1 Then Exit Sub
    
    'instanciate a class to validate value of a given cell
    Dim obInput As New clsInputValue
    Dim obValidationResult As clsValidationResult
    
    If obInput.InitializeValues(Target) Then
        'if InitializeValues returned True, proceed with validation
        
        'Debug.Print "Perform Validation - " & Target.Address
        
        'clean target cell on the Validate sheet before hand
        obInput.UpdateValidatedCell True
        
        'Debug.Print Target.Address
        
        'check if Validation Restults dictionary already contains an entry for the current cell
        If Not dictValidationResults.Exists(Target.Address) Then 'obValidationResult.ValidatedCellProperties.CellAddress
            'create a new instance of the obValidationResult object and pass it to the validation process
            Set obValidationResult = New clsValidationResult
            
            obInput.ValidateFieldValue obValidationResult
            
            'add a new entry Validation Restults to dictionary
            dictValidationResults.Add obValidationResult.ValidatedCellProperties.CellAddress, obValidationResult
        Else
            Set obValidationResult = dictValidationResults(Target.Address)
            obInput.ValidateFieldValue obValidationResult
            'update the value for an existing Validation Restults dictionary
            'Set dictValidationResults(obValidationResult.ValidatedCellProperties.CellAddress) = obValidationResult
        End If
        
        
        'OLD code
    '    'If obInput.ValidationErrors.ErrorCount > 0 Then
    '    If obValidationResult.ValidationErrors.ErrorCount > 0 Then
    '        MsgBox obValidationResult.ValidationErrors.toString, vbOKOnly, "Validation Error"
    '    End If
        
        'update target cell with the validated value
        obInput.UpdateValidatedCell
        
    End If
    
    'clean up internal reference and object itself
    'obInput.CleanObjectReferences
    Set obInput = Nothing

End Sub

Public Sub ValidateWholeWorksheet(Optional startCell As String = "A1", Optional AvoidWarningMessage As Boolean = False)
    OptimizeCode_Begin 'turn off visualization features of Excel while running this
    
        
    Dim iResponse As Integer
    Dim tStart As Date, tEnd As Date
    Dim mSeconds As Long, mMinutes As Integer, mHours As Integer
    Dim strTime As String
    
    If Not AvoidWarningMessage Then
        iResponse = MsgBox("The system is about to start validation of all values presented on the ""RawData"" spreadsheet." & _
                        "This also will update ""Validated"" sheet with validated values, thus data presented there will be modified." & _
                        vbCrLf & vbCrLf & "Do you want to proceed? If not, click 'Cancel'." & vbCrLf & vbCrLf _
                        & "Note: this process might take upto a minute, depending on the number of data entered on the spreadsheet.", _
                        vbOKCancel, "MSSM Data Entry File")
    Else
        iResponse = vbOK
    End If
    
    If iResponse = vbOK Then
    tStart = Now()
'Debug.Print tStart 'for test purposes only

        Dim iCols As Integer, iRows As Integer
        Dim rRng As Range, rCell As Range
        
        'clear all previous validation results
        Set dictValidationResults = Nothing
        
        With Worksheets(cRawDataWorksheetName)
            iCols = .UsedRange.Columns.Count 'number of actually used columns
            iRows = .UsedRange.Rows.Count 'number of actually used rows
            
            'identify range of actually used cells on the given spreadsheet
            Set rRng = .Range(startCell & ":" & Cells(iRows, iCols).Address)
            
            RemoveFormattingAndContents cRawDataWorksheetName
            RemoveFormattingAndContents cValidatedWorksheetName, , True
            
            For Each rCell In rRng.Cells
                'Debug.Print rCell.Address, rCell.Value
                
                'commented OLD code
                'check if ValidationResults dictionary has a key corresponding to the current cell; create one if it is absent
'                If Not dictValidationResults.Exists(rCell.Address) Then
'                    dictValidationResults.Add rCell.Address, Nothing 'set Nothing as a default value
'                End If

                'run validation for the given cell
    '            .Worksheet_Change_External rCell
                Validate_Cell_Value rCell
            Next rCell
        End With
        
        tEnd = Now()
'Debug.Print tEnd 'for test purposes only

        mSeconds = DateDiff("s", tStart, tEnd)
        mHours = mSeconds \ 3600
        mMinutes = (mSeconds - (mHours * 3600)) \ 60
        mSeconds = mSeconds - ((mHours * 3600) + (mMinutes * 60))
        
        If mHours > 0 Then strTime = mHours & " hours "
        If mMinutes > 0 Then strTime = strTime & mMinutes & " minutes "
        strTime = strTime & mSeconds & " seconds "

        Dim ValidErrStats As ValidationReportMsg
        ValidErrStats = AllValidatedCellsErrorReport()
        
        
        OptimizeCode_End 'turn back on visualization features of Excel

        'MsgBox "Validation process is completed.", vbInformation, "Validation of ""RawData"" sheet"
        MsgBox "Validation process is completed." & vbCrLf & "Time elapsed: " & strTime & vbCrLf & vbCrLf & "Validation Summary:" & vbCrLf & ValidErrStats.ValidationMessage, _
            ValidErrStats.MsgBoxStyle, "Validation of ""RawData"" sheet"
        
    Else
        OptimizeCode_End 'turn back on visualization features of Excel
    End If
    
End Sub

Private Sub RemoveFormattingAndContents(Optional sWorksheetName As String = "RawData", Optional startRow As String = "$2", Optional ClearContents As Boolean = False)
    
    Dim iCols As Integer, iRows As Integer
    Dim rRng As Range, rCell As Range, rExtraCol As Range
    Dim curVoidAutomatedValidation As Boolean
    
    'disable application events while modifying conditional formatting
    Application.EnableEvents = False
    On Error Resume Next
    
'    curVoidAutomatedValidation = bVoidAutomatedValidation 'save current status of bVoidAutomatedValidation into a temp variable
'    bVoidAutomatedValidation = True 'formating removal will trigger automated validation, to prevent that this flag is set temporarily true
            
    With Worksheets(sWorksheetName)
        
        .Range("A1").Select 'select the first cell to remove focus from any other cell that was previously used.
        
        'identify range of all cells except the header row (the first row)
        Set rRng = .Range(startRow & ":$" & .Cells.Rows.Count)
        
        rRng.Interior.Color = BackgroundColors.NoColor
        rRng.Font.Color = FontColors.Black
        
        If ClearContents Then
            '.Cells.ClearContents 'delete content out from the whole page (all cells)
            'rRng.ClearContents 'delete data only from the given range
            rRng.EntireRow.Delete 'entire rows starting from the given cell
            
            'delete any extra columns following the last column having title
            For Each rCell In .Range("$1:$1")
                'Debug.Print rCell.Address, rCell.Value
                
                If Len(Trim(rCell.value)) = 0 Then
                    .Range("$" & rCell.Column & ":$" & .Cells.Columns.Count).Delete
                    Exit For
                End If
            Next
        End If
        
        
    End With
    
    On Error GoTo 0
    'enable back application events
    Application.EnableEvents = True
'    bVoidAutomatedValidation = curVoidAutomatedValidation 'set bVoidAutomatedValidation back to the original value
    
End Sub

Public Sub ClearFormatingOfWorkbook_MenuCall()
    RemoveFormattingAndContents cRawDataWorksheetName, "$1"
    RemoveFormattingAndContents cValidatedWorksheetName, "$1"
End Sub

Public Sub ExportValidateSheet()
    Const numAttempts = 3
    
    Dim xPath As String, xWs As Worksheet
    Dim fileName As String, strResp As String
    Dim dirExists As Boolean, OFSO As FileSystemObject, fileFormat As Integer
    Dim fDialog As FileDialog, result As Integer, iTempCount As Integer
    Dim bManifests As Boolean, iResponse As Integer
    
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For Each xWs In ThisWorkbook.Sheets
        If xWs.Name = cValidatedWorksheetName Then 'identify Validated worksheet
        
            xPath = Application.ActiveWorkbook.Path
            fileName = xPath & "\" & xWs.Name & "_" & Format(Now(), "mmddyyyy_HHMMSS") & ".csv" 'xPath & "\" &
            
            Set fDialog = Application.FileDialog(msoFileDialogSaveAs)
             
            'Optional: FileDialog properties
            fDialog.Title = "Export Validated Sheet To File"
            fDialog.ButtonName = "Export As"
            fDialog.InitialFileName = fileName
            
            'MsgBox "The Export process supports only exporting data in comma separated value (.csv) or tab delimited text (.txt) formats." & vbCrLf & vbCrLf & "Please select one of these formats when choosing an export file format on the next step.", vbInformation, "Export File"
            MsgBox "The Export process supports only exporting data in comma separated value (.csv) format. Any other formats will be ignored by the system.", vbInformation, "Export File"

ShowDialog:
            result = fDialog.Show
             
            If result = -1 Then
'                'allow only .csv and tab delimited (.txt) files to go through
'                Select Case fDialog.FilterIndex
'                    Case 5
'                        fileFormat = 6 'csv file
'                    Case 12
'                        fileFormat = -4158 'tab delimited (.txt) file
'                    Case Else
'                        iTempCount = iTempCount + 1
'                        If iTempCount >= numAttempts Then
'                            MsgBox "Too many attempts. Exiting Exporting Process.", vbCritical
'                            Exit Sub
'                        End If
'
'                        MsgBox "Please select a comma separated value (.csv) or a tab delimited text (.txt) file formats and try again." _
'                            & vbCrLf & vbCrLf & "Attempt " & CStr(iTempCount) & " out of " & CStr(numAttempts), vbCritical, "Export File Format Error"
'                        GoTo ShowDialog:
'                End Select
            
                'Debug.Print fDialog.SelectedItems(1)
                'fileName = fDialog.SelectedItems(1)
                'force CSV format.
                fileName = Left(fDialog.SelectedItems(1), InStrRev(fDialog.SelectedItems(1), ".")) & "csv" ' this will replace any selected extension with "csv"
                fileFormat = 6 'force csv file format
                                    
                If fileName <> "" Then
                    'check if the folder exists
                    
                    Set OFSO = CreateObject("Scripting.FileSystemObject")
                    dirExists = OFSO.FolderExists(Left(fileName, InStrRev(fileName, "\") - 1))
                    Set OFSO = Nothing
                    
                    If dirExists Then
                        'copy Validated sheet data to memory
                        xWs.Cells.Copy
                        
                        'create a temp sheet to hold export data. This sheet won't have any code behind. This is needed to prevent copying VBA code to the exported file
                        Dim tempSheetName As String
                        tempSheetName = "ExportValidated_" & Format(Now(), "yyyymmdd_HHMMSS") 'temp sheet name
                        ThisWorkbook.Sheets.Add.Name = tempSheetName 'add the new temp sheet
                        ThisWorkbook.Sheets(tempSheetName).Cells.PasteSpecial Paste:=xlPasteAll 'copy all data from memory to the created sheet
                        
                        'delete columns that should be excluded from the export
                        RemoveColumnsExcludedFromExport ThisWorkbook.Sheets(tempSheetName), xWs
                        
                        ThisWorkbook.Sheets(tempSheetName).Copy 'copy the sheet as a Workbook. This will be used by SaveAs method
                        'xWs.Copy
                        
                        'delete any code behind from the copied worksheet
'                        With Application.ActiveWorkbook.VBProject.VBComponents(Application.ActiveSheet.CodeName).CodeModule
'                            .DeleteLines 1, .CountOfLines
'                        End With
                            
                        'for testing only
                        'Application.ActiveWorkbook.SaveAs xPath & "\" & xWs.name & "_" & Format(Now(), "mmddyyyy_HHMMSS") & ".csv", 6 ', , , , , , 1 ' xlUserResolution = 1 '_HHMMSS
                        'Application.ActiveWorkbook.SaveAs fileName, 6 'csv
                        'Application.ActiveWorkbook.SaveAs fileName, -4158 'tab
                        
                        'export the new workbook to a file and close it
                        Application.ActiveWorkbook.SaveAs fileName, fileFormat
                        Application.ActiveWorkbook.Close False
                        
                        ThisWorkbook.Sheets(tempSheetName).Delete 'delete temporary sheet
                        
                        'report successful export
                        MsgBox "Export of Validated sheet was successfully completed and the following file was created" & vbCrLf & fileName, _
                                vbOKOnly + vbInformation, "Export Validated Sheet"
                        
                        'prompt users to submit Manifest IDs
                        bManifests = CBool(GetConfigValue("Manifest_Prompt_OnExport"))
                        If bManifests Then
                            iResponse = MsgBox("Do you also want to submit Manifest IDs exported in this file to the database?" & _
                                        vbCrLf & vbCrLf & "If you want to proceed, click 'OK'. If not, click 'Cancel'.", _
                                        vbOKCancel + vbInformation, "Submitting Manfiest IDs")
                                        
                            If iResponse = vbOK Then
                                SubmitManifests True
                            End If
                        End If
                        
                    Else
                        'report bad path provided
                        MsgBox "Error writing to the provided path. It might not exist or not accessable. Verify the pass and try again", vbCritical, "Error saving export file"
                    End If
                End If
            
            End If
            
            Exit For
        End If
    Next

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Public Function GetFieldSettingsInstance(cellProperties As clsCellProperties, Optional updateVolatileSetting As Boolean = True, Optional fieldName As String = "")
    'TODO: this function should be called from all places where dictFieldSettings is validated for presense of an object for a given field name
    Dim oFieldSettings As clsFieldSettings
    
    If cellProperties Is Nothing And Len(Trim(fieldName)) > 0 Then
        'if cellProperties were not provided, but fieldName was given, the function will populate cellPropertis with bogus location info and would not update Volatile settings
        'this is required to get properties of a Field from the Field Setting page for cases not related to validation of the cell on the RawData sheet
        Set cellProperties = New clsCellProperties
        cellProperties.CellFieldName = fieldName
        cellProperties.CellAddress = "ZZ:10000" 'bogus number
        updateVolatileSetting = False
    End If
    
    'get the Field Settings for the selected field
    'populate dictFieldSettings dictionary with an entry for each of the Fields
    If Not dictFieldSettings.Exists(cellProperties.CellFieldName) Then
        Set oFieldSettings = New clsFieldSettings
        oFieldSettings.InitializeValues cellProperties
        dictFieldSettings.Add cellProperties.CellFieldName, oFieldSettings
    Else
        Set oFieldSettings = dictFieldSettings(cellProperties.CellFieldName)
        If updateVolatileSetting Then oFieldSettings.UpdateVolatileSettings cellProperties
    End If
    
    Set GetFieldSettingsInstance = oFieldSettings
    
End Function

'this function will loop through the FieldSettings Dictionary and check if any of the fields have to be excluded from the export. If such field found, the corresponded column will be deleted from the passed worksheet
Public Sub RemoveColumnsExcludedFromExport(tempWorksheet As Worksheet, sourceWorksheet As Worksheet)
    Dim rRng As Range, rCell As Range, rDropDown As Range ', rDropDownValues As Range
    Dim iCols As Integer
    Dim oFieldSettings As clsFieldSettings
    Dim cellProperties As clsCellProperties
    Dim curOffSet As Integer 'this number will track number of deleted columns and off-set the address of the collumns following the deleted one
    
    Const startCell = "A2" 'first cell containing field values
    
    With tempWorksheet
        iCols = .UsedRange.Columns.Count 'number of actually used columns
        'iRows = .UsedRange.Rows.Count 'number of actually used rows
        
        'identify range of actually used columns on the given spreadsheet. i.e. RawData. The Range will contain 1 row and all filled columns.
        Set rRng = sourceWorksheet.Range(startCell & ":" & Cells(.Range(startCell).Row, iCols).Address)
        
        'Loop through each cell (column) of the range, identify field assigned to this column and check if it marked as dropdown. If so, apply dropdown settings. If not clear out dropdown settings
        For Each rCell In rRng.Cells
            
            'Debug.Print rCell.Address
            
            Set cellProperties = New clsCellProperties
            cellProperties.InitializeValues rCell.Address

            'get the Field Settings for the selected field
            'populate dictFieldSettings dictionary with an entry for each of the Fields
            If Not dictFieldSettings.Exists(cellProperties.CellFieldName) Then
                Set oFieldSettings = New clsFieldSettings
                oFieldSettings.InitializeValues cellProperties
                dictFieldSettings.Add cellProperties.CellFieldName, oFieldSettings
            Else
                Set oFieldSettings = dictFieldSettings(cellProperties.CellFieldName)
                'oFieldSettings.UpdateVolatileSettings cellProperties 'in this sub we do not care about volatile values
            End If
'Debug.Print "Columns in 1st row range: " & rRng.Columns.Count, oFieldSettings.fieldName, rCell.Address, "Exclude: " & oFieldSettings.FieldExcludeFromExport, "curOffSet = " & curOffSet

            'if the current field is marked as Excluded From Export, drop the current column
            If oFieldSettings.FieldExcludeFromExport Then
                'Range(cellProperties.CellAddress).EntireColumn.Delete
                'Offset method below is used to compensate address changes of the current cell in case if any of the preceding columns were deleted. By default curOffSet = 0
                .Range(cellProperties.CellAddress).Offset(0, curOffSet).EntireColumn.Delete
                curOffSet = curOffSet - 1 'this will count number of deleted columns
            End If
            
            Set cellProperties = Nothing
        Next rCell
    
    End With
End Sub

'This function syncs fields listed in the FieldSettings sheets to RawData and Validated sheets. List of fields from FieldSettings will be transposed to the other sheets.
Public Sub SyncFieldsAccrossSheets()
    Dim iRows As Integer, iCols As Integer, rRng As Range, numFields As Integer
    Dim tempFieldsListRange As Variant
    Dim curVoidAutomatedValidation As Boolean
    
    With Worksheets(cSettingsWorksheetName)
        'iCols = .UsedRange.Columns.Count 'number of actually used columns
        iRows = .UsedRange.Rows.Count 'number of actually used rows
        numFields = iRows - Worksheets(cSettingsWorksheetName).Range(cFieldSettings_FirstFieldCell).Row + 1 'calculate number of fields in the list => Total - start row (with an adjustment)
        
        'identify range of actually used cells on the given spreadsheet
        Set rRng = .Range(cFieldSettings_FirstFieldCell & ":" & Cells(iRows, 1).Address)
        tempFieldsListRange = rRng.value
        
'        curVoidAutomatedValidation = bVoidAutomatedValidation 'save current status of bVoidAutomatedValidation into a temp variable
'        bVoidAutomatedValidation = True 'formating removal will trigger automated validation, to prevent that this flag is set temporarily true
        
        On Error Resume Next
        'disable all application events while columns are being synced
        Application.EnableEvents = False
        
        With Worksheets(cRawDataWorksheetName)
            'clean existing fields on RawData Sheet
            iCols = .UsedRange.Columns.Count 'number of actually used columns
            .Range(cRawData_FirstColumnCell & ":" & Cells(Range(cRawData_FirstColumnCell).Row, iCols).Address).ClearContents
            'apply new list of fields to RawData sheet
            .Range(cRawData_FirstColumnCell & ":" & Cells(Range(cRawData_FirstColumnCell).Row, numFields).Address) = Application.Transpose(tempFieldsListRange)
        End With
        
        With Worksheets(cValidatedWorksheetName)
            'clean existing fields on RawData Sheet
            iCols = .UsedRange.Columns.Count 'number of actually used columns
            Worksheets(cValidatedWorksheetName).Range(cValidated_FirstColumnCell & ":" & Cells(Range(cValidated_FirstColumnCell).Row, iCols).Address).ClearContents
            'apply new list of fields to Validated sheet
            Worksheets(cValidatedWorksheetName).Range(cValidated_FirstColumnCell & ":" & Cells(Range(cValidated_FirstColumnCell).Row, numFields).Address) = Application.Transpose(tempFieldsListRange)
        End With
        
        'set global flag ON to notify users about modifications applied to fields captions on RawData and Validated fields
        bFieldHeadersWereSynced = True
        
        'enable application events
        Application.EnableEvents = True
        On Error GoTo 0
        
'        bVoidAutomatedValidation = curVoidAutomatedValidation 'set bVoidAutomatedValidation back to the original value
        
'TODO - create a flag to notify user about fields updates when they naviagate ot RawData or Validated sheet. This should be a one time update.
        
    End With
End Sub

Public Sub NotifyUserAboutFieldSyncChanges()

    'this function will display a notification for users about possible changes to the column headers and their quantity as a result of updates on the FieldSetting sheet
    MsgBox "Note that the list of fields on ""FieldSettings"" sheet was updated and synced to the ""RawData"" and ""Validated"" sheets." & _
        vbCrLf & "As a result, headears of the columns and the quantity of columns might have changed." & _
        vbCrLf & vbCrLf & "Please check column headers on both pages and apply corrections to the data in associated columns, if needed." & _
        vbCrLf & vbCrLf & "You might want to run ""Validate 'RawData' Sheet"" command (CTRL+SHIFT+S) to make sure all Field Setting changes were properly applied.", vbInformation, "Field captions were synced"
    bFieldHeadersWereSynced = False
End Sub

'this function will apply Excel Data Validation settings to the fields marked as dropdowns, so list of expected values is displayed
Public Sub ApplyDropdownSettingsToCells(Optional sWorksheetName As String = "RawData", Optional RemoveValidationOnly As Boolean = False)
    
    Dim rRng As Range, rCell As Range, rDropDown As Range, rDropDownValues As Range
    Dim iCols As Integer
    Dim oFieldSettings As clsFieldSettings
    Dim cellProperties As clsCellProperties
    
    Const startCell = "A2" 'first cell containing field values
    
    With Worksheets(sWorksheetName)
        iCols = .UsedRange.Columns.Count 'number of actually used columns
        'iRows = .UsedRange.Rows.Count 'number of actually used rows
        
        'identify range of actually used columns on the given spreadsheet. i.e. RawData. The Range will contain 1 row and all filled columns.
        Set rRng = .Range(startCell & ":" & Cells(.Range(startCell).Row, iCols).Address)
        
        'Loop through each cell (column) of the range, identify field assigned to this column and check if it marked as dropdown. If so, apply dropdown settings. If not clear out dropdown settings
        For Each rCell In rRng.Cells
            
            'Debug.Print rCell.Address, rCell.Value
            
            Set cellProperties = New clsCellProperties
            cellProperties.InitializeValues rCell.Address

            'get the Field Settings for the selected field
            'populate dictFieldSettings dictionary with an entry for each of the Fields
            If Not dictFieldSettings.Exists(cellProperties.CellFieldName) Then
                Set oFieldSettings = New clsFieldSettings
                oFieldSettings.InitializeValues cellProperties
                dictFieldSettings.Add cellProperties.CellFieldName, oFieldSettings
            Else
                Set oFieldSettings = dictFieldSettings(cellProperties.CellFieldName)
                'oFieldSettings.UpdateVolatileSettings cellProperties 'in this sub we do not care about volatile values
            End If
            
            'Debug.Print oFieldSettings.fieldName
            
            'Debug.Print .Range(Cells(rCell.Row, rCell.Column).Address & ":" & Cells(rCell.EntireColumn.Rows.Count, rCell.Column).Address).Address 'rCell.EntireColumn.Rows.Count
            'set the range of the cells (of the current worksheet) to be updated with the Validation rulles
            Set rDropDown = .Range(Cells(rCell.Row, rCell.Column).Address & ":" & Cells(rCell.EntireColumn.Rows.Count, rCell.Column).Address)
            
            rDropDown.Validation.Delete 'delete previously existing validation
        
            If oFieldSettings.FieldDropDownBool Then 'check if this field is currently marked as a dropdown field
                
                On Error Resume Next
                'set the range containing list of possible values for the current field
                'if the range of values to be used for dropdown values is not valid, an error will be produced; assignment of the validation rules will be performed only if there were no errors.
                Set rDropDownValues = Worksheets(cDictionayWorksheetName).Range(GetRangeOfUsedCellsInColumn(Worksheets(cDictionayWorksheetName), oFieldSettings.FieldDropDownValueLookupRange))
                
                'old code
'                Set rDropDownValues = Worksheets(cDictionayWorksheetName).Range( _
'                                Worksheets(cDictionayWorksheetName).Range(oFieldSettings.FieldDropDownValueLookupRange).Cells(1).Address _
'                                & ":" & _
'                                Worksheets(cDictionayWorksheetName).Range(oFieldSettings.FieldDropDownValueLookupRange).Cells(1).Offset(Worksheets(cDictionayWorksheetName).Rows.Count - Range(oFieldSettings.FieldDropDownValueLookupRange).Cells(1).Row).End(xlUp).Address _
'                                )
                'Worksheets(cDictionayWorksheetName).Range(oFieldSettings.FieldDropDownValueLookupRange).Cells(1).Offset(Range(oFieldSettings.FieldDropDownValueLookupRange).Rows.Count - 1).End(xlUp).Address _

                'Debug.Print rDropDownValues.Address
                
                If Err.Number = 0 Then
                    'Apply validation rules to all cells in the column corresponding to the current field
                    If Not RemoveValidationOnly Then 'This flag controls if Validation has to be just removed.
                        rDropDown.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                            Formula1:="='" & cDictionayWorksheetName & "'!" & rDropDownValues.Address
                    End If
                End If
                On Error GoTo 0
                
            End If
            
            Set cellProperties = Nothing
        Next rCell
    
    End With
    
    'MsgBox "ApplyDropdownSettings has finished."

End Sub


'function returns a range of all used cells in a column based on the first cell provided
Public Function GetRangeOfUsedCellsInColumn(curWorkSheet As Worksheet, firstCell As String) As String

    With curWorkSheet.Range(firstCell).Cells(1)
        GetRangeOfUsedCellsInColumn = .Address & ":" & .Offset(curWorkSheet.Rows.Count - .Row).End(xlUp).Address
    End With
'Worksheets(cDictionayWorksheetName).Range(oFieldSettings.FieldDropDownValueLookupRange).Cells(1).Address _
'  & ":" & _
'Worksheets(cDictionayWorksheetName).Range(oFieldSettings.FieldDropDownValueLookupRange).Cells(1).Offset(Worksheets(cDictionayWorksheetName).Rows.Count - .Range(oFieldSettings.FieldDropDownValueLookupRange).Cells(1).Row).End(xlUp).Address _


End Function

Public Sub ValidateCurrentCell()
    If ActiveCell.Worksheet.Name = cRawDataWorksheetName Then
        'Debug.Print ActiveCell.Address
        Dim cellRow As Integer
    
        cellRow = ActiveCell.Row 'Right(ActiveCell.Address, Len(ActiveCell.Address) - InStrRev(ActiveCell.Address, "$"))
            
        'Do not perform validation for the captions row (the first row)
        If cellRow > 1 Then
            Validate_Cell_Value ActiveCell
        Else
            MsgBox "Caption row is detected - cannot continue with the validation!" & vbCrLf & vbCrLf & "Select a cell on the ""RawData"" sheet that has to be validated and call this action again.", vbCritical, "Validation of an Active Cell"
        End If
    Else
        MsgBox "Cannot continue with the validation!" & vbCrLf & vbCrLf & "Select a cell on the ""RawData"" sheet that has to be validated and call this action again.", vbCritical, "Validation of an Active Cell"
    End If
End Sub

Public Sub FieldDetailsRequest_Event()
    Dim validRes As ValidationReportMsg
    
    'check if the SHIFT+F1 combination was pressed on RawData or Validated worksheets. Other worksheets will be ignored
    If ActiveCell.Worksheet.Name = cRawDataWorksheetName Or ActiveCell.Worksheet.Name = cValidatedWorksheetName Then
        'Check if Validation results exists for the cell
        If Not dictValidationResults.Exists(ActiveCell.Address) Then
            MsgBox "Currently selected cell (" & ActiveCell.Address & ") was not validated during this session (since this file was opened) yet." & vbCrLf & vbCrLf & "Run validation first and then request this report again.", vbExclamation
        Else
            'MsgBox "View report for this cell (" & ActiveCell.Address & ") now.", vbInformation
            validRes = CellValidationReport(Worksheets(ActiveCell.Worksheet.Name).Range(ActiveCell.Address))
            MsgBox validRes.ValidationMessage, validRes.MsgBoxStyle, "Validation Results"
        End If
    End If
    
End Sub

Private Function CellValidationReport(curCell As Range) As ValidationReportMsg
    Dim oValidationResults As clsValidationResult
    Dim oFieldSettings As clsFieldSettings
    Dim outVal As ValidationReportMsg
    
    'load Validation object for the cell
    If dictValidationResults.Exists(curCell.Address) Then
        Set oValidationResults = dictValidationResults(curCell.Address)
        If dictFieldSettings.Exists(oValidationResults.ValidatedCellProperties.CellFieldName) Then
            Set oFieldSettings = dictFieldSettings(oValidationResults.ValidatedCellProperties.CellFieldName)
        Else
            'report that no Field Settings info was found
            outVal.ValidationMessage = "No Field Settings were found for the given field ('" & oValidationResults.ValidatedCellProperties.CellFieldName & "'). Please re-validate this cell."
            outVal.MsgBoxStyle = vbInformation
            CellValidationReport = outVal
            Exit Function
        End If
        
        'this is required to get proper values for  volatile values of FieldSetting properties (values that depends on values of other fields)
        oFieldSettings.UpdateVolatileSettings oValidationResults.ValidatedCellProperties
        
        'Proceed here if all needed information for the cell is present
        Dim sb As New StringBuilder
        
        sb.Append "VALIDATION SUMMARY FOR THE CELL - "
        sb.Append Replace(oValidationResults.ValidatedCellProperties.CellAddress, "$", "")
        sb.Append vbCrLf
        
        'validation results
        sb.Append vbCrLf & "VALIDATION RESULTS" & vbCrLf & "-------------------------------- " & vbCrLf
        sb.Append "Validation Status: "
        sb.Append ValidationOutcomeStatus_toString(oValidationResults.ValidationStatus) & vbCrLf
        sb.Append "Initial Value (RawData sheet): " & vbCrLf
        sb.Append oValidationResults.InitialValue & vbCrLf
        sb.Append "Validated Value (Validated sheet): " & vbCrLf
        sb.Append oValidationResults.ValidatedValue & vbCrLf
        
        If oValidationResults.ValidationErrors.ErrorCount > 0 Then
            'Report Errors
            sb.Append oValidationResults.ValidationErrors.toString
        End If
        
        'field settings
        sb.Append vbCrLf & "FIELD SETTINGS" & vbCrLf & "-------------------------------- " & vbCrLf
        sb.Append "Field Name: "
        sb.Append oFieldSettings.fieldName & vbCrLf
        sb.Append "Required: "
        sb.Append IIf(oFieldSettings.FieldRequiredBool, "Yes", "No") & vbCrLf
        sb.Append "Predefined values (dropdown field): "
        sb.Append IIf(oFieldSettings.FieldDropDownBool, "Yes", "No - Open Text") & vbCrLf
        sb.Append "Default Value: "
        sb.Append IIf(oFieldSettings.FieldDefaultBool, oFieldSettings.FieldDefaultValue, "No Default value") & vbCrLf
        sb.Append "Date Field: "
        sb.Append IIf(oFieldSettings.FieldDateType, "Yes", "No") & vbCrLf
        sb.Append "Numeric Only: "
        sb.Append IIf(oFieldSettings.FieldNumericOnlyType, "Yes", "No") & vbCrLf
        sb.Append "Calculated: "
        sb.Append IIf(oFieldSettings.FieldCalculated, "Yes", "No") & vbCrLf
        sb.Append "Triggers Calculation of other fields: "
        sb.Append IIf(oFieldSettings.FieldCalcTrigger, "Yes", "No") & vbCrLf
        If oFieldSettings.FieldCalcTrigger Then
            sb.Append "Calculation Overwrites Existing Value of the Target Field: "
            sb.Append IIf(oFieldSettings.FieldCalcOverwriteExistingVal, "Yes", "No") & vbCrLf
        End If
        sb.Append "Exclude From Export: "
        sb.Append IIf(oFieldSettings.FieldExcludeFromExport, "Yes", "No") & vbCrLf
        
        'MsgBox sb.toString, IIf(oValidationResults.ValidationErrors.ErrorCount > 0, vbExclamation, vbInformation), "Validation Results"
        outVal.ValidationMessage = sb.toString
        outVal.MsgBoxStyle = IIf(oValidationResults.ValidationErrors.ErrorCount > 0, vbExclamation, vbInformation)
        
        Set sb = Nothing
    Else
        'report that no validation info is available
        outVal.ValidationMessage = "No Validation Results were found. Please re-validate this cell."
        outVal.MsgBoxStyle = vbInformation
    End If
    
    CellValidationReport = outVal
    
    'outVal = AllValidatedCellsErrorReport()
    
End Function

Private Function AllValidatedCellsErrorReport() As ValidationReportMsg
    Dim oValidRes As clsValidationResult
    Dim oFieldSet As clsFieldSettings
    Dim outVal As ValidationReportMsg
    Dim errDict As New Dictionary, Key As Variant
    Dim msgOut As ValidationReportMsg, sb As StringBuilder
    
    If Not dictValidationResults Is Nothing Then 'check if validation results are available
    
        For Each Key In dictValidationResults.Keys 'loop through all validation results and check for reported errors
            Set oValidRes = dictValidationResults.Item(Key)
            If oValidRes.ValidationErrors.ErrorCount > 0 Then
                'group reported validation errors by field names; store total count of errors per the field name
                If errDict.Exists(oValidRes.ValidatedCellProperties.CellFieldName) Then
                    errDict(oValidRes.ValidatedCellProperties.CellFieldName) = errDict(oValidRes.ValidatedCellProperties.CellFieldName) + 1
                Else
                    errDict.Add oValidRes.ValidatedCellProperties.CellFieldName, 1
                End If
            End If
        Next
        
        Set sb = New StringBuilder
        
        'prepare output message
        If errDict.Count > 0 Then
            Dim rowCount As Integer
            For Each Key In errDict.Keys
                If rowCount > 0 Then
                    sb.Append vbCrLf
                End If
                sb.Append Key & ": " & errDict.Item(Key) & " errors"
                rowCount = rowCount + 1
            Next
            outVal.MsgBoxStyle = vbCritical
        Else
            sb.Append "No Errors were found."
            outVal.MsgBoxStyle = vbInformation
        End If
        
        outVal.ValidationMessage = sb.toString
        
        Set sb = Nothing
        
    Else
        'Return message informing that there was no validation results collected
        outVal.MsgBoxStyle = vbCritical
        outVal.ValidationMessage = "No validation results were found! Please run the validation procedure."
    End If
    
    AllValidatedCellsErrorReport = outVal
    Set errDict = Nothing
End Function

Public Function ApplyRegExToStr(strVal As String, strRegEx As String) As Object
    Dim regex As Object
    'Dim matches As Object, match As Object
    
    'intiate regex object and pass there the search pattern
    Set regex = CreateObject("VBScript.RegExp")
    With regex
      .Pattern = strRegEx
      .Global = True
    End With
         
    If regex.Test(strVal) Then
        'if pattern was found, the found field names are needed to be filled with the actual field values from Validated sheet
        Set ApplyRegExToStr = regex.Execute(strVal)
    Else
        Set ApplyRegExToStr = Nothing
    End If
    Set regex = Nothing
    
End Function

Private Function ValidationOutcomeStatus_toString(status As ValidationOutcomeStatus) As String
    Dim strOut As String
    
    Select Case status
        Case ValidationOutcomeStatus.CalculatedValueApplied
            strOut = "Calculated value Applied"
        Case ValidationOutcomeStatus.DefaultAssigned
            strOut = "Default value Applied"
        Case ValidationOutcomeStatus.MapConversionApplied
            strOut = "Mapped Value Applied"
        Case ValidationOutcomeStatus.ValidationError
            strOut = "Validation Error"
        Case ValidationOutcomeStatus.ValidationPassed
            strOut = "Passed Validation"
        Case Else
            strOut = "Unknown Status"
    End Select
    
    ValidationOutcomeStatus_toString = strOut
    
End Function

Private Function ValidationErrorStatus_toString(status As ValidationErrorStatus) As String
    Dim strOut As String
    
    Select Case status
        Case ValidationErrorStatus.CombinationOfErrors
            strOut = "Combination of errors"
        Case ValidationErrorStatus.NoErrors
            strOut = "No Errors"
        Case ValidationErrorStatus.RequiredFieldEmpty
            strOut = "Required field left empty"
        Case ValidationErrorStatus.UnexpectedValue
            strOut = "Unexpected Value"
        Case ValidationErrorStatus.IncorrectDate
            strOut = "Incorrect Date value"
        Case ValidationErrorStatus.FieldCalculationError
            strOut = "Error of Processing Calculated Filed"
        Case Else
            strOut = "Unknown Status"
    End Select
    
    ValidationErrorStatus_toString = strOut
    
End Function

Public Sub SwitchDropDownFunctionaltiyOnOff()
    Dim strCurMenuCaption As String
    
    'get current menu caption
    'strCurMenuCaption = cCustomMenu_SetDropdowonFunc & IIf(bVoidDropDownFunctionality, "ON", "OFF")
    strCurMenuCaption = GetSwitchableMenuCaption(bVoidDropDownFunctionality, cCustomMenu_SetDropdowonFunc, cCustomMenu_SetDropdowonFunc_ShortCut)
    
    'reset boolean flag
    bVoidDropDownFunctionality = IIf(bVoidDropDownFunctionality, False, True)
    
    'Update Menu caption
    Application.CommandBars("Worksheet Menu Bar").Controls(cCustomMenuName).Controls(cCustomMenu_SubMenuSettings).Controls(strCurMenuCaption).Caption = _
        GetSwitchableMenuCaption(bVoidDropDownFunctionality, cCustomMenu_SetDropdowonFunc, cCustomMenu_SetDropdowonFunc_ShortCut)
        'cCustomMenu_SetDropdowonFunc & IIf(bVoidDropDownFunctionality, "ON", "OFF") & "    CTRL+SHIFT+D"
    
    MsgBox "Dropdown functionality on ""RawData"" sheet was turned " & IIf(bVoidAutomatedValidation, "ON", "OFF") & "." & vbCrLf & vbCrLf & _
        "Note: Dropdown functionality can be switched through the ""Add-Ins/MSSM Menu/Settings/Set Dropdown Functionality"" menu. ", _
        vbInformation, "Automatic Validation Status"
End Sub

Public Sub SwitchValidationFunctionaltiyOnOff()
    Dim strCurMenuCaption As String
    
    'get current menu caption
    'strCurMenuCaption = cCustomMenu_SetValidationFunc & IIf(bVoidAutomatedValidation, "ON", "OFF")
    strCurMenuCaption = GetSwitchableMenuCaption(bVoidAutomatedValidation, cCustomMenu_SetValidationFunc, cCustomMenu_SetValidationFunc_ShortCut)
    
    'reset boolean flag
    bVoidAutomatedValidation = IIf(bVoidAutomatedValidation, False, True)
    
    'Update Menu caption
    Application.CommandBars("Worksheet Menu Bar").Controls(cCustomMenuName).Controls(cCustomMenu_SubMenuSettings).Controls(strCurMenuCaption).Caption = _
        GetSwitchableMenuCaption(bVoidAutomatedValidation, cCustomMenu_SetValidationFunc, cCustomMenu_SetValidationFunc_ShortCut)
        'cCustomMenu_SetValidationFunc & IIf(bVoidAutomatedValidation, "ON", "OFF") & "    CTRL+SHIFT+V"
    
    MsgBox "Automatic Validation functionality on ""RawData"" sheet was turned " & IIf(bVoidAutomatedValidation, "ON", "OFF") & "." & vbCrLf & vbCrLf & _
        "Note: Automatic validation functionality can be switched through the ""Add-Ins/MSSM Menu/Settings/Set Automatic Validation"" menu. ", _
        vbInformation, "Automatic Validation Status"
End Sub

'prepare menu caption for items that depends on the current status of boolean variable
Public Function GetSwitchableMenuCaption(bFlag As Boolean, sMenuCaption As String, sKeyShortCut As String)
    GetSwitchableMenuCaption = sMenuCaption & IIf(bFlag, "ON ", "OFF") & sKeyShortCut
End Function

Public Sub ShowAboutMessage()
    'this function will show an About message box. It can be invoked from the custom menu "About"

    Dim sb As New StringBuilder
    
    sb.Append cHelpTitle
    sb.Append vbCrLf & vbCrLf
    sb.Append "Version: "
    sb.Append cHelpVersion
    sb.Append vbCrLf & vbCrLf
    sb.Append cHelpDescription
    
    MsgBox sb.toString, , cHelpTitle
    
    Set sb = Nothing
End Sub

Public Sub RegisterCustomEvents()
    'create application level OnKey press event handlers
    Application.OnKey "+{F1}", "FieldDetailsRequest_Event" 'SHIFT+F1
    Application.OnKey "^+{s}", "ValidateWholeWorksheet" 'CTRL+SHIFT+S
    Application.OnKey "^+{c}", "ValidateCurrentCell" 'CTRL+SHIFT+C
    Application.OnKey "^+{e}", "ExportValidateSheet" 'CTRL+SHIFT+E
    Application.OnKey "^+{v}", "SwitchValidationFunctionaltiyOnOff" 'CTRL+SHIFT+V
    Application.OnKey "^+{d}", "SwitchDropDownFunctionaltiyOnOff" 'CTRL+SHIFT+D
    Application.OnKey "^+{h}", "HighlightDuplicates" 'CTRL+SHIFT+H
    'Application.OnKey "^{v}", "PasteAsSwitch" 'CTRL+V
    
End Sub

Public Sub UnRegisterCustomEvents()
    'un-assign OnKey press event handlers
    Application.OnKey "+{F1}" 'SHIFT+F1
    Application.OnKey "^+{s}" 'CTRL+SHIFT+S
    Application.OnKey "^+{c}" 'CTRL+SHIFT+C
    Application.OnKey "^+{e}" 'CTRL+SHIFT+E
    Application.OnKey "^+{v}" 'CTRL+SHIFT+V
    Application.OnKey "^+{d}" 'CTRL+SHIFT+D
    Application.OnKey "^+{h}" 'CTRL+SHIFT+H
    'Application.OnKey "^{v}" 'CTRL+V
End Sub


Public Sub LoadCustomMenus()
    Dim cmbBar As CommandBar
    Dim cmbControl As CommandBarControl
    Dim cmbSettings As CommandBarControl
    Dim cmbDBLink As CommandBarControl
    
    'add custom menus to Add-In ribon of Excel
    Set cmbBar = Application.CommandBars("Worksheet Menu Bar")
    
    'check if the custom menu already exists. If it exists, delete it; it will be recreated in the later code
    Dim i As Integer ', boolMenuExists As Boolean
    For i = cmbBar.Controls.Count To 1 Step -1
        If cmbBar.Controls.Item(i).Caption = cCustomMenuName Then
            'boolMenuExists = True
            cmbBar.Controls.Item(i).Delete
            Exit For
        End If
    Next
    
    'create menu bar entries
    Set cmbControl = cmbBar.Controls.Add(Type:=msoControlPopup, Temporary:=True) 'adds a menu item to the Menu Bar
    With cmbControl
        .Caption = cCustomMenuName 'names the menu item
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Validate ""RawData"" Sheet    CTRL+SHIFT+S" 'adds a description to the menu item
            .OnAction = "ValidateWholeWorksheet" 'runs the specified macro
            .FaceId = 501 '638 '1098 'assigns an icon to the dropdown
        End With
        With .Controls.Add(Type:=msoControlButton) 'adds a dropdown button to the menu item
            .Caption = "Validate Current Cell            CTRL+SHIFT+C" 'adds a description to the menu item
            .OnAction = "ValidateCurrentCell" 'runs the specified macro
            .FaceId = 501 '638 '1098 'assigns an icon to the dropdown
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Show Cell Validation Report          SHIFT+F1"
            .OnAction = "FieldDetailsRequest_Event"
            .FaceId = 18 '501
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Highlight Duplicates           CTRL+SHIFT+H"
            .OnAction = "HighlightDuplicates"
            .FaceId = 351 '501
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Exoport ""Validated"" Sheet   CTRL+SHIFT+E"
            .OnAction = "ExportValidateSheet"
            .FaceId = 638 '18
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Clear Validation Formating"
            .OnAction = "ClearFormatingOfWorkbook_MenuCall"
            .FaceId = 108
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Read Flatbed Scanner"
            .OnAction = "FBS_Scan"
            .FaceId = 485 '18
        End With
         
        'create sub menu "DB Link"
        Set cmbDBLink = .Controls.Add(Type:=msoControlPopup, Temporary:=True)
        With cmbDBLink
            .Caption = "Database Link"
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Submit Manifest IDs"
                .OnAction = "SubmitManifests"
                .FaceId = 3000
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Sync Dictionary values with Database"
                .OnAction = "LoadDictionaryValues"
                .FaceId = 3000
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Load Field Setting Profile"
                .OnAction = "LoadFieldSettings"
                .FaceId = 3000
            End With
        End With
        
        'create sub menu "Settings"
        Set cmbSettings = .Controls.Add(Type:=msoControlPopup, Temporary:=True)
        With cmbSettings
            .Caption = "Settings"
            With .Controls.Add(Type:=msoControlButton)
                '.Caption = cCustomMenu_SetValidationFunc & IIf(bVoidAutomatedValidation, "ON", "OFF")
                .Caption = GetSwitchableMenuCaption(bVoidAutomatedValidation, cCustomMenu_SetValidationFunc, cCustomMenu_SetValidationFunc_ShortCut)
                .OnAction = "SwitchValidationFunctionaltiyOnOff"
                .FaceId = 611
            End With
            With .Controls.Add(Type:=msoControlButton)
                '.Caption = cCustomMenu_SetDropdowonFunc & IIf(bVoidDropDownFunctionality, "ON", "OFF")
                .Caption = GetSwitchableMenuCaption(bVoidDropDownFunctionality, cCustomMenu_SetDropdowonFunc, cCustomMenu_SetDropdowonFunc_ShortCut)
                .OnAction = "SwitchDropDownFunctionaltiyOnOff"
                .FaceId = 611
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Sync Dropdown values to RawData sheet"
                .OnAction = "ApplyDropdownSettingsToCells"
                .FaceId = 3000
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Sync ""FieldSetting"" fields to data sheets"
                .OnAction = "SyncFieldsAccrossSheets"
                .FaceId = 3000
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "About " & cHelpTitle
                .OnAction = "ShowAboutMessage"
                '.FaceId = 3000
            End With
        End With
        
        
    
    End With
End Sub

Public Sub HighlightDuplicates(Optional rTargetCells As Range)
    Dim Target As Range, rCell As Range
    Dim firstRowRange As Range, rFirstRowCell As Range
    
    If rTargetCells Is Nothing Then
        Set rTargetCells = Selection 'ActiveCell
    End If
    
    'identify range consisting of cells of the first row of the passed range (or current selection)
    Set firstRowRange = Range(rTargetCells.Cells(1).Address, rTargetCells.Cells(1).Offset(0, rTargetCells.Columns.Count - 1).Address)
    
    'loop through each cell of the first row
    For Each rFirstRowCell In firstRowRange
        
        'identify range of all used cells in the column corresponding to the rFirstRowCell
        Set Target = Range(rFirstRowCell.Cells(1).Address, Cells(Rows.Count, rFirstRowCell.Column).End(xlUp))
            
        'loop through each cell of the identified column's range and utilize standard CountIf function to identify duplicates
        For Each rCell In Target
            If WorksheetFunction.CountIf(Target, rCell.value) > 1 Then
                'Debug.Print rCell.Value, "Duplicate"
                
                'highlight found duplicates
                rCell.Interior.Color = BackgroundColors.Yellow
                rCell.Font.Color = FontColors.DarkYellow
            End If
            
        Next
    
    Next
    
End Sub

'this sub will make sure that Ctrl+V by default inserts only values (without formulas, formatting, etc.)
Public Sub PasteAsSwitch()
    On Error Resume Next 'this will prevent a run-time error in case if Ctrl+V is pressed when buffer is empty
    'Debug.Print "PasteAsSwitch", Now()
    If bSetCtrlVPasteAsValues Then
        ActiveCell.PasteSpecial Paste:=xlPasteValues
    Else
        ActiveCell.PasteSpecial Paste:=xlPasteAll
    End If
    On Error GoTo 0
End Sub

Public Function GetConfigValue(Key As String) As Variant
    With Worksheets(cConfigWorksheetName)
        Dim fnr As Range
        Dim iRows As Integer
        
        iRows = .UsedRange.Rows.Count 'number of actually used rows
        
        'identify range of actually used cells on the given spreadsheet and apply Find function
        Set fnr = .Range(cConfig_FirstFieldCell & ":" & .Cells(iRows, 1).Address).Find(Key, LookIn:=xlValues) 'fnr will contain the cell matching the find criteria
        
        'if fnr is not Nothing, retrun the associate value
        If Not fnr Is Nothing Then
            GetConfigValue = fnr.Offset(0, 1).value 'it will return value of the cell located immediately on a right from the address located in fnr
        Else
            GetConfigValue = Null
        End If
    End With
End Function

Public Function SetConfigValue(Key As String, value As String) As Integer
    With Worksheets(cConfigWorksheetName)
        Dim fnr As Range
        Dim iRows As Integer
        
        iRows = .UsedRange.Rows.Count 'number of actually used rows
        
        'identify range of actually used cells on the given spreadsheet and apply Find function
        Set fnr = .Range(cConfig_FirstFieldCell & ":" & .Cells(iRows, 1).Address).Find(Key, LookIn:=xlValues) 'fnr will contain the cell matching the find criteria
        
        If Not fnr Is Nothing Then
            'if fnr is not Nothing, set the given Value for the requested Key
            fnr.Offset(0, 1).value = value 'it set the value of the found config cell to the passed Value
            
            SetConfigValue = 1
        Else
            'if the requested Key was not found, add a new entry for the key
            Set fnr = .Range(Cells(iRows, 1).Offset(1, 0).Address)
            fnr.value = Key
            fnr.Offset(0, 1).value = value
            
            SetConfigValue = 2
        End If
    End With
    
    Exit Function
    
'TODO - add error handler
End Function

'This function will return the location of the value of the requested field. It will use the cellProperties object to identify the correct row to be used
Public Function FindFieldAddress(fieldName As String, cellRow As String, Optional sWorksheetName As String = "RawData") As String
    'fieldName example: {MT_Project}
    
    fieldName = Replace(Replace(fieldName, "{", ""), "}", "") 'remove figure brackets from the name
    
    'This will find address of the cell to be evaluated
    With Worksheets(sWorksheetName)
        Dim fnr As Range, fcol As String
        
        'Find the supplied column name on the Validated sheet
        '   Used Range method and Cells object will retrun the cell of the actually used last column on the sheet
        Set fnr = .Range("A1:" & Cells(1, .UsedRange.Columns.Count).Address).Find(fieldName, LookIn:=xlValues)
        If Not fnr Is Nothing Then
            'find the column associated with the given field name
            'fcol = Replace(Left(fnr.Address, InStrRev(fnr.Address, "$")), "$", "")
            'identify an address of the cell that should be used for evaluation
            'FindFieldAddress = fcol & cellRow 'cValidatedWorksheetName & "!" &
            
            FindFieldAddress = Cells(cellRow, fnr.Column).Address
        Else
            FindFieldAddress = ""
        End If
    End With
    
End Function

Function GetUniqueValues(ws As Worksheet, col As Long) As Variant
    Dim data(), r As Long
    Dim dr As New Dictionary
    
    data = ws.UsedRange.Columns(col).Value2  'ws.UsedRange.value
    
    If UBound(data) > 1 Then 'make sure that the array has more then 1 member (the first member is the column name)
        For r = 2 To UBound(data)
            If Not dr.Exists(data(r, 1)) And Not data(r, 1) = Empty Then
                dr.Add data(r, 1), Empty
            End If
        Next r
    End If
    
    GetUniqueValues = dr.Keys()
End Function

'this function will retrieve a disctinct list values presented in the field (default: Manifest IDs) located on the given sheet (default: Validated sheet)
Public Function Get_DisticntValuesFromField(Optional fieldName As String = "MT_ManifestID", Optional worksheetName As String = cValidatedWorksheetName) As String
    Dim arr As Variant
    Dim val As String, out As String, r As Long
    Dim field_addr As String
    
    out = "" 'default output
    
    field_addr = FindFieldAddress(fieldName, 1)
    
    If Len(field_addr) > 0 Then 'if the given field was found, attempt to get values from it
        'the following is passing the worksheet and the column number (of ManifestID column) to be used for collecting unique values
        arr = GetUniqueValues(Worksheets(worksheetName), Range(field_addr).Column)
        
        For r = 0 To UBound(arr)
            out = out + arr(r)
            If r <> UBound(arr) Then
                out = out + ","
            End If
        Next
    
    End If
    
    Get_DisticntValuesFromField = out
End Function

Public Function Get_MiscSettingValue(field_name As String, setting_name As String, _
                                    Optional setting_delim As String = ";", _
                                    Optional value_delim As String = ":") As String
    
    Dim oFieldSettings As clsFieldSettings
    Dim misc As String
    Dim arList() As String, arSetting() As String
    Dim i As Integer
    Dim outVal As String
    
    outVal = "" 'default out value
    
    'get misc settings for the ManifestID field
    Set oFieldSettings = GetFieldSettingsInstance(Nothing, False, field_name) '"MT_ManifestID"
    If oFieldSettings.DataAvailable Then
        misc = oFieldSettings.FieldMiscSettings
    End If
    
    arList = Split(misc, setting_delim)
    For i = 0 To UBound(arList)
        arSetting = Split(arList(i), value_delim)
        If UBound(arSetting) > 0 Then 'check that setting info is not empty
            If Trim(arSetting(0)) = Trim(setting_name) Then
                outVal = arSetting(1)
            End If
        End If
    Next
    
    Get_MiscSettingValue = outVal
    
End Function

'this sub initiates copying collected scans to Dictionary tab.
'this will check "scanner" property of the misc_settings column to match name of the current worksheet with that. Process won't go through if there is no match.
Public Sub RequestCopyScansToDict(fld_name As String, ws_name As String, scansRngAddr As String, Optional miscVarName As String = "scanner")
    Dim oFieldSettings As clsFieldSettings
    
    If Get_MiscSettingValue(fld_name, miscVarName) = ws_name Then
        'set parameters for MT_Box Label field.
        Set oFieldSettings = GetFieldSettingsInstance(Nothing, False, fld_name)
        If oFieldSettings.DataAvailable And Len(Trim(oFieldSettings.FieldDropDownValueLookupRange)) Then
            'sourceRangeAddress = "D1" 'first cell of the column storring the list of Box IDs on the FlatbedScans tab
            CopyValuesToDictionarySheet scansRngAddr, oFieldSettings.FieldDropDownValueLookupRange, ws_name, True
        End If
        'CopyBoxIDsToDictionarySheet oFieldSettings.FieldDropDownValueLookupRange
        'CopyBoxIDsToDictionarySheet "BA3" ' - for test only
    End If
End Sub

'used from FlatbedScanner and HandledScanner sheets
'copy values (i.e. Box ID ) entered on the Flatbedscanner sheet to the Dictionary sheet to populate values for corresponede dropdown (i.e. MT_Box ID)
Public Sub CopyValuesToDictionarySheet(sourceRandgeAddress As String, targetRangeAddress As String, ws_name As String, Optional CopyUniqueValuesOnly As Boolean = False)

    With Worksheets(ws_name)
        
        Dim targetRange As Range, sourceRange As Range, tRange As Range
        Dim targetRangeStart As Range, targetRangeEnd As Range
        
        Set sourceRange = .Range(sourceRandgeAddress & ":" & .Range(sourceRandgeAddress).Offset(.Rows.Count - .Range(sourceRandgeAddress).Row).End(xlUp).Address) 'this source range will include all cells (in this column) located below the given cell
        
        With Worksheets(cDictionayWorksheetName)
            'clear target range
            Set targetRangeStart = .Range(targetRangeAddress)
            Set targetRangeEnd = .Range(targetRangeAddress).Offset(.Rows.Count - .Range(targetRangeAddress).Row, 0).End(xlUp)
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
            sourceRange.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=targetRange, unique:=CopyUniqueValuesOnly
            'copy the same information (as in previous row) to the 2nd column shifted to the right on 2 cells
            sourceRange.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=targetRange.Offset(0, 2), unique:=CopyUniqueValuesOnly
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
