VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
     
    'set comparision mode of the global dictionary - dictValidationResults
    dictValidationResults.CompareMode = TextCompare
    
    'set automatic calculation for the whole workbook - this will make sure that Automatic calculation is set as default
    Application.Calculation = xlCalculationAutomatic
    
    'set dropdown functionality based on the config settings
    bVoidAutomatedValidation = CBool(GetConfigValue("Dflt_bVoidAutomatedValidation")) 'True
    
    'register local events
    RegisterCustomEvents
    
    'create custom menus
    LoadCustomMenus
    
    'Prompt to run full validation on the file opening
'    Dim iResponse As Integer
'    iResponse = MsgBox("The system is about to start validating all values presented on the RawData spreadsheet. " & vbCrLf & vbCrLf & "Do you want to proceed? If not, click 'Cancel'." & vbCrLf & vbCrLf _
'                        & "Note: If cancel now, you can re-run this process later from the 'Add-In/MenuCommands/MSSM Menu/Validate RawData Sheet' menu.", _
'                        vbOKCancel, "MSSM Data Entry File")
'    If iResponse = vbOK Then
'        'on opening of the file, run validation of whole RawData worksheet. It will validate all currently existing data
'        ValidateWholeWorksheet cRawDataWorksheetName, , True
'    End If
     
End Sub

'Testing only ---->  test sub for testing loading of workbook
Public Sub TestLoad()
    Workbook_Open
End Sub
'Testing only ---->  test sub for testing closing of workbook
Public Sub TestBeforeClose()
    Workbook_BeforeClose False
End Sub
 
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next 'in case the menu item has already been deleted
    Application.CommandBars("Worksheet Menu Bar").Controls(cCustomMenuName).Delete 'delete the menu item
    
    UnRegisterCustomEvents 'remove registrations for custom events
    
    'Clean up memory - kill dictionary objects
    Set dictValidationResults = Nothing
    Set dictFieldSettings = Nothing
        
End Sub

