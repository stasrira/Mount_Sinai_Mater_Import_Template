VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsValidationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public ValidationErrors As clsFieldValidationErrors
Public InitialValue As String
Public ValidatedValue As String
Public ValidatedCellProperties As clsCellProperties
Public ValidationStatus As ValidationOutcomeStatus


Private Sub Class_Initialize()
    'for test only
'    Debug.Print "ValidationResult Initiated"
'    Debug.Print InitialValue
End Sub
