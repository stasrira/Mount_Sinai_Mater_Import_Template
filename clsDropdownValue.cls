VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDropdownValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strDropdownLookupRange As String
Private strErrorMessage As String
Private obError As clsFieldValidationErrors
'Private DropdownLookupRange As Range

'Private Const cRawValueColumn = 1
'Private Const cValidatedValueColumn = 3

Public Sub InitializeValues(vDropdownLookupRange As String, vErrorMessage As String, ByRef ErrorCollection As clsFieldValidationErrors)
    'strDropdownLookupRange = vDropdownLookupRange
    strDropdownLookupRange = GetRangeOfUsedCellsInColumn(Worksheets(cDictionayWorksheetName), vDropdownLookupRange)
    strErrorMessage = vErrorMessage
    Set obError = ErrorCollection
End Sub

Public Function VaildateValue(vFieldValue As String)
    'Using provided range (on cDictionayWorksheetName sheet), it will search for the given FieldValue and return the validated value
    'if nothing is found, it will report an error
    
    Dim rRng As Range

    'DropdownLookupRange = Worksheets(cDictionayWorksheetName).Range(vDropdownLookupRange)
    
    With Worksheets(cDictionayWorksheetName)
        If Not IsError(Application.match(vFieldValue, .Range(strDropdownLookupRange).Columns(1), 0)) Then
            Set rRng = .Range(.Range(strDropdownLookupRange).Cells(1), .Range(strDropdownLookupRange).Cells(.Range(strDropdownLookupRange).rows.Count).Offset(0, 2))
            VaildateValue = Application.WorksheetFunction.VLookup(vFieldValue, rRng, 3, False)
            'VaildateValue = Application.WorksheetFunction.VLookup(vFieldValue, .Range(strDropdownLookupRange), 3, False)
        Else
            'assign error message only if it is not blank. Otherwise keep the original value.
            If Len(Trim(strErrorMessage)) > 0 Then
                VaildateValue = strErrorMessage
            Else
                VaildateValue = vFieldValue
            End If
            obError.Add strErrorMessage, "Validation of provided dropdown value -'" & vFieldValue & "'", ValidationErrorStatus.UnexpectedValue
            'obError.Add strErrorMessage, "Validation of dropdown value #2 -'" & vFieldValue & "'" 'for test only, to be removed !!!!
        End If
    End With
    
End Function

