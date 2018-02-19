VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFieldValidationErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colReportedErrors As New Collection
Public ValidatedCellAddress As String

Public Sub Add(strErrorMessage As String, strProcessID As String, vValidationErrorStatus As ValidationErrorStatus)
    Dim clsErr As New clsFieldValidationError
    
    clsErr.ErrorMessage = strErrorMessage
    clsErr.InProcess = strProcessID
    clsErr.ErrorStatus = vValidationErrorStatus
    
    colReportedErrors.Add clsErr
    Set clsErr = Nothing
    
End Sub

Public Function ErrorCount() As Integer
    ErrorCount = colReportedErrors.Count
End Function

Public Function TotalValidationErrorStatus() As ValidationErrorStatus
    Dim errStatus As ValidationErrorStatus
    Dim iErr As clsFieldValidationError
    Dim countMessages As Integer
    
    errStatus = NoErrors
    
    For Each iErr In colReportedErrors
        If countMessages > 0 Then
            If errStatus <> iErr.ErrorStatus Then
                errStatus = ValidationErrorStatus.CombinationOfErrors
                Exit For
            End If
        Else
            errStatus = iErr.ErrorStatus
        End If
        
        countMessages = countMessages + 1
    Next
    
    TotalValidationErrorStatus = errStatus
End Function

Public Function toString(Optional bShowCellAddress As Boolean = False) As String
    Dim strOutput As String
    Dim iErr As clsFieldValidationError
    Dim countMessages As Integer
    
    Dim sb As New StringBuilder
    
    sb.Append vbCrLf & "VALIDATION ERRORS" & vbCrLf & "-------------------------------- " & vbCrLf
    
    If bShowCellAddress Then
        sb.Append "Validation erorrs for the cell - "
        sb.Append ValidatedCellAddress
        sb.Append vbCrLf & vbCrLf
    End If
    
    For Each iErr In colReportedErrors
        If countMessages > 0 Then
            sb.Append vbCrLf '".................." &
        End If
        sb.Append "Module: "
        sb.Append iErr.InProcess & vbCrLf
        sb.Append "Error: "
        sb.Append iErr.ErrorMessage & vbCrLf
        countMessages = countMessages + 1
    Next
    
    toString = sb.toString
    
    Set sb = Nothing
    
End Function
