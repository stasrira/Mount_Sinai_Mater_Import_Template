Attribute VB_Name = "mdlService"
Option Explicit

Sub ExportModules(Optional sPath As String = "", Optional sSubFolder As String = "")
'Dim sPath As String
Dim sFile As String
Dim vbp As Object ' VBProject
Dim comp As Object ' VBComponent
Dim wb As Object 'workbook
Dim i As Integer
    
    If Len(Trim(sPath)) = 0 Then sPath = Application.ActiveWorkbook.Path & "\"
    If Len(Trim(sSubFolder)) > 0 Then sPath = sPath & sSubFolder & "\"
    
    On Error Resume Next
    MkDir sPath
    On Error GoTo 0

    
    For Each wb In Workbooks
        Set vbp = wb.VBProject
        For Each comp In vbp.VBComponents
            sFile = ""
            Select Case comp.Type
            Case 1 ' vbext_ct_StdModule
                sFile = comp.Name & ".bas"
            Case 2 ' vbext_ct_ClassModule
                sFile = comp.Name & ".cls"
            Case 3 ' vbext_ct_MSForm
                sFile = comp.Name & ".frm"
                ' the frx will automatically be exported
            Case 100 ' vbext_ct_Document
                ' thisworkbook or sheet module
                ' this will re-import as a class module
                ' copy code to relevant object module then remove
                If comp.CodeModule.CountOfLines Then
                    sFile = comp.Name & ".cls"
                End If
            End Select
            If Len(sFile) Then
                comp.Export sPath & sFile
                i = i + 1
            End If
        Next
    Next
    
    MsgBox "Exporting of " & CStr(i) & " modules was successfully completed. " & vbCrLf _
            & "All exported files are located in: " & sPath, vbOKOnly, "Exporting code behind to files"
    
End Sub
