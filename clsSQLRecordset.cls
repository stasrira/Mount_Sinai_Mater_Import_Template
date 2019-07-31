VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSQLRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private clsConn As New clsSQLConnection

Public errProcessNameTitle As String 'this variable will be assigned with a default value at initiation that can be overwritten
Public errDefaultMsg As String 'this variable will be assigned with a default value at initiation that can be overwritten

Public Function GetRecordset(sqlStatement As String, Optional strConn As String = "") As ADODB.Recordset
    
    Dim conn As ADODB.Connection
    Dim err_str As String
    
    If Len(Trim(errProcessNameTitle)) > 0 Then clsConn.errProcessNameTitle = errProcessNameTitle
    
    Set conn = clsConn.OpenConnection(strConn)
    If Not conn Is Nothing Then
        Set rs = New ADODB.Recordset
        
        On Error GoTo err_recordset
        'fill recordset with data
        Set rs = conn.Execute(sqlStatement)
        On Error GoTo 0
        
        Set GetRecordset = rs
        
    Else
        Set GetRecordset = Nothing
    End If
    
    Exit Function
    
clean_up:
    ' Clean up
    If Not conn Is Nothing Then
        If CBool(conn.state And adStateOpen) Then conn.Close
        Set conn = Nothing
    End If
    Set clsConn = Nothing
    Set rs = Nothing
    
    Exit Function
    
err_recordset:
    err_str = errDefaultMsg & vbCrLf & vbCrLf & _
                "Detailed error description: " & vbCrLf & Err.Description
    
    MsgBox err_str, vbCritical, errProcessNameTitle
    
    GoTo clean_up
    Exit Function
    
End Function

Private Sub Class_Initialize()
    errProcessNameTitle = "Database communication" '"Retrieving Field Setting Profiles"
    errDefaultMsg = "Retrieving data from database generated an error. The process was aborted. Please contact your IT admin to resolve the issue."
End Sub

Private Sub Class_Terminate()
    Set rs = Nothing
    Set clsConn = Nothing
End Sub
