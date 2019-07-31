VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSQLConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public conn As ADODB.Connection

Public cofigConnVariableName As String 'this variable will be assigned with a default value at initiation that can be overwritten
Public errProcessNameTitle As String 'this variable will be assigned with a default value at initiation that can be overwritten
Public errDefaultMsg As String 'this variable will be assigned with a default value at initiation that can be overwritten
Dim errNoConfigValueMsg As String  'this variable will be assigned with a default value at initiation that can be overwritten

Dim sConnString As String


Public Function OpenConnection(Optional strConn As String = "") As ADODB.Connection
    Dim connStringConfigName As String
    Dim err_str As String
    
    If strConn = "" Then
        
        connStringConfigName = GetConfigValue(cofigConnVariableName)
        
        If Not IsNull(GetConfigValue(connStringConfigName)) Then
            sConnString = GetConfigValue(connStringConfigName)
        Else
            MsgBox errNoConfigValueMsg, vbCritical, errProcessNameTitle
            Exit Function
        End If
    Else
        sConnString = strConn
    End If
    
    'Create the Connection and Recordset objects.
    Set conn = New ADODB.Connection
    
    On Error GoTo err_connection
    'Open the connection
    conn.Open sConnString
    On Error GoTo 0
    
    Set OpenConnection = conn
    
    Exit Function
    
clean_up:
    ' Clean up
    If CBool(conn.state And adStateOpen) Then conn.Close
    Set conn = Nothing
    
    Set OpenConnection = Nothing
    
    Exit Function
    
err_connection:
    err_str = errDefaultMsg & vbCrLf & vbCrLf & _
                "Detailed error description: " & vbCrLf & Err.Description
    
    MsgBox err_str, vbCritical, errProcessNameTitle
    
    GoTo clean_up
    
End Function

Private Sub Class_Initialize()
    'assign default values
    cofigConnVariableName = "Conn_Dict_Current"
    errProcessNameTitle = "Database communication" '"Retrieving Field Setting Profiles"
    errNoConfigValueMsg = "This operation cannot be completed. Vefrify that connection string is provided in the configuration section of the application ('Conn_Dict_Current' variable)."
    errDefaultMsg = "The database cannot be reached or access is denied. Please contact your IT admin to resolve the issue."
End Sub

Private Sub Class_Terminate()
    'Clean up
    If Not conn Is Nothing Then
        If CBool(conn.state And adStateOpen) Then conn.Close
        Set conn = Nothing
    End If
End Sub
