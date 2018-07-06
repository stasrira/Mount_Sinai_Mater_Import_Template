Attribute VB_Name = "mdlDBRelated"
Public Sub LoadFieldSettings()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sConnString As String
    Dim DictTitlesRange As Range, c As Range
    Dim connStringConfigName As String
    Dim err_str As String
    
    connStringConfigName = GetConfigValue("Conn_Dict_Current")
    
    If Not IsNull(GetConfigValue(connStringConfigName)) Then
        sConnString = GetConfigValue(connStringConfigName)
    Else
        MsgBox "This operation cannot be completed. Vefrify that connection string is provided in the configuration section of the application.", vbCritical, "Loading Dictionary to Master Template"
        Exit Sub
    End If
    
    ' Create the Connection and Recordset objects.
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    On Error GoTo err_connection
    'Open the connection and execute.
    conn.Open sConnString
    On Error GoTo 0
    
    'TODO - show user a list of posstible profiles and last loaded profile name
    
    On Error GoTo err_recordset
    'fill recordset with data
    Set rs = conn.Execute(Replace(GetConfigValue("FieldSetting_Get_Statement"), "{{profile_name}}", GetConfigValue("FieldSetting_LastLoadedProfile")))
    On Error GoTo 0
    
    With Worksheets(cSettingsWorksheetName)
        'if returned recordset is not empty load received data for the current field
        'there is an expectation that range for the values form DB starts on the 3rd row under the field name and consists of 3 columns
        If Not rs.EOF Then
            'clean the area of insertion first; it will select all fields actually used in the first column (corresponding to the current field header) and offset to 2 columns to the right

            Set c = Range(GetConfigValue("FieldSetting_Range_First_Cell"))

            '.Range(c.Offset(1, 0).Address, c.offset(.usedrange.rows.count - c.row,.UsedRange.Columns.Count - c.Column).address).address
            .Range(c.Offset(1, 0).Address, c.Offset(.UsedRange.Rows.Count - c.Row, .UsedRange.Columns.Count - c.Column).Address).Clear
            
            'copy fresh set of dictionary data
            c.Offset(1, 0).CopyFromRecordset rs
            

        Else 'go here if DB does not return any data for the given profile
            'TODO - put handler for this case
            
        End If
    End With
    
clean_up:
    ' Clean up
    If CBool(conn.State And adStateOpen) Then conn.Close
    Set conn = Nothing
    Set rs = Nothing
    
    Exit Sub
    
err_connection:
    err_str = "The database cannot be reached or access denied. Please contact your IT admin to resolve the issue." & vbCrLf & vbCrLf & _
                "Detailed error description: " & vbCrLf & Err.Description
    
    MsgBox err_str, vbCritical, "Loading Dictionary to Master Template"
    
    GoTo clean_up
    Exit Sub
    
err_recordset:
    err_str = "Retrieving data from database generated an error. The process was aborted. Please contact your IT admin to resolve the issue." & vbCrLf & vbCrLf & _
                "Detailed error description: " & vbCrLf & Err.Description
    
    MsgBox err_str, vbCritical, "Loading Dictionary to Master Template"
    
    GoTo clean_up
    Exit Sub
    
End Sub

Public Sub LoadDictionaryValues()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sConnString As String
    Dim DictTitlesRange As Range, c As Range
    Dim updatedFields As New StringBuilder
    Dim notUpdatedFields As New StringBuilder
    Dim connStringConfigName As String
    Dim err_str As String
    
    'SSQLDBAT015001\TESTINS1
    'localhost\sqlexpress;
    
    ' Create the connection string.
'    sConnString = "Provider=SQLOLEDB;Data Source=localhost\sqlexpress;" & _
'                  "Initial Catalog=dw_motrpac;" & _
'                  "Integrated Security=SSPI;"
    
'    sConnString = "Provider=SQLOLEDB;Data Source=10.160.20.65\TESTINS1;" & _
'                  "Initial Catalog=dw_motrpac;" & _
'                  "Integrated Security=SSPI;"
    
    'connStringConfigName = "Conn_Dict_local" 'local connection string - Provider=SQLOLEDB;Data Source=localhost\sqlexpress; Initial Catalog=dw_motrpac; Integrated Security=SSPI;
    'connStringConfigName = "Conn_Dict_Mount_Sinai" 'Mount Sinai connection string - Provider=SQLOLEDB;Data Source=10.160.20.65\TESTINS1; Initial Catalog=dw_motrpac; Integrated Security=SSPI;
    
    connStringConfigName = GetConfigValue("Conn_Dict_Current")
    
    If Not IsNull(GetConfigValue(connStringConfigName)) Then
        sConnString = GetConfigValue(connStringConfigName)
    Else
        MsgBox "This operation cannot be completed. Vefrify that connection string is provided in the configuration section of the application.", vbCritical, "Loading Dictionary to Master Template"
        Exit Sub
    End If
 
    ' Create the Connection and Recordset objects.
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    With Worksheets(cDictionayWorksheetName)
        'set a range that covers first row with list of cells that have some dictionary info
'        Set DictTitlesRange = .Range("A1:" & Cells(1, .UsedRange.Columns.Count).Address)
        Set DictTitlesRange = .Range(GetConfigValue("Dict_DB_Title_Range_Start_Cell") & ":" & Cells(1, .UsedRange.Columns.Count).Address)
    
        If DictTitlesRange.Cells.Count > 0 Then
            
            On Error GoTo err_connection
            'Open the connection and execute.
            conn.Open sConnString
            
            On Error GoTo 0
            
            'loop through all fields listed as titles and search DB for dictionary info for these fields
            For Each c In DictTitlesRange.Cells
                'Debug.Print c.Address, c.Value
                
                If Len(Trim(c.Value)) > 0 Then
                    'if the field name is not empty, try to get data for it from the DB
                    'Set rs = conn.Execute("SELECT RawValue [Raw Value], iif(DefaultFlag = 1, '1','') [Default Flag], ValidatedValue [Validated Value] FROM dw_fw_dropdown_fields where FieldName = '" & Trim(c.Value) & "'")
                    '"SELECT RawValue [Raw Value], iif(DefaultFlag = 1, '1','') [Default Flag], ValidatedValue [Validated Value] FROM dw_fw_dropdown_fields where FieldName = '{{search_field_name}}'"
                    
                    On Error GoTo err_recordset
                    Set rs = conn.Execute(Replace(GetConfigValue("Dict_DB_Select_Statment"), "{{search_field_name}}", Trim(c.Value)))
                    On Error GoTo 0
                    
                    'if returned recordset is not empty load received data for the current field
                    'there is an expectation that range for the values form DB starts on the 3rd row under the field name and consists of 3 columns
                    If Not rs.EOF Then
                        'clean the area of insertion first; it will select all fields actually used in the first column (corresponding to the current field header) and offset to 2 columns to the right
'                        Debug.Print Range(c.Offset(.Rows.Count - c.Offset(2).Row).End(xlUp).Address).Offset(0, 2).Address
'                        Debug.Print Range(c.Offset(2, 0).Address, Range(c.Offset(.Rows.Count - c.Offset(2).Row).End(xlUp).Address).Offset(0, 2).Address).Address
                        .Range(c.Offset(2, 0).Address, .Range(c.Offset(.Rows.Count - c.Offset(2).Row).End(xlUp).Address).Offset(0, 2).Address).Clear
                        
                        'copy fresh set of dictionary data
                        c.Offset(2, 0).CopyFromRecordset rs
                        
                        'collect name of the successfully updated field
                        updatedFields.Append c.Value
                    Else 'go here if DB does not return any data for the given field
                        'collect name of the not updated field
                        notUpdatedFields.Append c.Value
                    End If
                End If
            Next
        Else
            'No dictionary fields available for update (i.e. Dictionary sheet is empty)
            MsgBox "Dictionary sheet does not contain any fields suitable for the database sync. Nothing was updated.", vbCritical, "Dictionary DB sync"
        End If
    End With
    
    updatedFields.Delimiter = ", "
    notUpdatedFields.Delimiter = ", "
    
    'display summary message to user
    MsgBox "Sync of dictionary values ran successfully! " & vbCrLf _
            & "**** Updated fields ****" & vbCrLf & Replace(updatedFields.toString, ", ", vbCrLf) & vbCrLf & vbCrLf _
            & "**** Not Updated fields ****" & vbCrLf & Replace(notUpdatedFields.toString, ", ", vbCrLf) _
            , vbInformation, "Dictionary DB sync"
            
clean_up:
    ' Clean up
    If CBool(conn.State And adStateOpen) Then conn.Close
    Set conn = Nothing
    Set rs = Nothing
    
    Exit Sub
    
err_connection:
    err_str = "The database cannot be reached or access denied. Please contact your IT admin to resolve the issue." & vbCrLf & vbCrLf & _
                "Detailed error description: " & vbCrLf & Err.Description
    
    MsgBox err_str, vbCritical, "Loading Dictionary to Master Template"
    
    GoTo clean_up
    Exit Sub
    
err_recordset:
    err_str = "Retrieving data from database generated an error. The process was aborted. Please contact your IT admin to resolve the issue." & vbCrLf & vbCrLf & _
                "Detailed error description: " & vbCrLf & Err.Description
    
    MsgBox err_str, vbCritical, "Loading Dictionary to Master Template"
    
    GoTo clean_up
    Exit Sub
            
End Sub

