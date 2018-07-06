Attribute VB_Name = "mdlDBRelated"

Public Function SelectFieldSettingProfile() As String
    'TODO - create UI to show list of profiles and select one
    SelectFieldSettingProfile = GetConfigValue("FieldSetting_LastLoadedProfile")
End Function

Public Sub LoadFieldSettings()
    
    Dim setting_profile As String
    Dim conn As ADODB.Connection
    Dim rs As ADODB.recordset
    Dim sConnString As String
    Dim DictTitlesRange As Range, c As Range
    Dim connStringConfigName As String
    Dim err_str As String
    
    Const msgTitle = "Loading Field Setting Profile to Master Template"
    
    setting_profile = SelectFieldSettingProfile() 'GetConfigValue("FieldSetting_LastLoadedProfile")
    
    'if no profile selected, exit sub
    If Len(Trim(setting_profile)) = 0 Then Exit Sub
    
    connStringConfigName = GetConfigValue("Conn_Dict_Current")
    
    If Not IsNull(GetConfigValue(connStringConfigName)) Then
        sConnString = GetConfigValue(connStringConfigName)
    Else
        MsgBox "This operation cannot be completed. Vefrify that connection string is provided in the configuration section of the application.", vbCritical, msgTitle
        Exit Sub
    End If
    
    ' Create the Connection and Recordset objects.
    Set conn = New ADODB.Connection
    Set rs = New ADODB.recordset
    
    On Error GoTo err_connection
    'Open the connection and execute.
    conn.Open sConnString
    On Error GoTo 0
    
    On Error GoTo err_recordset
    'fill recordset with data
    Set rs = conn.Execute(Replace(GetConfigValue("FieldSetting_Get_Statement"), "{{profile_name}}", setting_profile))
    On Error GoTo 0
    
    With Worksheets(cSettingsWorksheetName)
        'if returned recordset is not empty load received data for the current field
        'there is an expectation that range for the values form DB starts on the 3rd row under the field name and consists of 3 columns
        If Not rs.EOF Then
            'get the address of the fist cell of the range used on the page
            Set c = .Range(GetConfigValue("FieldSetting_Range_First_Cell"))
            
           'update captions for the newly loaded recordset
            LoadCaptionsForRecordset c, rs
            
            'clean the area of insertion first; it will select all fields actually used on the page; cleaning won't be applied to the first row containing column headers
            '.Range(c.Offset(1, 0).Address, c.offset(.usedrange.rows.count - c.row,.UsedRange.Columns.Count - c.Column).address).address
            .Range(c.Offset(1, 0).Address, c.Offset(.UsedRange.Rows.Count - c.Row, .UsedRange.Columns.Count - c.Column).Address).ClearContents
            
            'copy all information from the recordset to the page (starting with the second row)
            c.Offset(1, 0).CopyFromRecordset rs
            
            MsgBox "Loading of Field Setting profile '" & setting_profile & "' completed successfully!" & vbCrLf & vbCrLf & _
                    "Note: Column headers of the 'RawData' and 'Validated' tabs will be updated accordingly.", vbInformation, msgTitle
            
        Else 'go here if DB does not return any data for the given profile
            'TODO - put handler for this case
            MsgBox "Profile '" & setting_profile & "' was not found or no data was returned for it. Field Setting loading process was aborted!" & vbCrLf & "Please contact your IT admin to resolve the issue.", vbCritical, msgTitle
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
    
    MsgBox err_str, vbCritical, msgTitle
    
    GoTo clean_up
    Exit Sub
    
err_recordset:
    err_str = "Retrieving data from database generated an error. The process was aborted. Please contact your IT admin to resolve the issue." & vbCrLf & vbCrLf & _
                "Detailed error description: " & vbCrLf & Err.Description
    
    MsgBox err_str, vbCritical, msgTitle
    
    GoTo clean_up
    Exit Sub
    
End Sub

Sub LoadCaptionsForRecordset(firstCellOfHeaderRow As Range, rsData As ADODB.recordset)
    Dim i As Integer
    Dim r As Range
    
    With firstCellOfHeaderRow.Worksheet
        
        'Clear existing headers
        Set r = .Range(firstCellOfHeaderRow.Address, .Range(firstCellOfHeaderRow.Offset(0, .UsedRange.Columns.Count - firstCellOfHeaderRow.Column).Address))
        r.ClearContents
        
        'update headers on the page
        For i = 0 To rsData.Fields.Count - 1
            firstCellOfHeaderRow.Offset(0, i).Value = Replace(rsData.Fields(i).Name, "_", " ")
        Next
    End With
End Sub

'TODO: Is this function necessary?
Function ValidatePageHeaders(firstCellOfHeaderRow As Range, rsData As ADODB.recordset) As Boolean
    Dim i As Integer
    
    With firstCellOfHeaderRow.Worksheet
        'compare number of headers on a page and in recordset
        If (.UsedRange.Columns.Count - firstCellOfHeaderRow.Column) <> rsData.Fields.Count Then
            ValidatePageHeaders = False
            Exit Function
        End If
        
        'compare header captions between the page and the recordset
        For i = 0 To rsData.Fields.Count - 1
            If firstCellOfHeaderRow.Offset(0, i).Value <> Replace(rs.Data.Fields(i).Name, "_", " ") Then
                ValidatePageHeaders = False
                Exit Function
            End If
        Next
        
        ValidatePageHeaders = True
    End With
    
End Function

Public Sub LoadDictionaryValues()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.recordset
    Dim sConnString As String
    Dim DictTitlesRange As Range, c As Range
    Dim updatedFields As New StringBuilder
    Dim notUpdatedFields As New StringBuilder
    Dim connStringConfigName As String
    Dim err_str As String
    
    Const msgTitle = "Loading Dictionary to Master Template"
    
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
        MsgBox "This operation cannot be completed. Vefrify that connection string is provided in the configuration section of the application.", vbCritical, msgTitle
        Exit Sub
    End If
 
    ' Create the Connection and Recordset objects.
    Set conn = New ADODB.Connection
    Set rs = New ADODB.recordset
    
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
            MsgBox "Dictionary sheet does not contain any fields suitable for the database sync. Nothing was updated.", vbCritical, msgTitle
        End If
    End With
    
    updatedFields.Delimiter = ", "
    notUpdatedFields.Delimiter = ", "
    
    'display summary message to user
    MsgBox "Sync of dictionary values ran successfully! " & vbCrLf _
            & "**** Updated fields ****" & vbCrLf & Replace(updatedFields.toString, ", ", vbCrLf) & vbCrLf & vbCrLf _
            & "**** Not Updated fields ****" & vbCrLf & Replace(notUpdatedFields.toString, ", ", vbCrLf) _
            , vbInformation, msgTitle
            
clean_up:
    ' Clean up
    If CBool(conn.State And adStateOpen) Then conn.Close
    Set conn = Nothing
    Set rs = Nothing
    
    Exit Sub
    
err_connection:
    err_str = "The database cannot be reached or access denied. Please contact your IT admin to resolve the issue." & vbCrLf & vbCrLf & _
                "Detailed error description: " & vbCrLf & Err.Description
    
    MsgBox err_str, vbCritical, msgTitle
    
    GoTo clean_up
    Exit Sub
    
err_recordset:
    err_str = "Retrieving data from database generated an error. The process was aborted. Please contact your IT admin to resolve the issue." & vbCrLf & vbCrLf & _
                "Detailed error description: " & vbCrLf & Err.Description
    
    MsgBox err_str, vbCritical, msgTitle
    
    GoTo clean_up
    Exit Sub
            
End Sub

