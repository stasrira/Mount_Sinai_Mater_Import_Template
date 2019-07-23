Attribute VB_Name = "mdlDBRelated"
Option Explicit

Public Function PopulateFieldSettingProfilesList(ByRef cmb As ComboBox) As Boolean
    Dim lastLoadedProfile As String
    Dim clRs As New clsSQLRecordset
    
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim prof_details As clsFieldSettingProfile
    
    Const msgTitle = "Retrieving Field Setting Profiles"
    
    lastLoadedProfile = GetConfigValue("FieldSetting_LastLoadedProfile")
    
    clRs.errProcessNameTitle = msgTitle
    Set rs = clRs.GetRecordset(GetConfigValue("FieldSetting_Get_Profiles"))
    
    If Not rs Is Nothing Then  'if returned recordset is an object
        
        If Not rs.EOF Then 'if returned recordset is not empty load received data
            
            i = 0
            dictProfiles.RemoveAll
            cmb.Clear
            
            While Not rs.EOF
                Set prof_details = New clsFieldSettingProfile
                
                prof_details.Name = rs.Fields(1).value
                prof_details.ID = rs.Fields(0).value
                prof_details.Description = rs.Fields(2).value
                prof_details.Owner = rs.Fields(3).value
                prof_details.Created = rs.Fields(4).value
                
                dictProfiles.Add i, prof_details
                cmb.AddItem prof_details.Name 'rs.Fields(1).Value
                
                'select previously selected items as a default choise
                If prof_details.Name = lastLoadedProfile Then
                    cmb.ListIndex = i
                End If
                
                i = i + 1
                rs.MoveNext
            Wend
            
            PopulateFieldSettingProfilesList = True
        Else
            'no profiles were returned
            GoTo no_profiles
        End If
    Else
        'no profiles were returned
no_profiles:
        MsgBox "No profiles were returned from the database.", vbExclamation, msgTitle
        
        PopulateFieldSettingProfilesList = False
    End If
    
    Set clRs = Nothing
    
End Function

Public Sub LoadFieldSettings()
    
    Dim setting_profile As Integer
    Dim clRs As New clsSQLRecordset
    Dim rs As ADODB.Recordset
    Dim c As Range
    'Dim connStringConfigName As String
    'Dim err_str As String
    
    Const msgTitle = "Loading Field Setting Profile to Master Template"
    
    setting_profile = SelectFieldSettingProfile() 'GetConfigValue("FieldSetting_LastLoadedProfile")
    
    'if no profile selected, exit sub
    If setting_profile < 0 Then Exit Sub
    
    clRs.errProcessNameTitle = msgTitle
    Set rs = clRs.GetRecordset(Replace(GetConfigValue("FieldSetting_Get_Statement"), "{{profile_id}}", dictProfiles(setting_profile).ID))
    
    With Worksheets(cSettingsWorksheetName)
        
        If Not rs Is Nothing Then 'if returned recordset is an object
            
            If Not rs.EOF Then 'if returned recordset is not empty load received data
                'get the address of the fist cell of the range used on the page
                Set c = .Range(GetConfigValue("FieldSetting_Range_First_Cell"))
                
               'update captions for the newly loaded recordset
                LoadCaptionsForRecordset c, rs
                
                'clean the area of insertion first; it will select all fields actually used on the page; cleaning won't be applied to the first row containing column headers
                .Range(c.Offset(1, 0).Address, c.Offset(.UsedRange.Rows.Count - c.Row, .UsedRange.Columns.Count - c.Column).Address).ClearContents
                
                'copy all information from the recordset to the page (starting with the second row)
                c.Offset(1, 0).CopyFromRecordset rs
                
                'save name of the last loaded profile
                If SetConfigValue("FieldSetting_LastLoadedProfile", dictProfiles(setting_profile).Name) <= 0 Then
                    'TODO - make a decision what to do if the last loaded profile was not saved to the config section
                End If
                
                MsgBox "Loading of Field Setting profile '" & dictProfiles(setting_profile).Name & "' completed successfully!" & vbCrLf & vbCrLf & _
                        "Note: Column headers of the 'RawData' and 'Validated' tabs will be updated accordingly.", vbInformation, msgTitle
                
            Else 'go here if DB does not return any data for the given profile
                'MsgBox "Profile '" & dictProfiles(setting_profile) & "' was not found or no data was returned for it. Field Setting loading process was aborted!" & vbCrLf & "Please contact your IT admin to resolve the issue.", vbCritical, msgTitle
                GoTo empty_recordset
            End If
        Else
empty_recordset:
            MsgBox "Profile '" & dictProfiles(setting_profile).Name & "' was not found or no data was returned for it. Field Setting loading process was aborted!" & vbCrLf & "Please contact your IT admin to resolve the issue.", vbCritical, msgTitle
        End If
    End With
    
    Set clRs = Nothing
    
End Sub

Public Sub LoadDictionaryValues()
    Dim clRs As New clsSQLRecordset
    Dim rs As ADODB.Recordset
    Dim DictTitlesRange As Range, c As Range
    Dim targetRangeStart As Range, targetRangeEnd As Range
    Dim updatedFields As New StringBuilder
    Dim notUpdatedFields As New StringBuilder
    
    Const msgTitle = "Loading Dictionary to Master Template"
    
    'SSQLDBAT015001\TESTINS1
    'localhost\sqlexpress;
    
    ' Create the connection string.
    'connStringConfigName = "Conn_Dict_local" 'local connection string - Provider=SQLOLEDB;Data Source=localhost\sqlexpress; Initial Catalog=dw_motrpac; Integrated Security=SSPI;
    'connStringConfigName = "Conn_Dict_Mount_Sinai" 'Mount Sinai connection string - Provider=SQLOLEDB;Data Source=10.160.20.65\TESTINS1; Initial Catalog=dw_motrpac; Integrated Security=SSPI;
    
    clRs.errProcessNameTitle = msgTitle
    
    With Worksheets(cDictionayWorksheetName)
        'set a range that covers first row with list of cells that have some dictionary info
'        Set DictTitlesRange = .Range("A1:" & Cells(1, .UsedRange.Columns.Count).Address)
        Set DictTitlesRange = .Range(GetConfigValue("Dict_DB_Title_Range_Start_Cell") & ":" & Cells(1, .UsedRange.Columns.Count).Address)
    
        If DictTitlesRange.Cells.Count > 0 Then
            
            'loop through all fields listed as titles and search DB for dictionary info for these fields
            For Each c In DictTitlesRange.Cells
                'Debug.Print c.Address, c.Value
                
                If Len(Trim(c.value)) > 0 Then
                    'if the field name is not empty, try to get data for it from the DB
                    
                    Set rs = clRs.GetRecordset(Replace(GetConfigValue("Dict_DB_Select_Statment"), "{{search_field_name}}", Trim(c.value)))
                    
                    If rs Is Nothing Then Exit Sub 'exit sub if recordset is failed to instantiate
                    
                    'if returned recordset is not empty load received data for the current field
                    'there is an expectation that range for the values form DB starts on the 3rd row under the field name and consists of 3 columns
                    If Not rs.EOF Then
                        'clean the area of insertion first; it will select all fields actually used in the first column (corresponding to the current field header) and offset to 2 columns to the right
'                        Debug.Print Range(c.Offset(.Rows.Count - c.Offset(2).Row).End(xlUp).Address).Offset(0, 2).Address
'                        Debug.Print Range(c.Offset(2, 0).Address, Range(c.Offset(.Rows.Count - c.Offset(2).Row).End(xlUp).Address).Offset(0, 2).Address).Address
                        Set targetRangeStart = c.Offset(2, 0)
                        Set targetRangeEnd = .Range(c.Offset(.Rows.Count - c.Offset(2).Row).End(xlUp).Address).Offset(0, 2)
                        If targetRangeEnd.Row < targetRangeStart.Row Then
                            Set targetRangeEnd = targetRangeStart
                        End If
                        .Range(targetRangeStart.Address, targetRangeEnd.Address).Clear
                        '.Range(c.Offset(2, 0).Address, .Range(c.Offset(.Rows.Count - c.Offset(2).Row).End(xlUp).Address).Offset(0, 2).Address).Clear
                        
                        'copy fresh set of dictionary data
                        c.Offset(2, 0).CopyFromRecordset rs
                        
                        'collect name of the successfully updated field
                        updatedFields.Append c.value
                    Else 'go here if DB does not return any data for the given field
                        'collect name of the not updated field
                        notUpdatedFields.Append c.value
                    End If
                End If
            Next
        Else
            'No dictionary fields available for update (i.e. Dictionary sheet is empty)
            MsgBox "Dictionary sheet does not contain any fields suitable for the database sync. Nothing was updated.", vbCritical, msgTitle
            Exit Sub
        End If
    End With
    
    updatedFields.Delimiter = ", "
    notUpdatedFields.Delimiter = ", "
    
    'display summary message to user
    MsgBox "Sync of dictionary values ran successfully! " & vbCrLf _
            & "**** Updated fields ****" & vbCrLf & Replace(updatedFields.toString, ", ", vbCrLf) & vbCrLf & vbCrLf _
            & "**** Not Updated fields ****" & vbCrLf & Replace(notUpdatedFields.toString, ", ", vbCrLf) _
            , vbInformation, msgTitle
    
    Set clRs = Nothing
            
End Sub

Sub LoadCaptionsForRecordset(firstCellOfHeaderRow As Range, rsData As ADODB.Recordset)
    Dim i As Integer
    Dim r As Range
    
    With firstCellOfHeaderRow.Worksheet
        
        'Clear existing headers
        Set r = .Range(firstCellOfHeaderRow.Address, .Range(firstCellOfHeaderRow.Offset(0, .UsedRange.Columns.Count - firstCellOfHeaderRow.Column).Address))
        r.ClearContents
        
        'update headers on the page
        For i = 0 To rsData.Fields.Count - 1
            firstCellOfHeaderRow.Offset(0, i).value = Replace(rsData.Fields(i).Name, "_", " ")
        Next
    End With
End Sub

'TODO: Is this function necessary? Currently it is not in use.
'Function ValidatePageHeaders(firstCellOfHeaderRow As Range, rsData As ADODB.Recordset) As Boolean
'    Dim i As Integer
'
'    With firstCellOfHeaderRow.Worksheet
'        'compare number of headers on a page and in recordset
'        If (.UsedRange.Columns.Count - firstCellOfHeaderRow.Column) <> rsData.Fields.Count Then
'            ValidatePageHeaders = False
'            Exit Function
'        End If
'
'        'compare header captions between the page and the recordset
'        For i = 0 To rsData.Fields.Count - 1
'            If firstCellOfHeaderRow.Offset(0, i).value <> Replace(rs.data.Fields(i).Name, "_", " ") Then
'                ValidatePageHeaders = False
'                Exit Function
'            End If
'        Next
'
'        ValidatePageHeaders = True
'    End With
'
'End Function

Public Function SubmitManifests(study_id As String, Optional AvoidWarningMessage As Boolean = False)
    Const msgTitle = "Submitting Manifest IDs"
'    Const fieldName = "MT_ManifestID"
'    Const settingName = "study_id"
    
    Dim mics As String 'study_id As String
    Dim manifest_ids As String
    Dim user_id As String
    
    Dim sb As New StringBuilder
    Dim bSubmitStatus As Boolean
        
    Dim clRs As New clsSQLRecordset
    Dim rs As ADODB.Recordset
    Dim stored_proc As String
    
    Dim iResponse As Integer
    
    If Not AvoidWarningMessage Then
        iResponse = MsgBox("The system is about to start process of submitting Manifest IDs to the database." & _
                            vbCrLf & vbCrLf & "Do you want to proceed? If not, click 'Cancel'.", _
                            vbOKCancel, msgTitle)
    Else
        iResponse = vbOK
    End If
    
    If iResponse <> vbOK Then
        Exit Function
    End If
    
    'sb variable will contain the combined message that will be displayed at the end of this function.
    sb.Append "MANIFEST SUBMISSION STATUS:" & vbCrLf & "-------------------------------"
    
    'get comma delimited list of Manifest IDs
    manifest_ids = Get_DisticntValuesFromField(fieldName)
    
    'this assigment was moved to "ExportValidateSheet" function of the mdlGeneric.
'    'get misc settings for the ManifestID field
'    study_id = Get_MiscSettingValue(fieldName, settingName)
    
    'get current user id
    user_id = Environ("USERDOMAIN") + "\" + Environ("Username")
    
    'prepare stored procedure name to be called;
    stored_proc = GetConfigValue("Manifest_Add_New")
    
    'insert parameters to the procedure call
    stored_proc = Replace(stored_proc, "{{study_id}}", study_id)
    stored_proc = Replace(stored_proc, "{{manifest_ids}}", manifest_ids)
    stored_proc = Replace(stored_proc, "{{user_id}}", user_id)
    
    If Len(manifest_ids) > 0 Then
    
        'call stored procedure
        clRs.errProcessNameTitle = msgTitle
        Set rs = clRs.GetRecordset(stored_proc)
        
        If Not rs Is Nothing Then  'if returned recordset is an object
            
            If Not rs.EOF Then 'if returned recordset is not empty load received data

                While Not rs.EOF
    
                    sb.Append vbCrLf & vbCrLf & "Manifest: " & rs.Fields(0).value & vbCrLf
                    sb.Append "Status:   " & IIf(rs.Fields(1).value > 0, "OK", "Error") & vbCrLf
                    sb.Append "Details:  " & rs.Fields(2).value & vbCrLf
                   
                    rs.MoveNext
                Wend
                
                bSubmitStatus = True
            Else
                'no output was returned
                sb.Append vbCrLf & vbCrLf & "Something went wrong and the submission was not completed properly! No output was received from the database."
            End If
        Else
            'no output was returned
            sb.Append vbCrLf & vbCrLf & "Something went wrong and the submission was not completed properly! No output was received from the database."
        End If
        
        Set clRs = Nothing
    
    Else 'if no manifests were found
        sb.Append vbCrLf & vbCrLf & "No manifests were found on the Validated sheet available for submission."
    End If
    
    MsgBox sb.toString, IIf(bSubmitStatus, vbInformation, vbExclamation), msgTitle

End Function



