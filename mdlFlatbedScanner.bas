Attribute VB_Name = "mdlFlatbedScanner"
Option Explicit

'Moved to Config Tab
'Public Const FBS_HostIP = "10.90.121.149"
'Public Const FBS_ScpiPort = 2500
'Public Const Winsock_Ver = 512 'Version 1.1 (1*256 + 1) = 257; version 2.0 (2*256 + 0) = 512
'Public Const ReadScanDelay = 5000 'this is a delay to allow scanner to perfomr the actual scan, while the execution of the macro is on hold

'TODO - the following constants should be coming from the config file
Const Scanner_OKStatus = "OK"
Const Scanner_DataReady = "dataready"
Const Scanner_GetScanResults = "get scanresult"
Const Scanner_ReadRackID = "read rackid"
Const Scanner_ScanBox = "scan box"
Const Scanner_State = "state"
Const Scanner_LineEnd_Header = ",Line End,"
Const Scanner_LineEnd_Value = ",end text,"

'TODO - the following constants should be coming from the config file
Const msgStartScanning = "Scanning process is about to start. Please make sure that Tracxer application is running on Flatbed scanner PC."
Const msgReScanNote = "Please redo the scanning process."
Const msgVerifyTracxerNote = "Resolution Tip: please verify that Tracxer (scanner application) is running on the flatbed scanner computer."
Const msgScanCompletedStatus = "Scan Process was completed!"
Const msgScanCompletedAlert = "Scan request was successfully completed!"
Const msgScanFailedStatus = "Scanning process failed..."
Const msgDataReadinessAlert = "Scanner refused to confirm data readiness. Scan request was not completed."
Const msgReadingBoxStatus = "Scanning the box..."
Const msgReadingRackIDStatus = "Reading RackID..."
Const msgScanStartedStatus = "Scan process has started"
Const msgErrorOccuredStatus = "Error has occured. Scan Process was aborted..."
Const msgBoxTitle = "Flatbed Scanner"

'Moved to config tab
'Public Const ScanResultsTargetLocation = "A1"
'Public Const ScanResultsColumnsUsed = 4
'Public Const ScanResultsStatusLocation = "F2"

Enum Socket_command_type
    Start_sc = 1
    Open_sc = 2
    Send_sc = 3
    Recv_sc = 4
    Close_sc = 5
    End_sc = 6
End Enum

Type TCPIP_Response
    status As String
    command As String
    value As String
    error As String
End Type

Function get_hostname()
    
    'TODO - read IP address from config file
    get_hostname = GetConfigValue("FBS_HostIP") 'originally used const - FBS_HostIP
    
End Function

Sub Report_Status(strStatus As String, Optional vbColor As Long = vbBlack, Optional vbBackGroundColor As Long = vbYellow)
    Dim status_range As Range
    
    'TODO - store this address in the config section/file
    Set status_range = Worksheets(cFlatbedScansWorksheetName).Range(GetConfigValue("FBS_ScanResultsStatusLocation")) 'originally used const - ScanResultsStatusLocation
    
    status_range.value = strStatus
    status_range.Font.Color = vbColor
    status_range.Interior.Color = vbBackGroundColor
End Sub

Sub Clean_Scan_Results()
    Dim r As Range
    Dim ScanResultsTargetLocation As String, ScanResultsColumnsUsed As Integer
    
    ScanResultsTargetLocation = GetConfigValue("FBS_ScanResultsTargetLocation")
    ScanResultsColumnsUsed = CInt(GetConfigValue("FBS_ScanResultsColumnsUsed"))
    
    With Worksheets(cFlatbedScansWorksheetName) 'Sheets(2)
        'clear main scan area; identify number of actually used rows and use number of columns specified in ScanResultsColumnsUsed const
        Set r = .Range(ScanResultsTargetLocation & ":" & .Range(ScanResultsTargetLocation).Offset(.rows.Count - .Range(ScanResultsTargetLocation).row).End(xlUp).Offset(0, ScanResultsColumnsUsed - 1).Address)
        r.Clear
        
        Clean_Scan_Status
    End With
End Sub

Sub Clean_Scan_Status()
    Worksheets(cFlatbedScansWorksheetName).Range(GetConfigValue("FBS_ScanResultsStatusLocation")).Clear
End Sub

Sub FBS_Scan()
    
    Dim x As Long
    Dim recvBuf As String '* 4096 '*1024
    Dim rackId As String, strScnState As String
    Dim finalMsg As String, finalMsgStatus As VbMsgBoxStyle
    'Dim scanStatus As ScanReadStatus
    Dim scanStats As String
    Dim response As TCPIP_Response
    
    If MsgBox(msgStartScanning & vbCrLf & vbCrLf & _
        "Click ""OK"" to continue or ""Cancel"" to exit." _
        , vbOKCancel, msgBoxTitle) = vbCancel Then Exit Sub
    
    Application.EnableEvents = False
    'clean previou scan results
    Clean_Scan_Results
    Application.EnableEvents = True
    
    Report_Status msgScanStartedStatus
    
    If ValidateSocketOperation(StartIt, Start_sc) < 0 Then Exit Sub
'    If x < 0 Then
'        'TODO - report issue with starting the socket
'    End If
    
'    Call get_hostname
    
    If ValidateSocketOperation(OpenSocket(get_hostname(), GetConfigValue("FBS_ScpiPort")), Open_sc) < 0 Then Exit Sub 'Originally used const - FBS_ScpiPort
    
    Report_Status msgReadingRackIDStatus
    
    'Start scanning process
    'HELP: possible flatbed scanner commands: close,get layout,read rackid,get scanresult,help,scan box,scan tube,set layout,show layouts,state
    
    'Read Rackid from the scanner
    If ValidateSocketOperation(SendCommand(Scanner_ReadRackID), Send_sc, "Current SendCommand: " & Scanner_ReadRackID & "." & vbCrLf & msgVerifyTracxerNote) < 0 Then Exit Sub
'    x = SendCommand("scan box")
    Sleep (500) 'wait until scan is completed
    If ValidateSocketOperation(RecvAscii(recvBuf, 1024), Recv_sc, "Current SendCommand: " & Scanner_ReadRackID & ".") < 0 Then Exit Sub
'    x = RecvAscii(recvBuf, 1024) '1024
    
    'example of expected recvBuf: OK read rackid 4001206217
    'rackId = Read_Simple_Scanner_Output(recvBuf)
    response = ReadResponse(recvBuf, Scanner_ReadRackID)
    If response.status <> Scanner_OKStatus Then
        'TODO report an error and exit sub
        finalMsg = "There was an issue with reading the rack (box) id!" & vbCrLf & _
            "Here is the scanner's output (up to 100 characters): " & vbCrLf & Left(recvBuf, 100) & _
            IIf(Len(recvBuf) > 100, "...", "")
        GoTo failed_scan
    End If
    
    rackId = response.value
    
    Report_Status msgReadingBoxStatus
    
    If ValidateSocketOperation(SendCommand(Scanner_ScanBox), Send_sc, "Current SendCommand: " & Scanner_ScanBox & ".") < 0 Then Exit Sub
'    x = SendCommand("scan box")
    Sleep (CLng(GetConfigValue("FBS_ReadScanDelay"))) 'wait until scan is completed ' originally used const - ReadScanDelay
    If ValidateSocketOperation(RecvAscii(recvBuf, 1024), Recv_sc, "Current SendCommand: " & Scanner_ScanBox & ".") < 0 Then Exit Sub
'    x = RecvAscii(recvBuf, 1024) '1024
    
    response = ReadResponse(recvBuf, Scanner_ScanBox)
    If response.status <> Scanner_OKStatus Then
        'TODO: report an error and exit sub
        finalMsg = "There was an issue with performing scanning operation!" & vbCrLf & _
            "Here is the scanner's output (up to 100 characters): " & vbCrLf & Left(recvBuf, 100) & _
            IIf(Len(recvBuf) > 100, "...", "")
        GoTo failed_scan
    End If
    
    If ValidateSocketOperation(SendCommand(Scanner_State), Send_sc, "Current SendCommand: " & Scanner_State & ".") < 0 Then Exit Sub
'    x = SendCommand("state")
    If ValidateSocketOperation(RecvAscii(recvBuf, 1024), Recv_sc, "Current SendCommand: " & Scanner_State & ".") < 0 Then Exit Sub
'    x = RecvAscii(recvBuf, 1024) '1024
    
    response = ReadResponse(recvBuf, Scanner_State)
    If response.status <> Scanner_OKStatus Then
        'report an error and exit sub
        finalMsg = "There was an issue with receiving state of scan operation!" & vbCrLf & _
                "Here is the scanner's output (up to 100 characters): " & vbCrLf & Left(recvBuf, 100) & _
                IIf(Len(recvBuf) > 100, "...", "")
            GoTo failed_scan
    End If
    
    'example of expected recvBuf: OK state dataready
    'strScnState = Read_Simple_Scanner_Output(recvBuf)
    strScnState = response.value
    
    If strScnState = Scanner_DataReady Then
    
        If ValidateSocketOperation(SendCommand(Scanner_GetScanResults), Send_sc, "Current SendCommand: " & Scanner_GetScanResults & ".") < 0 Then Exit Sub
    '    'x = SendCommand("get scanresult")
    '    x = SendCommand(Scanner_GetScanResults)
        If ValidateSocketOperation(RecvAscii(recvBuf, 10240), Recv_sc, "Current SendCommand: " & Scanner_GetScanResults & ".") < 0 Then Exit Sub
    '    x = RecvAscii(recvBuf, 10240)
        
        response = ReadResponse(recvBuf, Scanner_GetScanResults)
        If response.status <> Scanner_OKStatus Then
            'report an error and exit sub
            finalMsg = "There was an issue with receiving scan results!" & vbCrLf & _
                "Here is the scanner's output (up to 100 characters): " & vbCrLf & Left(recvBuf, 100) & _
                IIf(Len(recvBuf) > 100, "...", "")
            GoTo failed_scan
        End If
        
        scanStats = PostScanResultsToPage(response.value, rackId)
        
        'read scan results and post them to a page
        ''scanStatus = Read_Scan_Results(recvBuf, rackId)
        
        
        If Len(scanStats) > 0 Then
            finalMsg = msgScanCompletedAlert & vbCrLf & vbCrLf & scanStats
            finalMsgStatus = vbInformation
            
            Report_Status msgScanCompletedStatus, , vbGreen
        Else
            finalMsg = "Parsing of the scanned information was not properly completed."
            GoTo failed_scan

        End If
    Else
        finalMsg = msgDataReadinessAlert & vbCrLf & msgReScanNote
        
failed_scan:
        finalMsgStatus = vbCritical
        Report_Status msgScanFailedStatus, , vbRed
    End If
    
    
    
    If ValidateSocketOperation(CloseConnection, Close_sc) < 0 Then Exit Sub
'    Call CloseConnection
    If ValidateSocketOperation(EndIt, End_sc) < 0 Then Exit Sub
'    Call EndIt
    
    MsgBox finalMsg, finalMsgStatus, msgBoxTitle
    
    Clean_Scan_Status
    
'Output examples
'    OK state idle
'    Connected
'
'    OK show layouts 16x24,8x12,6x8,4x6,3x4
'    Connected
'
'    OK get layout 8x12
'    Connected
'
'    OK Help
'    close,get layout,read rackid,get scanresult,help,scan box,scan tube,set layout,show layouts,state
'    Connected
    
End Sub

Function ValidateSocketOperation(ret_status As Long, cmd_type As Socket_command_type, Optional comment As String)
    Const constAdditMsg = vbCrLf & "Try to close and reopen the the Excel file you are working from (save any changes you did before closing)."
    
    Dim output_val As Integer
    Dim message As String, extra_com As String
    Dim bCloseConnection As Boolean, bEndIt As Boolean
    Dim x As Integer
    
    'if status is < 0, then an error has occured. Below code will implement handler for each case based on the type of the command that was called
    If ret_status < 0 Then
            
        output_val = -1
        
        If Len(Trim(comment)) > 0 Then extra_com = vbCrLf & comment
        
        Select Case cmd_type
            Case Start_sc
                message = "An error has occurred during starting the socket control!" & constAdditMsg
                
            Case Open_sc
                message = "An error has occurred during opening the socket control!" & constAdditMsg
                bEndIt = True
                
            Case Send_sc
                message = "An error has occurred during sending a request to the flatbed scanner!"
                bCloseConnection = True
                bEndIt = True
                
            Case Recv_sc
                message = "An error has occurred during receiving a response from the flatbed scanner!"
                bCloseConnection = True
                bEndIt = True
                
            Case Close_sc
                message = "An error has occurred during closing the socket control!" & constAdditMsg
                bEndIt = True
                
            Case End_sc
                message = "An error has occurred during ending the socket control activity!" & constAdditMsg
                
        End Select
        
        If bCloseConnection Then x = CloseConnection()
        If bEndIt Then x = EndIt()
        
        Report_Status msgErrorOccuredStatus, , vbRed
        
        MsgBox message & extra_com, vbCritical, msgBoxTitle
        
        Clean_Scan_Status
    End If
    
    ValidateSocketOperation = output_val
End Function

Function ReadResponse(input_data As String, orig_command As String) As TCPIP_Response
    Const const_command_replacement = "{{command}}"
    Dim input_arr() As String
    Dim output_val As String, returned_line As String
    Dim i As Integer
    Dim out As TCPIP_Response
    
    'Last not empty line, returned by the scanner, should be used to interpretate returned information. Other lines might contain some information not returned previously (i.e. "connected" status) and should be ignored.
    'The proper command line has the following structure:
    'The output is space delimited. First element is the status, second is the original command and the last element is the returned valule. In case if the original command has a space inside
    'that space should be ignored (i.e. "read rackid")
    'Examples:
    'OK read rackid 4001206217
    'OK state dataready
    
    If Len(input_data) > 0 Then
        'split returned information by the end of line characters
        input_arr = Split(input_data, Chr(13) & Chr(10))
        'loop through starting from the end to the first not empty value
        For i = UBound(input_arr) To LBound(input_arr) Step -1
            If Len(Trim(input_arr(i))) > 0 Then
                'save found returned response in the variable
                returned_line = input_arr(i)
                Exit For
            End If
        Next
        
        'replace the original command found in the response with command_replacement const ({{command}}) to avoid "inner command" spaces when splitting the string; it will replace only the first match.
        If InStr(returned_line, orig_command) > 0 Then
            returned_line = Replace(returned_line, orig_command, const_command_replacement, , 1)
            out.command = orig_command
        Else
            'Raise error that orginal command was not found
            out.error = "Expected command name (" & orig_command & ") was not found in the returned value!"
            GoTo exit_lab
        End If
        
        'split the returned response by " "
        input_arr = Split(returned_line, " ")
        
        If UBound(input_arr) > 0 Then 'verify that generated array has members
            For i = 0 To UBound(input_arr)
                If i = 0 Then
                    out.status = input_arr(i) 'save status to the output variable
                ElseIf i = 1 Then
                    out.command = out.command 'nothing should be done here
                Else
                    'concatenate all other array values into a single string; this will be a returned value
                    If Len(out.value) > 0 Then
                        out.value = out.value + " " + input_arr(i)
                    Else
                        out.value = out.value + input_arr(i) 'avoid an empty space for the first added value
                    End If
                End If
            Next
        Else
            'Raise error - nothing can be parsed out of the returned value
            out.error = "No data can be parsed out of returned value!"
            out.value = returned_line
            GoTo exit_lab
        End If
                
    Else
        'Raise error - NO Data Returned
        out.error = "No data was returned!"
        GoTo exit_lab
    End If
    
exit_lab:
    ReadResponse = out
End Function

Function PostScanResultsToPage(scan_result As String, rackId_val As String) As String
    Dim input_arr() As String
    Dim dest_range As Range, cur_range As Range
    Dim scanOK_cnt As Integer, scanNotOk_cnt As Integer
    Dim out_result As String, i As Integer
    
    'replace "Line End" (caption of the last column with vbCrLf
    scan_result = Replace(scan_result, Scanner_LineEnd_Header, vbCrLf)
    'replace "end text" (values of the last column with vbCrLf, also insert the rackId as a value of the column preceding the last one
    scan_result = Replace(scan_result, "," & Scanner_LineEnd_Value, "," & rackId_val & vbCrLf)
    
    Application.ScreenUpdating = False
    
    Set dest_range = Worksheets(cFlatbedScansWorksheetName).Range(GetConfigValue("FBS_ScanResultsTargetLocation"))
    
    'scan_result = "OK get scanresult Position,Tube ID,Status,Rack ID,Line End,A01,8019487649,OK,,end text,B01,8019487062,OK,,end text,C01,8019486823,OK,,end text,D01,8019487481,OK,,end text,E01,8019487593,OK,,end text,F01,8019486904,OK,,end text,G01,8019487421,OK,,end text,H01,8019487665,OK,,end text,A02,8019487356,OK,,end text,B02,8019487052,OK,,end text,C02,8019487343,OK,,end text,D02,8019487647,OK,,end text,E02,8019486896,OK,,end text,F02,8019487690,OK,,end text,G02,8019487442,OK,,end text,H02,8019487516,OK,,end text,A03,8019487112,OK,,end text,B03,8019487473,OK,,end text,C03,8019487315,OK,,end text,D03,8019486812,OK,,end text,E03,8019487496,OK,,end text,F03,8019487424,OK,,end text,G03,8019486925,OK,,end text,H03,8019487709,OK,,end text,A04,8019487368,OK,,end text,B04,8019487650,OK,,end text,C04,8019487127,OK,,end text,D04,8019487326,OK,,end text,E04,8019487575,OK,,end text,F04,8019487358,OK,,end text,G04,8019486918,OK,,end text,H04,8019487357,OK,,end text,"
    input_arr = Split(scan_result, vbCrLf)
    
    'loop through array of results and assign its values to the first cells of rows going down from the given start point (dest_range)
    For i = LBound(input_arr) To UBound(input_arr) - 1 'last element of the array is empty, since the very last row has a delimiter at the end
        'Debug.Print input_arr(i)
        
        Set cur_range = dest_range.Offset(i, 0)
        
        cur_range.value = input_arr(i)
        
        Application.DisplayAlerts = False 'this will prevent warning message (i.e. in case if some existing data presented in the columns being populated)
        
        'split text (located in the first cell of a row) to columns located on the same row
        If (Len(cur_range.value) > 0) Then 'proceed only if the value to be parsed is not blank
            cur_range.TextToColumns _
                Destination:=cur_range, _
                DataType:=xlDelimited, _
                TextQualifier:=xlTextQualifierNone, _
                ConsecutiveDelimiter:=False, _
                Tab:=False, _
                Semicolon:=False, _
                Comma:=True, _
                Space:=False, _
                Other:=False, _
                OtherChar:="", _
                FieldInfo:=Array( _
                Array(1, xlTextFormat), Array(2, xlTextFormat), _
                Array(3, xlTextFormat), Array(4, xlTextFormat))
                
            'if this is not the 1st (columns header) row, check reported status
            If i > 0 Then
                'Update counts of scanned and not scanned positions
                If cur_range.Offset(0, 2).Value2 = Scanner_OKStatus Then
                    scanOK_cnt = scanOK_cnt + 1
                Else
                    scanNotOk_cnt = scanNotOk_cnt + 1
                End If
            End If
        End If
        
        Application.DisplayAlerts = True ' revernt back supression of the Display Alerts
    Next
    
    out_result = "Number of successfully scanned positions: " & CStr(scanOK_cnt) & vbCrLf & _
            "Number of not scanned positions: " & CStr(scanNotOk_cnt)
    
    Application.ScreenUpdating = True
    
    PostScanResultsToPage = out_result
    
End Function

Function Load_FBS_Scan_Results()
    ImportFBSFile
End Function

Public Sub ImportFBSFile()
    Dim iResponse As Integer
    Dim importFileOutcome As Boolean
    Dim strFileToOpen As String
    Dim tStart As Date, tEnd As Date
    
    'confirm if user want to proceed.
    iResponse = MsgBox("The system is about to start importing Flatbed scanner result file to the 'FlatbedScans' tab. " _
                & vbCrLf & "- Any currently existing data will be overwritten wih the new data!" _
                & vbCrLf & vbCrLf & "Do you want to proceed? If not, click 'Cancel'." & vbCrLf & vbCrLf _
                & "Note: " _
                & vbCrLf & "- This process might take prolonged time, depending on the number of rows being imported. " _
                & vbCrLf & "- Some screen flickering might occur during the process. ", _
                vbOKCancel + vbInformation, "Master Check-in Template - FBS scan import")
    
    If iResponse <> vbOK Then
        'exit sub based on user's response
        Exit Sub
    End If
    
    'select a file to be loaded
    strFileToOpen = Application.GetOpenFilename _
        (Title:="Please choose a Flatbed scan file to open", _
        FileFilter:="CSV Files (.csv), *.csv")
    
    If strFileToOpen = "False" Then
        Exit Sub
    End If
    
    tStart = Now()
    'Debug.Print (tStart)
    
    importFileOutcome = ImportFile(strFileToOpen, Worksheets(cFlatbedScansWorksheetName))
    
    Application.CalculateFullRebuild
    
    If importFileOutcome Then
        'proceed here only if the file was successfully loaded
        Worksheets(cFlatbedScansWorksheetName).Activate 'bring focus to the "logs" tab
        Worksheets(cFlatbedScansWorksheetName).Cells(1, 1).Activate 'bring focus to the first cell on the sheet
        
        tEnd = Now()
        'Debug.Print (tEnd)
        
        MsgBox "FBS file loading was completed successfully." _
            & vbCrLf & "Execution time: " & getTimeLength(tStart, tEnd) _
            , vbInformation, "Master Check-in Template - FBS scan import"
    End If
    
End Sub

Private Function ImportFile(strFileToOpen As String, ws_target As Worksheet) As Boolean ', file_type_to_open As String
    'Dim strFileToOpen As String
    
    On Error GoTo ErrHandler 'commented to test, need to be uncommented
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim s As Worksheet
    Dim convToText_cols As Variant, i As Integer
    
    Set s = Worksheets(cTempLoadWrkSh)
    
    s.Cells.Clear 'delete everything on the target worksheet
    
    CopyDataFromFile s, strFileToOpen 'copy date of the main sheet from the source file to the temp_load sheet
    
    DeleteBlankRows s 'clean blank rows of just imported file on the temp_load sheet
    
    convToText_cols = Split(GetConfigValue("FBS_LoadFromFile_Cols_ConvertToText"), ",")
    If ArrLength(convToText_cols) > 1 Then
        For i = 0 To ArrLength(convToText_cols) - 1
            ConvertNumberValueToString CStr(convToText_cols(i)), s '    ConvertNumberValueToString "B", s
        Next
    End If
    
    CopySelectedColumnToTargetSheet s, ws_target, GetConfigValue("FBS_LoadFromFile_ColMapping_Position"), True ' '"A:A"
    CopySelectedColumnToTargetSheet s, ws_target, GetConfigValue("FBS_LoadFromFile_ColMapping_TubeID"), True ' '"B:B"
    CopySelectedColumnToTargetSheet s, ws_target, GetConfigValue("FBS_LoadFromFile_ColMapping_Status"), True ' '"C:C"
    CopySelectedColumnToTargetSheet s, ws_target, GetConfigValue("FBS_LoadFromFile_ColMapping_BoxID"), True ' '"D:D"
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    ImportFile = True
    Exit Function
    
ErrHandler:
    MsgBox Err.Description, vbCritical
ExitMark:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ImportFile = False
End Function

'This sub opens specified file and loads it contents to a specified worksheet
Private Sub CopyDataFromFile(ws_target As Worksheet, _
                    src_file_path As String, _
                    Optional src_worksheet_name As String = "")
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim src As Workbook
    Dim path As String
    
    ' OPEN THE SOURCE EXCEL WORKBOOK IN "READ ONLY MODE".
    Set src = Workbooks.Open(src_file_path, True, True)
    
    If src_worksheet_name = "" Then
        src_worksheet_name = src.Worksheets(1).Name
    End If
    
    src.Worksheets(src_worksheet_name).Cells.Copy 'copy into a clipboard
    ws_target.Cells.PasteSpecial Paste:=xlPasteAll 'paste to the worksheet
    Application.CutCopyMode = False 'clean clipboard
    
  
    ' CLOSE THE SOURCE FILE.
    src.Close False             ' FALSE - DON'T SAVE THE SOURCE FILE.
    Set src = Nothing
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbCritical
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub CopySelectedColumnToTargetSheet(source As Worksheet, Target As Worksheet, mapping As String, Optional convertToText As Boolean = False)
    Dim copy_cols() As String
    Dim src_col As Range, dst_col As Range
    Dim src_used_rows As Integer, dest_used_rows As Integer
    
    src_used_rows = source.UsedRange.rows.Count
    dest_used_rows = Target.UsedRange.rows.Count
    
    copy_cols = Split(mapping, ":")
    If ArrLength(copy_cols) > 1 Then
        Set src_col = source.Range(copy_cols(0) & "2:" & copy_cols(0) & CStr(src_used_rows))
        Set dst_col = Target.Range(copy_cols(1) & "2:" & copy_cols(1) & CStr(dest_used_rows))
    End If
    dst_col.Clear
    
    If convertToText Then
        src_col.NumberFormat = "@"
        dst_col.NumberFormat = "@"
    End If
        
    src_col.Cells.Copy 'copy into a clipboard
    dst_col.Cells.PasteSpecial Paste:=xlPasteValues 'paste to the worksheet
    
End Sub

Private Sub DeleteBlankRows(ws_target As Worksheet)
    Dim SourceRange As Range
    Dim EntireRow As Range
    Dim i As Long, non_blanks As Long, empty_strings As Long
 
    Set SourceRange = ws_target.UsedRange ' Cells.End(xlToLeft)
 
    If Not (SourceRange Is Nothing) Then
        'Application.ScreenUpdating = False
 
        For i = SourceRange.rows.Count To 1 Step -1
            Set EntireRow = SourceRange.Cells(i, 1).EntireRow
            non_blanks = Application.WorksheetFunction.CountA(EntireRow)
            empty_strings = Application.WorksheetFunction.CountIf(EntireRow, "")
            If non_blanks = 0 Or EntireRow.Cells.Count = empty_strings Then
                EntireRow.Delete
            'Else
                'Print ("Not blank row")
            End If
        Next
 
        'Application.ScreenUpdating = True
    End If
End Sub

'Adds apostrophe (') character in front of the value of each affected cell to make sure Excel recongnizes that value as text
Private Sub ConvertNumberValueToString(columnToUpdate As String, Target As Worksheet) ', column_to_fill As String, fill_rows_num As Integer
    Dim dst_col As Range
    Dim dest_used_rows As Integer
    Dim cell As Range
    
    'clean target column first
    dest_used_rows = Target.UsedRange.rows.Count
    Set dst_col = Target.Range(columnToUpdate & "2:" & columnToUpdate & CStr(dest_used_rows))
    
    For Each cell In dst_col
        If IsNumeric(cell.value) Then
            cell.value = "'" & cell.value
        End If
    Next
    
End Sub
