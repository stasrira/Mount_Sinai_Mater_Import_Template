Attribute VB_Name = "mdlSubAliquot"
Option Explicit

Public Sub CreateSubAliquotsAll()
    CreateSubAliquots , True
End Sub

Public Function CreateSubAliquots(Optional ByRef tRng As Range = Nothing, _
                                    Optional updateWholeSheet As Boolean = False, _
                                    Optional showConfirmMsg As Boolean = True, _
                                    Optional showValidateMsg As Boolean = True) As Range

    Const cMsgBoxTitle = "Create Sub-Aliquots"
    
    Dim r As Range, rfs As Range
    Dim i As Integer, cnt As Integer
    Dim fstCell As String, lstCell As String, lstUsedRngCell As String
    Dim lstUsedRngRow As Integer
    Dim wks As Worksheet
    Dim iResponse As Integer, inRowNum As Integer
    Dim fs As clsFieldSettings
    Dim cellProperties As clsCellProperties
    Dim field As Range, r1 As Range, colAutoFill As Collection, af As Variant
    Dim jVal As Dictionary, jNode As Variant, sn As Variant, jsonStr As String, updFields As Variant
    Dim msg_alert As String
    
    If tRng Is Nothing Then
        Set tRng = Selection
    End If
        
    Set wks = tRng.Worksheet
    
    If wks.Name <> cRawDataWorksheetName Then
        'MsgBox "Sub-Aliquots can be created only on the ""RawData"" worksheet. Please switch to ""RawData"" and try to run this operatoin again from there.", vbCritical, cMsgBoxTitle
        Set CreateSubAliquots = Nothing
        Exit Function
    End If
    
    If updateWholeSheet Then
        Set tRng = wks.UsedRange
    End If
    
    'get last row of the used range of cells on the current worksheet
    lstUsedRngCell = Split(wks.UsedRange.Address, ":")(1)
    lstUsedRngRow = Split(lstUsedRngCell, "$")(2)
    
    'Check that target range does not include field captions (first row)
    If Split(wks.Cells(tRng.row, 1).Address, "$")(2) = "1" Then
        If tRng.rows.Count > 1 And lstUsedRngRow > 1 Then
            'if requested target range has more than 1 row and
            'used range of the sheet has more than 1 row,
            'offset start position to one row down
            fstCell = wks.Cells(tRng.row, 1).Offset(1).Address
        Else
            'if target range has only 1 row (meaning that only the title row was selected) or used range consists of first row only, abort the operation
            MsgBox "Sub-Aliquots cannot be created for the captions row (first row of the worksheet). Select any other populated row to proceed.", vbCritical, cMsgBoxTitle
            Set CreateSubAliquots = Nothing
            Exit Function
        End If
    Else
        fstCell = wks.Cells(tRng.row, 1).Address
    End If
    
    'verify that the start of the region to be used for sub-aliquoting is not out of the used range of cells on the worksheet
    If Split(fstCell, "$")(2) > lstUsedRngRow Then
        'start of the region is outside of the used range, about operation
        MsgBox "Sub-Aliquots cannot be created for not populated rows. Select any populated row (except the first caption row) to proceed.", vbCritical, cMsgBoxTitle
        Exit Function
    End If
    
    lstCell = wks.Cells(tRng.row + tRng.rows.Count - 1, wks.UsedRange.Columns.Count).Address
    
    'verify that the end of the region to be used for sub-aliquoting is not out of the used range of cells on the worksheet
    If Split(lstCell, "$")(2) > lstUsedRngRow Then
        'end of the provided target range is outside of the used range on the worksheet
        lstCell = lstUsedRngCell
    End If
    
    're-set target range that will be used to create sub-aliqouts
    Set tRng = wks.Range(fstCell, lstCell)
    
    
    'get config settings for sub-aliquot processing
    Set fs = GetFieldSettingsInstance(Nothing, False, cConfigFieldPrefix & "SubALiquot")

    If fs.DataAvailable Then
        
        msg_alert = ""
        
        jsonStr = fs.FieldMiscSettings '"{'Process':'sub-aliquot','SubAliquotNumber':6,'UpdateFields':[{'Field':'MT_Sample ID','Update':1,'UpdateOriginal':'{MT_Sample ID}_0','AutoFill':1},{'Field':'MT_Vessel ID','Update':1,'UpdateOriginal':'{MT_Vessel ID}_10','AutoFill':0}]}"
        Set jVal = ParseJson(jsonStr)
    
        If jVal.Exists("SubAliquotNumber") Then
            inRowNum = jVal("SubAliquotNumber")
        End If
    
        If jVal.Exists("UpdateFields") Then
            Set updFields = jVal("UpdateFields")
    
    '        For Each jNode In updFields
    '            Debug.Print jNode("AutoFill"), jNode("UpdateOriginal"), jNode("Field")
    '        Next
        End If
    Else
        msg_alert = "Warning: Currently loaded profile has no sub-aliquot instructions specified and thus a default sub-aliquot processing will be applied!" _
                    & " Cancel the operation if you do not want to proceed." _
                    & vbCrLf & vbCrLf
        inRowNum = 1
        Set updFields = Nothing
    End If
    
    'Confirm that user want to proceed with sub-aliquot creation
    tRng.EntireRow.Select
    
    iResponse = MsgBox(msg_alert & _
                        "Creating sub-aliquots process is about to start. The system will create " & CStr(inRowNum) & " sub-aliqout(s) for each of the selected row(s)." & _
                        vbCrLf & vbCrLf & "Do you want to proceed? If not, click 'Cancel'.", _
                        vbOKCancel, cMsgBoxTitle)
    
    If iResponse <> vbOK Then
        Set CreateSubAliquots = Nothing
        Exit Function
    End If

    If inRowNum > 0 Then
        'count of rows in the target range
        cnt = tRng.rows.Count 'Columns(1).Cells.Count
        Set r = tRng.Cells(1)
        'fstCell = r.Address
        
        'loop through all cells of the first column
        For i = 1 To cnt
            If i > 1 Then
                'shift current cell to next cell (of the first column) in the target range
                Set r = r.Offset(1, 0)
            End If
            
            Set colAutoFill = New Collection
            
            If Not updFields Is Nothing Then
            
                For Each jNode In updFields
    '                Debug.Print jNode("AutoFill"), jNode("UpdateOriginal"), jNode("Field")
                    
                    Set cellProperties = New clsCellProperties
                    
                    Set r1 = wks.Range("A1", wks.Cells(1, wks.UsedRange.Columns.Count))
                    Set field = r1.Find(jNode("Field"), LookIn:=xlValues, LookAt:=xlWhole).Offset(r.row - 1)
                    If Not field Is Nothing Then
                        cellProperties.InitializeValues field.Address
                        field.value = fs.EvalCellValueWithRef(CStr(jNode("UpdateOriginal")), cellProperties, cRawDataWorksheetName)
                        If CStr(jNode("AutoFill")) = "1" Then
                            colAutoFill.Add field.Column
                        End If
                    End If
                Next
                
            End If
            
            'create sub-aliquots and return range of the affected cells
            Set rfs = InsertSubAliquotsPerRow(r, inRowNum)
            
            For Each af In colAutoFill
                AutoFillColumn rfs, CInt(af)
            Next
            'call autofill procedure passing there range of the sub-aliquots and the column number (of the range) where first cell of the column will be used to autofill rest of the cells of the column
            'AutoFillColumn rfs, 2
            'AutoFillColumn rfs, 4
            
        Next
        'lstCell = r.Address
        lstCell = wks.Cells(r.row + r.rows.Count - 1, wks.UsedRange.Columns.Count).Address
    End If
    

    Set r = wks.Range(fstCell, lstCell)
    'Debug.Print r.Address
    r.EntireRow.Select
    
    If showConfirmMsg Then
        MsgBox "Sub-aliqouts were successfully created. Affected rows are highlighted.", vbInformation, cMsgBoxTitle
    End If
    
    If showValidateMsg Then
        iResponse = MsgBox("Do you want to proceed with ""Validate RawData Sheet"" operation? This will validate all created sub-aliquot entries." & _
                            vbCrLf & vbCrLf & "If you do not want to proceed, click 'Cancel'.", _
                            vbOKCancel, cMsgBoxTitle)
        
        If iResponse = vbOK Then
            ValidateWholeWorksheet
        End If
    
    End If
    
    Set CreateSubAliquots = r
End Function

'fill series:
'Range("A10").AutoFill Destination:=Range("A10:A18"), Type:=xlFillSeries

Public Sub AutoFillColumn(updRng As Range, colNum As Integer)
    Dim col As Range
    Dim stCell As Range, endCell As Range
    Dim rows As Integer
    
    Set stCell = updRng.Columns(colNum).Cells(1)
    Set endCell = stCell.Offset(updRng.rows.Count - 1)
    
    stCell.AutoFill Destination:=Range(stCell, endCell), Type:=xlFillSeries
    
End Sub

Public Function InsertSubAliquotsPerRow(row As Range, inRowNum As Integer) As Range
    Dim r_ins As Range
    Dim j As Integer
    Dim wks As Worksheet
    Dim stCell As String, endCell As String
    
    Set wks = row.Worksheet
        
    'create a temp range where the first cell is the current cell of the target range and the temp range has number of row equal to number of rows to be inserted.
    Set r_ins = wks.Range(row.Address, row.Address)  'Range(r.Address, r.Offset(inRowNum - 1, 0))
    
    stCell = wks.Cells(r_ins.row, 1).Address
    
    For j = 1 To inRowNum
        r_ins.EntireRow.Copy
        'insert rows; this will insert as many rows as in the r_ins range.
        r_ins.EntireRow.Insert xlUp

    Next
    
    endCell = wks.Cells(r_ins.row, wks.UsedRange.Columns.Count).Address
    
    Set InsertSubAliquotsPerRow = wks.Range(stCell, endCell)
End Function

