Attribute VB_Name = "mdlSubAliquot"
Option Explicit

Public Function CreateSubAliquots(Optional inRowNum As Integer = 1, Optional ByRef tRng As Range = Nothing, Optional updateWholeSheet As Boolean = False) As Range
    Dim r As Range, rfs As Range
    Dim i As Integer, cnt As Integer
    Dim fstCell As String, lstCell As String
    Dim wks As Worksheet
    
    If tRng Is Nothing Then
        Set tRng = Selection
    End If
        
    Set wks = tRng.Worksheet
    
    If updateWholeSheet Then
        Set tRng = wks.UsedRange
    End If
        
    fstCell = wks.Cells(tRng.row, 1).Address
    lstCell = wks.Cells(tRng.row + tRng.rows.Count - 1, wks.UsedRange.Columns.Count).Address
    'Set tRng = wks.Range(wks.Cells(tRng.row, 1), wks.Cells(tRng.row + tRng.rows.Count - 1, wks.UsedRange.Columns.Count))
    Set tRng = wks.Range(fstCell, lstCell)

    
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
            
            'TODO: here should be a call to a procedure that will update some cells of the row whichi is the source for creating sub-aliquots
            'this is just for testing update first cells of the columns that will be autofill afterwords
            wks.Cells(r.row, 2).value = wks.Cells(r.row, 2).value & "_0"
            wks.Cells(r.row, 4).value = wks.Cells(r.row, 4).value & "_0"
            
            'create sub-aliquots and return range of the affected cells
            Set rfs = InsertSubAliquotsPerRow(r, inRowNum)
            
            'call autofill procedure passing there range of the sub-aliquots and the column number (of the range) where first cell of the column will be used to autofill rest of the cells of the column
            AutoFillColumn rfs, 2
            AutoFillColumn rfs, 4
            
        Next
        'lstCell = r.Address
        lstCell = wks.Cells(r.row + r.rows.Count - 1, wks.UsedRange.Columns.Count).Address
    End If
    

    Set r = wks.Range(fstCell, lstCell)
    Debug.Print r.Address
    r.EntireRow.Select
    
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

