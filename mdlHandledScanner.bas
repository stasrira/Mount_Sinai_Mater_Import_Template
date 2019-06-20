Attribute VB_Name = "mdlHandledScanner"
Option Explicit

Public Function GetWellColumn_For96Well(rowNum As Integer) As String
    Dim i As Integer

    Dim arrVal_96well() As String
    
    arrVal_96well() = Split("1,1,1,1,1,1,1,1,2,2,2,2,2,2,2,2,3,3,3,3,3,3,3,3,4,4,4,4,4,4,4,4,5,5,5,5,5,5,5,5,6,6,6,6,6,6,6,6,7,7,7,7,7,7,7,7,8,8,8,8,8,8,8,8,9,9,9,9,9,9,9,9,10,10,10,10,10,10,10,10,11,11,11,11,11,11,11,11,12,12,12,12,12,12,12,12,", ",", -1, vbBinaryCompare)
    
    While rowNum > UBound(arrVal_96well) - LBound(arrVal_96well)
        rowNum = rowNum - ((UBound(arrVal_96well) - LBound(arrVal_96well)))
    Wend
    
    If rowNum > 0 Then
        GetWellColumn_For96Well = arrVal_96well(rowNum - 1)
    Else
        GetWellColumn_For96Well = ""
    End If
End Function

'this function will extract 1 column out of the provided range; it might resize range (based on the provided parameter)to remove the 1st row, if the range includes it
Private Function GetColumnToEvaluate(ByVal Target As Range _
                                        , columnNum_toEval As Integer _
                                        , Optional avoidFirstRow As Boolean = True) As Range
    Dim i As Integer, rCol As Range
    
    'identify column to evaluate
    For i = 1 To Target.Columns.Count
        Set rCol = Target.Columns(i)
        If rCol.Cells(1).Column = columnNum_toEval Then 'identify range of cells belonging to the needed column number
            If avoidFirstRow Then 'check if avoid udpating first cell is set to True
                If rCol.Cells(1).Row = 1 Then 'check if the range includes the first row
                    If rCol.Rows.Count > 1 Then 'check if range contains more than one row
                        Set rCol = rCol.Offset(1, 0).Resize(rCol.Cells.Count - 1) 'shift range down by one row and decrease size by one row as well
                    Else
                        Set rCol = Nothing 'current range contains just one cell located in the header, abourt operation
                    End If
                End If
            End If
            Exit For
        Else
            Set rCol = Nothing
        End If
    Next
    
    Set GetColumnToEvaluate = rCol
End Function

'this function will update row/column values for hand scanned barcodes
Public Sub UpdatePositionValues(ByVal Target As Range, _
                                columnNum_toEval As Integer, _
                                Optional row_OffSetValue As Integer = 1, _
                                Optional col_OffSetValue As Integer = 2, _
                                Optional avoidFirstRow As Boolean = True)
    
    On Error GoTo err1
    
    Const formula_template1 = "=IF(LEN(TRIM({col_letter}{row_num}))>0, IF(MOD(ROW()-1, 8) = 0, CHAR(72), CHAR(72 - 8 + MOD(ROW()-1, 8))), """")"
    Const formula_template2 = "=IF(LEN(TRIM({col_letter}{row_num}))>0, GetWellColumn_For96Well(ROW()-1), """")"
    '=IF(LEN(TRIM({col_letter}{row_num}))>0, IF(MOD(ROW()-1, 8) = 0, CHAR(72), CHAR(72 - 8 + MOD(ROW()-1, 8))), "")
    '=IF(LEN(TRIM({col_letter}{row_num}))>0, GetWellColumn_For96Well(ROW()-1), "")
    
    Dim rCol As Range, c As Range
    Dim sForm1 As String, sForm2 As String
                    
    Set rCol = GetColumnToEvaluate(Target, columnNum_toEval, avoidFirstRow)
    
    sForm1 = formula_template1
    sForm1 = Replace(sForm1, "{col_letter}", Split(rCol.Cells(1).Address, "$")(1)) 'set column letter of the first cell of the range
    sForm1 = Replace(sForm1, "{row_num}", Split(rCol.Cells(1).Address, "$")(2)) 'set row number of the first cell of the range
    sForm2 = formula_template2
    sForm2 = Replace(sForm2, "{col_letter}", Split(rCol.Cells(1).Address, "$")(1)) 'set column letter of the first cell of the range
    sForm2 = Replace(sForm2, "{row_num}", Split(rCol.Cells(1).Address, "$")(2)) 'set row number of the first cell of the range
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    rCol.Offset(0, row_OffSetValue).Formula = sForm1
    rCol.Offset(0, col_OffSetValue).Formula = sForm2
    
err1:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub

'TO DELETE
Public Sub UpdatePositionValues_old(ByVal Target As Range, Optional row_OffSetValue As Integer = 1, Optional col_OffSetValue As Integer = 2)
    Dim c As Range

    With Target.Worksheet
        For Each c In Target
            If Len(Trim(c.value)) > 0 Then
                c.Offset(0, row_OffSetValue).value = IIf((c.Row - 1) Mod 8 = 0, Chr(72), Chr(72 - 8 + ((c.Row - 1) Mod 8)))
                c.Offset(0, col_OffSetValue).value = GetWellColumn_For96Well(c.Row - 1)
            Else
                c.Offset(0, row_OffSetValue).value = ""
                c.Offset(0, col_OffSetValue).value = ""
            End If
        Next
    End With
    
End Sub

Public Sub AssignConditionalFormatting_String(ByRef Target As Range, condition As String, fontColor As FontColors, BgColor As BackgroundColors, Optional clearPrevFormating As Boolean = False)

'apply conditional formatting
    If clearPrevFormating Then Target.FormatConditions.Delete
    Target.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=" & condition & ""
    Target.FormatConditions(Target.FormatConditions.Count).SetFirstPriority
    With Target.FormatConditions(1).Font
        .Color = fontColor '-16752384
        .TintAndShade = 0
    End With
    With Target.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = BgColor '13561798
        .TintAndShade = 0
    End With
    Target.FormatConditions(1).StopIfTrue = False
    
End Sub

'currently used from Flatbed scanner sheet
Public Sub AssignFormulaToCheckVesselIdMatch(ByVal Target As Range _
                                            , columnNum_toEval As Integer _
                                            , Optional col_OffSetValue As Integer = 3 _
                                            , Optional field_toMatch As String = "MT_Vessel ID" _
                                            , Optional avoidFirstRow As Boolean = True)
                                            
    'Example of a formula: IF(A4="", "",OR(COUNTIF(Validated!D:D,A4)>0,COUNTIF(Validated!D:D,A4)>0))
    Const formula_template = "=IF({barcode_col}{row_num}="""", """",OR(COUNTIF(Validated!{col_letter}:{col_letter},{barcode_col}{row_num})>0,COUNTIF(Validated!{col_letter}:{col_letter},{barcode_col}{row_num})>0))"
    
    Dim oFieldSettings As clsFieldSettings
    Dim sFormula As String, i As Integer
    Dim c As Range, cf As Range, rCol As Range

    On Error GoTo err1
    
    'get field setting properties of the field passed in the "field_toMatch" parameter (usualy MT_Vessel ID)
    Set oFieldSettings = GetFieldSettingsInstance(Nothing, False, field_toMatch)
    
    If oFieldSettings.DataAvailable Then
        
        Set rCol = GetColumnToEvaluate(Target, columnNum_toEval, avoidFirstRow)
        
        If Not rCol Is Nothing Then 'if required column was found, proceed here
        
            sFormula = formula_template
            sFormula = Replace(sFormula, "{col_letter}", oFieldSettings.FieldColumnNameOnRawData) 'set column letter of a column to be used as a search range of COUNTIF formula
            sFormula = Replace(sFormula, "{barcode_col}", Split(rCol.Cells(1).Address, "$")(1)) 'set current cell column letter
            sFormula = Replace(sFormula, "{row_num}", Split(rCol.Cells(1).Address, "$")(2)) 'set current cell row number
            
            Set cf = rCol.Offset(0, col_OffSetValue)
            
            Application.ScreenUpdating = False
            Application.EnableEvents = False
            
            cf.Formula = sFormula
            
            Application.ScreenUpdating = True
            Application.EnableEvents = True
            
            
            AssignConditionalFormatting_String cf, "TRUE", FontColors.DarkGreen, BackgroundColors.Green, True
            AssignConditionalFormatting_String cf, "FALSE", FontColors.DarkRed, BackgroundColors.LightRed, False
            
        End If
        
    End If
    
    Exit Sub
    
err1:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub

