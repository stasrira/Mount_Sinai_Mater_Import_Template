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

Public Sub GetPositionValues(ByVal Target As Range, Optional row_OffSetValue As Integer = 1, Optional col_OffSetValue As Integer = 2)
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


Public Sub AssignFormulaToCheckVesselIdMatch(ByVal Target As Range _
                                            , Optional col_OffSetValue As Integer = 3 _
                                            , Optional field_toMatch As String = "MT_Vessel ID")
                                            
    'Example of a formula: IF(A4="", "",OR(COUNTIF(Validated!D:D,A4)>0,COUNTIF(Validated!D:D,A4)>0))
    Const formula_template = "IF({barcode_col}{row_num}="""", """",OR(COUNTIF(Validated!{col_letter}:{col_letter},{barcode_col}{row_num})>0,COUNTIF(Validated!{col_letter}:{col_letter},{barcode_col}{row_num})>0))"
    
    Dim oFieldSettings As clsFieldSettings
    Dim sFormula As String
    Dim c As Range, cf As Range
    
    'disable application events while modifying conditional formatting
    Application.EnableEvents = False
    On Error Resume Next
    
    'get field setting properties of the field passed in the "field_toMatch" parameter (usualy MT_Vessel ID)
    Set oFieldSettings = GetFieldSettingsInstance(Nothing, False, field_toMatch)
    
    If oFieldSettings.DataAvailable Then
        With Target.Worksheet
            For Each c In Target
                'set formula for the current cell
                sFormula = formula_template
                sFormula = Replace(sFormula, "{col_letter}", oFieldSettings.FieldColumnNameOnRawData) 'set column letter of a column to be used as a search range of COUNTIF formula
                sFormula = Replace(sFormula, "{barcode_col}", Split(c.Address, "$")(1)) 'set current cell column letter
                sFormula = Replace(sFormula, "{row_num}", Split(c.Address, "$")(2)) 'set current cell row number
                Debug.Print sFormula 'for testing only
                'set cell reference to be updated with the formula
                Set cf = c.Offset(0, col_OffSetValue)
                If Len(Trim(c.value)) > 0 Then
                    If Not IsError(.Evaluate(sFormula)) Then
                        cf.value = .Evaluate(sFormula)
                        If cf.value Then
                            ApplyFormatingToCell cf, BackgroundColors.Green, FontColors.DarkGreen
                        Else
                            ApplyFormatingToCell cf, BackgroundColors.LightRed, FontColors.DarkRed
                        End If
                    Else
                        cf.value = ""
                        ApplyFormatingToCell cf, BackgroundColors.white, FontColors.Black
                    End If
                Else
                    cf.value = ""
                    ApplyFormatingToCell cf, BackgroundColors.white, FontColors.Black
                End If
            Next
        End With
        
    End If
    
    On Error GoTo 0
    'enable back application events
    Application.EnableEvents = True
    
End Sub
