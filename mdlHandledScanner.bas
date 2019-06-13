Attribute VB_Name = "mdlHandledScanner"
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
        
    If Target.Column = 1 And Target.Row > 1 Then
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
    End If
End Sub
