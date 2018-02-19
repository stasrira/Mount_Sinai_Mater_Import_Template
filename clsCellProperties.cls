VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCellProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CellAddress As String
'Public CellCollumn As String
Public cellRow As String
Public CellFieldName As String
Public CellConditionalInteriorColor As Long
Public CellConditionalFontColor As Long

Public Sub InitializeValues(vCellAddr As String)
    With Application.Worksheets(cRawDataWorksheetName).Range(vCellAddr)
        CellAddress = vCellAddr
    
    '    Dim Delim2 As Integer
    '    Delim2 = InStrRev(vCellAddr, "$")
    '   Right(vCellAddr, Len(vCellAddr) - Delim2)
        cellRow = .Row
    '    CellCollumn = Replace(Left(vCellAddr, Delim2), "$", "")
        
        'CellFieldName = Application.Worksheets(cRawDataWorksheetName).Range(CellCollumn + "1").Value
    
        CellFieldName = .Offset(-.Row + 1, 0).Value 'offset up to get to the first row (which should hold the field name)
    End With
    
End Sub
