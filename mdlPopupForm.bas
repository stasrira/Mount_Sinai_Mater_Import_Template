Attribute VB_Name = "mdlPopupForm"
Option Explicit

Enum FormUseCases
    FieldSettingProfile = 0
    ExportAssignmentSelection = 1
End Enum

Public popUpFormResponseIndex As Integer
Public formCurrentView As String 'this field will be checked from within frmSelection to turn on/off some of the features of the form
Public dictProfiles As New Dictionary
Public colExportItems As New Collection, a 'collection with key value, used to store list of export schemes available in the currently loaded profile

Public Function SelectFieldSettingProfile() As Integer
    
    popUpFormResponseIndex = -1 'set the default value
    
    If PrepareForm(FieldSettingProfile) Then
        frmSelection.Show
        'Debug.Print frmSelection.cmbProfileList.Value
        
    End If
    SelectFieldSettingProfile = popUpFormResponseIndex 'this value can be overwritten in the form frmSelection, if a selection was made there
End Function

Public Function SelectExportSchema() As Integer
    
    popUpFormResponseIndex = -1 'set the default value
    
    If PrepareForm(ExportAssignmentSelection) Then
        frmSelection.Show
        'Debug.Print frmSelection.cmbProfileList.Value
        
    End If
    SelectExportSchema = popUpFormResponseIndex 'this value can be overwritten in the form frmSelection, if a selection was made there
End Function

Public Function PrepareForm(use_case As FormUseCases) As Boolean
    
    formCurrentView = ""
    
    Select Case use_case
        Case FieldSettingProfile
            frmSelection.Caption = "Master Template Profiles"
            frmSelection.cmdLoad.Caption = "Load"
            frmSelection.cmdLoad.SetFocus
            frmSelection.Label1.Caption = "Select a profile to be loaded"
            frmSelection.txtDesc.Visible = True
            frmSelection.txtDesc.MultiLine = True
            frmSelection.txtDesc.Text = ""
            frmSelection.txtCurProfile.Visible = True
            
            formCurrentView = "FieldSettingProfile"

            PrepareForm = PopulateFieldSettingProfilesList(frmSelection.cmbProfileList)
        Case ExportAssignmentSelection
            frmSelection.Caption = "Master Template Exports"
            frmSelection.cmdLoad.Caption = "Select"
            frmSelection.cmdLoad.SetFocus
            frmSelection.Label1.Caption = "Select the export scheme to be used"
            frmSelection.txtDesc.Visible = True
            frmSelection.txtDesc.MultiLine = True
            frmSelection.txtDesc.Text = "Note: Default export format for profiles is comma separated value (.csv). " _
                & "This format will be forced for all profiles, except when profile specific conditions are " _
                & "provided through a special configuration field."
'            frmSelection.txtDesc.Text = "Note: The Export process supports only exporting data in comma separated value (.csv) format." _
'                                            & " Selecting any other file formats on the next screen will not honored by the system and the "".csv"" format will be used instead."
            frmSelection.txtCurProfile.Visible = False
            
            formCurrentView = "ExportAssignment"

            PrepareForm = PopulateExportList(frmSelection.cmbProfileList)
            
    End Select
End Function

Public Function PopulateExportList(ByRef cmb As ComboBox) As Boolean
    'read all Export Assignments from Field Settings worksheet
'    With Worksheets(cSettingsWorksheetName)
'        Dim rn As Range
'        Dim iRows As Integer
        Dim val_arr() As String
        Dim val_out() As Variant
        Dim i As Integer
'
'        'Dim fieldRowNum As Integer
'
'        iRows = .UsedRange.Rows.Count 'number of actually used rows
'
'        'identify range of actually used cells on the given spreadsheet for the ExportAssignment column
'        Set rn = .Range(cAddrExportAssignment & "2" & ":" & cAddrExportAssignment & iRows)
'        'concatenate all values from all cells of the range "rn" and split the recult into an array using "," as delimiter
'        val_arr = Split(Join(Application.WorksheetFunction.Transpose(rn), ","), ",")
'    End With

   'read all Export Assignments from Field Settings worksheet
   val_arr = GetFieldSettingPropertyVal_All(cAddrExportAssignment, ",")
        
    Set colExportItems = Nothing 'clear global collection
    
    'loop through all items of the array and add those to a collecton with the key.
    'This will keep only unique values and raise errors for duplicates - On Error Resume Next will ignore errors
    For Each a In val_arr
        'replace all blank values with "Default" schema
        If Len(Trim(a)) = 0 Then
            a = "Default"
        End If
        On Error Resume Next
        'Collections can be unique, as long as you use the second Key argument when adding items.
        'Key values must always be unique, and adding an item with an existing Key raises an error:hence the On Error Resume Next
        colExportItems.Add Trim(a), Trim(a)
        On Error GoTo 0
    Next
            
    'sort final collection of items and convert it back to array
    val_out = CollectionToArray(colExportItems) 'convert final collection to an array
    QuickSort val_out, LBound(val_out), UBound(val_out) - 1 'sort the array
    
    cmb.Clear
    Set colExportItems = Nothing 'clear global collection to load sorted list from scratch
    
    For i = 0 To UBound(val_out) - 1
        cmb.AddItem val_out(i)
        colExportItems.Add val_out(i), val_out(i)
        If val_out(i) = "Default" Then
            cmb.ListIndex = i
        End If
        
        'Debug.Print (val_out(i))
    Next
    
    PopulateExportList = True
    
End Function

'converts collection to an array
Public Function CollectionToArray(col As Collection) As Variant
    Dim a() As Variant
    ReDim a(0 To col.Count)
    Dim i As Long
    For i = 0 To col.Count - 1
        a(i) = col(i + 1)
    Next i
    CollectionToArray = a()
End Function

'implements quicksort algoritm, works with arrays
Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
    Dim pivot   As Variant
    Dim tmpSwap As Variant
    Dim tmpLow  As Long
    Dim tmpHi   As Long
    
    tmpLow = inLow
    tmpHi = inHi
    
    pivot = vArray((inLow + inHi) \ 2)
    
    While (tmpLow <= tmpHi)
       While (vArray(tmpLow) < pivot And tmpLow < inHi)
          tmpLow = tmpLow + 1
       Wend
    
       While (pivot < vArray(tmpHi) And tmpHi > inLow)
          tmpHi = tmpHi - 1
       Wend
    
       If (tmpLow <= tmpHi) Then
          tmpSwap = vArray(tmpLow)
          vArray(tmpLow) = vArray(tmpHi)
          vArray(tmpHi) = tmpSwap
          tmpLow = tmpLow + 1
          tmpHi = tmpHi - 1
       End If
    Wend
    
    If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub
