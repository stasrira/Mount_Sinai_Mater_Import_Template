VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBoxPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ScannedLocation As String
Public Row_p4 As String
Public Column_p5 As String

Public Sub SplitScannedLocationToPositions(ScanLocation As String)
    Dim regMatches As Object, match As Object
    
    ScannedLocation = ScanLocation
    
    Row_p4 = ""
    Column_p5 = ""
    
    Set regMatches = ApplyRegExToStr(ScanLocation, "\D{1,}") ' Get letter part of the scanned location ==> if location is "A01", this will get "A"
    For Each match In regMatches
        Row_p4 = match.value
        Exit For
    Next
    
    Set regMatches = ApplyRegExToStr(ScanLocation, "\d{1,}") ' Get letter part of the scanned location ==> if location is "A01", this will get "01"
    For Each match In regMatches
        Column_p5 = match.value
        Exit For
    Next
End Sub
