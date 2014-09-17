VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitTestModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private ThisProject As VBProject

Public Sub Add(Project As VBProject)
On Error GoTo ErrHandler

    Set ThisProject = Project
    
    Dim comp As VBComponent
    Set comp = Project.VBComponents.Add(vbext_ct_ClassModule)
    
    comp.Name = GetAvailableName("NewUnitTest")
    
    InsertTemplate comp
    
ExitSub:
    Set ThisProject = Nothing
    Exit Sub

ErrHandler:
    '@todo - log and report error
    Resume ExitSub
End Sub

Private Sub InsertTemplate(comp As VBComponent)
    
    comp.CodeModule.AddFromFile ProjectInfo.ResourcePath & "\NewUnitTestTemplate.bas"
    
End Sub

Private Function GetAvailableName(Name As String, Optional callCount As Long) As String

    Dim I As Long
    For I = 1 To ThisProject.VBComponents.Count
        If ThisProject.VBComponents.Item(I).Name = Name Then
            Name = GetAvailableName(Name & callCount + 1)
            Exit For
        End If
    Next I
    
    GetAvailableName = Name
End Function




