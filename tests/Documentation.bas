Attribute VB_Name = "Documentation"
Option Explicit

Public Sub AddNewUnitTest()
    ' Excel
    UnitTestModule.Add ThisWorkbook.VBProject
    
    ' Access
    'UnitTestModule.Add VBE.ActiveVBProject
    
    'Word
    'UnitTestModule.Add ThisDocument.VBProject
End Sub



