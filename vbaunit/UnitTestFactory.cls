VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitTestFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Function Create(ByVal Name As String, ByRef out As IOutput) As UnitTest
    Set Create = New UnitTest
    Create.Initialize Name, out
End Function
