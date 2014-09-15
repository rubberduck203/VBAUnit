VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AssertConditionTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private test As VBAUnit.UnitTest

Public Sub SetOutputStream(out As IOutput)
    Set test = VBAUnit.UnitTestFactory.Create(TypeName(Me), out)
End Sub

Private Sub Class_Terminate()
    test.Dispose
    Set test = Nothing
End Sub

Public Sub RunAllTests()
    IsTrueShouldPass
    IsTrueShouldFail
    IsFalseShouldPass
    IsFalseShouldFail
End Sub

Public Sub IsTrueShouldPass()
    With test.Assert
        .Name = "IsTrueShouldPass"
        .IsTrue True
    End With
End Sub

Public Sub IsTrueShouldFail()
    With test.Assert
        .Name = "IsTrueShouldFail"
        .IsTrue False, "Failure is success."
    End With
End Sub

Public Sub IsFalseShouldPass()
    With test.Assert
        .Name = "IsFalseShouldPass"
        .IsFalse False
    End With
End Sub

Public Sub IsFalseShouldFail()
    With test.Assert
        .Name = "IsFalseShouldFail"
        .IsFalse True, "Failure is success."
    End With
End Sub
