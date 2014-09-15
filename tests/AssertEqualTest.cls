VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AssertEqualTest"
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
    TwoEquatableObjectsAreEqual
    TwoNonEquatableObjectsAreEqual
    TwoIntegersAreEqual
    TwoDoublesAreEqual
    OneObjectAndOneValueShouldFail
End Sub

Public Sub TwoNonEquatableObjectsAreEqual()
    test.Assert.Name = "TwoNonEquatableObjectsAreEqual"
    
    Dim a As New MockNotEquatable
    Dim b As New MockNotEquatable
    Set b = a
    
    test.Assert.AreEqual a, b
End Sub

Public Sub TwoEquatableObjectsAreEqual()
    test.Assert.Name = "TwoEquatableObjectsAreEqual"
    
    Const str As String = "Hello"
    Dim a As New MockEquatable
    a.Name = str
    Dim b As New MockEquatable
    b.Name = str
    
    test.Assert.AreEqual a, b
End Sub

Public Sub OneObjectAndOneValueShouldFail()
    test.Assert.Name = "OneObjectAndOneValueShouldFail"
    
    Dim obj As New MockEquatable
    Const one As Integer = 1
    
    test.Assert.AreEqual obj, 1, "Failure is success."
    test.Assert.AreEqual 1, obj, "Failure is success."
    
End Sub

Public Sub TwoIntegersAreEqual()
    test.Assert.Name = "TwoIntegersAreEqual"
    test.Assert.AreEqual 1, 1
End Sub

Public Sub TwoDoublesAreEqual()
    test.Assert.Name = "TwoDoublesAreEqual"
    
    test.Assert.AreEqual 10#, 10#
End Sub
