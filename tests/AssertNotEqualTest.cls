VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AssertNotEqualTest"
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
    TwoEquatableObjectsAreNotEqual
    TwoNonEquatableObjectsAreNotEqual
    TwoIntegersAreNotEqual
    TwoDoublesAreNotEqual
    OneObjectAndOneValueAreNotEqual
End Sub

Public Sub TwoNonEquatableObjectsAreNotEqual()
    test.Assert.Name = "TwoNonEquatableObjectsAreNotEqual"
    
    Dim a As New MockNotEquatable
    Dim b As New MockNotEquatable
    
    test.Assert.NotEqual a, b
End Sub

Public Sub TwoEquatableObjectsAreNotEqual()
    test.Assert.Name = "TwoEquatableObjectsAreNotEqual"
    
    Dim a As New MockEquatable
    a.Name = "Hello"
    Dim b As New MockEquatable
    b.Name = "Goodbye"
    
    test.Assert.NotEqual a, b
End Sub

Public Sub OneObjectAndOneValueAreNotEqual()
    test.Assert.Name = "OneObjectAndOneValueAreNotEqual"
    
    Dim obj As New MockEquatable
    Const one As Integer = 1
    
    test.Assert.NotEqual obj, 1
    test.Assert.NotEqual 1, obj
    
End Sub

Public Sub TwoIntegersAreNotEqual()
    test.Assert.Name = "TwoIntegersAreNotEqual"
    
    test.Assert.NotEqual 1, 2
End Sub

Public Sub TwoDoublesAreNotEqual()
    test.Assert.Name = "TwoDoublesAreNotEqual"
    
    test.Assert.NotEqual 10#, 11#
End Sub

