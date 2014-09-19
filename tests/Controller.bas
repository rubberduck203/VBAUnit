Attribute VB_Name = "Controller"
Option Explicit

Private out As IOutput

Public Sub RunAllTests()
    
    'Set out = VBAUnit.LoggerFactory.Create
    Set out = VBAUnit.Console
              
    TestConditions
    out.PrintLine

    TestAreEquals
    out.PrintLine
    
    TestNotEquals
    out.PrintLine
    
End Sub

Public Sub TestConditions()
    Dim test As New AssertConditionTest
    test.SetOutputStream out
    
    test.RunAllTests
    'test.IsTrueShouldPass
    
End Sub

Public Sub TestAreEquals()
    Dim test As New AssertEqualTest
    test.SetOutputStream out
    
    test.RunAllTests
End Sub

Public Sub TestNotEquals()
    Dim test As New AssertNotEqualTest
    test.SetOutputStream out
    
    test.RunAllTests
End Sub
