VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDisposable

Private Type TUnitTest
    Name As String
    OutStream As IOutput
'    Assert As TestAssert
End Type

Private this As TUnitTest

Public WithEvents Assert As TestAssert
Attribute Assert.VB_VarHelpID = -1

Public Property Get Name() As String
    Name = this.Name
End Property

Friend Property Get OutStream() As IOutput
    Set OutStream = this.OutStream
End Property

'Public Property Get Assert() As Assert
'    Set Assert = this.Assert
'End Property

Friend Sub Initialize(Name As String, out As IOutput)
    this.Name = Name
    Set this.OutStream = out
    Set Assert = New TestAssert
    Set Assert.Parent = Me
End Sub

Public Sub Dispose()
    Set Assert = Nothing
 '   this.OutStream.Dispose
End Sub

Private Sub Assert_Completed(ByVal Result As TestResult)
    ReportResult Result
End Sub

Private Sub IDisposable_Dispose()
    Dispose
End Sub

Private Sub Class_Terminate()
    Dispose
End Sub

Private Sub ReportResult(Result As TestResult)
    Dim output As String
    output = this.Name & "." & Assert.Name & ": "
    
    Select Case Result.TestOutcome
        Case Failed
            output = output & "Fail"
        Case Succeeded
            output = output & "Pass"
        Case Else
            output = output & "Undetermined"
    End Select
    
    If Result.Message <> vbNullString Then
        output = output & ": " & Result.Message
    End If
    
    this.OutStream.PrintLine output
    
    Set Result = Nothing
End Sub
