VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Implements IOutput
Implements ITestOutput

Private Sub ITestOutput_ReportResult(result As TestResult)
    Dim output As String
    output = result.TestName & "." & result.MethodName & ": "
    
    Select Case result.TestOutcome
        Case Failed
            output = output & "Fail"
        Case Succeeded
            output = output & "Pass"
        Case Else
            output = output & "Undetermined"
    End Select
    
    If result.Message <> vbNullString Then
        output = output & ": " & result.Message
    End If
    
    PrintLine output
    
    Set result = Nothing
End Sub
Public Sub PrintLine(Optional ByVal output As Variant)
    VBEX.Console.PrintLine output
End Sub

Private Sub IOutput_PrintLine(Optional ByVal object As Variant)
    PrintLine object
End Sub

Public Sub Dispose()
End Sub

Private Sub IOutput_Dispose()
End Sub
