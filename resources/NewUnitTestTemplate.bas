Private test As VBAUnit.UnitTest

Friend Sub SetOutputStream(out As IOutput)
    Set test = VBAUnit.UnitTestFactory(TypeName(Me), out)
End Sub

Private Sub Class_Initialize()
    SetOutputStream VBEX.Console
End Sub

Private Sub Class_Terminate()
    test.Dispose
    Set test = Nothing
End Sub