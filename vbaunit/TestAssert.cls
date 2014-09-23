VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestAssert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type TAssert
    Parent As UnitTest
    Name As String
End Type

Private this As TAssert

Public Event Completed(ByVal result As TestResult)

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Let Name(ByVal value As String)
    this.Name = value
End Property

Friend Property Get Parent() As UnitTest
    Set Parent = this.Parent
End Property

Friend Property Set Parent(ByVal value As UnitTest)
    Set this.Parent = value
End Property

Public Sub IsTrue(condition As Boolean, Optional Message As String)

    Dim outcome As TestOutcome
    outcome = IIf(condition, Succeeded, Failed)
    
    RaiseEvent Completed(TestResult.Create(outcome, Message, this.Parent.Name, this.Name))
    
End Sub

Public Sub IsFalse(condition As Boolean, Optional Message As String)

    Dim outcome As TestOutcome
    outcome = IIf(condition, Failed, Succeeded)
    
    RaiseEvent Completed(TestResult.Create(outcome, Message, this.Parent.Name, this.Name))
    
End Sub

' have this reviewed
Public Sub AreEqual(object1 As Variant, object2 As Variant, Optional ByVal Message As String)
    
    If IsEquatable(object1) And IsEquatable(object2) Then
        IsTrue object1.Equals(object2), Message
        
    ElseIf IsObject(object1) Then
    
        If IsObject(object2) Then
            IsTrue (object1 Is object2), Message
        Else
            RaiseEvent Completed(TestResult.Create(Failed, Message, this.Parent.Name, this.Name))
        End If
        
    ElseIf IsObject(object2) Then
        RaiseEvent Completed(TestResult.Create(Failed, Message, this.Parent.Name, this.Name))
    Else
        IsTrue (object1 = object2), Message
    End If
    
End Sub

Public Sub NotEqual(object1 As Variant, object2 As Variant, Optional ByVal Message As String)

    If IsEquatable(object1) And IsEquatable(object2) Then
        IsFalse object1.Equals(object2), Message
        
    ElseIf IsObject(object1) Then
    
        If IsObject(object2) Then
            IsFalse (object1 Is object2), Message
        Else
            RaiseEvent Completed(TestResult.Create(Succeeded, Message, this.Parent.Name, this.Name))
        End If
        
    ElseIf IsObject(object2) Then
        RaiseEvent Completed(TestResult.Create(Succeeded, Message, this.Parent.Name, this.Name))
    Else
        IsFalse (object1 = object2), Message
    End If

End Sub

Private Function IsEquatable(object As Variant) As Boolean
    If IsObject(object) Then
        IsEquatable = (TypeOf object Is IEquatable)
    Else
        IsEquatable = False
    End If
End Function



