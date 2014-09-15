VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockEquatable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IEquatable

Public Name As String

Public Function Equals(ByVal other As MockEquatable) As Boolean
    Equals = (Name = other.Name)
End Function

Private Function IEquatable_Equals(ByVal other As Variant) As Boolean
    IEquatable_Equals = Equals(other)
End Function
