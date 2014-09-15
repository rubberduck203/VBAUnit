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
