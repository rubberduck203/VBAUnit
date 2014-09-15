VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Property Get Name() As String
    Name = "VBAUnit"
End Property

Public Property Get Version() As String
    Version = "0.0.1"
End Property

Public Property Get InstallationPath() As String
    Dim fso As New FileSystemObject
    Dim filePath As String
        
    filePath = Application.VBE.VBProjects(Name).fileName
    InstallationPath = fso.GetParentFolderName(filePath)
End Property

Friend Property Get ResourcePath() As String
    ResourcePath = InstallationPath & "\resources"
End Property
