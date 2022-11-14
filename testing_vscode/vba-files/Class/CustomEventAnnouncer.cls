VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomEventAnnouncer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

Public Event OnButton1Click()
Public Event OnButton2Click()

Public Sub RaiseCustomEvent(ByVal callerName As String)

If callerName = "button-1" Then RaiseEvent OnButton1Click
If callerName = "button-2" Then RaiseEvent OnButton2Click

End Sub

