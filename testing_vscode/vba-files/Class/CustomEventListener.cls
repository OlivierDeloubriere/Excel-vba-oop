VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomEventListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

Public announcer As CustomEventAnnouncer
Dim wks as Worksheet

Private Sub Class_Initialize()
    Set announcer = New CustomEventAnnouncer
End Sub
Private Sub announcer_OnButton1Click()
    MsgBox "button 1 clicked"
End Sub
Private Sub announcer_OnButton2Click()
    MsgBox "button 2 clicked"
End Sub

