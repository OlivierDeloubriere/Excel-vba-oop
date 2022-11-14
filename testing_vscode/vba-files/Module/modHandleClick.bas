Attribute VB_Name = "modHandleClick"
'@Folder("VBAProject")
Option Explicit
'namespace=vba-files/Module
Public Sub handleClickEvent()
    If IsError(Application.Caller) then 
        msgbox "Error in Sub handleClickEvent - Application.Caller not found", vbCritical
        Exit Sub
    End if
    on error resume next
    With New CustomEventListener
        .announcer.RaiseCustomEvent Application.Caller
    End With
End Sub