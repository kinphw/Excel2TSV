Attribute VB_Name = "ModCallMsgbox"
'namespace=vba-files\excel2tsv\sub\

Public Sub CallMsgbox(Msg As String)
    
    Dim WSH As Object
    Set WSH = CreateObject("WScript.Shell")
    WSH.Popup Msg, 1, "Title", vbInformation
    Set WSH = Nothing
    
End Sub
