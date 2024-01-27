Attribute VB_Name = "ModHelp"
'namespace=vba-files\excel2tsv\common\

Public Sub CallHelp(control As IRibbonControl)
    
    Dim Msg As String
    
    Msg = ""
    Msg = Msg + "엑셀파일들이 있는 폴더를 지정하면 모두 텍스트파일(tsv, cp949)로 변환합니다." + vbCrLf
    Msg = Msg + "하위폴더가 있는 경우 하위폴더도 순환합니다." + vbCrLf
    Msg = Msg + "버전 : v0.0.3" + vbCrLf
    Msg = Msg + "제작자 : sgt.Park"
    
    MsgBox Msg, , "Excel2TSV"

End Sub
