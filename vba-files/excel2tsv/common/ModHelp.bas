Attribute VB_Name = "ModHelp"
'namespace=vba-files\excel2tsv\common\

Public Sub CallHelp(control As IRibbonControl)
    
    Dim Msg As String
    
    Msg = ""
    Msg = Msg + "�������ϵ��� �ִ� ������ �����ϸ� ��� �ؽ�Ʈ����(tsv, cp949)�� ��ȯ�մϴ�." + vbCrLf
    Msg = Msg + "���������� �ִ� ��� ���������� ��ȯ�մϴ�." + vbCrLf
    Msg = Msg + "���� : v0.0.3" + vbCrLf
    Msg = Msg + "������ : sgt.Park"
    
    MsgBox Msg, , "Excel2TSV"

End Sub
