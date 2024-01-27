Attribute VB_Name = "Main"
'namespace=vba-files\excel2tsv\

Global gTarget As String


'Main
Public Sub RunMain(control As IRibbonControl)
    
    gTarget = ".xls" '���� ����� �����ϰ� ������ �� �κ��� ����
    
    Dim Arr() As String
    Dim Path As String
    Dim Ext As String
    
    Path = ModCommon.SelectFolder()
    If Path = "" Then
        MsgBox "�ƹ� ������ �������� �ʾҽ��ϴ�."
        Exit Sub '�������� ������ ����
    End If
    
    Call ModCommon.eachFolder(Path, Arr) 'Call by Reference
    
    If Len(Join(Arr)) = 0 Then: MsgBox "���� ���������� �����ϴ�.": Exit Sub '����ó�� �߰� 240127
    
    For i = LBound(Arr) To UBound(Arr)
        Call OpenAndSaveAsTSV(Arr(i))
    Next
    
    MsgBox "All Done"

End Sub
