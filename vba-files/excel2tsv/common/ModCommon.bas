Attribute VB_Name = "ModCommon"
'namespace=vba-files\excel2tsv\common\

Option Explicit

'������ ���ù���
Public Function SelectFolder()
    'PURPOSE: Have User Select a Folder Path and Store it to a variable
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    
    Dim FldrPicker As FileDialog
    Dim myFolder As String
    
    'Have User Select Folder to Save to with Dialog Box
      Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
    
      With FldrPicker
        .Title = "Select A Target Folder that has excel files"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function 'Check if user clicked cancel button
        myFolder = .SelectedItems(1) & "\"
      End With
      
    'Carry out rest of your code here....
    'MsgBox "Folder Path is: " & myFolder
    SelectFolder = myFolder
    
End Function

'�������� ��ͼ�ȯ�ϸ鼭 ��� ������ �迭�� �ν�
Public Function eachFolder(mainFolder As String, ByRef Arr As Variant)
   
    Dim fso As Scripting.FileSystemObject            'FileSystemObject�� ���� ����
    Dim fldName As Scripting.Folder                     '�������� ��ü���(�������� "\")�� ���� ����

    Dim subFolder As Scripting.Folder                  '�������� ���� ����
    Dim rngT As Range                                       '����� ��ġ ���� ����
    Dim wkBk As Workbook                                 '�� ������ ���� ����
    Dim wkSht As Worksheet                               '�� ��Ʈ�� ���� ����
    Dim FileName As Object                                 '�� ������ �̸�(Ȯ��������)�� ���� ����
   
    Set fso = New Scripting.FileSystemObject        '���ο� ���Ͻý��۰�ü�� ������
    Set fldName = fso.GetFolder(mainFolder)         '������ ������ �̸��� ������ ����
   
    'Dim Arr() As String
    
    For Each FileName In fldName.Files                 '������ �� ������ ��ȯ
        If InStr(FileName, gTarget) Then
            Call ModCommon.AppendArray(Arr, FileName)
        End If
    Next FileName
   
    For Each subFolder In fldName.SubFolders     '�� ������ ���������� ��ȯ
        eachFolder subFolder.Path, Arr                        '�� ������η� eachFolder�Լ�ȣ��
    Next subFolder
    
    Debug.Print "1"
    
End Function


''' PHW �Լ� : FOR ARRAY

'�迭�� ũ�⸦ ��ȯ�ϴ� �Լ� -> Integer
Public Function GetArraySize(Arr As Variant) As Integer

    GetArraySize = (UBound(Arr) - LBound(Arr) + 1)

End Function

'�迭�� Append�ϴ� �Լ�
'
Public Function AppendArray(ByRef Arr As Variant, Element As Variant) As Variant

    If Len(Join(Arr)) = 0 Then
        ReDim Preserve Arr(0 To 0)
        Arr(0) = Element
    Else
    
        ReDim Preserve Arr(LBound(Arr) To UBound(Arr) + 1)
        Arr(UBound(Arr)) = Element '�������� �߰���
    End If
    
    AppendArray = Arr

End Function

''' PHW �Լ� : FOR Ȯ���ں�ȯ

'XLS(X)�� �޾Ƽ� TSV�� ��ȯ
Public Function XlsToTsv(arg As Variant)

    arg = Replace(arg, ".xlsx", ".tsv")
    arg = Replace(arg, ".xlsm", ".tsv") '240127
    arg = Replace(arg, ".xlsb", ".tsv") '240127
    arg = Replace(arg, ".xls", ".tsv")
    
    XlsToTsv = arg

End Function

'Ȯ���� ����
Public Function RemoveTSV(FileName As Variant) As String

    RemoveTSV = Replace(FileName, ".tsv", "")

End Function
