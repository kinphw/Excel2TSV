Attribute VB_Name = "ModOpenAndSaveAsTSV"
'namespace=vba-files\excel2tsv\sub\

'������ ��� TSV�� ����
'Arg[1] Path : To save
Public Sub OpenAndSaveAsTSV(FileName As Variant)

    Dim FileNameOriginal As Variant
    
    FileNameOriginal = FileName
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim FileNameNew As String

    Set wb = Application.Workbooks.Open(FileName:=FileName)
    
    For Each ws In wb.Worksheets
        
        FileName = XlsToTsv(FileNameOriginal)
        FileName = RemoveTSV(FileName)
        FileName = FileName + "_" + ws.Name + ".tsv"
    
        ws.Activate
        
        Call DoSomething(ws) '��������� ����. ���� ���� �߰������ϰ��� �ϴ� ���
        
        Call ActiveWorkbook.SaveAs(FileName:=FileName, FileFormat:=xlText, CreateBackup:=False)
        
    Next
    
    wb.Close

End Sub
