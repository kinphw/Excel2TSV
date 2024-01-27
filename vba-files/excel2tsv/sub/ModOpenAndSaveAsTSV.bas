Attribute VB_Name = "ModOpenAndSaveAsTSV"
'namespace=vba-files\excel2tsv\sub\

'파일을 열어서 TSV로 저장
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
        
        Call DoSomething(ws) '사용자정의 조정. 추후 무언가 추가개발하고자 하는 경우
        
        Call ActiveWorkbook.SaveAs(FileName:=FileName, FileFormat:=xlText, CreateBackup:=False)
        
    Next
    
    wb.Close

End Sub
