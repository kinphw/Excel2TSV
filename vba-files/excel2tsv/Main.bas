Attribute VB_Name = "Main"
'namespace=vba-files\excel2tsv\

Global gTarget As String


'Main
Public Sub RunMain(control As IRibbonControl)
    
    gTarget = ".xls" '만약 대상을 변경하고 싶으면 이 부분을 수정
    
    Dim Arr() As String
    Dim Path As String
    Dim Ext As String
    
    Path = ModCommon.SelectFolder()
    If Path = "" Then
        MsgBox "아무 폴더도 선택하지 않았습니다."
        Exit Sub '선택하지 않으면 강종
    End If
    
    Call ModCommon.eachFolder(Path, Arr) 'Call by Reference
    
    If Len(Join(Arr)) = 0 Then: MsgBox "읽을 엑셀파일이 없습니다.": Exit Sub '오류처리 추가 240127
    
    For i = LBound(Arr) To UBound(Arr)
        Call OpenAndSaveAsTSV(Arr(i))
    Next
    
    MsgBox "All Done"

End Sub
