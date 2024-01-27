Attribute VB_Name = "ModCommon"
'namespace=vba-files\excel2tsv\common\

Option Explicit

'폴더를 선택받음
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

'하위폴더 재귀순환하면서 대상 파일을 배열로 인식
Public Function eachFolder(mainFolder As String, ByRef Arr As Variant)
   
    Dim fso As Scripting.FileSystemObject            'FileSystemObject를 넣을 변수
    Dim fldName As Scripting.Folder                     '상위폴더 전체경로(마지막이 "\")를 넣을 변수

    Dim subFolder As Scripting.Folder                  '하위폴더 넣을 변수
    Dim rngT As Range                                       '복사될 위치 넣을 변수
    Dim wkBk As Workbook                                 '각 파일을 넣을 변수
    Dim wkSht As Worksheet                               '각 시트를 넣을 변수
    Dim FileName As Object                                 '각 파일의 이름(확장자포함)을 넣을 변수
   
    Set fso = New Scripting.FileSystemObject        '새로운 파일시스템개체를 변수에
    Set fldName = fso.GetFolder(mainFolder)         '선택한 폴더의 이름을 변수에 넣음
   
    'Dim Arr() As String
    
    For Each FileName In fldName.Files                 '폴더내 각 파일을 순환
        If InStr(FileName, gTarget) Then
            Call ModCommon.AppendArray(Arr, FileName)
        End If
    Next FileName
   
    For Each subFolder In fldName.SubFolders     '각 폴더의 하위폴더를 순환
        eachFolder subFolder.Path, Arr                        '각 폴더경로로 eachFolder함수호출
    Next subFolder
    
    Debug.Print "1"
    
End Function


''' PHW 함수 : FOR ARRAY

'배열의 크기를 반환하는 함수 -> Integer
Public Function GetArraySize(Arr As Variant) As Integer

    GetArraySize = (UBound(Arr) - LBound(Arr) + 1)

End Function

'배열에 Append하는 함수
'
Public Function AppendArray(ByRef Arr As Variant, Element As Variant) As Variant

    If Len(Join(Arr)) = 0 Then
        ReDim Preserve Arr(0 To 0)
        Arr(0) = Element
    Else
    
        ReDim Preserve Arr(LBound(Arr) To UBound(Arr) + 1)
        Arr(UBound(Arr)) = Element '마지막에 추가함
    End If
    
    AppendArray = Arr

End Function

''' PHW 함수 : FOR 확장자변환

'XLS(X)를 받아서 TSV로 반환
Public Function XlsToTsv(arg As Variant)

    arg = Replace(arg, ".xlsx", ".tsv")
    arg = Replace(arg, ".xlsm", ".tsv") '240127
    arg = Replace(arg, ".xlsb", ".tsv") '240127
    arg = Replace(arg, ".xls", ".tsv")
    
    XlsToTsv = arg

End Function

'확장자 삭제
Public Function RemoveTSV(FileName As Variant) As String

    RemoveTSV = Replace(FileName, ".tsv", "")

End Function
