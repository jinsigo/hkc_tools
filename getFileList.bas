Attribute VB_Name = "getFileList"
Option Explicit
'============================================================
' 디렉토리 파일 리스트 가져오기
'============================================================
'
Sub getFileList()
'// [도구] - [참조] 에서 Microsoft Scripting Runtime 라이브러리 체크해야 함
    Dim FSO As New FileSystemObject
    Dim sDir As Folder      '// 찾을 폴더 변수 선언
    Dim fPath As Variant    '// 경로(Path) 변수 선언
    Dim fileExt As String   '// 파일확장자 변수 선언
    Dim i, n As Long
    Dim openMsg As String
    
    On Error Resume Next     '// 에러가 발생해도 계속 수행하라
    openMsg = "파일을 가져올 경로를 직접 지정하려면 Yes를 눌러주세요 " & vbCr & vbCr
    openMsg = openMsg & "현재 경로를 선택하려면 No를 눌러주세요" & vbCr
    openMsg = openMsg & "현재 Path : " & ThisWorkbook.path + "\"
    If MsgBox(openMsg, vbYesNo) = vbYes Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            .show
            fPath = .SelectedItems(1)   '// 선택될 폴더를 경로 변수에 저장
        End With
    Else
        fPath = ThisWorkbook.path + "\"     '// 엑셀 VBA 파일이 위치한 현재경로
    End If
    If Err.Number <> 0 Or fPath = False Then Exit Sub
    On Error GoTo 0
    
    fileExt = "*.*"   '// 찾고자 하는 파일 확장자
    Worksheets("검색결과").Select     '// 다른 시트가 선택되어 있어 잘못 기록되는 경우 방지 목적
    With Range("A1:C1")
        .Value = Array("디렉토리", "파일명", "중복검사")
        .HorizontalAlignment = xlCenter
    End With
    
    Range([A1], Cells(Rows.Count, "A").End(3)).Offset(1).Resize(, 3).ClearContents
    '// 화면에 뿌릴 영역 초기화
    
   
    Call makeFileList(fPath, fileExt)   '// 파일목록 만들기 호출
    Set sDir = FSO.GetFolder(fPath)
    Call subFolderFind(sDir, fileExt)   '// 서브폴더 찾기
    
    n = Cells(Rows.Count, "B").End(3).Row - 1
    If n = 0 Then
        MsgBox "파일이 없습니다"
    Else
        MsgBox n & " 개 파일리스트 검색완료"
    End If
End Sub

Sub subFolderFind(sDir As Folder, getExt As String)
    Dim subFolder As Folder
    
    On Error Resume Next
    For Each subFolder In sDir.SubFolders
        If subFolder.Files.Count > 0 Then
            Call makeFileList(subFolder.path, getExt)
        End If
            
        If subFolder.SubFolders.Count > 0 Then
            Call subFolderFind(subFolder, getExt)
        End If
    Next
End Sub

Sub makeFileList(fPath As Variant, getExt As String)
    Dim fName As String
    Dim SaveDir As Range
    
    fName = Dir(fPath & "\" & getExt)
    If fName <> "" Then
        Do
            Set SaveDir = Cells(Rows.Count, "A").End(3)(2)
            SaveDir.Value = fPath
            SaveDir.Offset(0, 1).Value = fName
            
            fName = Dir()
        Loop While fName <> ""
        Columns("A:B").AutoFit
    End If
End Sub

