Attribute VB_Name = "getFileList"
Option Explicit
'============================================================
' ���丮 ���� ����Ʈ ��������
'============================================================
'
Sub getFileList()
'// [����] - [����] ���� Microsoft Scripting Runtime ���̺귯�� üũ�ؾ� ��
    Dim FSO As New FileSystemObject
    Dim sDir As Folder      '// ã�� ���� ���� ����
    Dim fPath As Variant    '// ���(Path) ���� ����
    Dim fileExt As String   '// ����Ȯ���� ���� ����
    Dim i, n As Long
    Dim openMsg As String
    
    On Error Resume Next     '// ������ �߻��ص� ��� �����϶�
    openMsg = "������ ������ ��θ� ���� �����Ϸ��� Yes�� �����ּ��� " & vbCr & vbCr
    openMsg = openMsg & "���� ��θ� �����Ϸ��� No�� �����ּ���" & vbCr
    openMsg = openMsg & "���� Path : " & ThisWorkbook.path + "\"
    If MsgBox(openMsg, vbYesNo) = vbYes Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            .show
            fPath = .SelectedItems(1)   '// ���õ� ������ ��� ������ ����
        End With
    Else
        fPath = ThisWorkbook.path + "\"     '// ���� VBA ������ ��ġ�� ������
    End If
    If Err.Number <> 0 Or fPath = False Then Exit Sub
    On Error GoTo 0
    
    fileExt = "*.*"   '// ã���� �ϴ� ���� Ȯ����
    Worksheets("�˻����").Select     '// �ٸ� ��Ʈ�� ���õǾ� �־� �߸� ��ϵǴ� ��� ���� ����
    With Range("A1:C1")
        .Value = Array("���丮", "���ϸ�", "�ߺ��˻�")
        .HorizontalAlignment = xlCenter
    End With
    
    Range([A1], Cells(Rows.Count, "A").End(3)).Offset(1).Resize(, 3).ClearContents
    '// ȭ�鿡 �Ѹ� ���� �ʱ�ȭ
    
   
    Call makeFileList(fPath, fileExt)   '// ���ϸ�� ����� ȣ��
    Set sDir = FSO.GetFolder(fPath)
    Call subFolderFind(sDir, fileExt)   '// �������� ã��
    
    n = Cells(Rows.Count, "B").End(3).Row - 1
    If n = 0 Then
        MsgBox "������ �����ϴ�"
    Else
        MsgBox n & " �� ���ϸ���Ʈ �˻��Ϸ�"
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

