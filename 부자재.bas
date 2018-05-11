Attribute VB_Name = "������"
'
'======================================================================
' ���纰 ������ ��� ��Ȳ ��Ʈ �ۼ�
'======================================================================
' 2016.6.8 ������
'
Sub SplitTableIntoMultiSheets()
'
    Dim rngCell As Range, rngItem As Range
    Dim r As Range
    Dim i As Integer, intNum As Integer
    Dim c As Integer
    Dim bytCount As Byte
    Dim lngTotal As Long
    Dim shtSheet As Worksheet
    Dim msg As String

    Application.DisplayAlerts = False
    Set rngOrign = Selection
    If rngOrign.Rows.Count < 2 Then
        rngOrign.CurrentRegion.Select
        Set rngOrign = Selection
    Else
    End If
    
    Set shtOrigin = ActiveSheet
    Set rngTitle = rngOrign.Rows(1)
    Set rngData = Range(rngOrign.Rows(2), rngOrign.Rows(rngOrign.Rows.Count))
    Set rngClient = rngOrign.Columns(6)

    '��Ʈ �ʱ�ȭ
    If CheckUser = 0 Then Exit Sub
    sName = ActiveSheet.Name
    msg1 = "'" & sName & "' ��Ʈ�� ������ ��� ��Ʈ�� �����մϴ�."
    init = MsgBox(msg1, 3, cMenu)
    chk = 0
    
    If init = 6 And Worksheets.Count > 1 Then
         For Each s In Worksheets
            If s.Name <> sName Then s.Delete
         Next s
    ElseIf init = 2 Then
        Exit Sub
    Else
    
    End If
    
    '���ؿ� �Է�
    msg2 = "������ ��Ʈ�� ���ؼ�(��)�� ������ �ּ���" & Chr(10) & "������ _ ó���˴ϴ�."
    Set ss = Application.InputBox(msg2, cMenu, Type:=8)
    cKey = ss.Column
    
    '����
    For Each r In rngData.Rows
        strSplitter = r.Cells(1, cKey)
        If strSplitter = "" Then
            strSplitter = "Blank"
        Else
            strSplitter = Replace(strSplitter, " ", "_")
            strSplitter = RemoveSpecialChars(strSplitter)
            strSplitter = Left(strSplitter, 30)
        End If
        
        If IsSheet(strSplitter) Then
            '��Ʈ ����
            Sheets(strSplitter).Activate
            r.Rows(1).Copy
            ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Select
        Else
            '��Ʈ �ű� ���� ����
            ActiveWorkbook.Worksheets.Add After:=Worksheets(Worksheets.Count)
            With ActiveSheet
                .Name = strSplitter
                shtCount = shtCount + 1
                shtOrigin.Cells.Copy
                .Cells.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                rngTitle.Copy
                .Cells(1, 1).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            End With
        End If
        
        r.Rows(1).Copy
        ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Select
        ActiveSheet.Paste
        ActiveSheet.Cells(1, 1).Select
    Next r
    
    Application.DisplayAlerts = True
    shtOrigin.Activate
    rngOrign.Select
    msg = MsgBox(shtCount & " ���� ��Ʈ�� ����������ϴ�", 0, "��Ʈ ���� �Ϸ�")

End Sub


