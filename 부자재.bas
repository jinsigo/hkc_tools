Attribute VB_Name = "부자재"
'
'======================================================================
' 고객사별 부자재 재고 현황 시트 작성
'======================================================================
' 2016.6.8 이진성
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

    '시트 초기화
    If CheckUser = 0 Then Exit Sub
    sName = ActiveSheet.Name
    msg1 = "'" & sName & "' 시트를 제외한 모든 시트를 삭제합니다."
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
    
    '기준열 입력
    msg2 = "생성할 시트의 기준셀(열)을 선택해 주세요" & Chr(10) & "공백은 _ 처리됩니다."
    Set ss = Application.InputBox(msg2, cMenu, Type:=8)
    cKey = ss.Column
    
    '루프
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
            '시트 기존
            Sheets(strSplitter).Activate
            r.Rows(1).Copy
            ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Select
        Else
            '시트 신규 제목열 복사
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
    msg = MsgBox(shtCount & " 개의 시트가 만들어졌습니다", 0, "시트 생성 완료")

End Sub


