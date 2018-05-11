Attribute VB_Name = "MmainSub"
'
'======================================================================
' 찾기 함수
'======================================================================
'
Sub Go2Database()
'go2Cell() '160701 현재셀 값으로 데이타베이스 찾아가기
    Dim ws, sv As String
    sv = Selection.Value  '찾고자 하는 값
    
    Dim lngItem As Long
    Dim msg As String
 '
    Dim r As Integer
    Dim i, cn As Integer
    Dim rngDB As Range

'출력영역
    Set rngDB = ThisWorkbook.Sheets("S").Range("setup")   '데이타베이스 영역
    
'    Label1.Caption = "라벨1"
'    Label2.Caption = "라벨2"
'
    cn = rngDB.Rows.Count
    MsgBox rngDB.Address
    
    For i = 0 To cn
        With UserForm1.ListBox1
            .ColumnCount = 3
            .ColumnWidths = "150;120;100"
            .ColumnHeads = True
            .AddItem
            .List(i, 0) = rngDB.Cells(i, 1)    'DB 명칭
            .List(i, 1) = rngDB.Cells(i, 2)    '경로
            .List(i, 2) = rngDB.Cells(i, 3)    'cas
        End With
    Next i
 
 
    UserForm1.show
    Stop
    
    tmp = isOpenWrk(vbcWRLST) '원료리스트 활성화 여부
    If tmp = 0 Then
        wf = Application.GetOpenFilename("Excel Files,*.xls")
        Workbooks.Open FileName:=wf, UpdateLinks:=0
    End If
        Workbooks(vbcWRLST).Activate
        Range("Database").Find(What:=sv, LookIn:=xlValues).Activate
End Sub

Sub ingQueryDBInfo()
'선택셀로 DB정보 가져오기
    Dim wb, sv, r, smg As String
    Dim t, i, m As Integer
    Dim w As Workbooks
    Dim rs As Range
    
    sv = Selection.Value  '찾고자 하는 값
    owb = ActiveWorkbook.Name
   
    fn = edTrimPath(hkc_DB1, "\")
    wb = edTrimExtension(fn, ".")
   
     '원료리스트 활성화 여부
    If isOpenWrk(fn) = 0 Then
        MsgBox (fn & " 파일을 열겠습니다.")
        Workbooks.Open FileName:=hkc_DB1
    Else
    End If
    
    Set rs = Range(Workbooks(fn).Names("Database"))
    With rs
        r = .Find(What:=sv, LookIn:=xlValues).Row
        For i = 2 To 5
        msg = msg & Chr(10) & .Cells(r, i).Value
        Next i
    End With
    
    Workbooks(owb).Activate
    m = MsgBox(msg, vbOKCancel, sv)
    If m = 1 Then
        Selection.Cells(2, 1) = msg
    End If
End Sub

Function edTrimPath(f, spl) As String
    tmp = Split(f, spl)
    c = UBound(tmp, 1)
    edTrimPath = tmp(c)
End Function

Function edTrimExtension(f, spl) As String
    tmp = Split(f, spl)
    edTrimExtension = tmp(0)
End Function

'Asks for phrase to find then finds and marks within each cell everywhere it is found.
Sub skWordinCell()
    Dim rCell As Range, sToFind As String, iSeek As Long
    sToFind = InputBox("Enter Word / Phrase To Mark", "Criteria Request")
    If sToFind = "" Then MsgBox "Word / Phrase Required But Not Entered", , "Invalid Entry"
    
    For Each rCell In Selection 'can be any range or explicit (i.e. Range("A1:G6") instead of Selection)
    iSeek = InStr(1, rCell.Value, sToFind)
    Do While iSeek > 0
        With rCell.Characters(iSeek, Len(sToFind)).Font
            '.Name = "Arial"
            '.Size = 14
            '.Bold = True
            .Color = RGB(100, 100, 200)
        End With
        iSeek = InStr(iSeek + 1, rCell.Value, sToFind)
    Loop
  Next

End Sub

Sub Go2Today()
'
    Dim Today
    Dim c As Range
    
    Set c = Range("M:M").Find(What:=Date)
    c.Activate
        
End Sub

'
'======================================================================
' 보기(Show) 변경 함수
'======================================================================
'
Sub shGoNextSht() '150330
    ActiveSheet.Next.Select
End Sub

Sub shGoPrevSht() '150330
    ActiveSheet.Previous.Select
End Sub

Sub shGoPreface() '150330 test
    Application.GoTo Sheets("Preface").Range("a1"), True
End Sub

Sub shPrintPreview() '150330
    ActiveSheet.PrintPreview
End Sub

Sub shAddSht(inp As String) '120824/시트 추가
    
    Dim ws As Worksheet
    Dim chk As Integer
    chk = 0
    For Each ws In ActiveWorkbook.Sheets
        If ws.Name = inp Then chk = 1
    Next ws
    If chk = 0 Then
        ActiveWorkbook.Worksheets.Add After:=Worksheets(Worksheets.Count)
        ActiveSheet.Name = inp
    End If

End Sub

Sub shLockSht() '131108/jinsigo 시트 잠그기
'
    Cells.Locked = False
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlUp) + 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Locked = True
    ActiveSheet.Protect Password:="1", DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowInsertingColumns:=True, AllowInsertingRows:=True
    ActiveWorkbook.Save
    Selection.End(xlDown).Select

End Sub

Sub shUnLockSht() '131108/jinsigo 시트 풀기
'
    ActiveSheet.Unprotect Password:="1"
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Locked = False
    Selection.End(xlDown).Select
    
End Sub

Public Sub Clr_Sheet(inp As String) '120824/시트 내용 지우기
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        If ws.Name = inp Then
            ws.Cells.ClearContents
        End If
    Next ws
End Sub

Sub Del_Sheet(RefName As String)
'120825/시트 삭제
    Set wkb = ActiveWorkbook
    For Each ws In wkb.Names
        If ws = RefName Then wkb.Sheets(RefName).Delete
    Next ws
End Sub


Sub test()
   Del_Sheet ("IN")
End Sub

Sub shRmList()
'
' 매크로2 매크로
'

'

    Windows("원료LIST.xls").Activate
    With ActiveWindow
        .Width = 1450
        .Height = 470.25
        .Top = 1.75
        .Left = 143.5
        Range("A1").Select
    End With
'    Windows.Arrange ArrangeStyle:=xlTiled
End Sub

Sub arrange()
        
    ActiveWorksheet.Windows.arrange ArrangeStyle:=xlArrangeStyleVertical, _
        ActiveWorkbook:=False, SyncVertical:=False
        
    Workbooks("원료LIST.xls").Windows.arrange ArrangeStyle:=xlArrangeStyleVertical, _
        ActiveWorkbook:=False, SyncVertical:=False

    ActiveWindow.WindowState = xlMaximized
    ActiveWindow.WindowState = xlMinimized
    
    Windows.CompareSideBySideWith "원료LIST.xls"
    Windows.ResetPositionsSideBySide
    Windows.arrange ArrangeStyle:=xlVertical
    Windows("원료LIST.xls").Activate
    

End Sub

Sub DualViewButton_Click()
  Dim windowToPutOnTimeline As Window

  If Windows.Count = 1 Then
    ThisWorkbook.NewWindow
    Windows.arrange xlArrangeStyleHorizontal, True, False, False
    Set windowToPutOnTimeline = Windows(1)
    If Windows(1).Top < Windows(2).Top Then
      Set windowToPutOnTimeline = Windows(2)
    End If

    With windowToPutOnTimeline
      .Activate
      HorizontalTimelineSheet.Activate
      .DisplayGridlines = False
      .DisplayRuler = False
      .DisplayHeadings = False
      .DisplayWorkbookTabs = False
      '.EnableResize = False
    End With

    Windows(2).Activate 'go back to the right focus the user expects.

  Else
    If Windows(1).Top = Windows(2).Top Then
      Windows.arrange xlArrangeStyleHorizontal, True, False, False
    Else
      Windows.arrange xlArrangeStyleVertical, True, False, False
    End If
  End If
End Sub

'
'======================================================================
' 추출관련 함수
'======================================================================
'
Sub exCol_Filter2() '061026/처방추출기

    act_row = ActiveCell.Row
    act_col = ActiveCell.Column
    
    For cur_col = act_col To 220
    
        With Cells(act_row, cur_col)
            If .Value > 0 Then
                .EntireColumn.Hidden = False
            Else
                .EntireColumn.Hidden = True
            End If
            
        End With
    Next cur_col
   
End Sub

Sub Col_Filter() 'jinsigo/061027/처방검색 Macro

    r = ActiveCell.Row
    c = Range("list").Column
    Num_cols = Range("list").Columns.Count
    
    For cc = c To c + Num_cols
        With Cells(r, cc)
            If .Value <> "" Then
                .EntireColumn.Hidden = False
            Else
                .EntireColumn.Hidden = True
            End If
            
        End With
    Next cc
    
   
End Sub

Sub exActiveCell() 'jinsigo/061110/행추출 (현재 셀 기준)

    Selection.AutoFilter Field:=ActiveCell.Column, Criteria1:=ActiveCell.Value

End Sub


Sub exInput() ' 유형추출 Macro 06-12-07
'
    Kwds = Application.InputBox(prompt:="검색어 입력: ")
    Keywords1 = "=*" & Kwds & "*"
    Selection.AutoFilter Field:=ActiveCell.Column, Criteria1:=Keywords1, Operator:=xlAnd

End Sub


Sub exNoBlank() 'jinsigo/061012/행추출 (공백 아닌 셀 기준)

    Selection.AutoFilter Field:=ActiveCell.Column, Criteria1:="<>"

End Sub


Sub exShowAll() '데이터 자동필터에서 모두보기 06-11-10

    ActiveSheet.ShowAllData
End Sub

Sub 월변경()
'
' 월변경 매크로
'

'
Dim m As Integer
Dim Keywords1 As String
    Kwds = ActiveSheet.Range("I1").Value
    Keywords1 = "=*" & Kwds & "*"
    ActiveSheet.Range("$A$2:$A$20000").AutoFilter Field:=1, Criteria1:=Keywords1, Operator:=xlAnd
End Sub

Sub exSameBColor_HideColumns()
' 같은 바탕색 열 숨기기
' 160615
'
    '전체보기
    Range("list").EntireColumn.Hidden = False
    
    '선택보기
    Set rs = Selection
    rc = rs.Interior.Color
    r = rs.Row
    ln = Range("list").Columns.Count
    
    Set rss = Range(Cells(r, 1), Cells(r, ln))
    
    
    For Each c In rss.Columns
        If c.Interior.Color <> rc Then c.EntireColumn.Hidden = True
               
    Next c
End Sub

Sub exSameBColor_ShowColumns()
' 선택영역 모든 열 보이기
    Range("list").Columns.Select
    Selection.EntireColumn.Hidden = False
End Sub

Sub exSameBColor_HideRows()
' 같은 바탕색 행 숨기기
' 160615
'
    '전체보기
    Range("list").EntireRows.Hidden = False
    
    '선택보기
    Set rs = Selection
    cc = rs.Interior.Color
    c = rs.Column
    ln = Range("list").Rows.Count
    
    Set rss = Range(Cells(1, c), Cells(ln, c))
    
    
    For Each r In rss.Columns
        If r.Interior.Color <> cc Then r.EntireColumn.Hidden = True
               
    Next r
End Sub



'
'======================================================================
'/// 셀 편집 모듈 by Jinsigo ///
'======================================================================
'
Sub edMergeCell()
'120813/선택범위의 문자열 합치기
'150421 행 & 열 범위 확장
'
    Dim rs As Range  '입력셀
    Dim ro As Range  '출력셀
    Dim ss As String '구분자
    Dim st As String '병합 문자열
    Dim tmp As String
    Dim i As Range
  
    Set rs = Selection
    ss = InputBox("구분자를 입력해 주세요(기본값은 ',' 입니다).", "셀 병합하기", ",") '구분자 입력
    st = ""
    
    For Each i In rs.Rows
        For j = 1 To i.Columns.Count
            tmp = i.Cells(1, j)
            st = st + tmp & ss
        Next j
    Next i
    
    Set ro = Application.InputBox("값을 출력할 셀을 선택해 주세요", "셀 병합하기", Type:=8)
    If Right(st, Len(ss)) = ss Then st = Left(st, Len(st) - Len(ss))
    ro.Formula = st
End Sub

Sub edMergeRange() '선택범위 합치기 150520
    Dim rs As Range  '입력범위
    Dim ro As Range  '출력범위
    Dim cs As String '조건열1
    Dim co As String '조건열2
    Dim ns As Integer '입력열수
    Dim no As Integer '출력열수
    Dim vs, vo As String '입출력 값
    Dim ks, ko As String '같은 조건 유무 체크
    Dim chk As Integer ''같은 조건 유무 체크
  
    Set rs = Application.InputBox("병합할 영역을 선택해 주세요", "셀 병합하기:A", Type:=8)
        
    Set ro = Application.InputBox("병합할 영역을 선택해 주세요", "셀 병합하기:B", Type:=8)
    'If Right(st, Len(ss)) = ss Then st = Left(st, Len(st) - Len(ss))
    
    ns = rs.Columns.Count
    no = ro.Columns.Count
    chk = 0
    ro.Cells(1, no).Formula = rs.Cells(1, ns).Value
    For i = 2 To rs.Rows.Count
        cs = rs.Cells(i, 1)
        ks = rs.Cells(i, 2)
        vs = rs.Cells(i, ns)
        
        For j = 2 To ro.Rows.Count
            co = ro.Cells(j, 1)
            ko = ro.Cells(j, 2)
            vo = ro.Cells(j, no)
            If (cs = co) And (ks = ko) And (vs > 0) Then
                ro.Cells(j, no).Formula = rs.Cells(i, ns).Value
                chk = 0
                Exit For
            ElseIf ko = ks Then '코드만 다른경우
                chk = j + 1
            Else '코드 및 구분 모두 다른 경우
                
            End If
        Next j
        
        If chk And (vs > 0) Then
            ro.EntireRow(chk).Insert
            ro.Cells(chk, 1).Formula = rs.Cells(i, 1).Value
            ro.Cells(chk, 2).Formula = rs.Cells(i, 2).Value
            ro.Cells(chk, 3).Formula = rs.Cells(i, 3).Value
            ro.Cells(chk, no).Formula = rs.Cells(i, ns).Value
            chk = 0
           
        Else
        End If
    Next i
    

End Sub

Sub edSplitTXT2Cell()
'120813/선택셀의 문자열 나누기
    Dim rs As Range '입력셀
    Dim ro As Range '출력셀
    Dim ss As String '구분자
    Dim st  '나눈 문자열
    Dim nc As Integer '출력셀 행수
    
    Set rs = Selection
    ss = InputBox("구분자를 입력해 주세요(기본값은 ',' 입니다).", "셀 나누기", ",")
    st = Split(rs.Value, ss)
        
    Set ro = Application.InputBox("출력할 첫 셀을 선택해 주세요", "셀 나누기", Type:=8)
    nc = ro.Columns.Count
    
    For i = 0 To Round((UBound(st, 1) / nc), 1)
        For j = 1 To nc
            If (i * nc + j) <= UBound(st, 1) + 1 Then
                'MsgBox nc & ": " & i & ":" & j
                ro.Cells(i + 1, j).Formula = Trim(st((i * nc) + j - 1))
            End If
        Next j
    Next i
End Sub

Sub edSplitTXT2Cells()
'
'140923/문자나누기.다중셀
'150114/변수 변경, for문 j 추가.
    Dim inCells As Range '입력셀
    Dim rc As Integer '입력 행수
    Dim cc As Integer '입력 열수
    Dim outCells As Range '출력셀
    Dim sp As String '구분자
    Dim st  '나눈 문자열
    Dim stc '나눈 문자열수
    
    Set inCells = Selection
    rc = inCells.Rows.Count
    cc = inCells.Columns.Count
    
    sp = InputBox("구분자를 입력해 주세요(기본값은 ',' 입니다).", "셀 나누기", ",")
    Set outCells = Application.InputBox("출력할 첫 셀을 선택해 주세요", "셀 나누기", Type:=8)
    
    For i = 1 To rc
        st = Split(inCells.Cells(i, 1), sp)
        stc = UBound(st) + 1
        For j = 1 To stc
            '느낌표 왼쪽의 값은 split_value(0) 배열에 값이 입력되고 오른쪽은 split_value(1) 배열에 값이 입력됩니다.
           ' inCells.Offset(i - 1, 0).Formula = Trim(st(0)) '앞에 값을 뿌려줍니다.
            outCells.Offset(i - 1, j - 1).Formula = Trim(st(j - 1)) '뒤에 값을 뿌려줍니다.
        Next j
    Next i
End Sub


Sub edFillRng_NA2Null() '110415/선택범위의 #NA 값을 공백으로 변경하기 매크로
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Replace What:="#N/A", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

Sub edTxt2Char() '120119/특수기호 변환 key: Ctrl+Shift+A
    ActiveCell.Replace What:="-^", Replacement:="↑", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    ActiveCell.Replace What:="&^", Replacement:="↑"
    ActiveCell.Replace What:="&v", Replacement:="↓"
    ActiveCell.Replace What:="&>", Replacement:="→"
    ActiveCell.Replace What:="&<", Replacement:="←"
    ActiveCell.Replace What:="&c", Replacement:="℃"
    ActiveCell.Replace What:="&.", Replacement:="·"
End Sub

Sub edRmvSpecialChar()
    Set rs = Selection
    For Each c In Selection
        For i = 1 To 31
            c.Value = Replace(c.Value, Chr(i), "")
        Next i
    Next
    
End Sub

Function edGetNums(strIN)
'숫자만 읽기
    Dim RegExpObj As Object
    Dim NumStr As String

    Set RegExpObj = CreateObject("vbscript.regexp")
    With RegExpObj
        .Global = True
        .Pattern = "[^\d]+"
        NumStr = .Replace(strIN, vbNullString)
    End With

    edGetNums = NumStr
End Function

Sub edTrimRanges()
'150611
    Dim rs As Range
    
    Set rs = Selection
    For Each c In rs.Cells
            c.Value = Trim(c)
    Next c
    
End Sub



Sub edChange_Pic_Name() '150114/그림이름 바꾸기
    Dim shpC As Shape                                    '각각의 그림을 넣을 변수
    Dim rngShp As Range                                  '각 그림의 왼쪽위가 속한 영역을 넣을 변수
   
    For Each shpC In ActiveSheet.Shapes          '삭제영역내의 각 그림을 순환
        Set rngShp = shpC.TopLeftCell                 '각 그림의 왼쪽위지점이 속한 영역을 변수에 넣음
       
        If Not Intersect(Columns("B"), rngShp) Is Nothing Then   '그림과 B열이 겹치면
                shpC.Name = rngShp.Previous.Value    '각 그림의 이름을 A열의 이름으로 변경
        End If
    Next shpC
   
End Sub

Option Explicit

Sub fitPictureInCell() '[출처] (622) 선택영역내 그림을 각 셀의 크기에 맞추기 (엑셀 VBA 매크로)|작성자 니꾸
    Dim rngAll As Range                                    '선택영역을 넣을 변수
    Dim rngShp As Range                                  '각 그림의 왼쪽위가 속한 영역을 넣을 변수
    Dim shpC As Shape                                    '각각의 도형(shape)을 넣을 변수
    Dim rotationDegree As Integer                       '도형의 회전각도 넣을 변수
   
    Application.ScreenUpdating = False              '화면 업데이트 (일시)정지
   
    If Not TypeOf Selection Is Range Then           '만일 그림 등을 선택하거나 하였을 경우
        MsgBox "영역이 선택되지 않음", 64, "영역선택 오류"  '경고 메시지 출력
        Exit Sub                                                 '매크로 중단
    End If
   
    Set rngAll = Selection                                  '선택영역을 변수에 넣음
   
    For Each shpC In ActiveSheet.Shapes          '전체영역내 각 그림을 순환
        If shpC.Type = 13 Then                            '만일 각 도형이 그림이라면
            Set rngShp = shpC.TopLeftCell             '각 도형의 왼쪽위 영역을 변수에 넣음
           
            If rngShp.MergeCells Then                   'rngShp가 셀병합된 셀이라면
                Set rngShp = rngShp.MergeArea       '영역을 셀병합 영역으로 확장
            End If
           
            If Not Intersect(rngAll, rngShp) Is Nothing Then '각 도형이 전체영역에 포함되면
                rotationDegree = shpC.Rotation         '그림의 회전각을 변수에 넣음
               
                If rotationDegree = 90 Or rotationDegree = 270 Then '그림이 90도 or 270도 회전된 경우
               
                    With shpC                                   '각 그림으로 작업
                        .LockAspectRatio = msoFalse   '그림 좌우고정비율 해제
                        .Rotation = 0                           '그림 회전을 원상태로 돌려 놓음
                        .Height = rngShp.Width - 4        '그림 높이를 현재셀 크기  - 4
                        .Width = rngShp.Height - 4        '그림 폭을 현재셀 크기 - 4
                        .Left = rngShp.Left + (rngShp.Width - shpC.Width) / 2
                                                                    '그림 폭 가운데 위치가 셀의 중앙에 오도록 정렬
                        .Top = rngShp.Top + (rngShp.Height - shpC.Height) / 2
                                                                    '그림위쪽 가운데 위치가 셀의 중앙에 오도록 정렬
                        .Rotation = rotationDegree        '그림 회전 각도를 복원
                    End With
               
                Else
                    With shpC                                  '각 그림으로 작업
                        .LockAspectRatio = msoFalse  '그림 좌우고정비율 해제
                        .Left = rngShp.Left + 2             '그림왼쪽위치를 셀의 왼쪽위 + 2
                        .Top = rngShp.Top + 2             '그림위쪽 위치를  셀의 왼쪽위 위치 + 2
                        .Height = rngShp.Height - 4      '그림 높이를 현재셀 크기  - 4
                        .Width = rngShp.Width - 4        '그림 폭을 현재셀 크기 - 4
                    End With
                End If
            End If
        End If
    Next shpC
   
    Set rngAll = Nothing                                      '개체변수 초기화(메모리 비우기)
End Sub
   
'
'======================================================================
' 시트 추가 편집
'======================================================================
'
Sub AddFormatCondition()
    With ActiveSheet.Range("A1:A10").FormatConditions _
        .Add(xlCellValue, xlEqual, "vba팀")
        .Font.Color = vbRed
    End With
 ''
End Sub

Sub ModifyFormatCondition()
    With ActiveSheet.Range("M1:M1000").FormatConditions(1)
        .Modify xlCellValue, xlEqual, "A"
          .ThemeColor = xlThemeColorAccent1

    End With
    
End Sub

Sub DeleteFormatCodition()
    ActiveSheet.Range("M1:M1000").FormatConditions.Delete
End Sub

'
'======================================================================
' 찾기 및 정보 함수
'======================================================================
'
Function isOpenName(nn As String) As Integer
'이름정의 존재 여부 체크 150318
    Dim nam As Name
    isOpenName = 0
    
    For Each nam In ActiveWorkbook.Names
        If nam.Name = nn Then
            isOpenName = isOpenName + 1
        Else
            isOpenName = isOpenName
        End If
    Next nam
    
End Function

Function isOpenWrk(wb) As Integer
'워크시트 열림 여부 체크 150318
    Dim wrk As Workbook
    isOpenWrk = 0
    
    For Each wrk In Workbooks
        If wrk.Name = wb Then
            'MsgBox wrk.name & "가 이미 열려 있습니다."
            isOpenWrk = isOpenWrk + 1
        Else
            isOpenWrk = isOpenWrk
        End If
    Next wrk
    
End Function

Sub isInfoModule()
 Debug.Print Application.VBE.ActiveVBProject.VBComponents(1).CodeModule.CountOfLines
 Debug.Print Application.VBE.ActiveVBProject.VBComponents(1).CodeModule.Name
End Sub

Sub isMatchRow()
' 선택영역의 첫행을 필드명의로 DB 번
    st = Selection.Value
    'r = Application.Match(st, "[원료LIST.xls]!database").rows(1), 0)
    MsgBox r
End Sub


Sub isCallReadValue()
    Dim strPath As String
    Dim strFile As String
    Dim strSheet As String
    Dim strAddress As String
    Dim sht As Worksheet
    Dim r As Long
    Dim c As Integer
    
    Application.ScreenUpdating = True
    Set sht = Worksheets.Add
    ActiveWindow.DisplayGridlines = False
    strPath = ThisWorkbook.path
    strFile = "지점별실적.xls"
    strSheet = "Sheet1"
    
    For r = 1 To 11
        For c = 1 To 5
            strAddress = Cells(r, c).Address
            If ReadValue(strPath, strFile, strSheet, strAddress) = 0 Then Exit For
            Cells(r, c) = ReadValue(strPath, strFile, strSheet, strAddress)
        Next c
    Next r
    
    Selection.AutoFormat Format:=xlRangeAutoFormatList1, Number:=True, Font:= _
        True, Alignment:=True, Border:=True, Pattern:=True, Width:=True
    Rows("1:2").Insert Shift:=xlDown
    
    With ActiveCell
        ActiveSheet.Buttons.Add(.Left, .Top, .Width * 2, .Height).Select
        With Selection
            .Caption = "<<돌아가기"
            .OnAction = "GoBack"
        End With
    End With
    Range("a1").Select
    MsgBox "자료를 모두 읽어들였습니다", vbInformation, "작업 종료//Exceller"
End Sub

Function isReadValue(path, file, sht, rng) As Variant
    Dim msg As String
    Dim strTemp As String
    
    If Trim(Right(path, 1)) <> "\" Then path = path & "\"
    If Dir(path & file) = "" Then
        ReadValue = "해당 파일이 없습니다"
        Exit Function
    End If
    msg = "'" & path & "[" & file & "]" & sht & "'!" & Range(rng).Range("a1").Address(, , xlR1C1)
    ReadValue = ExecuteExcel4Macro(msg)
End Function

Public Function IsFormula(c)
'함수 여부를 체크 150518
    IsFormula = c.HasFormula
   
End Function

'
'======================================================================
' 파일 입출력 함수
'======================================================================
'
Function rPreset(nfile As String) As Range
' 이름정의표로 파일 열기     150401
    
    Dim nam, path, file, sht As String
    Dim rng As Range
    Dim rs As Range
    Dim c As Range
    
    Set rs = Workbooks("HKC.xlsm").Sheets("R").Range("a:k") '참고 영역
    
    For Each c In rs.Rows
        If c.Cells(1, 1).Value = nfile Then
            i = c.Row
            Set rs = c.Cells(i, 1).Resize(1, 5)
            MsgBox "i= " & i
        Else
            MsgBox "해당 이름이 정의되지 않았습니다."
        End If
        
    Next
    MsgBox rs.Address
    With rs.Rows(i)
        nam = .Cells(1, 1).Value
        path = .Cells(1, 2).Value
        file = .Cells(1, 3).Value
        sht = .Cells(1, 4).Value
        rng = .Cells(1, 5).Value
        ActiveWorkbook.Names.Add Name:=nam, RefersToR1C1:=rng
    
    End With
    If Trim(Right(path, 1)) <> "\" Then path = path & "\"
    If Dir(path & file) = "" Then
        MsgBox "해당 파일이 없습니다"
        Exit Function
    End If
    
    rPreset = "'" & path & "[" & file & "]" & sht & "'!" & rng
       
End Function

Function ioOpenWrk(wb) As Integer
'엑셀파일 열기 150403
    tmp = isOpenWrk(wb)
    If tmp = 0 Then Workbooks.Open wb
    ioOpenWrk = tmp
End Function




Function ioOpenBydNames() As Integer
'이름정의영역.이름으로 엑셀파일 열기 150403
    Dim nwbk As String
    Dim rs As Range
    
    Set rs = rdefNames
    For Each i In rs
        ndir = i.Cells(1, 2).Value  ' 경로
        nwbk = i.Cells(1, 3).Value  ' 파일명
        nvis = i.Cells(1, 6).Value  ' 1:숨기기 0:보이기
        nrdo = i.Cells(1, 7).Value '1: 읽기전용
        'MsgBox ndir & nwbk & i
        If i.Cells(1, 1).Interior.ColorIndex <> -4142 And isOpenWrk(nwbk) = 0 Then
            Workbooks.Open FileName:=(ndir & "\" & nwbk), UpdateLinks:=0, ReadOnly:=nrdo
            If nvis = 1 Then ActiveWindow.Visible = False
        End If
    Next i
            
End Function

Sub ioOpendBySelection()
'선택영역 이름으로 파일열기 150604
    
    Dim rs As Range
    Dim rd As Range
    Dim nwbk As String
    Dim dn As String
    Dim tmp  As String
    
    Set rs = Selection
    Set rd = rdefNames
    
    For Each s In rs
        tmp = s.Cells(1, 1).Value
        If isOpenWrk(tmp) = 0 Then
            For Each i In rd
                If i.Cells(1, 1).Value = tmp Then
                    ndir = i.Cells(1, 2).Value  ' 경로
                    nwbk = i.Cells(1, 3).Value  ' 파일명
                    nvis = i.Cells(1, 6).Value  ' 1:숨기기 0:보이기
                    Workbooks.Open FileName:=(ndir & "\" & nwbk), ReadOnly:=1, UpdateLinks:=0
                    If nvis = 1 Then ActiveWindow.Visible = False
                End If
            Next i
        End If
    Next s
End Sub


Sub ioOpendByActiveCell2()
'이름정의영역.이름으로 엑셀파일 열기 150403
    Dim nwbk As String
    Dim rs As Range
    
    dn = Selection.Value
    Set rs = rdefNames
    
    i = rs.Columns(1).Find(dn).Row - 1
    MsgBox i
    ndir = rs.Cells(i, 2).Value  ' 경로
    nwbk = rs.Cells(i, 3).Value  ' 파일명
    nvis = rs.Cells(i, 6).Value  ' 1:숨기기 0:보이기
    'MsgBox ndir & nwbk & i
    If isOpenWrk(nwbk) = 0 Then
        Workbooks.Open FileName:=(ndir & "\" & nwbk), UpdateLinks:=0, ReadOnly:=1
        If nvis = 1 Then ActiveWindow.Visible = False
    End If
            
End Sub



Sub isOpenXL()
    st = "원료LIST"
    MsgBox rPreset(st)
    Workbooks.Open rPreset(st).Cells(1, 1).Value
End Sub
Sub ioLoadModules() '150212/jinsigo 모듈 가져오기
    
    Dim st(10) As String
    
    ChDir ("D:\APP\VBA\BAS\")
    path = "D:\APP\VBA\BAS\"
    
    With ThisWorkbook
        r = ActiveCell.Row
        For i = 1 To 5
            st(i) = .ActiveSheet.Cells(r, i).Value  ' sub,desc,date,module,file
        Next i
        
    
        MsgBox st(5) & st(4) & st(3) & st(2) & st(1)
        fn = path & st(1) & ".BAS"
        MsgBox fn
        .VBProject.VBComponents(st(4)).Import FileName:=fn
    End With
End Sub

Sub ioDumpModules()  '150212/jinsigo 모듈 버리기

    ThisWorkbook.VBProject.VBComponents.Remove ActiveWorkbook.VBProject("HKC").VBComponents("module1")

End Sub

Sub IOBAS()

    Dim wb As Excel.Workbook
    Dim VBProj As VBIDE.VBComponent
    Dim lVbComp As Long
    Dim i As Long
    
    wb = xlapp.ActiveWorkbook
    lVbComp = wb.VBProject.VBComponents.Count

'Loop existing modules and get Sheet1
    For i = 1 To lVbComp
        Debug.Print wb.VBProject.VBComponents(i).Name
        If wb.VBProject.VBComponents(i).Name = "Sheet1" Then
            VBProj = wb.VBProject.VBComponents(i)
            Exit For
        End If
    Next

'Get existing code in Sheet1 code module
    For i = 1 To VBProj.CodeModule.CountOfLines
        Debug.Print VBProj.CodeModule.Lines(i, 1)
    Next

'Add an event to the Sheet1 code module

    VBProj.CodeModule.CreateEventProc "Activate", "Worksheet"
End Sub


Sub ioGo2Ref()
'
' 함수 찾아 가기
'

'
    st = Selection.Value
    Application.GoTo Reference:=st
End Sub



'*****
' Source Code: esKillErrName.esAPI  가져옴15.06.26

'Private Const MAX_PATH As Integer = 255
'Private Declare Function GetSystemDirectory& Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long)
'Private Declare Function GetWindowsDirectory& Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long)
'Private Declare Function GetTempDir Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
        
Function isReturnTempDir()
'임시 폴더명 반환 Returns Temp Folder Name
Dim strTempDir As String
Dim lngx As Long
    strTempDir = String$(MAX_PATH, 0)
    lngx = GetTempDir(MAX_PATH, strTempDir)
    If lngx <> 0 Then
        isReturnTempDir = Left$(strTempDir, lngx)
    Else
        isReturnTempDir = ""
    End If
End Function

Function isReturnSysDir()
'시스템 폴더명 반환 (C:\WinNT\System32)
Dim strSysDirName As String
Dim lngx As Long
    strSysDirName = String$(MAX_PATH, 0)
    lngx = GetSystemDirectory(strSysDirName, MAX_PATH)
    If lngx <> 0 Then
        isReturnSysDir = Left$(strSysDirName, lngx)
    Else
        isReturnSysDir = ""
    End If
End Function
Function dhReturnWinDir()
'OS 폴더명 반환 (C:\Win95)
Dim strWinDirName As String
Dim lngx As Long
    strWinDirName = String$(MAX_PATH, 0)
    lngx = GetWindowsDirectory(strWinDirName, MAX_PATH)
    If lngx <> 0 Then
        isReturnWinDir = Left$(strWinDirName, lngx)
    Else
        isReturnWinDir = ""
    End If
End Function
'*****

Sub GetSysInfo()
' 컴퓨터 환경 변수 읽어오기
    Dim rs As Range: Set rs = Sheets.Add.Range("A1")
    Dim i As Integer: i = 1
    Dim j As Integer
    Dim st As String
    rs.Cells(1, 1) = "operating system environment variables"
    Do
        st = Environ(i)
        If st = "" Then Exit Do
        j = InStr(1, st, "=")
        rs.Offset(i, 0).Value = Left(st, j)
        rs.Offset(i, 1).Value = Mid(st, j + 1, Len(st) - j)
        i = i + 1
    Loop Until st = ""
End Sub





Sub copyFormula()
'
' 매크로3 매크로
'

'
    Dim wd, ws As Object
    Dim rs, rd As Range
    
    Set ws = ActiveWindow
    Set wd = Windows("성분표기작성5.0.xlsm")
    Set rs = Selection
    
'    Set rd = wd.Sheets("작성").Range("A6")
    
    'Windows("20150828100214.CDVSK.xls").Activate
    'Range("E2").Select
    Application.Run "성분표기작성5.0.xlsm!Clear_Data"
    
    ws.Activate
    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    wd.Activate
    Range("A6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ws.Activate
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    wd.Activate

    Range("A6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="JU", Replacement:="JU-", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=True, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Range("B6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range(Selection, Selection.End(xlDown)).Select
    For Each r In Selection
        r.Cells(1, 1) = r.Cells(1, 1) * 100
        
    Next r
    
    
End Sub


Sub mkSheets()
'
' 전성분 시트 작성
' 2015.08.28

'
    Dim wd, ws As Object
    Dim rs, rd As Range
    
    Set ws = ActiveWindow
    Set wd = Windows("성분표기작성5.0.xlsm")
    Set rs = Selection
    Set rd = Worksheets("전성분").Range("A6")
    
    'Windows("20150828100214.CDVSK.xls").Activate
    'Range("E2").Select
    ws.Activate
    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    wd.Activate
    Range("A6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ws.Activate
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    
    wd.Activate

    Range("A6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="JU", Replacement:="JU-", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=True, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Range("B6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range(Selection, Selection.End(xlDown)).Select
    For Each r In Selection
        r.Cells(1, 1) = r.Cells(1, 1) * 100
        
    Next r
    
    
End Sub



Function rdefNames() As Range
'이름정의 영역 가져오기150403
    Dim sc As Range '시작셀
    'Set sc = Workbooks("HKC.xlsm").Sheets("R").Range("setup")
    'Set rdefNames = Range(sc, sc.End(xlUp))
    Set rdefNames = Workbooks("HKC.xlsm").Sheets("S").Range("setup")
    
End Function

Sub ioDefines()
'사용자정의 영역(색상 부분) 이름정의 150406

    Dim rs As Range
    Dim nam As String
    Dim ck As Integer
    'Dim ref As Range
    Set rs = Workbooks("HKC.xlsm").Sheets("S").Range("setup")
    Set wbk = ActiveSheet
    For Each i In rs.Rows
        With i
        nam = .Cells(1, 1).Value
        path = .Cells(1, 2).Value
        file = .Cells(1, 3).Value
        sht = .Cells(1, 4).Value
        rng = .Cells(1, 5).Value
        
        ref = "='" & file & "'!" & rng
        
        ck = isOpenName(nam)
        If ck = 1 Then ActiveWorkbook.Names(nam).Delete

        If i.Cells(1, 1).Interior.ColorIndex <> -4142 Then
            ActiveWorkbook.Names.Add Name:=nam, RefersToR1C1:=ref
        End If
                   
        End With
    Next i
    wbk.Activate
    
End Sub

Sub Cells2defNames()
' 리스트 to 이름정의 150403
'
    Dim ndef As String
    Dim ndir As String
    Dim nwbk As String
    Dim nsht As String
    Dim nrng As String
    Dim ncom As String
    Dim rdef As String '참고영역(경로포함)
    Dim wbk As Workbook
    Dim r As Range
        
    Set rs = Sheets("R").Range("A2:F2", Cells(2, 1).End(xlDown))

    For Each r In rs.Rows
        ndef = r.Cells(1, 1).Value
        ndir = r.Cells(1, 2).Value
        nwbk = r.Cells(1, 3).Value
        nsht = r.Cells(1, 4).Value
        nrng = r.Cells(1, 5).Value
        ncom = r.Cells(1, 6).Value
        rdef = "[" & ndir & "\" & nwbk & "]" & nsht & "!" & nrng
        'MsgBox rdef
        
        'If isOpenwbk(nwbk) = 0 Then Workbooks.Open Filename:=ndir & "\" & nwbk
        
        With ActiveWorkbook
                .Activate
                If isOpenName(ndef) > 0 Then .Names(ndef).Delete
                .Names.Add Name:=ndef, RefersTo:=rdef
        End With
        
    Next
End Sub

Sub redefineNames()

    Set rs = Sheets("S").Range("A2:F2", Cells(2, 1).End(xlDown))

    With ActiveWorkbook.Names("데이타")
        .Name = "데이타2"
        .RefersToR1C1 = "='D:\Documents and Settings\이진성\My Documents\1.원료성분\데이타.xls'!list"
        .Comment = ""
    End With


    For Each r In rs.Rows
        ndef = r.Cells(1, 1).Value
        ndir = r.Cells(1, 2).Value
        nwbk = r.Cells(1, 3).Value
        nsht = r.Cells(1, 4).Value
        nrng = r.Cells(1, 5).Value
        ncom = r.Cells(1, 6).Value
        rdef = "[" & ndir & "\" & nwbk & "]" & nsht & "!" & nrng
        'MsgBox rdef
        
        'If isOpenwbk(nwbk) = 0 Then Workbooks.Open Filename:=ndir & "\" & nwbk
        
        With ActiveWorkbook
                .Activate
                If isOpenName(ndef) > 0 Then .Names(ndef).Delete
                .Names.Add Name:=ndef, RefersTo:=rdef
        End With
        
    Next

End Sub

