Attribute VB_Name = "전성분작성"
'
'======================================================================
' 전성분 작성 매크로 ver 5.0 130330
'======================================================================
'
'
'
'
Option Explicit

Public rs As Range ' 선택영역 (제목 포함)
Public ra As Range ' 선택영역 (데이타 영역만)
Public rb As Range ' Base.xls 참고영역
Public rm As Range ' 원료LIST.xls 참고영역
Public st As String ' 상태값
Public c1, c2, c3, c4, c5, c6, c7, c8, c9, c10 As Integer '열번호값
Dim i, j, nk_cn, nc_cn, ni_cn, rb_cn, next_row, opt As Integer
Dim bc, bbc As String '코드값
Dim nk() As String
Dim ni() As String
Dim nc() As String
Dim bv, tmp, temr As Variant '함량값
Dim temp, temc As String
Dim dir1, dir2, temq As String '경로


Sub ingInit()
' 변수값 정의

'On Error GoTo ErrorHandler
    Dim c(100) As Integer
    Dim t() As Variant

    'Set rs = ActiveSheet.Range("list") ' 선택영역
    Set rb = Workbooks("Base.xls").Sheets("BASE").Range("$A1:$IV20000") '원료영역
    Set rm = Workbooks("원료LIST.xls").Sheets("원료LIST").Range("database") '원료영역

    c(1) = 1  ' 코드
    c(2) = 1  ' 함량
    c(3) = 1  ' 구분(상태표시)
    c(4) = 2  ' 원료명
    c(5) = 3  ' 전성분표준화명
    c(6) = 18  ' 조성비
    c(7) = 5  ' INCI명
    c(8) = 4  ' 규격
    c(9) = 0  ' 배합한도
    c(10) = 10
    c(11) = 5 '기능1
    c(12) = 7 '기능2

    't = Array("원료코드", "함량", "구분", "원료명", "전성분", "조성비", "INCI", "규격", "배합한도", "CAS No.", "HS코드")
    i = 2

'    If Mid(rs.Cells(i, 1).Value, 3, 1) <> "-" Then rs.Columns(1).Replace what:="JU", Replacement:="JU-", SearchOrder:=xlByColumns, MatchCase:=True


ErrorHandler:
    'MsgBox Err.Number
    Select Case Err.Number    ' 오류 번호를 계산합니다.

        Case 1004    ' "파일이 이미 열려 있습니다" 오류입니다.
            rs.Cells(i, c5) = "해당 코드가 없습니다."
            ' 열린 파일을 닫습니다.
        Case Else
            ' 여기서 다른 상황을 다룹니다.
    End Select
'    Resume Next

End Sub


Sub ingAuto()
'
'
'
    ingSplitBase
    ingQueryDB
    ingSplitIng
    ingSortByIng
    ingMergeByIng
    ingSortByVol
End Sub

Sub ingClearData()
'
    ingInit
    rs.Offset(1, 0).Resize(rs.Rows.Count - 1, rs.Columns.Count).Clear
    rs.Offset(1, 0).Resize(rs.Rows.Count - 1, rs.Columns.Count).RowHeight = 12

End Sub
Sub ingMakeSheet()
' 선택영역에서 시트 생성하고 코드값 가져오기
'2016-06-24
    Dim rs, ro, r, t As Range
    Dim dc As String
    Dim dv As Variant
    Dim cs As Integer
    Dim ti As Variant

    ti = Array("원료코드", "원료명", "구분", "원료명", "전성분", "조성비", "INCI", "규격", "배합한도", "CAS No.", "HS코드")

    Set rs = Selection
    cs = rs.Columns.Count

' 시트 만들기
    shAddSht ("전성분")
    Worksheets("전성분").Activate
    Set ro = Worksheets("전성분").Range("A6")

' 제목 넣기
'    Call LoadArray(ro, ti)
                ro.Offset(0, 0) = "원료코드"
                ro.Offset(0, 1) = "원료명"
                ro.Offset(0, 2) = "INCI"
                ro.Offset(0, 3) = "CAS No."
                ro.Offset(0, 4) = "함량(w/w%)"
                ro.Offset(0, 5) = "조성비"
                ro.Offset(0, 6) = "실함량(w/w%)"
                ro.Offset(0, 7) = "규격"
                ro.Offset(0, 8) = "Function"

 '데이타 넣기
    i = 1
    For Each r In rs.Rows
            dc = r.Cells(1, 1)
            dv = r.Cells(1, cs).Value
            If IsNumeric(dv) And (dv > 0) Then
                ro.Offset(i, 0) = dc
                ro.Offset(i, 1) = "=VLOOKUP(" & ro.Offset(i, 0).Address & ",원료LIST,2,0)"
                ro.Offset(i, 2) = "=VLOOKUP(" & ro.Offset(i, 0).Address & ",원료LIST,5,0)"
                ro.Offset(i, 3) = "=VLOOKUP(" & ro.Offset(i, 0).Address & ",원료LIST,16,0)"
                ro.Offset(i, 4) = Format(dv, "#,##0.000")
                ro.Offset(i, 5) = "=VLOOKUP(" & ro.Offset(i, 0).Address & ",원료LIST,18,0)"
                ro.Offset(i, 6) = ""
                ro.Offset(i, 7) = "=VLOOKUP(" & ro.Offset(i, 0).Address & ",원료LIST,4,0)"
                ro.Offset(i, 8) = "=VLOOKUP(" & ro.Offset(i, 0).Address & ",원료LIST,8,0)"
                i = i + 1
            Else

            End If
    Next r
    'Set ro = ro.Resize(i, 8)
'서식 변경
    ro.Select
    ActiveCell.CurrentRegion.Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "표4"
    ActiveSheet.ListObjects("표4").TableStyle = "jss1"
    Range("표4[#All]").Select
    Range("I20").Activate
    With Selection.Font
        .Name = "맑은 고딕"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
'    ro.EntireColumn.AutoFit
    ro.Offset(i, 4).Cells(1, 1).Formula = "=SUM(" & Selection.Columns(5).Address & ")"
    'MsgBox ro.Address
End Sub
Sub ingOpenDB()
' 원료LIST & Base 읽어오기
' 2013-5-14 / 이진성

    dir1 = Range("C2").Value
    dir2 = Range("C3").Value

    Workbooks.Open FileName:=dir1, UpdateLinks:=0, ReadOnly:=1
    Workbooks.Open FileName:=dir2, UpdateLinks:=0, ReadOnly:=1

    Windows("성분표기작성4.1.xls").Activate

End Sub


Sub ingSplitBase()
' JU-BASE코드 풀어주기 ver2.0 by 이진성
' 2011-11-03 이진성
' 2012-12-13 현 시트에서 모든 작업(코드풀기,집계)이 진행되도록 수정
' 2015-06-17 선택영역에서 베이스 풀기(범용)
'
    Dim rs, r As Range
    Dim i, j, k As Integer
    Dim tmp As Variant


    Set rs = Selection
    i = 1
'- 베이스 풀기 ----------------------------------------------------------------------

    For Each r In rs.Rows
    With r
        bc = .Cells(1, 1).Value '베이스 코드값
        bv = .Cells(1, 5).Value '베이스 함량값

        If Left(bc, 5) = "JU-BS" And Mid(bc, 6, 4) <> "9999" Then
        ' BS코드이면
            bbc = Replace(bc, "-", "_")
            Set rb = Workbooks("Base.xls").Sheets("BASE").Range(bbc)
            For j = (rb.Rows.Count - 1) To 1 Step -1
                .Rows(2).Insert
                .Rows(2).Interior.ColorIndex = Null
                'Set rs = .Resize(.Rows.Count + 1, .Columns.Count)
                .Cells(2, 1).Value = rb.Offset(j, 0).Value '데이타 코드값
                tmp = rb.Offset(j, 1).Value  '참고영역 함량비
                .Cells(2, 5).Value = bv * tmp(1, 1) '베이스 함량값 * 참고영역 함량비


'                For k = 3 To .Columns.Count
'                    .Cells(2, k).Formula = .Cells(1, k).Formula
'                Next k

            Next j
            .Rows(1).Interior.Color = RGB(0, 255, 0)
        End If
    End With

    Next r
    rs.Select
Exit Sub

End Sub

Sub ingSplitBase2()
Attribute ingSplitBase2.VB_ProcData.VB_Invoke_Func = "g\n14"
'
    Dim c As Range '제목
    Dim rsai As Integer '기준범위 시작 열번호
    Dim rsax As Integer '기준범위 a 열번호
    Dim rsbx As Integer '기준범위 b 열번호
    Dim rbax As Integer '참조범위 a 열번호
    Dim rbbx As Integer '참조범위 b 열번호
    Dim rsas As String '기준범위 a 키값(코드)
    Dim rsbs As String '기준범위 b 키값(함량)
    Dim rbcc As Integer '참조범위 행수


    ' 초기값 설정
'    On Error GoTo ErrorHandler
    Set rs = Selection
    '머리글 행 체크
    i = 1
    rsai = rs.Cells(1, 1).Column
    For Each c In rs.Rows(1).Cells
        If c.Cells(1, 1).Value = "함량" Then rsbx = c.Column - rsai + 1
        If c.Cells(1, 1).Value = "코드" Then rsax = c.Column - rsai + 1
        If rsbx + rsax Then i = 2
    Next
    '
    ioOpendFile ("Base")
    'MsgBox "rsai:rsax:rsbx = " & rsai & "," & rsax & "," & rsbx

    Do While rs.Cells(i, 1) <> ""

        rs.Rows(i).Select
        '기준값
        rsas = rs.Cells(i, rsax).Value
        rsbs = rs.Cells(i, rsbx).Value

        'MsgBox rsas & "," & rsbs

        If rsas = "" Then Exit Do
        If Left(rsas, 5) = "JU-BS" And Mid(rsas, 6, 4) <> "9999" Then
            '참조 영역 설정
            Set rb = rBase(rsas)
            rbcc = rb.Rows.Count
            '참조 영역 가져오기
            rs.Offset(i, rsax - 1).Resize(rbcc, rs.Columns.Count).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
            rs.Cells(i, 1).Resize(1, rs.Columns.Count).Copy
            rs.Cells(i + 1, 1).PasteSpecial Paste:=xlPasteFormulas

            rb.Offset(0, 0).Resize(rbcc, 1).Copy
            rs.Cells(i + 1, rsax).PasteSpecial Paste:=xlPasteValues

            rb.Offset(0, 1).Resize(rbcc, 1).Copy
            rs.Cells(i + 1, rsbx).PasteSpecial Paste:=xlPasteValues
            rs.Select

            '함량 계산
            For j = i + 1 To i + rbcc
                rs.Cells(j, rsbx) = rs.Cells(j, rsbx) * rs.Cells(i, rsbx)
            Next j
            'rs.Rows(i).Delete
            i = i + rbcc
        End If
        i = i + 1
        'Stop
    Loop
Exit Sub
ErrorHandler:
    'MsgBox Err.Number
    Select Case Err.Number    ' 오류 번호를 계산합니다.

        Case 91    ' "파일이 이미 열려 있습니다" 오류입니다.
            MsgBox "에러:" & Err.Number
            ' 열린 파일을 닫습니다.
        Case Else
            ' 여기서 다른 상황을 다룹니다.
    End Select
    Resume Next

End Sub
Function rBase(st As String) As Range
' 찾는 BASE코드 영역 가져오기
' 150401
'
    Dim cc As Range
    Dim rb As Range
    Set rb = Workbooks("Base.xls").Sheets("BASE").Cells '참고 영역

    For Each cc In rb
     'Set cc = rb.Find(st) '위치 찾기
        If cc = st And cc.Interior.ColorIndex <> -4142 Then
            MsgBox "base위치: " & cc.Address & " 색번호:" & cc.Interior.ColorIndex
            Exit For
        End If
    Next
    Set rBase = Range(cc.Cells(2, 1), cc.End(xlDown)) '찾는 Base코드 영역
End Function

Sub ingQueryDB() ' 전성분 조회

'    On Error GoTo ErrorHandler
    ingInit


    For j = 2 To rs.Columns.Count
        tn = rs.Cells(1, j).Value
        tt(j) = WorksheetFunction.Match(tn, rm.Rows(1), 0)
        MsgBox tt(j)
    Next


    Set ra = Range(rs.Rows(2), rs.End(xlDown))

    For Each i In ra.Rows

        i.Rows(1).Select
        i.Interior.ColorIndex = 2
        bc = i.Cells(1, c1).Value '베이스 코드값

        For j = 1 To i.Columns.Count
            If i.i.Cells(1, 1).Interior.ColorIndex = -4142 Then
                i.Cells(1, j).Value = WorksheetFunction.VLookup(bc, rm, c(j), 0)
            End If
        Next j





    Next

ErrorHandler:
    'MsgBox Err.Number
    Select Case Err.Number    ' 오류 번호를 계산합니다.

        Case 1004    ' "파일이 이미 열려 있습니다" 오류입니다.
            i.Cells(1, c5) = "해당 코드가 없습니다."
            ' 열린 파일을 닫습니다.
        Case Else
            ' 여기서 다른 상황을 다룹니다.
'            Resume Next
    End Select

End Sub

Sub ingQueryDbName() ' DB 이름 위치 확인 150330 미완성
' 테이블 첫행의 타이틀값 읽어오기 기
    Dim se, db As Range
    Dim ti(100) As Integer
    Set se = Selection.Rows(1)
    Set db = "원료LIST.xls!database"
    'tit = Selection.Rows(1)
    For i = 1 To db.Columns.Count
        MsgBox i
    Next i
End Sub
Sub ingCheckSum()
' 조성비 합=1 여부 확인 160310
    Dim c As Range
    Dim ck, cl As Variant
    Dim iSeek As Long
    Dim iStart, iLen As Integer

    Set rs = Selection ' 선택영역
    Set rb = Workbooks("원료성분.xlsx").Sheets("성분사전").Range("INCI") '원료영역
    Set rm = Workbooks("원료LIST.xls").Sheets("원료LIST").Range("database") '원료영역

    For Each c In Selection
' 문자 트림
        If c.Value = "" Then GoTo Blank
        ni = Split(c.Value, "(and)")
        ni_cn = UBound(ni, 1)
        c.Value = ""
        For i = 0 To ni_cn - 1
            c.Value = c.Value + WorksheetFunction.Proper(Trim(ni(i))) & " (and) "
        Next i
        c.Value = c.Value & WorksheetFunction.Proper(Trim(ni(ni_cn)))

        c.Value = WorksheetFunction.Substitute(c.Text, "(Ci", "(CI")

 '오자 검색
        ni = Split(c.Value, "(and)")
        ni_cn = UBound(ni, 1)
        tmp = ""
        iStart = 1
        For j = 0 To ni_cn
            tmp = WorksheetFunction.Proper(Trim(ni(j)))
            tmp = WorksheetFunction.Substitute(tmp, "Ci", "CI")
            ck = Application.VLookup(tmp, rb, 1, 0)
            cl = Application.VLookup(tmp, rb, 14, 0)
            iSeek = InStr(iStart, c.Value, tmp)
            iLen = Len(tmp)

            With c.Characters(iSeek, iLen).Font
            If IsError(ck) Then
               ' MsgBox "Err: " & iSeek & ":" & iLen & " -- " & c.Value & ":" & tmp

                    .Bold = False
                    .Color = RGB(255, 0, 0)
                    .Name = "맑은 고딕"
                    .Size = 10
  '                  .Strikethrough = False
  '                  .Superscript = False
  '                  .Subscript = False
  '                  .OutlineFont = False
  '                  .Shadow = False
  '                  .Underline = xlUnderlineStyleNone
  '                  .TintAndShade = 0
  '                  .ThemeFont = xlThemeFontMinor

                iStart = iStart + iLen

            Else
               ' MsgBox "ok: " & iSeek & ":" & iLen & " -- " & c.Value & ":" & tmp
                    .Bold = False
                    .Color = RGB(0, 0, 0)
                    .Name = "맑은 고딕"
                    .Size = 10

                iStart = iStart + iLen
            End If
            End With
            'if c.characters(iSeek+iLen+1,4,"(and)") then c.Characters(iSeek, iLen).Text = tmp

        Next j

End Sub
Sub ingCheckSpell()
Attribute ingCheckSpell.VB_ProcData.VB_Invoke_Func = " \n14"
' 성분 오자 검정 150528
' Proper->문자 트림->CI포멧
' 160310 CI 번호 수정


    Dim c As Range
    Dim ck, cl As Variant
    Dim iSeek As Long
    Dim iStart, iLen As Integer

    Set rs = Selection ' 선택영역
    Set rb = Workbooks("원료성분.xlsx").Sheets("성분사전").Range("INCI") '원료영역
    Set rm = Workbooks("원료LIST.xls").Sheets("원료LIST").Range("database") '원료영역

    For Each c In Selection
' 문자 트림
        If c.Value = "" Then GoTo Blank
        ni = Split(c.Value, "(and)")
        ni_cn = UBound(ni, 1)
        c.Value = ""
        For i = 0 To ni_cn - 1
            c.Value = c.Value + WorksheetFunction.Proper(Trim(ni(i))) & " (and) "
        Next i
        c.Value = c.Value & WorksheetFunction.Proper(Trim(ni(ni_cn)))

        c.Value = WorksheetFunction.Substitute(c.Text, "(Ci", "(CI")

 '오자 검색
        ni = Split(c.Value, "(and)")
        ni_cn = UBound(ni, 1)
        tmp = ""
        iStart = 1
        For j = 0 To ni_cn
            tmp = WorksheetFunction.Proper(Trim(ni(j)))
            tmp = WorksheetFunction.Substitute(tmp, "Ci", "CI")
            ck = Application.VLookup(tmp, rb, 1, 0)
            cl = Application.VLookup(tmp, rb, 14, 0)
            iSeek = InStr(iStart, c.Value, tmp)
            iLen = Len(tmp)

            With c.Characters(iSeek, iLen).Font
            If IsError(ck) Then
               ' MsgBox "Err: " & iSeek & ":" & iLen & " -- " & c.Value & ":" & tmp

                    .Bold = False
                    .Color = RGB(255, 0, 0)
                    .Name = "맑은 고딕"
                    .Size = 10
  '                  .Strikethrough = False
  '                  .Superscript = False
  '                  .Subscript = False
  '                  .OutlineFont = False
  '                  .Shadow = False
  '                  .Underline = xlUnderlineStyleNone
  '                  .TintAndShade = 0
  '                  .ThemeFont = xlThemeFontMinor

                iStart = iStart + iLen

            Else
               ' MsgBox "ok: " & iSeek & ":" & iLen & " -- " & c.Value & ":" & tmp
                    .Bold = False
                    .Color = RGB(0, 0, 0)
                    .Name = "맑은 고딕"
                    .Size = 10

                iStart = iStart + iLen
            End If
            End With
            'if c.characters(iSeek+iLen+1,4,"(and)") then c.Characters(iSeek, iLen).Text = tmp

        Next j

Blank:
    Next c

End Sub

Sub ingSplitIng()
'성분나누기
'150617 선택영역

'-----------------------------------------------------------------------
    Dim i, j, k As Integer
    Dim r, w, ws As Range
    Dim ai() As String
    Dim ak() As String
    Dim ac()  As String
    Dim bc As String
    Dim bv As Variant
    Dim ai_cn, ak_cn, ac_cn, cn As Integer

    Set rs = Selection
    Set ws = Sheets("B").Range("A2")
    i = 1

    For Each r In rs.Rows


        bc = r.Cells(1, 1).Value '코드값
        bv = r.Cells(1, 2).Value '함량값
        ak = Split(r.Cells(1, 4).Value, "·")   '전성분명
        ai = Split(r.Cells(1, 5).Value, "(and)")  'INCI명
        ac = Split(r.Cells(1, 6), "/")    '조성비

        cn = 0
        ak_cn = UBound(ak, 1)  '전성분명 수
        ai_cn = UBound(ai, 1)  'INCI명 수
        ac_cn = UBound(ac, 1)  '조성비 수

        If ai_cn >= ak_cn Then cn = ai_cn


        Set w = ws.Rows(i)
        w.Cells(1, 2).Value = r.Cells(1, 3).Value '원료명
        w.Cells(1, 5).Value = r.Cells(1, 2).Value '함량

        If cn >= 1 Then  '다중 성분인 경우
                For j = 0 To cn
                    w.Cells(1 + j, 1).Value = r.Cells(1, 1).Value & "." & j '코드명
                    w.Cells(1 + j, 3).Value = Trim(ak(j)) '전성분명
                    w.Cells(1 + j, 4).Value = Trim(ai(j)) 'inci명


                    If ac_cn = cn Then
                        w.Cells(1 + j, 6).Value = Trim(ac(j)) '조성비

'                        If WorksheetFunction.IsNumber(r.Cells(2, 5)) Then r.Cells(2, 5).Font.Color = RGB(255, 0, 0)
'                        If WorksheetFunction.IsNumber(r.Cells(2, 6)) Then r.Cells(2, 5).Font.Color = RGB(255, 0, 0)
'
'                        If WorksheetFunction.IsNumber(r.Cells(1, ac(i))) Then
'                            r.Cells(2, 7).Value = r.Cells(2, 5).Value * r.Cells(2, 6).Value '실함량 계산

                    Else
'                        r.Cells(1, 10).Value = "※조성비:" & r.Cells(i, c6).Value
                        w.Cells(1 + j, 6).Font.Color = RGB(255, 0, 0)
                        tmp = Format(1 / (cn + 1), "#0.000")
                        w.Cells(1 + j, 6).Value = tmp
'                        r.Cells(1, c2).Value = r.Cells(i, c2).Value * tmp
'                        r.Cells(1, c4).Value = r.Cells(i, c4) & "(?)" '원료명(함량비)
'                        r.Cells(1, c6).Font.Color = RGB(255, 0, 0)
                    End If

                        'r.Cells(1, c3).Value = r.Cells(1, opt) & "(" & r.Cells(1, c2) & ")." & r.Cells(i, c3) '구분
                    w.Cells(1 + j, 7).Formula = "=" & w.Cells(1, 5).Address & "*" & w.Cells(1 + j, 6).Address

                Next j
                w.Cells(1 + j, 10) = cn & ":" & ak_cn & ":" & ai_cn & ":" & ac_cn
                i = i + cn + 1
            'r.Rows(1).Interior.Color = RGB(100, 100, 100)

        Else
                    w.Cells(1, 1).Value = r.Cells(1, 1).Value & "." & "00" '코드명
                    w.Cells(1, 3).Value = Trim(ak(0)) '전성분명
                    w.Cells(1, 4).Value = Trim(ai(0)) 'inci명
                    w.Cells(1, 6).Value = 1 '조성비
                    w.Cells(1, 7).Formula = "=" & w.Cells(1, 5).Address & "*" & w.Cells(1, 6).Address
                    i = i + 1
        End If


        'r.Rows(1).Interior.Color = RGB(255, 255, 0)
        'MsgBox "row =" & i

    Next r


End Sub

Sub ingSplitIng3()
'성분나누기
'150617 선택영역
 '   c1 = 1  ' 코드
 '   c2 = 2  ' 함량
 '   c3 = 4  ' 구분(상태표시)
 '   c4 = 3  ' 원료명
 '   c5 = 5  ' 전성분표준화명
 '   c6 = 6  ' 조성비
 '   c7 = 7  ' INCI명
 '   c8 = 8  ' 규격
 '   c9 = 9  ' 배합한도

'-----------------------------------------------------------------------
    Dim i, j, k As Integer
    Dim r As Range
    Dim ai() As String
    Dim ak() As String
    Dim ac()  As String
    Dim bc As String
    Dim bv As Variant
    Dim ai_cn, ak_cn, ac_cn As Integer

    Set rs = Selection

    For Each r In rs.Rows


        bc = r.Cells(1, 1).Value '코드값
        bv = r.Cells(1, 5).Value '함량값
        ai = Split(r.Cells(1, 3).Value, "(and)")  'INCI명
        ak = Split(r.Cells(1, 4), "/")   'CAS
        ac = Split(r.Cells(1, 6), "/")    '조성비

        ai_cn = UBound(ai, 1)   'INCI수
        ak_cn = UBound(ak, 1)   'CAS
        ac_cn = UBound(ac, 1)   '조성비수


        If ai_cn > 1 Then  '다중 성분인 경우

            For j = ai_cn To 0 Step -1

                    r.Rows(2).Insert
                    r.Cells(2, 3).Value = Trim(ai(j)) 'inci
                    If ak_cn = ai_cn Then r.Cells(2, 4).Value = Trim(ak(j)) 'cas
                    If ac_cn = ai_cn Then r.Cells(2, 6).Value = Trim(ac(j)) '조성비
                    'r.Cells(2, 7).Formula = r.Cells(2, 5).Address * r.Cells(2.6).Address
                    If WorksheetFunction.IsNumber(r.Cells(2, 5)) Then r.Cells(2, 5).Font.Color = RGB(255, 0, 0)
                    If WorksheetFunction.IsNumber(r.Cells(2, 6)) Then r.Cells(2, 5).Font.Color = RGB(255, 0, 0)

                    If WorksheetFunction.IsNumber(r.Cells(1, ac(i))) Then
                        r.Cells(2, 7).Value = r.Cells(2, 5).Value * r.Cells(2, 6).Value '실함량 계산

                    Else
'                        r.Cells(1, 10).Value = "※조성비:" & r.Cells(i, c6).Value
'                        r.Cells(1, 10).Font.Color = RGB(255, 0, 0)
'                        tmp = Format(1 / (ak_cn + 1), "#0.000")
'                        r.Cells(1, c6).Value = tmp
'                        r.Cells(1, c2).Value = r.Cells(i, c2).Value * tmp
'                        r.Cells(1, c4).Value = r.Cells(i, c4) & "(?)" '원료명(함량비)
'                        r.Cells(1, c6).Font.Color = RGB(255, 0, 0)
                    End If

                        'r.Cells(1, c3).Value = r.Cells(1, opt) & "(" & r.Cells(1, c2) & ")." & r.Cells(i, c3) '구분

            Next j
            r.Rows(1).Interior.Color = RGB(100, 100, 100)

        End If
        r.Rows(1).Interior.Color = RGB(255, 255, 0)
        'MsgBox "row =" & i

    Next r


End Sub

Sub ingSplitIng2()
' 성분나누기
 '   c1 = 1  ' 코드
 '   c2 = 2  ' 함량
 '   c3 = 4  ' 구분(상태표시)
 '   c4 = 3  ' 원료명
 '   c5 = 5  ' 전성분표준화명
 '   c6 = 6  ' 조성비
 '   c7 = 7  ' INCI명
 '   c8 = 8  ' 규격
 '   c9 = 9  ' 배합한도
    ingInit
'-----------------------------------------------------------------------
    i = 2

    opt = Range("D2").Value

    Do While rs.Cells(i, c1) <> ""

        rs.Rows(i).Select
        bc = rs.Cells(i, c1).Value '코드값
        bv = rs.Cells(i, c2).Value '함량값

        nk = Split(rs.Cells(i, c5), "·")   '전성분명
        nc = Split(rs.Cells(i, c6), "/")    '조성비
        ni = Split(rs.Cells(i, c7).Value, "(and)")  'INCI명
        nk_cn = UBound(nk, 1)   '전성분수
        nc_cn = UBound(nc, 1)   '조성비수
        ni_cn = UBound(ni, 1)   'INCI수

        If nk_cn >= 1 Then

            For j = nk_cn To 0 Step -1

                rs.Cells.Rows(i + 1).Insert
                rs.Cells(i + 1, c5).Value = Trim(nk(j))
                If nk_cn = UBound(nc, 1) Then rs.Cells(i + 1, c6).Value = Trim(nc(j))
                If nk_cn = UBound(ni, 1) Then rs.Cells(i + 1, c7).Value = Trim(ni(j))
                rs.Cells(i + 1, c1).Value = rs.Cells(i + 1, opt)

                If WorksheetFunction.IsNumber(rs.Cells(i + 1, c6)) Then
                    rs.Cells(i + 1, c2).Value = rs.Cells(i, c2).Value * rs.Cells(i + 1, c6).Value
                    rs.Cells(i + 1, c4).Value = rs.Cells(i, c4) & "(" & Format(nc(j), "#0.0%") & ")"  '원료명(함량비)
                Else
                    rs.Cells(i + 1, 10).Value = "※조성비:" & rs.Cells(i, c6).Value
                    rs.Cells(i + 1, 10).Font.Color = RGB(255, 0, 0)
                    tmp = Format(1 / (nk_cn + 1), "#0.000")
                    rs.Cells(i + 1, c6).Value = tmp
                    rs.Cells(i + 1, c2).Value = rs.Cells(i, c2).Value * tmp
                    rs.Cells(i + 1, c4).Value = rs.Cells(i, c4) & "(?)" '원료명(함량비)
                    rs.Cells(i + 1, c6).Font.Color = RGB(255, 0, 0)
                End If

                    rs.Cells(i + 1, c3).Value = rs.Cells(i + 1, opt) & "(" & rs.Cells(i + 1, c2) & ")." & rs.Cells(i, c3) '구분

            Next j

                rs.Rows(i).Delete
        Else
                rs.Cells(i, c1).Value = rs.Cells(i, opt)
                rs.Cells(i, c3).Value = rs.Cells(i, opt) & "(" & rs.Cells(i, c2) & ")." & rs.Cells(i, c3)    '구분
                rs.Cells(i, c4).Value = rs.Cells(i, c4) & "(100%)"  '원료명(함량비)

        End If

        i = i + nk_cn + 1
        'MsgBox "row =" & i

    Loop


End Sub

Sub ingSortByIng()
'
' 매크로1 매크로
'


    With ActiveWorkbook.ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A6:A6"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("B6:B6"), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

        .SetRange Range("성분코드")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Sub ingSortByVol()
'
' 매크로1 매크로
'
    With ActiveWorkbook.ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("B6:B6"), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange Range("성분코드")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Sub ingMergeIng()

' 전성분 합치기  by 이진성
' 2012-12-17

'
    init
    Set rs = ActiveSheet.Range("성분코드")


 '   c1 = 1  ' 코드
 '   c2 = 2  ' 함량
 '   c3 = 4  ' 구분(상태표시)
 '   c4 = 3  ' 원료명
 '   c5 = 5  ' 전성분표준화명
 '   c6 = 6  ' 조성비
 '   c7 = 7  ' INCI명
 '   c8 = 8  ' 규격
 '   c9 = 9  ' 배합한도

'- 전성분 병합 ----------------------------------------------------------------------
    i = 2

    'bc = rs.Cells(1, c1).Value

    Do While rs.Cells(i, c1) <> ""

        rs.Rows(i).Select
        st = rs.Cells(i, c3).Value '상태 구분

        If rs.Cells(i, c1) = rs.Cells(i + 1, c1) Then

            'rs.Cells(i, c1) = rs.Cells(i, c1) & Chr(10) & rs.Cells(i + 1, c1)
            rs.Cells(i, c2) = rs.Cells(i, c2) + rs.Cells(i + 1, c2) 'temr 함량값
            rs.Cells(i, c3) = rs.Cells(i, c3) & Chr(10) & rs.Cells(i + 1, c3)
            rs.Cells(i, c4) = rs.Cells(i, c4) & Chr(10) & rs.Cells(i + 1, c4)
            rs.Cells(i, c5) = rs.Cells(i, c5)
            rs.Cells(i, c7) = rs.Cells(i, c7)

            rs.Cells.Rows(i + 1).Delete


        Else

            i = i + 1

        End If



    Loop
End Sub

Sub ingOpenAll()
' 원료리스트 모두 열기 150402
    Workbooks.Open rPreset("원료LIST")
    'Workbooks.Open rPreset("원료Base")
End Sub

Sub edMergeTbl1()
'
' 매크로2 매크로
'

'
    Dim ro As Range
    Dim rs As String
    rs = Selection.Address(ReferenceStyle:=xlR1C1)
    'Selection.ClearContents
    Set ro = Range("K5")
    ro.Consolidate Sources:=rs, Function:=xlSum, TopRow:=False, LeftColumn:=True, CreateLinks:=False

    ro.CurrentRegion.Select
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("K5:K27") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("K5:L27")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


End Sub
