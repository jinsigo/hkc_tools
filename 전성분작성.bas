Attribute VB_Name = "�������ۼ�"
'
'======================================================================
' ������ �ۼ� ��ũ�� ver 5.0 130330
'======================================================================
'
'
'
'
Option Explicit

Public rs As Range ' ���ÿ��� (���� ����)
Public ra As Range ' ���ÿ��� (����Ÿ ������)
Public rb As Range ' Base.xls ������
Public rm As Range ' ����LIST.xls ������
Public st As String ' ���°�
Public c1, c2, c3, c4, c5, c6, c7, c8, c9, c10 As Integer '����ȣ��
Dim i, j, nk_cn, nc_cn, ni_cn, rb_cn, next_row, opt As Integer
Dim bc, bbc As String '�ڵ尪
Dim nk() As String
Dim ni() As String
Dim nc() As String
Dim bv, tmp, temr As Variant '�Է���
Dim temp, temc As String
Dim dir1, dir2, temq As String '���


Sub ingInit()
' ������ ����

'On Error GoTo ErrorHandler
    Dim c(100) As Integer
    Dim t() As Variant
    
    'Set rs = ActiveSheet.Range("list") ' ���ÿ���
    Set rb = Workbooks("Base.xls").Sheets("BASE").Range("$A1:$IV20000") '���῵��
    Set rm = Workbooks("����LIST.xls").Sheets("����LIST").Range("database") '���῵��
    
    c(1) = 1  ' �ڵ�
    c(2) = 1  ' �Է�
    c(3) = 1  ' ����(����ǥ��)
    c(4) = 2  ' �����
    c(5) = 3  ' ������ǥ��ȭ��
    c(6) = 18  ' ������
    c(7) = 5  ' INCI��
    c(8) = 4  ' �԰�
    c(9) = 0  ' �����ѵ�
    c(10) = 10
    c(11) = 5 '���1
    c(12) = 7 '���2
    
    't = Array("�����ڵ�", "�Է�", "����", "�����", "������", "������", "INCI", "�԰�", "�����ѵ�", "CAS No.", "HS�ڵ�")
    i = 2
    
'    If Mid(rs.Cells(i, 1).Value, 3, 1) <> "-" Then rs.Columns(1).Replace what:="JU", Replacement:="JU-", SearchOrder:=xlByColumns, MatchCase:=True
    
    
ErrorHandler:
    'MsgBox Err.Number
    Select Case Err.Number    ' ���� ��ȣ�� ����մϴ�.
    
        Case 1004    ' "������ �̹� ���� �ֽ��ϴ�" �����Դϴ�.
            rs.Cells(i, c5) = "�ش� �ڵ尡 �����ϴ�."
            ' ���� ������ �ݽ��ϴ�.
        Case Else
            ' ���⼭ �ٸ� ��Ȳ�� �ٷ�ϴ�.
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
' ���ÿ������� ��Ʈ �����ϰ� �ڵ尪 ��������
'2016-06-24
    Dim rs, ro, r, t As Range
    Dim dc As String
    Dim dv As Variant
    Dim cs As Integer
    Dim ti As Variant
    
    ti = Array("�����ڵ�", "�����", "����", "�����", "������", "������", "INCI", "�԰�", "�����ѵ�", "CAS No.", "HS�ڵ�")
    
    Set rs = Selection
    cs = rs.Columns.Count
    
' ��Ʈ �����
    shAddSht ("������")
    Worksheets("������").Activate
    Set ro = Worksheets("������").Range("A6")
    
' ���� �ֱ�
'    Call LoadArray(ro, ti)
                ro.Offset(0, 0) = "�����ڵ�"
                ro.Offset(0, 1) = "�����"
                ro.Offset(0, 2) = "INCI"
                ro.Offset(0, 3) = "CAS No."
                ro.Offset(0, 4) = "�Է�(w/w%)"
                ro.Offset(0, 5) = "������"
                ro.Offset(0, 6) = "���Է�(w/w%)"
                ro.Offset(0, 7) = "�԰�"
                ro.Offset(0, 8) = "Function"
                
 '����Ÿ �ֱ�
    i = 1
    For Each r In rs.Rows
            dc = r.Cells(1, 1)
            dv = r.Cells(1, cs).Value
            If IsNumeric(dv) And (dv > 0) Then
                ro.Offset(i, 0) = dc
                ro.Offset(i, 1) = "=VLOOKUP(" & ro.Offset(i, 0).Address & ",����LIST,2,0)"
                ro.Offset(i, 2) = "=VLOOKUP(" & ro.Offset(i, 0).Address & ",����LIST,5,0)"
                ro.Offset(i, 3) = "=VLOOKUP(" & ro.Offset(i, 0).Address & ",����LIST,16,0)"
                ro.Offset(i, 4) = Format(dv, "#,##0.000")
                ro.Offset(i, 5) = "=VLOOKUP(" & ro.Offset(i, 0).Address & ",����LIST,18,0)"
                ro.Offset(i, 6) = ""
                ro.Offset(i, 7) = "=VLOOKUP(" & ro.Offset(i, 0).Address & ",����LIST,4,0)"
                ro.Offset(i, 8) = "=VLOOKUP(" & ro.Offset(i, 0).Address & ",����LIST,8,0)"
                i = i + 1
            Else
                
            End If
    Next r
    'Set ro = ro.Resize(i, 8)
'���� ����
    ro.Select
    ActiveCell.CurrentRegion.Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = "ǥ4"
    ActiveSheet.ListObjects("ǥ4").TableStyle = "jss1"
    Range("ǥ4[#All]").Select
    Range("I20").Activate
    With Selection.Font
        .Name = "���� ���"
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
' ����LIST & Base �о����
' 2013-5-14 / ������

    dir1 = Range("C2").Value
    dir2 = Range("C3").Value
    
    Workbooks.Open FileName:=dir1, UpdateLinks:=0, ReadOnly:=1
    Workbooks.Open FileName:=dir2, UpdateLinks:=0, ReadOnly:=1
    
    Windows("����ǥ���ۼ�4.1.xls").Activate

End Sub


Sub ingSplitBase()
' JU-BASE�ڵ� Ǯ���ֱ� ver2.0 by ������
' 2011-11-03 ������
' 2012-12-13 �� ��Ʈ���� ��� �۾�(�ڵ�Ǯ��,����)�� ����ǵ��� ����
' 2015-06-17 ���ÿ������� ���̽� Ǯ��(����)
'
    Dim rs, r As Range
    Dim i, j, k As Integer
    Dim tmp As Variant
    
    
    Set rs = Selection
    i = 1
'- ���̽� Ǯ�� ----------------------------------------------------------------------
        
    For Each r In rs.Rows
    With r
        bc = .Cells(1, 1).Value '���̽� �ڵ尪
        bv = .Cells(1, 5).Value '���̽� �Է���
        
        If Left(bc, 5) = "JU-BS" And Mid(bc, 6, 4) <> "9999" Then
        ' BS�ڵ��̸�
            bbc = Replace(bc, "-", "_")
            Set rb = Workbooks("Base.xls").Sheets("BASE").Range(bbc)
            For j = (rb.Rows.Count - 1) To 1 Step -1
                .Rows(2).Insert
                .Rows(2).Interior.ColorIndex = Null
                'Set rs = .Resize(.Rows.Count + 1, .Columns.Count)
                .Cells(2, 1).Value = rb.Offset(j, 0).Value '����Ÿ �ڵ尪
                tmp = rb.Offset(j, 1).Value  '������ �Է���
                .Cells(2, 5).Value = bv * tmp(1, 1) '���̽� �Է��� * ������ �Է���
                
                
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
    Dim c As Range '����
    Dim rsai As Integer '���ع��� ���� ����ȣ
    Dim rsax As Integer '���ع��� a ����ȣ
    Dim rsbx As Integer '���ع��� b ����ȣ
    Dim rbax As Integer '�������� a ����ȣ
    Dim rbbx As Integer '�������� b ����ȣ
    Dim rsas As String '���ع��� a Ű��(�ڵ�)
    Dim rsbs As String '���ع��� b Ű��(�Է�)
    Dim rbcc As Integer '�������� ���
    
    
    ' �ʱⰪ ����
'    On Error GoTo ErrorHandler
    Set rs = Selection
    '�Ӹ��� �� üũ
    i = 1
    rsai = rs.Cells(1, 1).Column
    For Each c In rs.Rows(1).Cells
        If c.Cells(1, 1).Value = "�Է�" Then rsbx = c.Column - rsai + 1
        If c.Cells(1, 1).Value = "�ڵ�" Then rsax = c.Column - rsai + 1
        If rsbx + rsax Then i = 2
    Next
    '
    ioOpendFile ("Base")
    'MsgBox "rsai:rsax:rsbx = " & rsai & "," & rsax & "," & rsbx
    
    Do While rs.Cells(i, 1) <> ""
    
        rs.Rows(i).Select
        '���ذ�
        rsas = rs.Cells(i, rsax).Value
        rsbs = rs.Cells(i, rsbx).Value
        
        'MsgBox rsas & "," & rsbs
        
        If rsas = "" Then Exit Do
        If Left(rsas, 5) = "JU-BS" And Mid(rsas, 6, 4) <> "9999" Then
            '���� ���� ����
            Set rb = rBase(rsas)
            rbcc = rb.Rows.Count
            '���� ���� ��������
            rs.Offset(i, rsax - 1).Resize(rbcc, rs.Columns.Count).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
            rs.Cells(i, 1).Resize(1, rs.Columns.Count).Copy
            rs.Cells(i + 1, 1).PasteSpecial Paste:=xlPasteFormulas
            
            rb.Offset(0, 0).Resize(rbcc, 1).Copy
            rs.Cells(i + 1, rsax).PasteSpecial Paste:=xlPasteValues
            
            rb.Offset(0, 1).Resize(rbcc, 1).Copy
            rs.Cells(i + 1, rsbx).PasteSpecial Paste:=xlPasteValues
            rs.Select
            
            '�Է� ���
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
    Select Case Err.Number    ' ���� ��ȣ�� ����մϴ�.
    
        Case 91    ' "������ �̹� ���� �ֽ��ϴ�" �����Դϴ�.
            MsgBox "����:" & Err.Number
            ' ���� ������ �ݽ��ϴ�.
        Case Else
            ' ���⼭ �ٸ� ��Ȳ�� �ٷ�ϴ�.
    End Select
    Resume Next
    
End Sub
Function rBase(st As String) As Range
' ã�� BASE�ڵ� ���� ��������
' 150401
'
    Dim cc As Range
    Dim rb As Range
    Set rb = Workbooks("Base.xls").Sheets("BASE").Cells '���� ����
    
    For Each cc In rb
     'Set cc = rb.Find(st) '��ġ ã��
        If cc = st And cc.Interior.ColorIndex <> -4142 Then
            MsgBox "base��ġ: " & cc.Address & " ����ȣ:" & cc.Interior.ColorIndex
            Exit For
        End If
    Next
    Set rBase = Range(cc.Cells(2, 1), cc.End(xlDown)) 'ã�� Base�ڵ� ����
End Function

Sub ingQueryDB() ' ������ ��ȸ

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
        bc = i.Cells(1, c1).Value '���̽� �ڵ尪
              
        For j = 1 To i.Columns.Count
            If i.i.Cells(1, 1).Interior.ColorIndex = -4142 Then
                i.Cells(1, j).Value = WorksheetFunction.VLookup(bc, rm, c(j), 0)
            End If
        Next j
        

    
        

    Next
    
ErrorHandler:
    'MsgBox Err.Number
    Select Case Err.Number    ' ���� ��ȣ�� ����մϴ�.
    
        Case 1004    ' "������ �̹� ���� �ֽ��ϴ�" �����Դϴ�.
            i.Cells(1, c5) = "�ش� �ڵ尡 �����ϴ�."
            ' ���� ������ �ݽ��ϴ�.
        Case Else
            ' ���⼭ �ٸ� ��Ȳ�� �ٷ�ϴ�.
'            Resume Next
    End Select
    
End Sub
        
Sub ingQueryDbName() ' DB �̸� ��ġ Ȯ�� 150330 �̿ϼ�
' ���̺� ù���� Ÿ��Ʋ�� �о���� ��
    Dim se, db As Range
    Dim ti(100) As Integer
    Set se = Selection.Rows(1)
    Set db = "����LIST.xls!database"
    'tit = Selection.Rows(1)
    For i = 1 To db.Columns.Count
        MsgBox i
    Next i
End Sub
Sub ingCheckSum()
' ������ ��=1 ���� Ȯ�� 160310
    Dim c As Range
    Dim ck, cl As Variant
    Dim iSeek As Long
    Dim iStart, iLen As Integer
    
    Set rs = Selection ' ���ÿ���
    Set rb = Workbooks("���Ἲ��.xlsx").Sheets("���л���").Range("INCI") '���῵��
    Set rm = Workbooks("����LIST.xls").Sheets("����LIST").Range("database") '���῵��
    
    For Each c In Selection
' ���� Ʈ��
        If c.Value = "" Then GoTo Blank
        ni = Split(c.Value, "(and)")
        ni_cn = UBound(ni, 1)
        c.Value = ""
        For i = 0 To ni_cn - 1
            c.Value = c.Value + WorksheetFunction.Proper(Trim(ni(i))) & " (and) "
        Next i
        c.Value = c.Value & WorksheetFunction.Proper(Trim(ni(ni_cn)))
        
        c.Value = WorksheetFunction.Substitute(c.Text, "(Ci", "(CI")
        
 '���� �˻�
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
                    .Name = "���� ���"
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
                    .Name = "���� ���"
                    .Size = 10

                iStart = iStart + iLen
            End If
            End With
            'if c.characters(iSeek+iLen+1,4,"(and)") then c.Characters(iSeek, iLen).Text = tmp
 
        Next j
    
End Sub
Sub ingCheckSpell()
Attribute ingCheckSpell.VB_ProcData.VB_Invoke_Func = " \n14"
' ���� ���� ���� 150528
' Proper->���� Ʈ��->CI����
' 160310 CI ��ȣ ����


    Dim c As Range
    Dim ck, cl As Variant
    Dim iSeek As Long
    Dim iStart, iLen As Integer
    
    Set rs = Selection ' ���ÿ���
    Set rb = Workbooks("���Ἲ��.xlsx").Sheets("���л���").Range("INCI") '���῵��
    Set rm = Workbooks("����LIST.xls").Sheets("����LIST").Range("database") '���῵��
    
    For Each c In Selection
' ���� Ʈ��
        If c.Value = "" Then GoTo Blank
        ni = Split(c.Value, "(and)")
        ni_cn = UBound(ni, 1)
        c.Value = ""
        For i = 0 To ni_cn - 1
            c.Value = c.Value + WorksheetFunction.Proper(Trim(ni(i))) & " (and) "
        Next i
        c.Value = c.Value & WorksheetFunction.Proper(Trim(ni(ni_cn)))
        
        c.Value = WorksheetFunction.Substitute(c.Text, "(Ci", "(CI")
        
 '���� �˻�
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
                    .Name = "���� ���"
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
                    .Name = "���� ���"
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
'���г�����
'150617 ���ÿ���
 
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
        
        
        bc = r.Cells(1, 1).Value '�ڵ尪
        bv = r.Cells(1, 2).Value '�Է���
        ak = Split(r.Cells(1, 4).Value, "��")   '�����и�
        ai = Split(r.Cells(1, 5).Value, "(and)")  'INCI��
        ac = Split(r.Cells(1, 6), "/")    '������
        
        cn = 0
        ak_cn = UBound(ak, 1)  '�����и� ��
        ai_cn = UBound(ai, 1)  'INCI�� ��
        ac_cn = UBound(ac, 1)  '������ ��
         
        If ai_cn >= ak_cn Then cn = ai_cn
        
       
        Set w = ws.Rows(i)
        w.Cells(1, 2).Value = r.Cells(1, 3).Value '�����
        w.Cells(1, 5).Value = r.Cells(1, 2).Value '�Է�
       
        If cn >= 1 Then  '���� ������ ���
                For j = 0 To cn
                    w.Cells(1 + j, 1).Value = r.Cells(1, 1).Value & "." & j '�ڵ��
                    w.Cells(1 + j, 3).Value = Trim(ak(j)) '�����и�
                    w.Cells(1 + j, 4).Value = Trim(ai(j)) 'inci��

                    
                    If ac_cn = cn Then
                        w.Cells(1 + j, 6).Value = Trim(ac(j)) '������
                    
'                        If WorksheetFunction.IsNumber(r.Cells(2, 5)) Then r.Cells(2, 5).Font.Color = RGB(255, 0, 0)
'                        If WorksheetFunction.IsNumber(r.Cells(2, 6)) Then r.Cells(2, 5).Font.Color = RGB(255, 0, 0)
'
'                        If WorksheetFunction.IsNumber(r.Cells(1, ac(i))) Then
'                            r.Cells(2, 7).Value = r.Cells(2, 5).Value * r.Cells(2, 6).Value '���Է� ���
                        
                    Else
'                        r.Cells(1, 10).Value = "��������:" & r.Cells(i, c6).Value
                        w.Cells(1 + j, 6).Font.Color = RGB(255, 0, 0)
                        tmp = Format(1 / (cn + 1), "#0.000")
                        w.Cells(1 + j, 6).Value = tmp
'                        r.Cells(1, c2).Value = r.Cells(i, c2).Value * tmp
'                        r.Cells(1, c4).Value = r.Cells(i, c4) & "(?)" '�����(�Է���)
'                        r.Cells(1, c6).Font.Color = RGB(255, 0, 0)
                    End If
                        
                        'r.Cells(1, c3).Value = r.Cells(1, opt) & "(" & r.Cells(1, c2) & ")." & r.Cells(i, c3) '����
                    w.Cells(1 + j, 7).Formula = "=" & w.Cells(1, 5).Address & "*" & w.Cells(1 + j, 6).Address
                    
                Next j
                w.Cells(1 + j, 10) = cn & ":" & ak_cn & ":" & ai_cn & ":" & ac_cn
                i = i + cn + 1
            'r.Rows(1).Interior.Color = RGB(100, 100, 100)
        
        Else
                    w.Cells(1, 1).Value = r.Cells(1, 1).Value & "." & "00" '�ڵ��
                    w.Cells(1, 3).Value = Trim(ak(0)) '�����и�
                    w.Cells(1, 4).Value = Trim(ai(0)) 'inci��
                    w.Cells(1, 6).Value = 1 '������
                    w.Cells(1, 7).Formula = "=" & w.Cells(1, 5).Address & "*" & w.Cells(1, 6).Address
                    i = i + 1
        End If
        
 
        'r.Rows(1).Interior.Color = RGB(255, 255, 0)
        'MsgBox "row =" & i

    Next r
    

End Sub

Sub ingSplitIng3()
'���г�����
'150617 ���ÿ���
 '   c1 = 1  ' �ڵ�
 '   c2 = 2  ' �Է�
 '   c3 = 4  ' ����(����ǥ��)
 '   c4 = 3  ' �����
 '   c5 = 5  ' ������ǥ��ȭ��
 '   c6 = 6  ' ������
 '   c7 = 7  ' INCI��
 '   c8 = 8  ' �԰�
 '   c9 = 9  ' �����ѵ�

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
        
        
        bc = r.Cells(1, 1).Value '�ڵ尪
        bv = r.Cells(1, 5).Value '�Է���
        ai = Split(r.Cells(1, 3).Value, "(and)")  'INCI��
        ak = Split(r.Cells(1, 4), "/")   'CAS
        ac = Split(r.Cells(1, 6), "/")    '������
        
        ai_cn = UBound(ai, 1)   'INCI��
        ak_cn = UBound(ak, 1)   'CAS
        ac_cn = UBound(ac, 1)   '�������
        
               
        If ai_cn > 1 Then  '���� ������ ���
        
            For j = ai_cn To 0 Step -1
                
                    r.Rows(2).Insert
                    r.Cells(2, 3).Value = Trim(ai(j)) 'inci
                    If ak_cn = ai_cn Then r.Cells(2, 4).Value = Trim(ak(j)) 'cas
                    If ac_cn = ai_cn Then r.Cells(2, 6).Value = Trim(ac(j)) '������
                    'r.Cells(2, 7).Formula = r.Cells(2, 5).Address * r.Cells(2.6).Address
                    If WorksheetFunction.IsNumber(r.Cells(2, 5)) Then r.Cells(2, 5).Font.Color = RGB(255, 0, 0)
                    If WorksheetFunction.IsNumber(r.Cells(2, 6)) Then r.Cells(2, 5).Font.Color = RGB(255, 0, 0)
                    
                    If WorksheetFunction.IsNumber(r.Cells(1, ac(i))) Then
                        r.Cells(2, 7).Value = r.Cells(2, 5).Value * r.Cells(2, 6).Value '���Է� ���
                        
                    Else
'                        r.Cells(1, 10).Value = "��������:" & r.Cells(i, c6).Value
'                        r.Cells(1, 10).Font.Color = RGB(255, 0, 0)
'                        tmp = Format(1 / (ak_cn + 1), "#0.000")
'                        r.Cells(1, c6).Value = tmp
'                        r.Cells(1, c2).Value = r.Cells(i, c2).Value * tmp
'                        r.Cells(1, c4).Value = r.Cells(i, c4) & "(?)" '�����(�Է���)
'                        r.Cells(1, c6).Font.Color = RGB(255, 0, 0)
                    End If
                        
                        'r.Cells(1, c3).Value = r.Cells(1, opt) & "(" & r.Cells(1, c2) & ")." & r.Cells(i, c3) '����
    
            Next j
            r.Rows(1).Interior.Color = RGB(100, 100, 100)
        
        End If
        r.Rows(1).Interior.Color = RGB(255, 255, 0)
        'MsgBox "row =" & i

    Next r
    

End Sub

Sub ingSplitIng2()
' ���г�����
 '   c1 = 1  ' �ڵ�
 '   c2 = 2  ' �Է�
 '   c3 = 4  ' ����(����ǥ��)
 '   c4 = 3  ' �����
 '   c5 = 5  ' ������ǥ��ȭ��
 '   c6 = 6  ' ������
 '   c7 = 7  ' INCI��
 '   c8 = 8  ' �԰�
 '   c9 = 9  ' �����ѵ�
    ingInit
'-----------------------------------------------------------------------
    i = 2
        
    opt = Range("D2").Value
    
    Do While rs.Cells(i, c1) <> ""
        
        rs.Rows(i).Select
        bc = rs.Cells(i, c1).Value '�ڵ尪
        bv = rs.Cells(i, c2).Value '�Է���
        
        nk = Split(rs.Cells(i, c5), "��")   '�����и�
        nc = Split(rs.Cells(i, c6), "/")    '������
        ni = Split(rs.Cells(i, c7).Value, "(and)")  'INCI��
        nk_cn = UBound(nk, 1)   '�����м�
        nc_cn = UBound(nc, 1)   '�������
        ni_cn = UBound(ni, 1)   'INCI��
        
        If nk_cn >= 1 Then

            For j = nk_cn To 0 Step -1
            
                rs.Cells.Rows(i + 1).Insert
                rs.Cells(i + 1, c5).Value = Trim(nk(j))
                If nk_cn = UBound(nc, 1) Then rs.Cells(i + 1, c6).Value = Trim(nc(j))
                If nk_cn = UBound(ni, 1) Then rs.Cells(i + 1, c7).Value = Trim(ni(j))
                rs.Cells(i + 1, c1).Value = rs.Cells(i + 1, opt)
                
                If WorksheetFunction.IsNumber(rs.Cells(i + 1, c6)) Then
                    rs.Cells(i + 1, c2).Value = rs.Cells(i, c2).Value * rs.Cells(i + 1, c6).Value
                    rs.Cells(i + 1, c4).Value = rs.Cells(i, c4) & "(" & Format(nc(j), "#0.0%") & ")"  '�����(�Է���)
                Else
                    rs.Cells(i + 1, 10).Value = "��������:" & rs.Cells(i, c6).Value
                    rs.Cells(i + 1, 10).Font.Color = RGB(255, 0, 0)
                    tmp = Format(1 / (nk_cn + 1), "#0.000")
                    rs.Cells(i + 1, c6).Value = tmp
                    rs.Cells(i + 1, c2).Value = rs.Cells(i, c2).Value * tmp
                    rs.Cells(i + 1, c4).Value = rs.Cells(i, c4) & "(?)" '�����(�Է���)
                    rs.Cells(i + 1, c6).Font.Color = RGB(255, 0, 0)
                End If
                    
                    rs.Cells(i + 1, c3).Value = rs.Cells(i + 1, opt) & "(" & rs.Cells(i + 1, c2) & ")." & rs.Cells(i, c3) '����

            Next j
                
                rs.Rows(i).Delete
        Else
                rs.Cells(i, c1).Value = rs.Cells(i, opt)
                rs.Cells(i, c3).Value = rs.Cells(i, opt) & "(" & rs.Cells(i, c2) & ")." & rs.Cells(i, c3)    '����
                rs.Cells(i, c4).Value = rs.Cells(i, c4) & "(100%)"  '�����(�Է���)
                
        End If
        
        i = i + nk_cn + 1
        'MsgBox "row =" & i

    Loop
    

End Sub

Sub ingSortByIng()
'
' ��ũ��1 ��ũ��
'

        
    With ActiveWorkbook.ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A6:A6"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("B6:B6"), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

        .SetRange Range("�����ڵ�")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Sub ingSortByVol()
'
' ��ũ��1 ��ũ��
'
    With ActiveWorkbook.ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("B6:B6"), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange Range("�����ڵ�")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Sub ingMergeIng()

' ������ ��ġ��  by ������
' 2012-12-17

'
    init
    Set rs = ActiveSheet.Range("�����ڵ�")
    
    
 '   c1 = 1  ' �ڵ�
 '   c2 = 2  ' �Է�
 '   c3 = 4  ' ����(����ǥ��)
 '   c4 = 3  ' �����
 '   c5 = 5  ' ������ǥ��ȭ��
 '   c6 = 6  ' ������
 '   c7 = 7  ' INCI��
 '   c8 = 8  ' �԰�
 '   c9 = 9  ' �����ѵ�
    
'- ������ ���� ----------------------------------------------------------------------
    i = 2

    'bc = rs.Cells(1, c1).Value
    
    Do While rs.Cells(i, c1) <> ""

        rs.Rows(i).Select
        st = rs.Cells(i, c3).Value '���� ����
              
        If rs.Cells(i, c1) = rs.Cells(i + 1, c1) Then
        
            'rs.Cells(i, c1) = rs.Cells(i, c1) & Chr(10) & rs.Cells(i + 1, c1)
            rs.Cells(i, c2) = rs.Cells(i, c2) + rs.Cells(i + 1, c2) 'temr �Է���
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
' ���Ḯ��Ʈ ��� ���� 150402
    Workbooks.Open rPreset("����LIST")
    'Workbooks.Open rPreset("����Base")
End Sub

Sub edMergeTbl1()
'
' ��ũ��2 ��ũ��
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


