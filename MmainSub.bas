Attribute VB_Name = "MmainSub"
'
'======================================================================
' ã�� �Լ�
'======================================================================
'
Sub Go2Database()
'go2Cell() '160701 ���缿 ������ ����Ÿ���̽� ã�ư���
    Dim ws, sv As String
    sv = Selection.Value  'ã���� �ϴ� ��
    
    Dim lngItem As Long
    Dim msg As String
 '
    Dim r As Integer
    Dim i, cn As Integer
    Dim rngDB As Range

'��¿���
    Set rngDB = ThisWorkbook.Sheets("S").Range("setup")   '����Ÿ���̽� ����
    
'    Label1.Caption = "��1"
'    Label2.Caption = "��2"
'
    cn = rngDB.Rows.Count
    MsgBox rngDB.Address
    
    For i = 0 To cn
        With UserForm1.ListBox1
            .ColumnCount = 3
            .ColumnWidths = "150;120;100"
            .ColumnHeads = True
            .AddItem
            .List(i, 0) = rngDB.Cells(i, 1)    'DB ��Ī
            .List(i, 1) = rngDB.Cells(i, 2)    '���
            .List(i, 2) = rngDB.Cells(i, 3)    'cas
        End With
    Next i
 
 
    UserForm1.show
    Stop
    
    tmp = isOpenWrk(vbcWRLST) '���Ḯ��Ʈ Ȱ��ȭ ����
    If tmp = 0 Then
        wf = Application.GetOpenFilename("Excel Files,*.xls")
        Workbooks.Open FileName:=wf, UpdateLinks:=0
    End If
        Workbooks(vbcWRLST).Activate
        Range("Database").Find(What:=sv, LookIn:=xlValues).Activate
End Sub

Sub ingQueryDBInfo()
'���ü��� DB���� ��������
    Dim wb, sv, r, smg As String
    Dim t, i, m As Integer
    Dim w As Workbooks
    Dim rs As Range
    
    sv = Selection.Value  'ã���� �ϴ� ��
    owb = ActiveWorkbook.Name
   
    fn = edTrimPath(hkc_DB1, "\")
    wb = edTrimExtension(fn, ".")
   
     '���Ḯ��Ʈ Ȱ��ȭ ����
    If isOpenWrk(fn) = 0 Then
        MsgBox (fn & " ������ ���ڽ��ϴ�.")
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
' ����(Show) ���� �Լ�
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

Sub shAddSht(inp As String) '120824/��Ʈ �߰�
    
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

Sub shLockSht() '131108/jinsigo ��Ʈ ��ױ�
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

Sub shUnLockSht() '131108/jinsigo ��Ʈ Ǯ��
'
    ActiveSheet.Unprotect Password:="1"
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Locked = False
    Selection.End(xlDown).Select
    
End Sub

Public Sub Clr_Sheet(inp As String) '120824/��Ʈ ���� �����
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        If ws.Name = inp Then
            ws.Cells.ClearContents
        End If
    Next ws
End Sub

Sub Del_Sheet(RefName As String)
'120825/��Ʈ ����
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
' ��ũ��2 ��ũ��
'

'

    Windows("����LIST.xls").Activate
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
        
    Workbooks("����LIST.xls").Windows.arrange ArrangeStyle:=xlArrangeStyleVertical, _
        ActiveWorkbook:=False, SyncVertical:=False

    ActiveWindow.WindowState = xlMaximized
    ActiveWindow.WindowState = xlMinimized
    
    Windows.CompareSideBySideWith "����LIST.xls"
    Windows.ResetPositionsSideBySide
    Windows.arrange ArrangeStyle:=xlVertical
    Windows("����LIST.xls").Activate
    

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
' ������� �Լ�
'======================================================================
'
Sub exCol_Filter2() '061026/ó�������

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

Sub Col_Filter() 'jinsigo/061027/ó��˻� Macro

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

Sub exActiveCell() 'jinsigo/061110/������ (���� �� ����)

    Selection.AutoFilter Field:=ActiveCell.Column, Criteria1:=ActiveCell.Value

End Sub


Sub exInput() ' �������� Macro 06-12-07
'
    Kwds = Application.InputBox(prompt:="�˻��� �Է�: ")
    Keywords1 = "=*" & Kwds & "*"
    Selection.AutoFilter Field:=ActiveCell.Column, Criteria1:=Keywords1, Operator:=xlAnd

End Sub


Sub exNoBlank() 'jinsigo/061012/������ (���� �ƴ� �� ����)

    Selection.AutoFilter Field:=ActiveCell.Column, Criteria1:="<>"

End Sub


Sub exShowAll() '������ �ڵ����Ϳ��� ��κ��� 06-11-10

    ActiveSheet.ShowAllData
End Sub

Sub ������()
'
' ������ ��ũ��
'

'
Dim m As Integer
Dim Keywords1 As String
    Kwds = ActiveSheet.Range("I1").Value
    Keywords1 = "=*" & Kwds & "*"
    ActiveSheet.Range("$A$2:$A$20000").AutoFilter Field:=1, Criteria1:=Keywords1, Operator:=xlAnd
End Sub

Sub exSameBColor_HideColumns()
' ���� ������ �� �����
' 160615
'
    '��ü����
    Range("list").EntireColumn.Hidden = False
    
    '���ú���
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
' ���ÿ��� ��� �� ���̱�
    Range("list").Columns.Select
    Selection.EntireColumn.Hidden = False
End Sub

Sub exSameBColor_HideRows()
' ���� ������ �� �����
' 160615
'
    '��ü����
    Range("list").EntireRows.Hidden = False
    
    '���ú���
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
'/// �� ���� ��� by Jinsigo ///
'======================================================================
'
Sub edMergeCell()
'120813/���ù����� ���ڿ� ��ġ��
'150421 �� & �� ���� Ȯ��
'
    Dim rs As Range  '�Է¼�
    Dim ro As Range  '��¼�
    Dim ss As String '������
    Dim st As String '���� ���ڿ�
    Dim tmp As String
    Dim i As Range
  
    Set rs = Selection
    ss = InputBox("�����ڸ� �Է��� �ּ���(�⺻���� ',' �Դϴ�).", "�� �����ϱ�", ",") '������ �Է�
    st = ""
    
    For Each i In rs.Rows
        For j = 1 To i.Columns.Count
            tmp = i.Cells(1, j)
            st = st + tmp & ss
        Next j
    Next i
    
    Set ro = Application.InputBox("���� ����� ���� ������ �ּ���", "�� �����ϱ�", Type:=8)
    If Right(st, Len(ss)) = ss Then st = Left(st, Len(st) - Len(ss))
    ro.Formula = st
End Sub

Sub edMergeRange() '���ù��� ��ġ�� 150520
    Dim rs As Range  '�Է¹���
    Dim ro As Range  '��¹���
    Dim cs As String '���ǿ�1
    Dim co As String '���ǿ�2
    Dim ns As Integer '�Է¿���
    Dim no As Integer '��¿���
    Dim vs, vo As String '����� ��
    Dim ks, ko As String '���� ���� ���� üũ
    Dim chk As Integer ''���� ���� ���� üũ
  
    Set rs = Application.InputBox("������ ������ ������ �ּ���", "�� �����ϱ�:A", Type:=8)
        
    Set ro = Application.InputBox("������ ������ ������ �ּ���", "�� �����ϱ�:B", Type:=8)
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
            ElseIf ko = ks Then '�ڵ常 �ٸ����
                chk = j + 1
            Else '�ڵ� �� ���� ��� �ٸ� ���
                
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
'120813/���ü��� ���ڿ� ������
    Dim rs As Range '�Է¼�
    Dim ro As Range '��¼�
    Dim ss As String '������
    Dim st  '���� ���ڿ�
    Dim nc As Integer '��¼� ���
    
    Set rs = Selection
    ss = InputBox("�����ڸ� �Է��� �ּ���(�⺻���� ',' �Դϴ�).", "�� ������", ",")
    st = Split(rs.Value, ss)
        
    Set ro = Application.InputBox("����� ù ���� ������ �ּ���", "�� ������", Type:=8)
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
'140923/���ڳ�����.���߼�
'150114/���� ����, for�� j �߰�.
    Dim inCells As Range '�Է¼�
    Dim rc As Integer '�Է� ���
    Dim cc As Integer '�Է� ����
    Dim outCells As Range '��¼�
    Dim sp As String '������
    Dim st  '���� ���ڿ�
    Dim stc '���� ���ڿ���
    
    Set inCells = Selection
    rc = inCells.Rows.Count
    cc = inCells.Columns.Count
    
    sp = InputBox("�����ڸ� �Է��� �ּ���(�⺻���� ',' �Դϴ�).", "�� ������", ",")
    Set outCells = Application.InputBox("����� ù ���� ������ �ּ���", "�� ������", Type:=8)
    
    For i = 1 To rc
        st = Split(inCells.Cells(i, 1), sp)
        stc = UBound(st) + 1
        For j = 1 To stc
            '����ǥ ������ ���� split_value(0) �迭�� ���� �Էµǰ� �������� split_value(1) �迭�� ���� �Էµ˴ϴ�.
           ' inCells.Offset(i - 1, 0).Formula = Trim(st(0)) '�տ� ���� �ѷ��ݴϴ�.
            outCells.Offset(i - 1, j - 1).Formula = Trim(st(j - 1)) '�ڿ� ���� �ѷ��ݴϴ�.
        Next j
    Next i
End Sub


Sub edFillRng_NA2Null() '110415/���ù����� #NA ���� �������� �����ϱ� ��ũ��
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Replace What:="#N/A", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

Sub edTxt2Char() '120119/Ư����ȣ ��ȯ key: Ctrl+Shift+A
    ActiveCell.Replace What:="-^", Replacement:="��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    ActiveCell.Replace What:="&^", Replacement:="��"
    ActiveCell.Replace What:="&v", Replacement:="��"
    ActiveCell.Replace What:="&>", Replacement:="��"
    ActiveCell.Replace What:="&<", Replacement:="��"
    ActiveCell.Replace What:="&c", Replacement:="��"
    ActiveCell.Replace What:="&.", Replacement:="��"
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
'���ڸ� �б�
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



Sub edChange_Pic_Name() '150114/�׸��̸� �ٲٱ�
    Dim shpC As Shape                                    '������ �׸��� ���� ����
    Dim rngShp As Range                                  '�� �׸��� �������� ���� ������ ���� ����
   
    For Each shpC In ActiveSheet.Shapes          '������������ �� �׸��� ��ȯ
        Set rngShp = shpC.TopLeftCell                 '�� �׸��� ������������ ���� ������ ������ ����
       
        If Not Intersect(Columns("B"), rngShp) Is Nothing Then   '�׸��� B���� ��ġ��
                shpC.Name = rngShp.Previous.Value    '�� �׸��� �̸��� A���� �̸����� ����
        End If
    Next shpC
   
End Sub

Option Explicit

Sub fitPictureInCell() '[��ó] (622) ���ÿ����� �׸��� �� ���� ũ�⿡ ���߱� (���� VBA ��ũ��)|�ۼ��� �ϲ�
    Dim rngAll As Range                                    '���ÿ����� ���� ����
    Dim rngShp As Range                                  '�� �׸��� �������� ���� ������ ���� ����
    Dim shpC As Shape                                    '������ ����(shape)�� ���� ����
    Dim rotationDegree As Integer                       '������ ȸ������ ���� ����
   
    Application.ScreenUpdating = False              'ȭ�� ������Ʈ (�Ͻ�)����
   
    If Not TypeOf Selection Is Range Then           '���� �׸� ���� �����ϰų� �Ͽ��� ���
        MsgBox "������ ���õ��� ����", 64, "�������� ����"  '��� �޽��� ���
        Exit Sub                                                 '��ũ�� �ߴ�
    End If
   
    Set rngAll = Selection                                  '���ÿ����� ������ ����
   
    For Each shpC In ActiveSheet.Shapes          '��ü������ �� �׸��� ��ȯ
        If shpC.Type = 13 Then                            '���� �� ������ �׸��̶��
            Set rngShp = shpC.TopLeftCell             '�� ������ ������ ������ ������ ����
           
            If rngShp.MergeCells Then                   'rngShp�� �����յ� ���̶��
                Set rngShp = rngShp.MergeArea       '������ ������ �������� Ȯ��
            End If
           
            If Not Intersect(rngAll, rngShp) Is Nothing Then '�� ������ ��ü������ ���ԵǸ�
                rotationDegree = shpC.Rotation         '�׸��� ȸ������ ������ ����
               
                If rotationDegree = 90 Or rotationDegree = 270 Then '�׸��� 90�� or 270�� ȸ���� ���
               
                    With shpC                                   '�� �׸����� �۾�
                        .LockAspectRatio = msoFalse   '�׸� �¿�������� ����
                        .Rotation = 0                           '�׸� ȸ���� �����·� ���� ����
                        .Height = rngShp.Width - 4        '�׸� ���̸� ���缿 ũ��  - 4
                        .Width = rngShp.Height - 4        '�׸� ���� ���缿 ũ�� - 4
                        .Left = rngShp.Left + (rngShp.Width - shpC.Width) / 2
                                                                    '�׸� �� ��� ��ġ�� ���� �߾ӿ� ������ ����
                        .Top = rngShp.Top + (rngShp.Height - shpC.Height) / 2
                                                                    '�׸����� ��� ��ġ�� ���� �߾ӿ� ������ ����
                        .Rotation = rotationDegree        '�׸� ȸ�� ������ ����
                    End With
               
                Else
                    With shpC                                  '�� �׸����� �۾�
                        .LockAspectRatio = msoFalse  '�׸� �¿�������� ����
                        .Left = rngShp.Left + 2             '�׸�������ġ�� ���� ������ + 2
                        .Top = rngShp.Top + 2             '�׸����� ��ġ��  ���� ������ ��ġ + 2
                        .Height = rngShp.Height - 4      '�׸� ���̸� ���缿 ũ��  - 4
                        .Width = rngShp.Width - 4        '�׸� ���� ���缿 ũ�� - 4
                    End With
                End If
            End If
        End If
    Next shpC
   
    Set rngAll = Nothing                                      '��ü���� �ʱ�ȭ(�޸� ����)
End Sub
   
'
'======================================================================
' ��Ʈ �߰� ����
'======================================================================
'
Sub AddFormatCondition()
    With ActiveSheet.Range("A1:A10").FormatConditions _
        .Add(xlCellValue, xlEqual, "vba��")
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
' ã�� �� ���� �Լ�
'======================================================================
'
Function isOpenName(nn As String) As Integer
'�̸����� ���� ���� üũ 150318
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
'��ũ��Ʈ ���� ���� üũ 150318
    Dim wrk As Workbook
    isOpenWrk = 0
    
    For Each wrk In Workbooks
        If wrk.Name = wb Then
            'MsgBox wrk.name & "�� �̹� ���� �ֽ��ϴ�."
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
' ���ÿ����� ù���� �ʵ���Ƿ� DB ��
    st = Selection.Value
    'r = Application.Match(st, "[����LIST.xls]!database").rows(1), 0)
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
    strFile = "����������.xls"
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
            .Caption = "<<���ư���"
            .OnAction = "GoBack"
        End With
    End With
    Range("a1").Select
    MsgBox "�ڷḦ ��� �о�鿴���ϴ�", vbInformation, "�۾� ����//Exceller"
End Sub

Function isReadValue(path, file, sht, rng) As Variant
    Dim msg As String
    Dim strTemp As String
    
    If Trim(Right(path, 1)) <> "\" Then path = path & "\"
    If Dir(path & file) = "" Then
        ReadValue = "�ش� ������ �����ϴ�"
        Exit Function
    End If
    msg = "'" & path & "[" & file & "]" & sht & "'!" & Range(rng).Range("a1").Address(, , xlR1C1)
    ReadValue = ExecuteExcel4Macro(msg)
End Function

Public Function IsFormula(c)
'�Լ� ���θ� üũ 150518
    IsFormula = c.HasFormula
   
End Function

'
'======================================================================
' ���� ����� �Լ�
'======================================================================
'
Function rPreset(nfile As String) As Range
' �̸�����ǥ�� ���� ����     150401
    
    Dim nam, path, file, sht As String
    Dim rng As Range
    Dim rs As Range
    Dim c As Range
    
    Set rs = Workbooks("HKC.xlsm").Sheets("R").Range("a:k") '���� ����
    
    For Each c In rs.Rows
        If c.Cells(1, 1).Value = nfile Then
            i = c.Row
            Set rs = c.Cells(i, 1).Resize(1, 5)
            MsgBox "i= " & i
        Else
            MsgBox "�ش� �̸��� ���ǵ��� �ʾҽ��ϴ�."
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
        MsgBox "�ش� ������ �����ϴ�"
        Exit Function
    End If
    
    rPreset = "'" & path & "[" & file & "]" & sht & "'!" & rng
       
End Function

Function ioOpenWrk(wb) As Integer
'�������� ���� 150403
    tmp = isOpenWrk(wb)
    If tmp = 0 Then Workbooks.Open wb
    ioOpenWrk = tmp
End Function




Function ioOpenBydNames() As Integer
'�̸����ǿ���.�̸����� �������� ���� 150403
    Dim nwbk As String
    Dim rs As Range
    
    Set rs = rdefNames
    For Each i In rs
        ndir = i.Cells(1, 2).Value  ' ���
        nwbk = i.Cells(1, 3).Value  ' ���ϸ�
        nvis = i.Cells(1, 6).Value  ' 1:����� 0:���̱�
        nrdo = i.Cells(1, 7).Value '1: �б�����
        'MsgBox ndir & nwbk & i
        If i.Cells(1, 1).Interior.ColorIndex <> -4142 And isOpenWrk(nwbk) = 0 Then
            Workbooks.Open FileName:=(ndir & "\" & nwbk), UpdateLinks:=0, ReadOnly:=nrdo
            If nvis = 1 Then ActiveWindow.Visible = False
        End If
    Next i
            
End Function

Sub ioOpendBySelection()
'���ÿ��� �̸����� ���Ͽ��� 150604
    
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
                    ndir = i.Cells(1, 2).Value  ' ���
                    nwbk = i.Cells(1, 3).Value  ' ���ϸ�
                    nvis = i.Cells(1, 6).Value  ' 1:����� 0:���̱�
                    Workbooks.Open FileName:=(ndir & "\" & nwbk), ReadOnly:=1, UpdateLinks:=0
                    If nvis = 1 Then ActiveWindow.Visible = False
                End If
            Next i
        End If
    Next s
End Sub


Sub ioOpendByActiveCell2()
'�̸����ǿ���.�̸����� �������� ���� 150403
    Dim nwbk As String
    Dim rs As Range
    
    dn = Selection.Value
    Set rs = rdefNames
    
    i = rs.Columns(1).Find(dn).Row - 1
    MsgBox i
    ndir = rs.Cells(i, 2).Value  ' ���
    nwbk = rs.Cells(i, 3).Value  ' ���ϸ�
    nvis = rs.Cells(i, 6).Value  ' 1:����� 0:���̱�
    'MsgBox ndir & nwbk & i
    If isOpenWrk(nwbk) = 0 Then
        Workbooks.Open FileName:=(ndir & "\" & nwbk), UpdateLinks:=0, ReadOnly:=1
        If nvis = 1 Then ActiveWindow.Visible = False
    End If
            
End Sub



Sub isOpenXL()
    st = "����LIST"
    MsgBox rPreset(st)
    Workbooks.Open rPreset(st).Cells(1, 1).Value
End Sub
Sub ioLoadModules() '150212/jinsigo ��� ��������
    
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

Sub ioDumpModules()  '150212/jinsigo ��� ������

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
' �Լ� ã�� ����
'

'
    st = Selection.Value
    Application.GoTo Reference:=st
End Sub



'*****
' Source Code: esKillErrName.esAPI  ������15.06.26

'Private Const MAX_PATH As Integer = 255
'Private Declare Function GetSystemDirectory& Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long)
'Private Declare Function GetWindowsDirectory& Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long)
'Private Declare Function GetTempDir Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
        
Function isReturnTempDir()
'�ӽ� ������ ��ȯ Returns Temp Folder Name
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
'�ý��� ������ ��ȯ (C:\WinNT\System32)
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
'OS ������ ��ȯ (C:\Win95)
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
' ��ǻ�� ȯ�� ���� �о����
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
' ��ũ��3 ��ũ��
'

'
    Dim wd, ws As Object
    Dim rs, rd As Range
    
    Set ws = ActiveWindow
    Set wd = Windows("����ǥ���ۼ�5.0.xlsm")
    Set rs = Selection
    
'    Set rd = wd.Sheets("�ۼ�").Range("A6")
    
    'Windows("20150828100214.CDVSK.xls").Activate
    'Range("E2").Select
    Application.Run "����ǥ���ۼ�5.0.xlsm!Clear_Data"
    
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
' ������ ��Ʈ �ۼ�
' 2015.08.28

'
    Dim wd, ws As Object
    Dim rs, rd As Range
    
    Set ws = ActiveWindow
    Set wd = Windows("����ǥ���ۼ�5.0.xlsm")
    Set rs = Selection
    Set rd = Worksheets("������").Range("A6")
    
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
'�̸����� ���� ��������150403
    Dim sc As Range '���ۼ�
    'Set sc = Workbooks("HKC.xlsm").Sheets("R").Range("setup")
    'Set rdefNames = Range(sc, sc.End(xlUp))
    Set rdefNames = Workbooks("HKC.xlsm").Sheets("S").Range("setup")
    
End Function

Sub ioDefines()
'��������� ����(���� �κ�) �̸����� 150406

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
' ����Ʈ to �̸����� 150403
'
    Dim ndef As String
    Dim ndir As String
    Dim nwbk As String
    Dim nsht As String
    Dim nrng As String
    Dim ncom As String
    Dim rdef As String '������(�������)
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

    With ActiveWorkbook.Names("����Ÿ")
        .Name = "����Ÿ2"
        .RefersToR1C1 = "='D:\Documents and Settings\������\My Documents\1.���Ἲ��\����Ÿ.xls'!list"
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

