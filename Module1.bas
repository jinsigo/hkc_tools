Attribute VB_Name = "Module1"
Option Explicit
Option Compare Text

Dim xclass As New Class1 'Ŭ������� ���
Dim strCut As String, rngMulti As Range
Dim lngMinRow As Long, lngMinColumn As Long
Dim lngMaxRow As Long, lngMaxColumn As Long
Dim blnMode As Boolean
Public rngUndo As Range, varUndo As Variant
 '������Ҹ� ���� �ʿ�
#If Win64 Then
  Private Declare PtrSafe Function ImmGetContext Lib "imm32.dll" _
    (ByVal hwnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function ImmSetConversionStatus Lib "imm32.dll" _
    (ByVal himc As LongPtr, ByVal dw1 As Long, ByVal dw2 As Long) As Long
  Private Declare PtrSafe Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
  Private Declare Function ImmGetContext Lib "imm32.dll" _
    (ByVal hwnd As Long) As Long
  Private Declare Function ImmSetConversionStatus Lib "imm32.dll" (ByVal _
    himc As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If
 '�ѱ۸������� ���� �ʿ�

Sub Addin_Open_MyAddin(Optional x As Boolean)
Dim cmdCont As CommandBarControl
    On Error Resume Next
    Set xclass.App = Application 'Ŭ������� ���
    Call ImmSetConversionStatus(ImmGetContext(FindWindow("XLMAIN", _
      Application.Caption)), &H1, &H0) 'IME_CMODE_HANGEUL, IME_SMODE_NONE
    With Application
      With .CommandBars("Tools")
        Set cmdCont = .FindControl(Tag:="My_Addin2015")
        If Not cmdCont Is Nothing Then cmdCont.Delete
        With .Controls.Add(Type:=msoControlPopup, Before:=1, Temporary:=True)
          .Caption = "My_Addin2015"
          .Tag = "My_Addin2015"
          .OnAction = "menu_enable"
          With .Controls.Add(Type:=msoControlPopup)
            .Caption = "���־��� ����Ű"
            .Tag = "���־��� ����Ű"
            With .Controls.Add(Type:=msoControlButton)
              .FaceId = 40
              .Caption = "�Ʒ��� ��������"
              .ShortcutText = "Alt+PageDown"
              .OnAction = "End1_Cell"
            End With
            With .Controls.Add(msoControlButton)
              .FaceId = 39
              .Caption = "������ ��������"
              .ShortcutText = "Alt+PageUp"
              .OnAction = "End2_Cell"
            End With
            With .Controls.Add(msoControlButton)
              .Caption = "�������� �缳��"
              .ShortcutText = "Ctrl+Alt+End"
              .OnAction = "Sheet_Refresh"
            End With
            With .Controls.Add(msoControlButton)
              .BeginGroup = True
              .Caption = "������(����,��) �̵�"
              .ShortcutText = "Ctrl+Alt+����"
              .OnAction = "'Offset_Select " & True & ", " & True & "'"
            End With
            With .Controls.Add(msoControlButton)
              .Caption = "���ù��� Ȯ��"
              .ShortcutText = "Ctrl+Shift+Alt+����"
              .OnAction = "'Tot_Select " & True & ", " & True & "'"
            End With
            With .Controls.Add(msoControlButton)
              .Caption = "���߼� ���� ���"
              .ShortcutText = "Ctrl+Shift+Alt+-"
              .OnAction = "Cancel_EndSelect"
            End With
            With .Controls.Add(msoControlButton)
              .BeginGroup = True
              .FaceId = 368
              .Caption = "G/ǥ�� ����"
              .ShortcutText = "Ctrl+[Shift]+��"
              .OnAction = "'General_Format " & True & "'"
            End With
            With .Controls.Add(msoControlButton)
              .FaceId = 376
              .Caption = "��Ģ�����ϱ�"
              .ShortcutText = "Alt+F10"
              .OnAction = "Calculate_Num"
            End With
            With .Controls.Add(Type:=msoControlButton)
              .BeginGroup = True
              .FaceId = 183
              .Caption = "�����(����) ã��"
              .ShortcutText = "Ctrl+Shift+F"
              .OnAction = "Special_Find"
            End With
            With .Controls.Add(Type:=msoControlButton)
              .FaceId = 202
              .Caption = "����(�ٸ�)�� ã��"
              .ShortcutText = "Ctrl+Alt+F"
              .OnAction = "Diff_Select"
            End With
            With .Controls.Add(Type:=msoControlButton)
              .FaceId = 29
              .Caption = "�ߺ��� ã��"
              .ShortcutText = "Ctrl+Alt+D"
              .OnAction = "Duplicated_Range"
            End With
            With .Controls.Add(Type:=msoControlButton)
              .FaceId = 564
              .Caption = "������ �ٲٱ�"
              .ShortcutText = "Ctrl+Shift+H"
              .OnAction = "Safe_Replace"
            End With
            With .Controls.Add(Type:=msoControlButton)
              .BeginGroup = True
              .Caption = "����(����)�� ����"
              .ShortcutText = "Ctrl+Shift+C"
              .OnAction = "'Special_Copy " & False & "'"
            End With
            With .Controls.Add(Type:=msoControlButton)
              .Caption = "���ռ� �߶󳻱�"
              .ShortcutText = "Ctrl+Shift+X"
              .OnAction = "'Special_Copy " & True & "'"
            End With
            With .Controls.Add(Type:=msoControlButton)
              .Caption = "���߼� �ٿ��ֱ�"
              .ShortcutText = "Ctrl+Shift[Alt]+V"
              .OnAction = "Special_Paste"
            End With
            With .Controls.Add(Type:=msoControlButton)
              .FaceId = 535
              .Caption = "���ϱ� �����ϱ�"
              .ShortcutText = "Ctrl+Alt+A"
              .OnAction = "Paste_AddValue"
            End With
            With .Controls.Add(Type:=msoControlButton)
              .BeginGroup = True
              .FaceId = 309
              .Caption = "���� �ϰ�����(����)"
              .ShortcutText = "Ctrl+Shift+I"
              .OnAction = "Insert_Text"
            End With
            With .Controls.Add(Type:=msoControlButton)
              .FaceId = 653
              .Caption = "���� �����ϱ�"
              .ShortcutText = "Ctrl+Shift+N"
              .OnAction = "Input_Serialnum"
            End With
            With .Controls.Add(Type:=msoControlButton)
              .Caption = "����ä���(�����)"
              .ShortcutText = "Ctrl+Shift+B"
              .OnAction = "BlankCell_Input"
            End With
            With .Controls.Add(Type:=msoControlButton)
              .Caption = "����ä���"
              .ShortcutText = "Ctrl+Alt+B"
              .OnAction = "Formula_MultiInput"
            End With
            With .Controls.Add(msoControlButton)
              .BeginGroup = True
              .FaceId = 3
              .Caption = "������ ����"
              .ShortcutText = "Ctrl+S"
              .OnAction = "Safe_Save"
            End With
            With .Controls.Add(msoControlButton)
              .Caption = "�μ�ݺ��� ����"
              .ShortcutText = "Shift+F11"
              .OnAction = "Print_Title"
            End With
            With .Controls.Add(msoControlButton)
              .Caption = "Ʋ���� ����"
              .ShortcutText = "Shift+Alt+F11"
              .OnAction = "Freeze_Panes"
            End With
            With .Controls.Add(msoControlButton)
              .BeginGroup = True
              .FaceId = 487
              .Caption = "���� ����ǥ��"
              .ShortcutText = "F12"
              .OnAction = "Total_Info"
            End With
          End With
          With .Controls.Add(Type:=msoControlButton)
            .BeginGroup = True
            .Caption = "�糡���� �����"
            .ShortcutText = "Ctrl+Shift+T"
            .OnAction = "Trim_Text"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .FaceId = 47
            .Caption = "���ɹ��� �����"
            .ShortcutText = "Ctrl+Shift+Del"
            .OnAction = "Del_Ascii"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .FaceId = 288
            .Caption = "���� �κ� ���󺯰�"
            .OnAction = "Color_Text"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .BeginGroup = True
            .FaceId = 125
            .Caption = "��¥���� ����"
            .ShortcutText = "Ctrl+Alt+Y"
            .OnAction = "Date_Format"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .Caption = "(�ݿø�)�Լ� �ٷ�����"
            .ShortcutText = "Ctrl+Shift+R"
            .OnAction = "Function_Evaluate"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .Caption = "�Ҽ�(��)�� ����"
            .OnAction = "Fraction_Select"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .BeginGroup = True
            .FaceId = 127
            .Caption = "����(�ߺ�)������ �����ϱ�"
            .ShortcutText = "Ctrl+[Shift]+Alt+O"
            .OnAction = "'Select_UniqData " & True & "'"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .FaceId = 464
            .Caption = "�ߺ��� (������) �ջ��ϱ�"
            .OnAction = "Del_Repetition"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .FaceId = 295
            .Caption = "�������ڸ�ŭ �����"
            .OnAction = "Insert_BlankRowCol"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .FaceId = 633
            .Caption = "�� ���� ���� ���߱�"
            .OnAction = "TwoArea_Adjust"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .Caption = "���뺸�� ����"
            .ShortcutText = "Ctrl+Shift+M"
            .OnAction = "Multi_Murge"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .BeginGroup = True
            .FaceId = 246
            .Caption = "���̺� ���º�ȯ"
            .OnAction = "Table_Conform"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .FaceId = 545
            .Caption = "�ٹٲ޼� ��и�"
            .OnAction = "Str_Split"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .FaceId = 541
            .Caption = "��(��)���� �ϰ��ø���"
            .ShortcutText = "Ctrl+Alt+H"
            .OnAction = "RowHightColWidth"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .FaceId = 620
            .Caption = "������ �����Ͽ� �����ϱ�"
            .OnAction = "DataBase_Split"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .FaceId = 159
            .Caption = "���� ���չ��� �����ϱ�"
            .OnAction = "Wb_Combine"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .FaceId = 455
            .Caption = "�������� �����ϱ�"
            .OnAction = "Make_RecoveryNewBook"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .BeginGroup = True
            .FaceId = 503
            .Caption = "���ڼ��� ����"
            .ShortcutText = "Ctrl+Alt+S"
            .OnAction = "StrAndNum_Sort"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .FaceId = 163
            .Caption = "������ ����"
            .ShortcutText = "Ctrl+Alt+R"
            .OnAction = "Randomize_Sort"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .BeginGroup = True
            .FaceId = 628
            .Caption = "������ ����"
            .ShortcutText = "Shift+Alt+L"
            .OnAction = "Filter_Reverse"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .FaceId = 499
            .Caption = "�������� ����"
            .ShortcutText = "Ctrl+Alt+L"
            .OnAction = "Color_Filter"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .BeginGroup = True
            .FaceId = 160
            .Caption = "���������� �μ�"
            .ShortcutText = "Ctrl+Shift+P"
            .OnAction = "Print_SelectPage"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .BeginGroup = True
            .Caption = "Addin2015 ����"
            .OnAction = "Help_Text"
            .FaceId = 273 '�����
          End With
        End With
      End With
    End With
    Onkey_Make
    Randomize
End Sub

Sub Onkey_Make(Optional x As Boolean)
    With Application
      .OnKey "%{PGDN}", "End1_Cell"
      .OnKey "%{PGUP}", "End2_Cell"
      .OnKey "^%{END}", "Sheet_Refresh"
      .OnKey "^%{RIGHT}", "'Offset_Select " & True & ", " & True & "'"
      .OnKey "^%{LEFT}", "'Offset_Select " & False & ", " & True & "'"
      .OnKey "^%{DOWN}", "'Offset_Select " & True & ", " & False & "'"
      .OnKey "^%{UP}", "'Offset_Select " & False & ", " & False & "'"
      .OnKey "^+%{RIGHT}", "'Tot_Select " & True & ", " & True & "'"
      .OnKey "^+%{LEFT}", "'Tot_Select " & False & ", " & True & "'"
      .OnKey "^+%{DOWN}", "'Tot_Select " & True & ", " & False & "'"
      .OnKey "^+%{UP}", "'Tot_Select " & False & ", " & False & "'"
      .OnKey "^+%{-}", "Cancel_EndSelect"
      .OnKey "+ ", "'EntireRowCol_Select " & True & "'"
      .OnKey "^ ", "'EntireRowCol_Select " & False & "'"
      .OnKey "^{BS}", "'General_Format " & True & "'"
      .OnKey "^+{BS}", "'General_Format " & False & "'"
      .OnKey "%{F10}", "Calculate_Num"
      .OnKey "^+f", "Special_Find"
      .OnKey "^%f", "Diff_Select"
      .OnKey "^%d", "Duplicated_Range"
      .OnKey "^+h", "Safe_Replace"
      .OnKey "^+c", "'Special_Copy " & False & "'"
      .OnKey "^+x", "'Special_Copy " & True & "'"
      .OnKey "^+v", "Special_Paste"
      .OnKey "^%v", "'Special_Paste " & True & "'"
      .OnKey "^%a", "Paste_AddValue"
      .OnKey "^+i", "Insert_Text"
      .OnKey "^+n", "Input_Serialnum"
      .OnKey "^+b", "BlankCell_Input"
      .OnKey "^%b", "Formula_MultiInput"
      .OnKey "^s", "Safe_Save"
      .OnKey "+{F11}", "Print_Title"
      .OnKey "+%{F11}", "Freeze_Panes"
      .OnKey "{F12}", "Total_Info"
      .OnKey "^+t", "Trim_text"
      .OnKey "^+{DEL}", "Del_Ascii"
      .OnKey "^%y", "Date_Format"
      .OnKey "^+r", "Function_Evaluate"
      .OnKey "^%o", "'Select_UniqData " & True & "'"
      .OnKey "^+%o", "'Select_UniqData " & False & "'"
      .OnKey "^+m", "Multi_Murge"
      .OnKey "^%h", "RowHightColWidth"
      .OnKey "^%s", "StrAndNum_Sort"
      .OnKey "^%r", "Randomize_Sort"
      .OnKey "+%l", "Filter_Reverse"
      .OnKey "^%l", "Color_Filter"
      .OnKey "^+p", "Print_SelectPage"
    End With
End Sub

Sub Addin_Close_MyAddin(Optional x As Boolean)
 '�߰������ ������ �޴�����
Dim cmdCont As CommandBarControl
    On Error Resume Next
    With Application
      Set cmdCont = .CommandBars("Tools").FindControl(Tag:="My_Addin2015")
      If Not cmdCont Is Nothing Then cmdCont.Delete
      .OnKey "%{PGDN}"
      .OnKey "%{PGUP}"
      .OnKey "^%{END}"
      .OnKey "^%{RIGHT}"
      .OnKey "^%{LEFT}"
      .OnKey "^%{DOWN}"
      .OnKey "^%{UP}"
      .OnKey "^+%{RIGHT}"
      .OnKey "^+%{LEFT}"
      .OnKey "^+%{DOWN}"
      .OnKey "^+%{UP}"
      .OnKey "^+%{-}"
      .OnKey "+ "
      .OnKey "^ "
      .OnKey "^{BS}"
      .OnKey "^+{BS}"
      .OnKey "%{F10}"
      .OnKey "^+f"
      .OnKey "^%f"
      .OnKey "^%d"
      .OnKey "^+h"
      .OnKey "^+c"
      .OnKey "^+x"
      .OnKey "^+v"
      .OnKey "^%v"
      .OnKey "^%a"
      .OnKey "^+i"
      .OnKey "^+n"
      .OnKey "^+b"
      .OnKey "^%b"
      .OnKey "^s"
      .OnKey "+{F11}"
      .OnKey "+%{F11}"
      .OnKey "{F12}"
      .OnKey "^+t"
      .OnKey "^+{DEL}"
      .OnKey "^%y"
      .OnKey "^+r"
      .OnKey "^%o"
      .OnKey "^+%o"
      .OnKey "^+m"
      .OnKey "^%h"
      .OnKey "^%s"
      .OnKey "^%r"
      .OnKey "+%l"
      .OnKey "^%l"
      .OnKey "^+p"
      .CommandBars("MyAddinFirst").Delete
      .CommandBars("MyAddinSecond").Delete
    End With
End Sub

Private Sub Menu_Enable()
 '�ƹ��� ��Ʈ�� ���� ��� �޴��� ��Ȱ��ȭ
Dim stChk As Worksheet, lngI As Long
    On Error GoTo Err_Step
    Set stChk = ActiveSheet
    If stChk Is Nothing Then
      With Application.CommandBars("Tools").Controls("My_Addin2015")
        For lngI = 1 To .Controls.Count - 1
          .Controls(lngI).Enabled = False
        Next lngI
      End With
    Else
      With Application.CommandBars("Tools").Controls("My_Addin2015")
        For lngI = 1 To .Controls.Count - 1
          .Controls(lngI).Enabled = True
        Next lngI
      End With
    End If
Err_Step:
End Sub

Private Sub End1_Cell()
 '����Ű Alt+PageDown
 '�����͹����� �ż�����
Dim rngSelec As Range, rngCell As Range, rngTarget As Range
Dim lngI As Long
    On Error Resume Next
    Set rngSelec = Selection
    Set rngCell = Cells(ActiveCell.SpecialCells(xlLastCell).Row + 1, ActiveCell.Column)
    Application.GoTo rngCell
    '�Ʒ��ʼ��� ȭ���̵�
    Set rngTarget = Range(rngSelec.Areas(1), _
      Cells(rngSelec.Areas(1).SpecialCells(xlLastCell).Row, rngSelec.Areas(1).Column))
    For lngI = 2 To rngSelec.Areas.Count
      Set rngTarget = Union(rngTarget, Range(rngSelec.Areas(lngI), _
        Cells(rngSelec.Areas(lngI).SpecialCells(xlLastCell).Row, rngSelec.Areas(lngI).Column)))
    Next lngI
    ActiveSheet.ScrollArea = rngCell.Address
    rngTarget.Select
    ActiveSheet.ScrollArea = ""
    Application.OnRepeat "", ""
End Sub

Private Sub End2_Cell()
 '����Ű Alt+PageUp
 '�����͹����� �ż�����
Dim rngSelec As Range, rngCell As Range, rngTarget As Range
Dim lngI As Long
    On Error Resume Next
    Set rngSelec = Selection
    Set rngCell = Cells(ActiveCell.Row, ActiveCell.SpecialCells(xlLastCell).Column + 1)
    Application.GoTo rngCell
    Set rngTarget = Range(rngSelec.Areas(1), _
      Cells(rngSelec.Areas(1).Row, rngSelec.Areas(1).SpecialCells(xlLastCell).Column))
    For lngI = 2 To rngSelec.Areas.Count
      Set rngTarget = Union(rngTarget, Range(rngSelec.Areas(lngI), _
        Cells(rngSelec.Areas(lngI).Row, rngSelec.Areas(lngI).SpecialCells(xlLastCell).Column)))
    Next lngI
    ActiveSheet.ScrollArea = rngCell.Address
    rngTarget.Select
    ActiveSheet.ScrollArea = ""
    Application.OnRepeat "", ""
End Sub

Private Sub Sheet_Refresh()
 '����Ű Ctrl+Alt+End
 '������ ���������� �������ϹǷν� ���ʿ��� ������뿡 ���� �����ؼ�
Dim rngNow As Range, rngLastcell As Range
Dim rngRow As Range, rngCol As Range
Dim lngR As Long, lngC As Long
Dim objDraw As Object
    On Error GoTo Err_Step
    Set rngNow = Selection
    Set rngLastcell = ActiveSheet.UsedRange
    If rngLastcell.Address = "$A$1" Then
      Exit Sub
    End If
    If Application.CountA(rngLastcell.SpecialCells(xlLastCell)(1).EntireRow) = 0 Then
      If ActiveSheet.FilterMode = False Then
        rngLastcell.SpecialCells(xlLastCell).Select
        If MsgBox("��������(" & rngLastcell.SpecialCells(xlLastCell).Address _
          & "��)���� ���ĸ� �ֽ��ϴ�. �� ���ʿ��� ������ �����Ͽ� " & _
          "���������� �����Ϳ��������� ����Ͻðڽ��ϱ�?", _
          vbYesNo + vbInformation) = vbYes Then
          Set rngRow = rngLastcell.Find(What:="*", After:=rngLastcell(1), _
            SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
          If rngRow Is Nothing Then
            Set rngRow = Range("A1")
            Set rngCol = Range("A1")
          Else
            Set rngCol = rngLastcell.Find(What:="*", After:=rngLastcell(1), _
              SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
          End If
          rngRow.EntireRow.Select
          lngR = Selection.Rows.Count
          rngCol.EntireColumn.Select
          lngC = Selection.Columns.Count
          With Range(Cells(rngRow.Row + lngR, 1), _
            Cells(rngRow.Row + lngR, 1).End(xlDown)).EntireRow
            .Clear
            .Delete
          End With
          With Range(Cells(1, rngCol.Column + lngC), _
            Cells(1, rngCol.Column + lngC).End(xlToRight)).EntireColumn
             .Clear
             .Delete
          End With
          If ActiveSheet.DrawingObjects.Count Then
            Set rngLastcell = ActiveSheet.UsedRange.SpecialCells(xlLastCell) _
              (ActiveSheet.UsedRange.SpecialCells(xlLastCell).Cells.Count)
            For Each objDraw In ActiveSheet.DrawingObjects
              If objDraw.Top >= rngLastcell(2).Top Or objDraw.Left >= rngLastcell(1, 2).Left Then
                If objDraw.Placement = xlMoveAndSize Then
                  objDraw.Delete
                End If
              ElseIf objDraw.Width = 0 And objDraw.Height = 0 Then
                objDraw.Delete
              End If
            Next objDraw
          End If
        End If
      End If
    End If
    Set rngLastcell = ActiveSheet.UsedRange.SpecialCells(xlLastCell) _
      (ActiveSheet.UsedRange.SpecialCells(xlLastCell).Cells.Count)
    rngLastcell.Select
    MsgBox "���������� " & rngLastcell.Address & " (��)�� �����Ǿ����ϴ�." _
      , vbInformation
    rngNow.Select
    Intersect(rngNow, Range(rngNow(1), rngNow.SpecialCells(xlLastCell))).Select
Err_Step:
    Application.OnRepeat "", ""
End Sub

Sub Offset_Select(x As Boolean, Y As Boolean)
 '����Ű Shift+Alt+Right[Left,Down,Up]
 '�����͹��� ����
Dim strAddress As String
    On Error GoTo Err_Step
    If Selection.Areas(1).MergeCells Then
      strAddress = Selection.Areas(1).Address
      Range(strAddress).UnMerge
    End If
    If x Then
      If Y Then
        Selection.Offset(0, 1).Select
      Else
        Selection.Offset(1, 0).Select
      End If
    Else
      If Y Then
        Selection.Offset(0, -1).Select
      Else
        Selection.Offset(-1, 0).Select
      End If
    End If
    If Not strAddress = vbNullString Then
      Range(strAddress).Merge
    End If
    Application.OnRepeat "", ""
Err_Step:
End Sub

Sub Tot_Select(x As Boolean, Y As Boolean)
 '����Ű Ctrl+Shift+Alt+Right[Left,Down,Up]
 '�����͹��� ����
Dim lngArea As Long, lngRow As Long, lngCol As Long, lngI As Long
Dim rngArea As Range, rngTemp As Range
Dim strAddress As String
    On Error GoTo Err_Step
    lngArea = Selection.Areas.Count
    If Selection.Areas(1).MergeCells Then
      strAddress = Selection.Areas(1).Address
      Range(strAddress).UnMerge
    End If
    If x Then
      If Y Then
        Set rngArea = Intersect(Selection, Selection.Areas(lngArea).EntireRow)
        lngI = rngArea.Areas.Count
        If lngI > 1 Then
          lngCol = rngArea.Areas(lngI)(1).Column - _
            rngArea.Areas(lngI - 1)(1).Column
        End If
        If lngCol > 0 Then
          Set rngArea = Selection.Offset(0, lngCol)
          Set rngArea = Union(Selection, rngArea)
        Else
          Set rngArea = Union(Selection, Selection.Offset(0, 1))
        End If
      Else
        Set rngArea = Intersect(Selection, Selection.Areas(lngArea).EntireColumn)
        lngI = rngArea.Areas.Count
        If lngI > 1 Then
          lngRow = rngArea.Areas(lngI)(1).Row - _
            rngArea.Areas(lngI - 1)(1).Row
        End If
        If lngRow > 0 Then
          Set rngArea = Selection.Offset(lngRow, 0)
          Set rngArea = Union(Selection, rngArea)
        Else
          Set rngArea = Union(Selection, Selection.Offset(1, 0))
        End If
      End If
    Else
      If Y Then
        Set rngArea = Intersect(Selection, Selection.Areas(lngArea).EntireRow)
        lngI = rngArea.Areas.Count
        If lngI > 1 Then
          lngCol = Selection.Areas(lngArea)(1).Column
        End If
        If lngCol > 0 Then
          Set rngArea = Intersect(Selection, _
            Range(Cells(1, 1), Cells(1, lngCol - 1)).EntireColumn)
        Else
          If Selection.Columns.Count > 1 Then
            Set rngArea = Intersect(Selection, Selection.Offset(0, 1)).Offset(0, -1)
          Else
            Set rngArea = Selection.Offset(0, -1)
          End If
        End If
      Else
        Set rngArea = Intersect(Selection, Selection.Areas(lngArea).EntireColumn)
        lngI = rngArea.Areas.Count
        If lngI > 1 Then
          lngRow = Selection.Areas(lngArea)(1).Row
        End If
        If lngRow > 0 Then
          Set rngArea = Intersect(Selection, _
            Range(Cells(1, 1), Cells(lngRow - 1, 1)).EntireRow)
        Else
          If Selection.Rows.Count > 1 Then
            Set rngArea = Intersect(Selection, Selection.Offset(1, 0)).Offset(-1, 0)
          Else
            Set rngArea = Selection.Offset(-1, 0)
          End If
        End If
      End If
    End If
    If Not strAddress = vbNullString Then
      Range(strAddress).Merge
    End If
    Set rngTemp = rngArea.Areas(rngArea.Areas.Count)
    Application.GoTo rngTemp.Cells(rngTemp.Cells.Count)
    ActiveSheet.ScrollArea = ActiveCell.Address
    rngArea.Select
    rngArea.Areas(rngArea.Areas.Count)(1).Activate
    ActiveSheet.ScrollArea = ""
Err_Step:
    Application.OnRepeat "", ""
End Sub

Private Sub Cancel_EndSelect()
Dim rngTarget As Range, rngRe As Range
Dim lngI As Long, lngCount As Long
    On Error GoTo Err_Step
    Set rngTarget = Selection
    lngCount = rngTarget.Areas.Count
    If lngCount > 1 Then
      Set rngRe = rngTarget.Areas(1)
      For lngI = 2 To lngCount - 1
        Set rngRe = Union(rngRe, rngTarget.Areas(lngI))
      Next lngI
      rngRe.Select
      rngRe.Areas(lngI - 1)(1).Activate
    End If
Err_Step:
End Sub

Sub EntireRowCol_Select(x As Boolean)
 '��ü ��(��)�� �����Ͽ� �ݴϴ�.
Dim rngTemp As Range
    On Error GoTo Err_Step
    Set rngTemp = ActiveCell
    Application.ScreenUpdating = False
    If x Then
      Selection.EntireRow.Select
    Else
      Selection.EntireColumn.Select
    End If
    rngTemp.Activate
    Application.ScreenUpdating = True
Err_Step:
End Sub

Sub General_Format(x As Boolean)
 '����Ű Ctrl+[Shift]+��
 '���ڿ�ǥ������ ���� �Ϲ�ǥ���������� �ż��ϰ� �ٲߴϴ�.
Dim rngSelect As Range, rngArea As Range
    On Error GoTo Err_Step
    If x Then
      Selection.NumberFormatLocal = "G/ǥ��"
      Application.SendKeys "{F2}"
    Else
      Selection.NumberFormatLocal = "G/ǥ��"
      Set rngSelect = Intersect(Selection, ActiveSheet.UsedRange)
      For Each rngArea In rngSelect
        rngArea = rngArea.Formula
      Next rngArea
    End If
Err_Step:
    Application.OnRepeat "", ""
End Sub

Private Sub Calculate_Num()
 '����Ű Alt+F10
 '���� ���ϱ�, ����, ������, ���ϱ� ���� ����
Dim rngConstant As Range, rngEacharea As Range
Dim varQues
    On Error GoTo Err_Step
    Set rngUndo = Intersect(Selection, Selection.SpecialCells(xlCellTypeVisible))
    If rngUndo.Areas.Count = 1 Then
      varUndo = rngUndo.Formula
    End If
    Set rngConstant = Intersect(rngUndo, _
      Selection.SpecialCells(xlCellTypeConstants, 5))
    varQues = InputBox("�����ϰ� ���� ��ȣ �� ���ڸ� �Է��ϼ���.", Default:="*1000")
    If varQues = vbNullString Then Exit Sub
    If varQues Like "[/,*]0" Then
      MsgBox "0�� ���ϰų� �������Ͻø� ����մϴ�.", vbInformation
      Exit Sub
    End If

    '�����Ͽ��ٿ��ֱ� ��ɻ���� ������ ����
    For Each rngEacharea In rngConstant.Areas
      rngEacharea = Application.Evaluate(rngEacharea.Address & varQues)
    Next rngEacharea
    If rngUndo.Areas.Count = 1 Then
      Application.OnUndo "���� ���", "Action_Undo"
    End If
Err_Step:
    If Err.Number = 1004 Then
      MsgBox "������ ������ ����� �����ϴ�.", vbInformation
    End If
    Application.OnRepeat "", ""
End Sub

Private Sub Print_Title()
 '����Ű Shift+F11
 '�ݺ��� ���� �� ������ �Ҽ��ֵ���
    On Error GoTo Err_Step
    If ActiveSheet.PageSetup.PrintTitleRows <> "" Then
      ActiveSheet.PageSetup.PrintTitleRows = ""
      MsgBox "�μ�ݺ����� �����Ǿ����ϴ�.", vbInformation
    Else
      ActiveSheet.PageSetup.PrintTitleRows = Selection.EntireRow.Address
      MsgBox "�μ�ݺ����� " & Selection.EntireRow.Address & _
          " �� �����Ǿ����ϴ�.", vbInformation
    End If
Err_Step:
    Application.OnRepeat "", ""
End Sub

Private Sub Freeze_Panes()
 '����Ű Shift+Alt+F11
 'Ʋ���� ���� �� ������ �Ҽ��ֵ���
    On Error GoTo Err_Step
    If ActiveWindow.FreezePanes = True Then
      ActiveWindow.FreezePanes = False
      MsgBox "Ʋ������ �����Ǿ����ϴ�.", vbInformation
    Else
      ActiveWindow.FreezePanes = True
      MsgBox "Ʋ������ �����Ǿ����ϴ�.", vbInformation
    End If
Err_Step:
    Application.OnRepeat "", ""
End Sub

Private Sub Total_Info()
 '����Ű F12
 '���ÿ����� ���� ������ �����ݴϴ�.
Dim rngVisible As Range, rngEach As Range
Dim varArr As Variant
Dim varSum As Variant, varMod As Variant
Dim colTemp As New Collection
Dim lngRow As Long, lngCol As Long, lngCount As Long
Dim strInfo As String
    On Error GoTo Err_Step
    Set rngVisible = Intersect(Selection, _
        Selection.SpecialCells(xlCellTypeVisible))
    On Error Resume Next
    varSum = Application.Sum(rngVisible)
    If IsNumeric(varSum) Then
      varMod = varSum - Int(varSum)
      If varMod = 0 Then
        strInfo = strInfo & "�հ� : " & _
          Application.Text(Application.Sum(rngVisible), "#,##0") & vbCr
      Else
        strInfo = strInfo & "�հ� : " & _
          Application.Text(Application.Sum(rngVisible), "#,##0.############") & vbCr
      End If
    End If
    strInfo = strInfo & " (" & _
      Application.Text(Application.Sum(rngVisible) * 0.3025, "#,##0.0") & "��, " & _
      Application.Text(Application.Sum(rngVisible) / 0.3025, "#,##0.0") & "��)" & vbCr
    If Selection.Areas.Count = 2 Then
      strInfo = strInfo & " ���� : " & _
        Application.Sum(Intersect(Selection.Areas(1), rngVisible)) - _
        Application.Sum(Intersect(Selection.Areas(2), rngVisible)) & vbCr
      strInfo = strInfo & " ���ϱ� : " & _
        Application.Sum(Intersect(Selection.Areas(1), rngVisible)) * _
        Application.Sum(Intersect(Selection.Areas(2), rngVisible)) & vbCr
    End If
    strInfo = strInfo & "���ü� : " & rngVisible.Cells.Count & " ("
    strInfo = strInfo & Intersect(Selection, Selection(1).EntireColumn).Cells.Count & "��*"
    strInfo = strInfo & Intersect(Selection, Selection(1).EntireRow).Cells.Count & "��)" & vbCr
    strInfo = strInfo & "���� : " & Application.CountA(rngVisible) & vbCr
    strInfo = strInfo & "���ڰ��� : " & Application.Count(rngVisible) & vbCr
    If rngVisible.Cells.Count <= 10000 Then
      Err = 0
      If rngVisible.Cells.Count = 1 Then
        If rngVisible.Value <> vbNullString Then
          lngCount = 1
        End If
      Else
        For Each rngEach In rngVisible.Areas
          If rngEach.Cells.Count = 1 Then
            If rngEach.Value <> vbNullString Then
              colTemp.Add 0, CStr(rngEach.Value)
              If Err > 0 Then
                Err = 0
              Else
                lngCount = lngCount + 1
              End If
            End If
          Else
            varArr = rngEach.Value
            For lngRow = 1 To UBound(varArr, 1)
              For lngCol = 1 To UBound(varArr, 2)
                If varArr(lngRow, lngCol) <> vbNullString Then
                  colTemp.Add 0, CStr(varArr(lngRow, lngCol))
                  If Err > 0 Then
                    Err = 0
                  Else
                    lngCount = lngCount + 1
                  End If
                End If
              Next lngCol
            Next lngRow
          End If
        Next rngEach
      End If
      strInfo = strInfo & "���������ͼ� : " & lngCount & vbCr
    End If
    If rngVisible.Cells.Count = 1 Then
      strInfo = strInfo & "���ڿ����� : " & Len(rngVisible) & vbCr
    End If
    strInfo = strInfo & "��� : " & Application.Average(rngVisible) & vbCr
    strInfo = strInfo & "�ִ밪 : " & Application.Max(rngVisible) & vbCr
    strInfo = strInfo & "�ּҰ� : " & Application.Min(rngVisible)
    If ActiveSheet.FilterMode Then
      With Intersect(Names("'" & ActiveSheet.Name & "'!_FilterDatabase") _
        .RefersToRange.EntireRow, ActiveCell.EntireColumn)
        strInfo = strInfo & vbCr & vbCr & .Cells.Count - 1 & "�� �� " & _
         .SpecialCells(xlCellTypeVisible).Count - 1 & "���� ���ڵ尡 ���͵�"
      End With
    End If
    If ActiveWorkbook.Styles.Count > 300 Then
      strInfo = strInfo & vbCr & vbCr & "(�߿�)" & vbCr & ActiveWorkbook.Name & _
        " ���Ͽ� " & vbCr & ActiveWorkbook.Styles.Count & _
        "���� ��Ÿ���� �ֽ��ϴ�." & vbCr & _
        "��Ÿ���� �����ϰ� ���� ��� �ɰ��� ������ �߱�� �� �ֽ��ϴ�." & vbCr & _
        "MyAddin�� �ִ� ""�������� �����ϱ�"" ������� " & vbCr & _
        "���������� ���� ����ñ� �ٶ��ϴ�.(���±���!!)"
    End If
    MsgBox strInfo, vbInformation
Err_Step:
    Application.OnRepeat "", ""
End Sub

Private Sub Special_Find()
 '����Ű Ctrl+Shift+F
 '������� ã�� ���� ã���ݴϴ�.
    On Error GoTo Err_Step
    If Selection.Areas.Count > 1 Then
      MsgBox "���߹��������� ������ �� ���� ����Դϴ�.", vbInformation
      Exit Sub
    End If
    Join_Find.show
Err_Step:
End Sub

Private Sub Diff_Select()
 '����Ű Ctrl+Alt+F
 'ActiveCell�� ���� �ٸ��� �Ǵ� ���������� ã���ݴϴ�.
Dim colTemp As New Collection
Dim rngTarget As Range, rngEach As Range, rngUnion As Range
Dim blnChk As Boolean
Dim lngRow As Long, lngCol As Long
Dim varArr As Variant
    On Error GoTo Err_Step
    Set rngTarget = Intersect(Selection, Selection.SpecialCells(xlCellTypeVisible))
    Select Case MsgBox("�������� ã������ ����, Ȱ������ ���� �ٸ����� ã������ �ƴϿ��� Ŭ���ϼ���.", _
      vbInformation + vbYesNoCancel)
    Case vbYes
      On Error Resume Next
      Set rngUnion = rngTarget.SpecialCells(xlCellTypeConstants, 16)
      If Err > 0 Then
        Err.Clear
        Set rngUnion = rngTarget.SpecialCells(xlCellTypeFormulas, 16)
      Else
        Set rngUnion = Union(rngUnion, rngTarget.SpecialCells(xlCellTypeFormulas, 16))
      End If
      rngUnion.Select
    Case vbNo
      colTemp.Add 0, CStr(ActiveCell.Value)
      On Error Resume Next
      For Each rngEach In rngTarget.Areas
        If rngEach.Cells.Count = 1 Then
          colTemp.Add 0, CStr(rngEach.Value)
          If Err > 0 Then
            Err.Clear
          Else
            If blnChk Then
              Set rngUnion = Union(rngUnion, rngEach)
            Else
              Set rngUnion = rngEach
              blnChk = True
            End If
            colTemp.Remove CStr(rngEach.Value)
          End If
        Else
          varArr = rngEach.Value
          For lngRow = 1 To UBound(varArr, 1)
            For lngCol = 1 To UBound(varArr, 2)
              colTemp.Add 0, CStr(varArr(lngRow, lngCol))
              If Err > 0 Then
                Err.Clear
              Else
                If blnChk Then
                  Set rngUnion = Union(rngUnion, rngEach(lngRow, lngCol))
                Else
                  Set rngUnion = rngEach(lngRow, lngCol)
                  blnChk = True
                End If
                colTemp.Remove CStr(varArr(lngRow, lngCol))
              End If
            Next lngCol
          Next lngRow
        End If
      Next rngEach
      rngUnion.Select
    End Select
Err_Step:
    Application.OnRepeat "", ""
End Sub

Private Sub Duplicated_Range()
 '����Ű Ctrl+Alt+D
 '�ߺ��Ǵ� ������ �����Ͽ� �ݴϴ�.
Dim colData As New Collection
Dim rngTarget As Range, rngEach As Range
Dim varInput As Variant, varOut() As Variant
Dim lngRow As Long, lngCol As Long
Dim lngI As Long, lngJ As Long, lngK As Long
Dim lngTmp As Long, lngTemp As Long
    On Error GoTo Err_Step
    Set rngTarget = Intersect(Selection, Selection.SpecialCells(xlCellTypeVisible))
    With colData
      On Error Resume Next
      For Each rngEach In rngTarget.Areas
        varInput = rngEach.Value
        If IsArray(varInput) Then
          lngRow = UBound(varInput, 1)
          lngCol = UBound(varInput, 2)
          For lngI = 1 To lngRow
            For lngJ = 1 To lngCol
              If Not varInput(lngI, lngJ) = vbNullString Then
                lngK = lngK + 1
                .Add Array(varInput(lngI, lngJ), 1, lngK, _
                  "1_" & STRFORSORT(varInput(lngI, lngJ), 8)), CStr(varInput(lngI, lngJ))
                If Err > 0 Then
                  lngTmp = .Item(CStr(varInput(lngI, lngJ)))(1)
                  lngTemp = .Item(CStr(varInput(lngI, lngJ)))(2)
                  .Remove CStr(varInput(lngI, lngJ))
                  .Add Array(varInput(lngI, lngJ), lngTmp + 1, lngTemp, _
                    "0_" & STRFORSORT(varInput(lngI, lngJ), 8)), CStr(varInput(lngI, lngJ))
                  Err.Clear
                End If
              End If
            Next lngJ
          Next lngI
        Else
          If Not varInput = vbNullString Then
            lngK = lngK + 1
            .Add Array(varInput, 1, lngK, _
              "1_" & STRFORSORT(varInput, 8)), CStr(varInput)
            If Err > 0 Then
              lngTmp = .Item(CStr(varInput))(1)
              lngTemp = .Item(CStr(varInput))(2)
              .Remove CStr(varInput)
              .Add Array(varInput, lngTmp + 1, lngTemp, _
                "0_" & STRFORSORT(varInput, 8)), CStr(varInput)
              Err.Clear
            End If
          End If
        End If
      Next rngEach
      ReDim varOut(1 To .Count, 1 To 5)
      For lngI = 1 To .Count
        varOut(lngI, 1) = .Item(lngI)(0)
        varOut(lngI, 4) = .Item(lngI)(1)
        varOut(lngI, 3) = Application.Text(varOut(lngI, 4), "????0")
        varOut(lngI, 5) = .Item(lngI)(3)
      Next lngI
      Quick_Sort varOut, 5, True, 5, 1, lngI - 1
      With Uniq_Items
        .Label3.Caption = "��ü�� : " & rngTarget.Cells.Count & _
          "�� / ���������� : " & lngI - 1 & "��"
        .ListBox1.List = varOut
        .ListBox1.ListIndex = 0
        .ListBox1.SetFocus
        .show vbModeless
      End With
    End With
Err_Step:
    Application.OnRepeat "", ""
End Sub

Private Sub Safe_Replace()
 '����Ű Ctrl+Shift+H
 'ã��-�ٲٱ⸦ �����ϰ� �� �� �ֵ��� �����ݴϴ�.
Dim strChk As String
    On Error GoTo Err_Step
    strChk = ActiveCell.Text
    Safe_Rep.show
Err_Step:
End Sub

Function AscExtract(strInput As Variant)
 'chrW�� �ۼ��� ���ڿ��� �ƽ�Ű���ڷ� ��ȯ�մϴ�.
Dim strTemp As String
Dim varMatch As Variant
Dim lngI As Long
    With CreateObject("vbscript.regexp")
      .Global = True
      .Pattern = "chr[Ww]\(\d+\)"
      Set varMatch = .Execute(strInput)
      For lngI = 0 To varMatch.Count - 1
        strTemp = varMatch(lngI)
        .Pattern = "\d+"
        strTemp = .Execute(strTemp)(0)
        strInput = Application.Substitute(strInput, "chrw(" & strTemp & ")", ChrW(strTemp))
      Next lngI
    End With
    AscExtract = strInput
End Function

Private Sub Insert_Text()
 '����Ű Ctrl+Shift+I
 '���ڿ��� �鿩����, ���� ���� �ϰ��� �� �� ��
Dim strChk As String
    On Error GoTo Err_Step
    strChk = ActiveCell.Text
    Text_Ins.show
Err_Step:
End Sub

Private Sub Color_Text()
 '���ڿ��� �Ϻκ��� �˻��Ͽ� ������ �����մϴ�.
Dim rngText As Range
    On Error GoTo Err_Step
    Set rngText = Selection.SpecialCells(xlCellTypeConstants, 2)
    Text_Color.show
    Exit Sub
Err_Step:
    If Err.Number = 1004 Then
      MsgBox "������ ������ ���ڰ� �����ϴ�.", vbInformation
    End If
End Sub

Private Sub Trim_Text()
 '����Ű Ctrl+Shift+T
 '���ڿ� �糡�� ���ʿ��� ������ �����ݴϴ�.
Dim rngText As Range, rngTemp  As Range, strTemp As String
Dim lngI As Long, lngJ As Long, lngK As Long, lngL As Long, lngQues As Long
Dim varTemp As Variant
    On Error GoTo Err_Step
    Application.EnableCancelKey = xlErrorHandler
    Set rngUndo = Intersect(Selection, Selection.SpecialCells(xlCellTypeVisible))
    If rngUndo.Areas.Count = 1 Then
      varUndo = rngUndo.Formula
    End If
    Set rngText = Intersect(rngUndo, _
      Selection.SpecialCells(xlCellTypeConstants, 2))
    lngQues = MsgBox("���ڿ��� Trim�մϴ�. ���ڿ� �޺κи� Trim�ҷ��� �ƴϿ��� " _
      & "Ŭ���ϼ���.", vbYesNoCancel + vbInformation)
    Select Case lngQues
    Case vbYes
      For lngI = 1 To rngText.Areas.Count
        If rngText.Areas(lngI).Cells.Count = 1 Then
          rngText.Areas(lngI).Value = Application.Trim(rngText.Areas(lngI).Value)
        Else
          varTemp = rngText.Areas(lngI).Value
          rngText.Areas(lngI) = Evaluate("IF(" & _
            rngText.Areas(lngI).Address & "=" & rngText.Areas(lngI).Address & _
            ",TRIM(" & rngText.Areas(lngI).Address & "))")
          For lngJ = 1 To UBound(varTemp, 1)
            For lngK = 1 To UBound(varTemp, 2)
              If Len(varTemp(lngJ, lngK)) > 255 Then
                rngText.Areas(lngI)(lngJ, lngK) = _
                  Application.Trim(varTemp(lngJ, lngK))
              End If
            Next lngK
          Next lngJ
        End If
      Next lngI
    Case vbNo
      For lngI = 1 To rngText.Areas.Count
        If rngText.Areas(lngI).Cells.Count = 1 Then
          strTemp = Application.Trim(rngText.Areas(lngI).Value)
          lngJ = Application.Find(Left(strTemp, 1), rngText.Areas(lngI).Value)
          rngText.Areas(lngI).Value = String(lngJ - 1, " ") & strTemp
        Else
          varTemp = rngText.Areas(lngI).Value
          rngText.Areas(lngI) = Evaluate("IF(" & rngText.Areas(lngI).Address _
            & "=" & rngText.Areas(lngI).Address & _
            ",REPT("" "",FIND(LEFT(TRIM(" & rngText.Areas(lngI).Address & "),1)," & _
            rngText.Areas(lngI).Address & ")-1)&TRIM(" & rngText.Areas(lngI).Address & "))")
          For lngJ = 1 To UBound(varTemp, 1)
            For lngK = 1 To UBound(varTemp, 2)
              If Len(varTemp(lngJ, lngK)) > 255 Then
                strTemp = Application.Trim(varTemp(lngJ, lngK))
                lngL = Application.Find(Left(strTemp, 1), varTemp(lngJ, lngK))
                rngText.Areas(lngI)(lngJ, lngK).Value = String(lngL - 1, " ") & strTemp
              End If
            Next lngK
          Next lngJ
        End If
      Next lngI
    End Select
    If rngUndo.Areas.Count = 1 Then
      Application.OnUndo "�糡���� ����� ���", "Action_Undo"
    End If
Err_Step:
    If Err.Number = 18 Then
      MsgBox "����ڿ� ���� ��ҵǾ����ϴ�.", vbInformation
    ElseIf Err.Number = 1004 Then
      MsgBox "������ ������ ���ڰ� �����ϴ�.", vbInformation
    End If
    Application.OnRepeat "", ""
End Sub

Private Sub Del_Ascii()
 '����Ű Ctrl+Shift+Del
 '��ǥ����, ���ɹ��ڸ� ����ϴ�.
Dim rngDB As Range, rngEach As Range
Dim varTemp As Variant
Dim lngI As Long, lngJ As Long
    On Error GoTo Err_Step
    Application.EnableCancelKey = xlErrorHandler
    If Selection.Cells.Count = 1 Then
      Set rngDB = ActiveSheet.UsedRange
    Else
      Set rngDB = Selection
    End If
    If MsgBox("��ǥ, ����, �������ڸ� �����մϴ�.", vbInformation + vbYesNo) = vbNo Then
      Exit Sub
    End If
    Application.ScreenUpdating = False
    With rngDB
      .Replace What:=ChrW(13), Replacement:=vbNullString, LookAt:=xlPart
      .Replace What:=ChrW(160), Replacement:=" "
    End With
    Set rngDB = Intersect(rngDB.SpecialCells(xlCellTypeVisible), _
      rngDB.SpecialCells(xlCellTypeConstants))
    For Each rngEach In rngDB.Areas
      varTemp = rngEach.Value
      If IsArray(varTemp) Then
        For lngI = 1 To UBound(varTemp, 1)
          For lngJ = 1 To UBound(varTemp, 2)
            varTemp(lngI, lngJ) = StrConv(varTemp(lngI, lngJ), vbNarrow)
          Next lngJ
        Next lngI
      Else
        varTemp = StrConv(varTemp, vbNarrow)
      End If
      rngEach = varTemp
    Next rngEach
    Application.ScreenUpdating = True
    MsgBox "��ǥ,����, �������ڸ� �����Ͽ����ϴ�.", vbInformation
Err_Step:
    If Err.Number = 18 Then
      MsgBox "����ڿ� ���� ��ҵǾ����ϴ�.", vbInformation
    End If
    Application.OnRepeat "", ""
End Sub

Private Sub Fraction_Select()
 '�Ҽ�(��)�� �� ������ �ƴѼ��� ������ �� �ֵ��� �����ݴϴ�.
Dim rngTarget As Range, rngEach As Range, rngUnion As Range
Dim varEach As Variant, varRound As Variant
Dim lngRow As Long, lngColumn As Long
Dim lngI As Long, lngJ As Long
Dim blnUnion As Boolean
    On Error GoTo Err_Step
    Set rngTarget = Intersect(Selection, Selection.SpecialCells(xlCellTypeVisible), _
      Selection.SpecialCells(xlCellTypeConstants, 1))
    For Each rngEach In rngTarget.Areas
      varEach = rngEach.Value
      varRound = Application.Round(rngEach, 0)
      If IsArray(varRound) Then
         lngRow = UBound(varRound, 1)
         lngColumn = UBound(varRound, 2)
         For lngI = 1 To lngRow
           For lngJ = 1 To lngColumn
             If varEach(lngI, lngJ) <> varRound(lngI, lngJ) Then
               If blnUnion Then
                 Set rngUnion = Union(rngUnion, rngEach(lngI, lngJ))
               Else
                 Set rngUnion = rngEach(lngI, lngJ)
                 blnUnion = True
               End If
             End If
           Next lngJ
         Next lngI
      Else
        If varEach <> varRound Then
          If blnUnion Then
            Set rngUnion = Union(rngUnion, rngEach)
          Else
            Set rngUnion = rngEach
            blnUnion = True
          End If
        End If
      End If
    Next rngEach
    If rngUnion Is Nothing Then
      MsgBox "�м�(������ �ƴ� ��)�� ���� �����ϴ�!!!", vbInformation
      Exit Sub
    Else
      rngUnion.Select
    End If
Err_Step:
    If Err.Number = 1004 Then
      MsgBox "���ù����� ���ڻ���� �����ϴ�.", vbInformation
    End If
    Application.OnRepeat "", ""
End Sub

Private Sub Input_Serialnum()
 '����Ű Ctrl+Shift+N
 '���ù����� ������ �����մϴ�.
Dim rngArea As Range, rngEach As Range
Dim varTemp As Variant
Dim lngChk As Long, lngI As Long, lngJ As Long
Dim blnChk As Boolean
Dim dblJ As Double
Dim strFormat As String
    On Error GoTo Err_Step
    Application.EnableCancelKey = xlErrorHandler
    Set rngUndo = Intersect(Selection(1).EntireColumn, _
      Selection.SpecialCells(xlCellTypeVisible))
    If rngUndo.Cells.Count = 1 Then
      Set rngUndo = Intersect(Selection(1).EntireRow, _
        Selection.SpecialCells(xlCellTypeVisible))
      blnChk = True
    End If
    If rngUndo.Areas.Count = 1 Then
      varUndo = rngUndo.Formula
    End If
    If MsgBox("������ �����ұ��?", _
      vbInformation + vbOKCancel) = vbCancel Then
      Exit Sub
    End If
    If rngUndo(1) = vbNullString Then
      lngChk = 0
      dblJ = 0
    ElseIf IsNumeric(rngUndo(1)) Then
      lngChk = 0
      dblJ = rngUndo(1) - 1
    ElseIf IsNumeric(Left(rngUndo(1).Text, 1)) Then
      For lngI = 1 To Len(rngUndo(1).Text)
        If Not Mid(rngUndo(1).Text & " ", lngI + 1, 1) Like "#" Then
          lngChk = 1
          dblJ = Left(rngUndo(1).Text, lngI) - 1
          strFormat = String(lngI, "0")
          Exit For
        End If
      Next lngI
    ElseIf IsNumeric(Right(rngUndo(1).Text, 1)) Then
      For lngI = Len(rngUndo(1).Text) To 1 Step -1
        If Not Mid(" " & rngUndo(1).Text, lngI, 1) Like "#" Then
          lngChk = 2
          dblJ = Mid(rngUndo(1).Text, lngI) - 1
          strFormat = String(Len(rngUndo(1).Text) - lngI + 1, "0")
          Exit For
        End If
      Next lngI
    End If
    For Each rngArea In rngUndo.Areas
      varTemp = rngArea.Value
      If IsArray(varTemp) Then
        lngI = 0
        For Each rngEach In rngArea
          lngI = lngI + 1
          If rngEach.MergeArea(1).Address = rngEach.Address Then
            dblJ = dblJ + 1
            Select Case lngChk
            Case 0
              varTemp(IIf(blnChk, 1, lngI), IIf(blnChk, lngI, 1)) = dblJ
            Case 1
              For lngJ = 1 To Len(rngEach.Text) + 1
                If Not Mid(rngEach.Text & " ", lngJ, 1) Like "#" Then
                  varTemp(IIf(blnChk, 1, lngI), IIf(blnChk, lngI, 1)) = _
                    Format(dblJ, strFormat) & Mid(rngEach.Text, lngJ)
                  Exit For
                End If
              Next lngJ
            Case Else
              For lngJ = Len(rngEach.Text) + 1 To 1 Step -1
                If Not Mid(" " & rngEach.Text, lngJ, 1) Like "#" Then
                  varTemp(IIf(blnChk, 1, lngI), IIf(blnChk, lngI, 1)) = _
                    Left(rngEach.Text, lngJ - 1) & Format(dblJ, strFormat)
                  Exit For
                End If
              Next lngJ
            End Select
          End If
        Next rngEach
        rngArea = varTemp
      Else
        dblJ = dblJ + 1
        Select Case lngChk
        Case 0
          rngArea = dblJ
        Case 1
          For lngJ = 1 To Len(rngArea.Text) + 1
            If Not Mid(rngArea.Text & " ", lngJ, 1) Like "#" Then
              rngArea = Format(dblJ, strFormat) & Mid(rngArea.Text, lngJ)
              Exit For
            End If
          Next lngJ
        Case Else
          For lngJ = Len(rngArea.Text) + 1 To 1 Step -1
            If Not Mid(" " & rngArea.Text, lngJ, 1) Like "#" Then
              rngArea = Left(rngArea.Text, lngJ - 1) & Format(dblJ, strFormat)
              Exit For
            End If
          Next lngJ
        End Select
      End If
    Next rngArea
    If rngUndo.Areas.Count = 1 Then
      Application.OnUndo "�������� ���", "Action_Undo"
    End If
Err_Step:
    If Err.Number = 18 Then
      MsgBox "����ڿ� ���� ��ҵǾ����ϴ�.", vbInformation
    End If
    Application.OnRepeat "", ""
End Sub

Private Sub BlankCell_Input()
 '����Ű Ctrl+Shift+B
 '������ ���ʵ����ͷ� ä��ų�, ä������������ �Ʒ����� ����ϴ�.
Dim rngTmp As Range, rngTemp As Range
Dim lngRow As Long, lngCol As Long, lngI As Long, lngJ As Long
Dim blnCheck As Boolean
    On Error GoTo Err_Step
    Application.EnableCancelKey = xlErrorHandler
    Set rngUndo = Intersect(Selection, Selection.SpecialCells(xlCellTypeVisible))
    If rngUndo.Areas.Count = 1 Then
      varUndo = rngUndo.Formula
    End If
    If Selection.Cells.Count = 1 Then
      Exit Sub
    ElseIf rngUndo.Areas.Count > 1 Then
      MsgBox "���߿���(�Ǵ� ���ռ�)������ �� ����� ����� �� �����ϴ�.", vbInformation
      Exit Sub
    ElseIf IsNull(rngUndo.MergeCells) Or rngUndo.MergeCells Then
      MsgBox "���߿���(�Ǵ� ���ռ�)������ �� ����� ����� �� �����ϴ�.", vbInformation
      Exit Sub
    End If
    On Error Resume Next
    Set rngTmp = rngUndo.SpecialCells(xlCellTypeBlanks)
    If rngTmp Is Nothing Then
      On Error GoTo Err_Step
      If MsgBox("���ʵ����Ϳ� ���� ��� �Ʒ��ʵ����͸� ����ϴ�.", _
        vbInformation + vbOKCancel) = vbCancel Then
        Exit Sub
      End If
      lngRow = rngUndo.Rows.Count
      lngCol = rngUndo.Columns.Count
      For lngJ = 1 To lngCol
        blnCheck = False
        For lngI = lngRow To 2 Step -1
          If rngUndo(lngI, lngJ) = rngUndo(lngI - 1, lngJ) Then
            If Not blnCheck Then
              Set rngTemp = rngUndo(lngI, lngJ)
              blnCheck = True
            End If
          Else
            If blnCheck Then
              Range(rngUndo(lngI + 1, lngJ), rngTemp).ClearContents
              blnCheck = False
            End If
          End If
        Next lngI
        If blnCheck Then
          Range(rngUndo(lngI + 1, lngJ), rngTemp).ClearContents
          blnCheck = False
        End If
      Next lngJ
    Else
      On Error GoTo Err_Step
      If MsgBox("������ ���ʵ����ͷ� ä��ϴ�.", _
        vbInformation + vbOKCancel) = vbCancel Then
        Exit Sub
      End If
      rngTmp.NumberFormatLocal = "G/ǥ��"
      rngTmp.FormulaR1C1 = "=R[-1]C"
      rngUndo.Copy
      rngUndo.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
      Application.CutCopyMode = False
    End If
    If rngUndo.Areas.Count = 1 Then
      Application.OnUndo "����ä���(�����) ���", "Action_Undo"
    End If
Err_Step:
    If Err.Number = 18 Then
      MsgBox "����ڿ� ���� ��ҵǾ����ϴ�.", vbInformation
    End If
    Application.OnRepeat "", ""
End Sub

Private Sub Formula_MultiInput()
 '����Ű Ctrl+Alt+B
 'ù°���� ������ ���ľ��� �Ʒ���(VLOOKP�� ��� ������)���� ä���ݴϴ�.
Dim rngSelect As Range, rngFirst As Range, rngEach As Range
Dim lngI As Long, lngCount As Long, lngSplit As Long
Dim strFormula() As Variant, strFx As String
Dim varSplit As Variant
    On Error GoTo Err_Step
    Application.EnableCancelKey = xlErrorHandler
    Set rngUndo = Intersect(Selection, Selection.SpecialCells(xlCellTypeVisible))
    If rngUndo.Areas.Count = 1 Then
      varUndo = rngUndo.Formula
    End If
    If Intersect(rngUndo, rngUndo(1).EntireColumn).Cells.Count = 1 Then
      strFx = Selection(1).Formula
      If StrComp(Left(strFx, 9), "=VLOOKUP(", 1) = 0 Then
        If MsgBox("VLOOKUP�� ������ �����ʼ��� ä��ϴ�.", _
          vbInformation + vbOKCancel) = vbCancel Then
          Exit Sub
        End If
      Else
        Exit Sub
      End If
      varSplit = Split(strFx, ",")
      lngSplit = UBound(varSplit) - 1
      lngCount = rngUndo.Cells.Count
      ReDim strFormula(1 To 1, 1 To lngCount)
      strFormula(1, 1) = strFx
      For lngI = 2 To lngCount
        varSplit(lngSplit) = varSplit(lngSplit) + 1
        strFormula(1, lngI) = Join(varSplit, ",")
      Next lngI
      For Each rngEach In rngUndo.Areas
        rngEach.Value = strFormula
      Next rngEach
    Else
      If MsgBox("ù°���� ������ �Ʒ������� ä��ϴ�.", _
        vbInformation + vbOKCancel) = vbCancel Then
        Exit Sub
      End If
      Set rngSelect = Selection
      Set rngFirst = Intersect(rngSelect.SpecialCells(xlCellTypeVisible), _
        Intersect(rngSelect.EntireRow, rngSelect.EntireColumn)(1).EntireRow)
      For Each rngEach In rngFirst.Areas
        rngEach.Copy
        Intersect(rngSelect.SpecialCells(xlCellTypeVisible), _
          rngEach.EntireColumn).PasteSpecial Paste:=xlPasteFormulas
      Next rngEach
      Application.CutCopyMode = False
      rngSelect.Select
    End If
    If rngUndo.Areas.Count = 1 Then
      Application.OnUndo "����ä��� ���", "Action_Undo"
    End If
Err_Step:
    If Err.Number = 18 Then
      MsgBox "����ڿ� ���� ��ҵǾ����ϴ�.", vbInformation
    End If
    Application.OnRepeat "", ""
End Sub

Private Sub Date_Format()
 '����Ű Ctrl+Alt+Y
 '��¥�� ǥ���԰� ���ÿ� ��¥ǥ���������� �ٲ��ݴϴ�.
Dim rngEach As Range, rngTarget As Range
Dim varDate As Variant
Dim strTemp As String
Dim lngI As Long, lngJ As Long, lngK As Long
    On Error GoTo Err_Step
    Application.EnableCancelKey = xlErrorHandler
    Set rngUndo = Intersect(Selection, Selection.SpecialCells(xlCellTypeVisible))
    If rngUndo.Areas.Count = 1 Then
      varUndo = rngUndo.Formula
    End If
    If MsgBox("���ÿ����� ��¥�������� �ϰ������� �����մϴ�.", _
      vbInformation + vbOKCancel) = vbCancel Then
      Exit Sub
    End If
    On Error Resume Next
    Set rngTarget = Intersect(rngUndo, ActiveSheet.UsedRange)
    If rngTarget Is Nothing Then
      Selection.NumberFormatLocal = "yyyy-mm-dd"
      Exit Sub
    End If
    With Application
      .ScreenUpdating = False
      For Each rngEach In rngTarget.Areas
        If rngEach.Cells.Count = 1 Then
          varDate = rngEach.Resize(2, 1).Formula
          lngK = 1
        Else
          varDate = rngEach.Formula
          lngK = 0
        End If
        For lngI = 1 To UBound(varDate, 1) - lngK
          For lngJ = 1 To UBound(varDate, 2)
            strTemp = varDate(lngI, lngJ)
            If Not IsError(varDate(lngI, lngJ)) Then
              If varDate(lngI, lngJ) <> vbNullString Then
                strTemp = .Substitute(.Trim(varDate(lngI, lngJ)), ".", "-")
                If Right(strTemp, 1) = "-" Then
                  strTemp = Left(strTemp, Len(strTemp) - 1)
                End If
                If IsDate(strTemp) Then
                  varDate(lngI, lngJ) = DateValue(strTemp) * 1
                Else
                  If IsNumeric(varDate(lngI, lngJ)) Then
                    If varDate(lngI, lngJ) >= 100000 Then
                      varDate(lngI, lngJ) = DateValue(.Text(varDate(lngI, lngJ), "##00-00-00"))
                    ElseIf varDate(lngI, lngJ) > 100 And varDate(lngI, lngJ) <= 1231 Then
                      varDate(lngI, lngJ) = DateValue(.Text(varDate(lngI, lngJ), "00-00"))
                    End If
                  End If
                End If
              End If
            End If
            If varDate(lngI, lngJ) < 0 Then
              varDate(lngI, lngJ) = strTemp
            End If
          Next lngJ
        Next lngI
        rngEach.NumberFormatLocal = "yyyy-mm-dd"
        rngEach = varDate
      Next rngEach
      .ScreenUpdating = True
    End With
    If rngUndo.Areas.Count = 1 Then
      Application.OnUndo "��¥ǥ������ ���", "Action_Undo"
    End If
Err_Step:
    If Err.Number = 18 Then
      MsgBox "����ڿ� ���� ��ҵǾ����ϴ�.", vbInformation
    End If
    Application.OnRepeat "", ""
End Sub

Private Sub Function_Evaluate()
 '����Ű Ctrl+Shift+R
 '���� ������ �Լ��� �ٷ� �����մϴ�.
Dim rngNum As Range, rngEach As Range
Dim varQues(0 To 2) As String
Dim lngI As Long, lngInsu As Long
    On Error GoTo Err_Step
    Application.EnableCancelKey = xlErrorHandler
    Set rngUndo = Intersect(Selection, Selection.SpecialCells(xlCellTypeVisible))
    If rngUndo.Areas.Count = 1 Then
      varUndo = rngUndo.Formula
    End If
    Set rngNum = Intersect(rngUndo, _
      Selection.SpecialCells(xlCellTypeConstants))
    lngInsu = InStr(rngNum(1).Text, ".")
    If lngInsu > 0 Then
      lngInsu = Len(rngNum(1).Text) - InStr(rngNum(1).Text, ".")
    End If
    varQues(2) = InputBox("����Լ� �� �μ��� �Է��ϼ���.", _
      Default:="Round, " & lngInsu)
    lngI = InStr(varQues(2) & ",", ",")
    varQues(0) = Left$(varQues(2), lngI - 1)
    varQues(1) = Mid$(varQues(2), lngI + 1)
    For Each rngEach In rngNum.Areas
      If Trim$(varQues(1)) = vbNullString Then
        rngEach = Evaluate("if(" & rngEach.Address & "=" & rngEach.Address & "," & _
          Trim$(varQues(0)) & "(" & rngEach.Address & "))")
      Else
        rngEach = Evaluate("if(" & rngEach.Address & "=" & rngEach.Address & "," & _
          Trim$(varQues(0)) & "(" & rngEach.Address & "," & varQues(1) & "))")
      End If
    Next rngEach
    If rngUndo.Areas.Count = 1 Then
      Application.OnUndo "�Լ����� ���", "Action_Undo"
    End If
Err_Step:
    If Err.Number = 18 Then
      MsgBox "����ڿ� ���� ��ҵǾ����ϴ�.", vbInformation
    ElseIf Err.Number = 1004 Then
      MsgBox "������ ������ ����� �����ϴ�.", vbInformation
    End If
    Application.OnRepeat "", ""
End Sub

Private Sub Paste_AddValue()
 '����Ű Ctrl+Alt+A
 'Ȱ�������� ���Ѱ��� ǥ���մϴ�.
Dim colTemp As New Collection
Dim rngC As Range
Dim varSum As Variant
    On Error GoTo Err_Step
    Set rngUndo = Intersect(Selection, Selection.SpecialCells(xlCellTypeVisible))
    If rngUndo.Areas.Count = 1 Then
      varUndo = rngUndo.Formula
    End If
    If Application.CutCopyMode Then
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd, SkipBlanks _
        :=True, Transpose:=False
    Else
      On Error Resume Next
      varSum = 0
      For Each rngC In rngUndo
        colTemp.Add 0, CStr(rngC.Address)
        If Err > 0 Then
          Err.Clear
        Else
          varSum = varSum + rngC.Value
          If Err > 0 Then Exit Sub
        End If
      Next rngC
      Select Case MsgBox("Ȱ������ �հ踦 ǥ���ϰ� �������� ������?", _
        vbYesNoCancel + vbInformation)
      Case vbYes
        rngUndo.ClearContents
        ActiveCell = varSum
      Case vbNo
        ActiveCell = varSum
      End Select
    End If
    If rngUndo.Areas.Count = 1 Then
      Application.OnUndo "���ϱ� �����ϱ� ���", "Action_Undo"
    End If
Err_Step:
    Application.OnRepeat "", ""
End Sub

Sub Action_Undo(Optional x As Boolean)
 '���� ���
    rngUndo = varUndo
End Sub

Sub Special_Copy(x As Boolean)
 '����Ű Ctrl+Shift+C[X]
 '���߼�(���ռ�)�� ���簡 �����ϵ��� ��
Dim rngMurge As Range, rng_Area As Range
Dim lngTemp As Long
    On Error GoTo Err_Step
    If x Then
      If Selection.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
        MsgBox "����(����)���������� �߶󳻱� ����� �������� �ʽ��ϴ�.", vbInformation
        Exit Sub
      End If
      Set rngMurge = Selection
      lngMinColumn = rngMurge(1).Column
      lngMaxColumn = rngMurge(rngMurge.Cells.Count).Column
      lngMinRow = rngMurge(1).Row
      lngMaxRow = rngMurge(rngMurge.Cells.Count).Row
      strCut = rngMurge.Address
      rngMurge.Cut
    Else
      blnMode = True
      If Selection.Cells.Count > 1 Then
        Set rngMulti = Intersect(Selection.SpecialCells(xlCellTypeVisible), _
          Range(Cells(1, 1), ActiveSheet.UsedRange))
        lngMinColumn = Rows(1).Cells.Count
        lngMaxColumn = 1
        lngMinRow = Columns(1).Cells.Count
        lngMaxRow = 1
        Set rngMurge = Intersect(rngMulti(1).EntireColumn, rngMulti.EntireRow)
        For Each rng_Area In rngMurge.Areas
          lngTemp = rng_Area.Row
          If lngMinRow > lngTemp Then lngMinRow = lngTemp
          lngTemp = rng_Area.Cells(rng_Area.Cells.Count).Row
          If lngMaxRow < lngTemp Then lngMaxRow = lngTemp
        Next rng_Area
        Set rngMurge = Intersect(rngMulti(1).EntireRow, rngMulti.EntireColumn)
        For Each rng_Area In rngMurge.Areas
          lngTemp = rng_Area.Column
          If lngMinColumn > lngTemp Then lngMinColumn = lngTemp
          lngTemp = rng_Area.Cells(rng_Area.Cells.Count).Column
          If lngMaxColumn < lngTemp Then lngMaxColumn = lngTemp
        Next rng_Area
        Set rngMurge = Range(Cells(lngMinRow, lngMinColumn), _
          Cells(lngMaxRow, lngMaxColumn))
        On Error Resume Next
        rngMurge.Copy
        '���ͻ��¿��� �����ϰ��� �ϴ� ���߻��Ǵ� ������ ���
        If Err.Number = 1004 Then
          rngMurge(1).Copy
        End If
      Else
        Set rngMulti = Selection
        lngMinColumn = rngMulti.Column
        lngMaxColumn = lngMinColumn
        lngMinRow = rngMulti.Row
        lngMaxRow = lngMinRow
        Selection.Copy
      End If
    End If
Err_Step:
    Application.OnRepeat "", ""
End Sub

Sub Special_Paste(Optional x As Boolean)
 '����Ű Ctrl+Shift(Alt)+V
 '���߼�(���ռ�)�� �ٿ��ֱⰡ �����ϵ��� ��
Dim strAddress As String
Dim rngCopy As Range, rngPaste As Range
Dim rngTemp As Range, rngEach As Range
Dim lngI As Long, lngJ As Long, lngvarMaxRow As Long, lngvarMaxCol As Long
Dim varPaste() As Variant, varTemp As Variant
Dim varRow() As Long, varColumn() As Long
    On Error GoTo Err_Step
    If Application.CutCopyMode = xlCut Then
      With ThisWorkbook.Sheets(1)
        If strCut = vbNullString Then
          ActiveSheet.Paste
        Else
          If Selection(1).MergeCells Then
            strAddress = InputBox("�ٿ����� ���ۼ� �ּҸ� �Է��ϼ���", _
              Default:=Selection(1).Address(0, 0))
            If strAddress = vbNullString Then
              Exit Sub
            End If
            Set rngPaste = Range(strAddress)
          Else
            Set rngPaste = Selection(1)
          End If
          .Paste .Range(strCut)
          rngPaste.Resize(lngMaxRow - lngMinRow + 1, _
            lngMaxColumn - lngMinColumn + 1).Select
          Selection.UnMerge
          .Range(strCut).Cut
          ActiveSheet.Paste rngPaste
          strCut = vbNullString
        End If
      End With
      Application.CutCopyMode = False
    ElseIf Application.CutCopyMode = xlCopy Then
      If blnMode = True Then
        Set rngCopy = Range(rngMulti.Parent.Cells(lngMinRow, lngMinColumn), _
          rngMulti.Parent.Cells(lngMaxRow, lngMaxColumn))
        ReDim varRow(rngCopy.Row To rngCopy(rngCopy.Rows.Count, 1).Row, 1 To 1)
        ReDim varColumn(rngCopy.Column To _
          rngCopy(1, rngCopy.Columns.Count).Column, 1 To 1)
        If rngCopy.Columns(1).Cells.Count = 1 Then
          varRow(rngCopy.Row, 1) = 1
          lngvarMaxRow = 1
        Else
          Set rngTemp = rngCopy.Columns(1).SpecialCells(xlCellTypeVisible)
          lngI = 0
          For Each rngEach In rngTemp
             varRow(rngEach.Row, 1) = 1
             lngI = lngI + 1
          Next rngEach
          lngvarMaxRow = lngI
        End If
        If rngCopy.Rows(1).Cells.Count = 1 Then
          varColumn(rngCopy.Column, 1) = 1
          lngvarMaxCol = 1
        Else
          Set rngTemp = rngCopy.Rows(1).SpecialCells(xlCellTypeVisible)
          lngI = 0
          For Each rngEach In rngTemp
             varColumn(rngEach.Column, 1) = 1
             lngI = lngI + 1
          Next rngEach
          lngvarMaxCol = lngI
        End If
        Set rngPaste = Selection(1)
        Set rngTemp = Range(rngPaste, _
          Cells(Columns(1).Cells.Count, rngPaste.Column)).SpecialCells(xlCellTypeVisible)
        lngI = 0
        For Each rngEach In rngTemp
          lngI = lngI + 1
          If lngI = lngvarMaxRow Then
            lngvarMaxRow = rngEach.Row
            Exit For
          End If
        Next rngEach
        Set rngTemp = Range(rngPaste, _
          Cells(rngPaste.Row, Rows(1).Cells.Count)).SpecialCells(xlCellTypeVisible)
        lngI = 0
        For Each rngEach In rngTemp
          lngI = lngI + 1
          If lngI = lngvarMaxCol Then
            lngvarMaxCol = rngEach.Column
            Exit For
          End If
        Next rngEach
        Set rngPaste = Range(rngPaste, Cells(lngvarMaxRow, lngvarMaxCol))
        rngPaste.Select
        ReDim varPaste(rngPaste.Row To lngvarMaxRow, rngPaste.Column To lngvarMaxCol)
        If rngPaste.Columns(1).Cells.Count = 1 Then
          varRow(rngCopy.Row, 1) = rngPaste.Row
        Else
          Set rngTemp = rngPaste.Columns(1).SpecialCells(xlCellTypeVisible)
          lngI = rngCopy.Row - 1
          For Each rngEach In rngTemp
            Do
              lngI = lngI + 1
              If varRow(lngI, 1) = 1 Then
                varRow(lngI, 1) = rngEach.Row
                Exit Do
              End If
            Loop
          Next rngEach
        End If
        If rngPaste.Rows(1).Cells.Count = 1 Then
          varColumn(rngCopy.Column, 1) = rngPaste.Column
        Else
          Set rngTemp = rngPaste.Rows(1).SpecialCells(xlCellTypeVisible)
          lngI = rngCopy.Column - 1
          For Each rngEach In rngTemp
            Do
              lngI = lngI + 1
              If varColumn(lngI, 1) = 1 Then
                varColumn(lngI, 1) = rngEach.Column
                Exit Do
              End If
            Loop
          Next rngEach
        End If
        If x Then
          For Each rngEach In rngMulti
            varPaste(varRow(rngEach.Row, 1), _
              varColumn(rngEach.Column, 1)) = rngEach.Formula
          Next rngEach
        Else
          For Each rngEach In rngMulti
            varPaste(varRow(rngEach.Row, 1), _
              varColumn(rngEach.Column, 1)) = rngEach.Value
          Next rngEach
        End If
        If rngPaste.Cells.Count = 1 Then
          Set rngTemp = rngPaste
        Else
          Set rngTemp = rngPaste.SpecialCells(xlCellTypeVisible)
        End If
        For Each rngEach In rngTemp.Areas
           If IsArray(rngEach) Then
             varTemp = rngEach.Formula
             For lngI = 1 To UBound(varTemp, 1)
               For lngJ = 1 To UBound(varTemp, 2)
                 If varPaste(rngEach.Row + lngI - 1, rngEach.Column + lngJ - 1) _
                   <> vbNullString Then
                   varTemp(lngI, lngJ) = varPaste(rngEach.Row + lngI - 1, _
                     rngEach.Column + lngJ - 1)
                 End If
               Next lngJ
             Next lngI
             rngEach = varTemp
           Else
             If varPaste(rngEach.Row, rngEach.Column) <> vbNullString Then
               rngEach = varPaste(rngEach.Row, rngEach.Column)
             End If
           End If
        Next rngEach
        Application.CutCopyMode = False
      Else
        With Intersect(Selection, Selection.SpecialCells(xlCellTypeVisible))
          If x Then
            .PasteSpecial Paste:=xlPasteFormulas
          Else
            .PasteSpecial Paste:=xlPasteValues
          End If
        End With
      End If
    Else
      If x Then
        ActiveSheet.PasteSpecial Format:="HTML", NoHTMLFormatting:=True
      End If
    End If
    blnMode = False
    Application.OnRepeat "", ""
    Exit Sub
Err_Step:
    ThisWorkbook.Sheets(1).UsedRange.Clear
    Application.OnRepeat "", ""
End Sub

Private Sub Multi_Murge()
 '����Ű Ctrl+Shift+M
 '�����ս� ������ �״�� �����ϸ鼭 ����
Static strGubun As Variant
Static blnGubun As Boolean
Dim rngEacharea As Range, varEacharea As Variant
Dim strGubunja As Variant, strTemp As String, lngI As Long, lngJ As Long
    On Error GoTo Err_Step
    strTemp = ActiveCell.Text
    If Not blnGubun Then
      strGubun = "ChrW(10)"
    End If
    strGubun = Application.InputBox("�ؽ�Ʈ�� ���� �����ڸ� �Է��ϼ���." & vbCr & _
      "(�ƽ�Ű���ڴ� ChrW(10) ó�� �Է��ϼ���!!)", "������", strGubun)
    If VarType(strGubun) = vbBoolean Then
      Exit Sub
    ElseIf LCase(Left$(strGubun, 5)) = "chrw(" Then
      strGubunja = ChrW(Val(Mid$(strGubun, 6)))
    Else
      strGubunja = strGubun
    End If
    For Each rngEacharea In Selection.Areas
      strTemp = vbNullString
      varEacharea = rngEacharea.Value
      If IsArray(varEacharea) Then
        For lngI = 1 To UBound(varEacharea, 1)
          For lngJ = 1 To UBound(varEacharea, 2)
            If varEacharea(lngI, lngJ) <> vbNullString Then
              strTemp = strTemp & varEacharea(lngI, lngJ) & strGubunja
            End If
          Next lngJ
        Next lngI
      Else
        strTemp = varEacharea & strGubunja
      End If
      With rngEacharea
        .ClearContents
        .MergeCells = True
        If strTemp <> vbNullString Then
          .Value = Left$(strTemp, Len(strTemp) - Len(strGubunja))
        End If
      End With
    Next rngEacharea
    blnGubun = True
Err_Step:
    Application.OnRepeat "", ""
End Sub

Private Sub Table_Conform()
 '���̺��� ���¸� ��ȯ�Ͽ� �ݴϴ�.
Dim rngVisible As Range
    On Error GoTo Err_Step
    Set rngVisible = Intersect(Selection, ActiveSheet.UsedRange, _
      Selection.SpecialCells(xlCellTypeVisible))
    If rngVisible.Areas.Count > 1 Then
      MsgBox "���߿���(������ ����)�� �ִ� ��쿡�� ������ �Ұ��մϴ�.", vbInformation
      Exit Sub
    ElseIf rngVisible.Cells.Count = 1 Then
      Exit Sub
    End If
    Table_Change.show
Err_Step:
End Sub

Private Sub Str_Split()
 '�� ���� ���๮�ڷ� �Էµ� ������ �� ���� �и��մϴ�.
Dim rngTemp As Range
Dim lngCount As Long, lngColumns As Long
Dim lngI As Long, lngJ As Long, lngK As Long
Dim lngTemp As Long, lngRange As Long
Dim varTemp As Variant
    On Error GoTo Err_Step
    Application.EnableCancelKey = xlErrorHandler
    If Selection.Areas.Count > 1 Then
      MsgBox "���߿��� ���¿����� ������ �� ���� ����Դϴ�.", vbInformation
      Exit Sub
    Else
      If MsgBox("�ٹٲ޼��� ������ ������ �и��ϴ� �۾��� �����մϴ�.", _
        vbInformation + vbOKCancel) = vbCancel Then
        Exit Sub
      End If
    End If
    Set rngTemp = Intersect(Selection, ActiveSheet.UsedRange)
    rngTemp.MergeCells = False
    lngCount = rngTemp.Columns(1).Cells.Count
    lngColumns = rngTemp.Rows(1).Cells.Count
    With CreateObject("vbscript.regexp")
      .Global = True
      .Pattern = Chr(10)
      For lngI = lngCount To 1 Step -1
        For lngJ = 1 To lngColumns
          If InStr(rngTemp(lngI, lngJ), Chr(10)) > 0 Then
            lngTemp = Application.Max(lngTemp, Len(rngTemp(lngI, lngJ)) _
              - Len(.Replace(rngTemp(lngI, lngJ), "")))
          End If
        Next lngJ
        If lngTemp > 0 Then
          Range(rngTemp(lngI + 1, 1), rngTemp(lngI + lngTemp, 1)).EntireRow.Insert
          lngRange = lngRange + lngTemp
          lngTemp = 0
          For lngJ = 1 To lngColumns
            varTemp = Split(rngTemp(lngI, lngJ), Chr(10))
            For lngK = 1 To UBound(varTemp) + 1
              rngTemp(lngI, lngJ)(lngK, 1) = varTemp(lngK - 1)
            Next lngK
          Next lngJ
        End If
      Next lngI
    End With
    Selection.Resize(Selection.Rows.Count + lngRange, _
      Selection.Columns.Count).Select
Err_Step:
    If Err.Number = 18 Then
      MsgBox "����ڿ� ���� ��ҵǾ����ϴ�.", vbInformation
    End If
    Application.OnRepeat "", ""
End Sub

Private Sub Del_Repetition()
 '�ߺ��࿡ ���� (������) �ջ����� �����մϴ�.
Dim colData1 As New Collection, colData2 As New Collection
Dim rngTarget As Range, rngComp As Range
Dim rngEach As Range, rngUnion As Range
Dim lngRow As Long, lngCol As Long, lngTo As Long
Dim lngI As Long, lngJ As Long, lngK As Long
Dim lngTmp As Long
Dim varTemp As Variant, varTmp As Variant
Dim strTemp As String
Dim blnChk As Boolean, blnDel As Boolean, blnStay As Boolean
Dim varArr1() As Variant, varArr2() As Variant
    On Error GoTo Err_Step
    Application.EnableCancelKey = xlErrorHandler
    Set rngTarget = Intersect(Selection, Selection.SpecialCells(xlCellTypeVisible))
    lngCol = rngTarget.Areas(1).Columns.Count
    Set rngComp = Application.InputBox("�ջ��ϱ⸦ ���ϴ� ������ �����ϼ���." & vbCr & _
      "(�Ѽ��� �����Ͽ��� �ش翭�� ���� �ջ��� �մϴ�.)", Title:="�ߺ����ջ�", Type:=8)
    Set rngComp = Intersect(rngComp, rngComp(1).EntireRow)
    If Intersect(rngTarget, rngComp) Is Nothing Then
      Select Case MsgBox("�ջ� �� �ߺ����� �����ұ��?", vbYesNoCancel + vbInformation)
      Case vbNo
        blnStay = True
      Case vbCancel
        Exit Sub
      End Select
      lngTo = rngComp.Cells.Count
      ReDim varArr1(1 To lngTo)
      ReDim varArr2(1 To lngTo)
      lngI = 0
      For Each rngEach In rngComp
        lngI = lngI + 1
        varArr1(lngI) = rngEach.Column
      Next rngEach
    Else
      blnDel = True
    End If
    On Error Resume Next
    With colData1
      For lngI = 1 To rngTarget.Areas.Count
        varTemp = rngTarget.Areas(lngI).Value
        If IsArray(varTemp) Then
          lngRow = UBound(varTemp, 1)
        Else
          lngRow = 1
        End If
        For lngJ = 1 To lngRow
          strTemp = vbNullString
          If IsArray(varTemp) Then
            For lngK = 1 To lngCol
              strTemp = strTemp & Chr(1) & varTemp(lngJ, lngK)
            Next lngK
          Else
            strTemp = strTemp & Chr(1) & varTemp
          End If
          If blnDel Then
            .Add Array(1, 1), strTemp
          Else
            For lngK = 1 To lngTo
              varArr2(lngK) = Cells(rngTarget.Areas(lngI).Cells(lngJ, 1).Row, varArr1(lngK))
            Next lngK
            .Add Array(1, varArr2), strTemp
          End If
          If Err > 0 Then
            lngTmp = .Item(strTemp)(0)
            varTmp = .Item(strTemp)(1)
            If blnDel Then
              .Remove strTemp
              .Add Array(lngTmp + 1, 1), strTemp
            Else
              For lngK = 1 To lngTo
                If IsNumeric(varTmp(lngK) + varArr2(lngK)) Then
                  varArr2(lngK) = varTmp(lngK) + varArr2(lngK)
                Else
                  varArr2(lngK) = varTmp(lngK) & ";" & varArr2(lngK)
                End If
              Next lngK
              .Remove strTemp
              .Add Array(lngTmp + 1, varArr2), strTemp
            End If
            colData2.Add 1, strTemp
            Err.Clear
          End If
        Next lngJ
      Next lngI
      Err.Clear
      For lngI = 1 To rngTarget.Areas.Count
        varTemp = rngTarget.Areas(lngI).Value
        If IsArray(varTemp) Then
          lngRow = UBound(varTemp, 1)
        Else
          lngRow = 1
        End If
        For lngJ = 1 To lngRow
          strTemp = vbNullString
          If IsArray(varTemp) Then
            For lngK = 1 To lngCol
              strTemp = strTemp & Chr(1) & varTemp(lngJ, lngK)
            Next lngK
          Else
            strTemp = strTemp & Chr(1) & varTemp
          End If
          If Not blnDel Then
            For lngK = 1 To lngTo
              varArr2(lngK) = Cells(rngTarget.Areas(lngI).Cells(lngJ, 1).Row, varArr1(lngK))
            Next lngK
          End If
          colData2.Add 1, strTemp
          If Err > 0 Then
            varArr2 = .Item(strTemp)(1)
            If colData2.Item(strTemp) = 1 Then
              If Not blnDel Then
                For lngK = 1 To lngTo
                  Cells(rngTarget.Areas(lngI).Cells(lngJ, 1).Row, varArr1(lngK)) = varArr2(lngK)
                Next lngK
              End If
              colData2.Remove strTemp
              colData2.Add 2, strTemp
            Else
              If blnStay Then
                For lngK = 1 To lngTo
                  Cells(rngTarget.Areas(lngI).Cells(lngJ, 1).Row, varArr1(lngK)).ClearContents
                Next lngK
              Else
                If Not blnChk Then
                  Set rngUnion = rngTarget.Areas(lngI).Cells(lngJ, 1)
                  blnChk = True
                Else
                  Set rngUnion = Union(rngUnion, rngTarget.Areas(lngI).Cells(lngJ, 1))
                End If
              End If
            End If
            Err.Clear
          End If
        Next lngJ
      Next lngI
    End With
    If Not blnStay Then
      rngUnion.EntireRow.Delete
    End If
Err_Step:
    If Err.Number = 18 Then
      MsgBox "����ڿ� ���� ��ҵǾ����ϴ�.", vbInformation
    End If
    Application.OnRepeat "", ""
End Sub

Private Sub Insert_BlankRowCol()
 '�ش�Ǵ� ���ڸ�ŭ ��(��)�� �����մϴ�.
Dim rngTarget As Range
Dim blnChk As Boolean
Dim varCount As Variant, varQues As Variant
Dim lngCount As Long, lngI As Long
    On Error GoTo Err_Step
    Application.EnableCancelKey = xlErrorHandler
    Set rngTarget = Intersect(Selection, _
      Selection(1).EntireColumn.SpecialCells(xlCellTypeVisible))
    If rngTarget.Cells.Count = 1 Then
      Set rngTarget = Intersect(Selection, _
        Selection(1).EntireRow.SpecialCells(xlCellTypeVisible))
      blnChk = True
    End If
    If rngTarget.Areas.Count > 1 Then
      MsgBox "����(����)���������� ������ ���� �ʴ� ����Դϴ�.", vbInformation
      Exit Sub
    ElseIf rngTarget.Cells.Count = 1 Then
      Exit Sub
    End If
    varCount = rngTarget.Value
    lngCount = rngTarget.Cells.Count
    Set rngTarget = rngTarget(1)
    If Not IsNumeric(varCount(1, 1) & vbNullString) Then
      varQues = InputBox("�����ϰ��� �ϴ� " & IIf(blnChk, "��", "��") & "���� ���� �����ϼ���.", Default:=2)
      If Not IsNumeric(varQues) Then Exit Sub
      If blnChk Then
        For lngI = 1 To lngCount
          varCount(1, lngI) = varQues
        Next lngI
      Else
        For lngI = 1 To lngCount
          varCount(lngI, 1) = varQues
        Next lngI
      End If
    ElseIf MsgBox("�ش�Ǵ� ���ڸ�ŭ ��ü" & IIf(blnChk, "��", "��") & "�� �����մϴ�.", _
      vbInformation + vbOKCancel) = vbCancel Then
      Exit Sub
    End If
    If blnChk Then
      For lngI = lngCount To 1 Step -1
        If Round(varCount(1, lngI), 0) >= 2 Then
          rngTarget.Range(Cells(1, lngI + 1), _
            Cells(1, lngI + Round(varCount(1, lngI), 0) - 1)).EntireColumn.Insert
        End If
      Next lngI
    Else
      For lngI = lngCount To 1 Step -1
        If Round(varCount(lngI, 1), 0) >= 2 Then
          rngTarget.Range(Cells(lngI + 1, 1), _
            Cells(lngI + Round(varCount(lngI, 1), 0) - 1, 1)).EntireRow.Insert
        End If
      Next lngI
    End If
Err_Step:
    If Err.Number = 18 Then
      MsgBox "����ڿ� ���� ��ҵǾ����ϴ�.", vbInformation
    ElseIf Err.Number = 1004 Then
      MsgBox "������ ������ ������ ���� ������ �����Ǿ����ϴ�.", vbInformation
    End If
    Application.OnRepeat "", ""
End Sub

Private Sub TwoArea_Adjust()
 '������ �ΰ��� ������ ������ ����ϴ�.
Dim rngAarea As Range, rngBarea As Range
Dim varAarea As Variant, varBarea As Variant
Dim varARearea() As Variant, varBRearea() As Variant
Dim lngAcheck As Long, lngBcheck As Long
Dim lngArowbound As Long, lngBrowbound As Long
Dim lngAcolbound As Long, lngBcolbound As Long
Dim lngArowst As Long, lngBrowst As Long
Dim lngRemake As Long, lngI As Long, lngJ As Long
    On Error GoTo Err_Step
    Application.EnableCancelKey = xlErrorHandler
    If Selection.Areas.Count <> 2 Then
      MsgBox "������ �ٸ� �� ���� �񱳼��� �����Ͽ��� �մϴ�.", vbInformation
      Exit Sub
    End If
    
    lngAcheck = Selection.Areas(1)(1).Row
    lngBcheck = Selection.Areas(2)(1).Row
    If lngAcheck < lngBcheck Then
      lngAcheck = lngBcheck
    Else
      lngBcheck = lngAcheck
    End If
    
    Set rngAarea = Intersect(Selection.Areas(1).CurrentRegion, _
      Range(Cells(lngAcheck, 1), ActiveCell.SpecialCells(xlLastCell)))
    Set rngBarea = Intersect(Selection.Areas(2).CurrentRegion, _
      Range(Cells(lngBcheck, 1), ActiveCell.SpecialCells(xlLastCell)))
    
    If Intersect(rngAarea.EntireColumn, rngBarea.EntireColumn) Is Nothing Then
      If MsgBox("�ΰ� ������ ���� ���߱⸦ �����ϱ��?", vbInformation + vbYesNo) = vbNo Then
        Exit Sub
      End If
    Else
      MsgBox "������ �ٸ� �� ���� �񱳼��� �����Ͽ��� �մϴ�.", vbInformation
      Exit Sub
    End If
    
    With rngAarea
      varAarea = .Value
      lngArowbound = UBound(varAarea, 1)
      lngAcolbound = UBound(varAarea, 2)
      lngAcheck = Selection.Areas(1)(1).Column - .Column + 1
    End With
    With rngBarea
      varBarea = .Value
      lngBrowbound = UBound(varBarea, 1)
      lngBcolbound = UBound(varBarea, 2)
      lngBcheck = Selection.Areas(2)(1).Column - .Column + 1
    End With
    
    ReDim varARearea(1 To lngArowbound + lngBrowbound, 1 To lngAcolbound)
    ReDim varBRearea(1 To lngArowbound + lngBrowbound, 1 To lngBcolbound)
    
    lngArowst = 1
    lngBrowst = 1
    lngRemake = 1
    Do
      If varAarea(lngArowst, lngAcheck) > varBarea(lngBrowst, lngBcheck) Then
        For lngI = 1 To lngBcolbound
          varBRearea(lngRemake, lngI) = varBarea(lngBrowst, lngI)
        Next lngI
        lngBrowst = lngBrowst + 1
      ElseIf varAarea(lngArowst, lngAcheck) < varBarea(lngBrowst, lngBcheck) Then
        For lngI = 1 To lngAcolbound
          varARearea(lngRemake, lngI) = varAarea(lngArowst, lngI)
        Next lngI
        lngArowst = lngArowst + 1
      Else
        For lngI = 1 To lngAcolbound
          varARearea(lngRemake, lngI) = varAarea(lngArowst, lngI)
        Next lngI
        For lngI = 1 To lngBcolbound
          varBRearea(lngRemake, lngI) = varBarea(lngBrowst, lngI)
        Next lngI
        lngArowst = lngArowst + 1
        lngBrowst = lngBrowst + 1
      End If
      lngRemake = lngRemake + 1
    Loop While (lngArowst <= lngArowbound) And (lngBrowst <= lngBrowbound)
    If lngArowst <= lngArowbound Then
      For lngI = lngArowst To lngArowbound
        For lngJ = 1 To lngAcolbound
          varARearea(lngRemake, lngJ) = varAarea(lngI, lngJ)
        Next lngJ
        lngRemake = lngRemake + 1
      Next lngI
    End If
    If lngBrowst <= lngBrowbound Then
      For lngI = lngBrowst To lngBrowbound
        For lngJ = 1 To lngBcolbound
          varBRearea(lngRemake, lngJ) = varBarea(lngI, lngJ)
        Next lngJ
        lngRemake = lngRemake + 1
      Next lngI
    End If
    rngAarea(1).Resize(lngRemake, lngAcolbound) = varARearea
    rngBarea(1).Resize(lngRemake, lngBcolbound) = varBRearea
    
Err_Step:
    If Err.Number = 18 Then
      MsgBox "����ڿ� ���� ��ҵǾ����ϴ�.", vbInformation
    End If
    Application.OnRepeat "", ""
End Sub

Sub Select_UniqData(blnUniq As Boolean)
 '������ �Ǵ� �ߺ��� �����͸� �����Ͽ� �ݴϴ�.
Dim colData1 As New Collection, colData2 As New Collection
Dim rngTarget As Range, rngComp As Range, rngUnion As Range
Dim lngRow As Long, lngCol As Long
Dim lngI As Long, lngJ As Long, lngK As Long
Dim varTemp As Variant
Dim strTemp As String
Dim blnChk As Boolean, blnSame As Boolean
    On Error GoTo Err_Step
    Set rngTarget = Intersect(Selection, Selection.SpecialCells(xlCellTypeVisible))
    Set rngComp = Application.InputBox(IIf(blnUniq, "����", "�ߺ�") & _
      "���� �񱳴�� �������� �����ϼ���.", _
      Title:=IIf(blnUniq, "����", "�ߺ�") & "������ ����", _
      Default:=rngTarget.Areas(1).Address, Type:=8)
    Set rngComp = Intersect(rngComp, rngComp.SpecialCells(xlCellTypeVisible))
    If rngTarget.Address = rngComp.Address Then
      blnSame = True
    End If
    lngCol = rngComp.Areas(1).Columns.Count
    On Error Resume Next
    With colData1
      For lngI = 1 To rngComp.Areas.Count
        varTemp = rngComp.Areas(lngI).Value
        If IsArray(varTemp) Then
          lngRow = UBound(varTemp, 1)
        Else
          lngRow = 1
        End If
        For lngJ = 1 To lngRow
          strTemp = vbNullString
          If IsArray(varTemp) Then
            For lngK = 1 To lngCol
              strTemp = strTemp & Chr(1) & varTemp(lngJ, lngK)
            Next lngK
          Else
            strTemp = strTemp & Chr(1) & varTemp
          End If
          .Add 1, strTemp
          If blnSame Then
            If Err.Number > 0 Then
              colData2.Add 1, strTemp
              Err.Clear
            End If
          End If
        Next lngJ
      Next lngI
      Err.Clear
      For lngI = 1 To rngTarget.Areas.Count
        varTemp = rngTarget.Areas(lngI).Value
        If IsArray(varTemp) Then
          lngRow = UBound(varTemp, 1)
        Else
          lngRow = 1
        End If
        For lngJ = 1 To lngRow
          strTemp = vbNullString
          If IsArray(varTemp) Then
            For lngK = 1 To lngCol
              strTemp = strTemp & Chr(1) & varTemp(lngJ, lngK)
            Next lngK
          Else
            strTemp = strTemp & Chr(1) & varTemp
          End If
          If blnSame Then
            colData2.Add 1, strTemp
            If (Err.Number = 0) = blnUniq Then
              If Not blnChk Then
                Set rngUnion = rngTarget.Areas(lngI) _
                  .Range(Cells(lngJ, 1), Cells(lngJ, lngK - 1))
                blnChk = True
              Else
                Set rngUnion = Union(rngUnion, _
                  rngTarget.Areas(lngI).Range(Cells(lngJ, 1), Cells(lngJ, lngK - 1)))
              End If
            End If
          Else
            .Add 1, strTemp
            If (Err.Number = 0) = blnUniq Then
              If Not blnChk Then
                Set rngUnion = rngTarget.Areas(lngI) _
                  .Range(Cells(lngJ, 1), Cells(lngJ, lngK - 1))
                blnChk = True
              Else
                Set rngUnion = Union(rngUnion, _
                  rngTarget.Areas(lngI).Range(Cells(lngJ, 1), Cells(lngJ, lngK - 1)))
              End If
            End If
            If Err.Number = 0 Then
              .Remove strTemp
            End If
          End If
          Err.Clear
        Next lngJ
      Next lngI
    End With
    rngUnion.Select
Err_Step:
    Application.OnRepeat "", ""
End Sub

Private Sub RowHightColWidth()
 '����Ű Ctrl+Alt+H
 '��(��)������ �ϰ������� �����ϴ�.
Dim blnChk As Boolean
Dim rngTarget As Range, rng_Each As Range
Dim lngQues As Variant
    On Error GoTo Err_Step
    Application.EnableCancelKey = xlErrorHandler
    Set rngTarget = Intersect(ActiveSheet.UsedRange, Selection(1).EntireColumn, Selection)
    If rngTarget.Cells.Count = 1 Then
      blnChk = True
      Set rngTarget = Intersect(ActiveSheet.UsedRange, Selection(1).EntireRow, Selection)
    End If
    lngQues = InputBox("���ÿ����� " & IIf(blnChk, "���ʺ�", "�����") & _
      "�� �ϰ������� �����ϴ�.", Default:=5)
    If lngQues = vbNullString Then Exit Sub
    For Each rng_Each In rngTarget
      If blnChk Then
        rng_Each.ColumnWidth = rng_Each.ColumnWidth + lngQues
      Else
        rng_Each.RowHeight = rng_Each.RowHeight + lngQues
      End If
    Next rng_Each
Err_Step:
    Application.OnRepeat "", ""
End Sub

Private Sub DataBase_Split()
 '�ϳ��� ��Ʈ�� �뷮���� �ִ� �����͸� �����Ͽ� �������� ���Ϸ� �����Ͽ� �ݴϴ�.
Dim colData As New Collection
Dim stTarget As Worksheet
Dim strFolder As String
Dim lngI As Long, lngCount As Long, lngChk As Long
Dim rngEach As Range, rngTarget As Range
Dim rngUnion As Range, rngTemp As Range
Dim varTarget As Variant, varWbname As Variant
Dim strWbname As String, strPassword As String
    On Error GoTo Err_Step
    Application.EnableCancelKey = xlErrorHandler
    Set stTarget = ActiveSheet
    Set rngTarget = Intersect(stTarget.UsedRange, Selection, Selection(1).EntireColumn)
    varTarget = rngTarget.Value
    strFolder = ActiveWorkbook.path
    If strFolder = "" Then
      strFolder = Application.DefaultFilePath
    End If
    strWbname = InputBox("�����ҷ��� ���ϸ�[;��й�ȣ]�� �����ϼ���.", Default:="����")
    If strWbname = vbNullString Then Exit Sub
    varWbname = Split(strWbname, ";")
    If UBound(varWbname) = 1 Then
      strPassword = varWbname(1)
    End If
    On Error Resume Next
    With colData
      For lngI = 1 To UBound(varTarget)
        .Add varTarget(lngI, 1), CStr(varTarget(lngI, 1))
        If Err > 0 Then
          Err.Clear
        End If
      Next lngI
      lngCount = .Count
      For lngI = 1 To lngCount
        Set rngUnion = Nothing
        lngChk = 0
        For Each rngEach In rngTarget
          If rngEach = .Item(lngI) Then
            If lngChk = 0 Then
              Set rngUnion = rngEach
              lngChk = lngChk + 1
            Else
              Set rngUnion = Union(rngUnion, rngEach)
              lngChk = lngChk + 1
            End If
          End If
        Next rngEach
        If lngChk > 0 Then
          stTarget.Copy
          Set rngTemp = Selection
          rngUnion.EntireRow.Copy Selection(1).EntireRow(1)
          Intersect(rngTemp, rngTemp.Offset(lngChk, 0)).EntireRow.Delete
          ActiveCell.Select
          ActiveWorkbook.SaveAs FileName:=strFolder & "\" & Format(Date, "yymmdd") & "_" & _
            .Item(lngI) & varWbname(0), _
            Password:=IIf(UBound(varWbname) = 1, strPassword, ""), _
            WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
          ActiveWindow.Close
        End If
      Next lngI
    End With
    MsgBox strFolder & " ������" & vbCr & lngCount & _
      "���� ������ ���������� �����߽��ϴ�.", vbInformation
    Exit Sub
Err_Step:
    If Err.Number = 18 Then
      MsgBox "����ڿ� ���� ��ҵǾ����ϴ�.", vbInformation
    End If
End Sub

Private Sub Wb_Combine()
 '���� ���չ����� �����մϴ�.
Dim strChk As String
    On Error GoTo Err_Step
    strChk = ActiveCell.Value
    Comb_Wb.show
Err_Step:
End Sub

Sub Make_RecoveryNewBook(Optional blnMsg As Boolean)
 '�������� ���������� �����Ͽ� �ݴϴ�.
Dim wbActive As Workbook
Dim lngI As Long
    On Error GoTo Err_Step
    Application.EnableCancelKey = xlErrorHandler
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set wbActive = ActiveWorkbook
    With wbActive
      If blnMsg = False Then
        If MsgBox("������ ������ �ݰ�, �������� ��������(��ũ�� ����)��" _
          & vbCr & "�����Ͻðڽ��ϱ�?", _
          vbInformation + vbYesNo) = vbNo Then Exit Sub
      End If
      CommandBars("Exit Design Mode").Controls(1).Reset
      For lngI = 1 To .Worksheets.Count
        .Worksheets(lngI).Unprotect
        If Worksheets(lngI).Visible = xlSheetVeryHidden Then
          Worksheets(lngI).Visible = xlSheetHidden
        End If
        If .Worksheets(lngI).FilterMode Then
          .Worksheets(lngI).ShowAllData
        End If
      Next lngI
    End With
    Application.OnTime Now + 0.00001, "Xml_Tempmake"
    wbActive.Worksheets.Copy
    '����2007������ ��Ȱ�ϰ� ������� �ʾ� �Ҽ� ���� ������ �и���
    Exit Sub
Err_Step:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "������ ������ �־� ���󺹱����� ���߽��ϴ�.", vbCritical
End Sub

Private Sub Xml_Tempmake()
Dim wbOld As Workbook, wbNew As Workbook
Dim strPath As String, strOldPath As String, strName As String
Dim naTemp As Name
Dim styTemp As Style
Dim rngTemp As Range
Dim lngI As Long, lngJ As Long
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    With ActiveWorkbook
      For lngI = 1 To .Worksheets.Count
        .Worksheets(lngI).Cells.Clear
        Set rngTemp = .Worksheets(lngI).UsedRange
      Next lngI
      Set rngTemp = Nothing
      .Worksheets.Copy
      .Close SaveChanges:=False
    End With
    strPath = Application.DefaultFilePath
    With ActiveWorkbook
      .SaveAs FileName:=strPath & "\TempRecovery.tmp", _
        FileFormat:=xlXMLSpreadsheet, CreateBackup:=False
      .Close SaveChanges:=False
    End With
    Set wbOld = ActiveWorkbook
    strName = wbOld.Name
    strOldPath = wbOld.path
    If strOldPath = "" Then strOldPath = strPath
    TextStrimConvert strPath & "\TempRecovery.tmp"
    Set wbNew = Workbooks.Open(FileName:=strPath & "\TempRecovery.tmp")
    With wbNew
      For lngI = 1 To .Sheets.Count
        wbOld.Worksheets(.Sheets(lngI).Name).Cells.Copy .Worksheets(lngI).Cells(1, 1)
        With wbNew.Worksheets(lngI)
          For lngJ = .DrawingObjects.Count To 1 Step -1
            With wbNew.Worksheets(lngI).DrawingObjects(lngJ)
              If .Height + .Width = 0 Then
                .Delete
              Else
                .OnAction = ""
              End If
            End With
          Next lngJ
        End With
      Next lngI
      .ChangeLink Name:=strName, NewName:=.Name, Type:=xlExcelLinks
      .Worksheets.Copy
      Application.EnableEvents = True
      .Close SaveChanges:=False
      Kill strPath & "\TempRecovery.tmp"
    End With
    For Each styTemp In ActiveWorkbook.Styles
      If Not styTemp.BuiltIn Then
        styTemp.Delete
      End If
    Next styTemp
    For Each naTemp In ActiveWorkbook.Names
      If Not naTemp.Visible = False Then
        naTemp.Visible = True
      End If
      naTemp.Delete
    Next naTemp
    ChDir strOldPath
    wbOld.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox "���������� �������� ���������� �����Ͽ����ϴ�. " _
      & vbCr & "�ٸ��̸�(����Ⱑ��)���� �����Ͻñ� �ٶ��ϴ�.", vbInformation
    Application.OnRepeat "", ""
End Sub

Private Sub TextStrimConvert(strPathName As String)
Dim objStream As Object
Dim varSplit As Variant
Dim strString As String
Dim lngI As Long, lngSp As Long
    On Error GoTo Err_Step
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "UTF-8"
    objStream.LoadFromFile strPathName
    strString = objStream.ReadText
    objStream.Close
    varSplit = Split(strString, "</ExcelWorkbook>")
    strString = ""
    For lngI = 0 To UBound(varSplit) - 1
      lngSp = InStr(varSplit(lngI), "<ExcelWorkbook")
      strString = strString & Left(varSplit(lngI), lngSp - 1)
    Next lngI
    strString = strString & varSplit(lngI)
    varSplit = Split(strString, "</Names>")
    strString = ""
    For lngI = 0 To UBound(varSplit) - 1
      lngSp = InStr(varSplit(lngI), "<Names")
      strString = strString & Left(varSplit(lngI), lngSp - 1)
    Next lngI
    strString = strString & varSplit(lngI)
    objStream.Open
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "UTF-8"
    objStream.WriteText strString
    objStream.SaveToFile strPathName, 2
    Set objStream = Nothing
Err_Step:
End Sub

Private Sub StrAndNum_Sort()
 '����Ű Ctrl+Alt+S
 '���ڼ��ڸ� �����մϴ�.
Dim rngTemp As Range
    On Error GoTo Err_Step
    Set rngTemp = Intersect(ActiveSheet.UsedRange, _
      Selection.SpecialCells(xlCellTypeVisible))
    If rngTemp.Columns(1).Cells.Count = 1 Then
      MsgBox "������ ������ ������ �� �����ϼ���.", vbInformation
      Exit Sub
    ElseIf rngTemp.Areas.Count > 1 Then
      MsgBox "���߹��������� ������ �� ���� ����Դϴ�.", vbInformation
      Exit Sub
    End If
    Numtext_Sort.show
Err_Step:
End Sub

Private Sub Randomize_Sort()
 '����Ű Ctrl+Alt+R
 '������ ������ �����մϴ�.
Dim rngTemp As Range, rngTmp As Range
Dim strAddress As String
Dim varInputdata As Variant, varTemp1() As Variant, varTemp2() As Variant
Dim blnFormula() As Boolean
Dim lngChk As Long, lngRow As Long, lngCol As Long
Dim lngI As Long, lngJ As Long, lngK As Long, lngL As Long
    On Error GoTo Err_Step
    Set rngTemp = Intersect(ActiveSheet.UsedRange, _
      Selection.SpecialCells(xlCellTypeVisible))
    If rngTemp.Areas.Count > 1 Then
      MsgBox "���߿���(����) ���¿����� ������ �� ���� ����Դϴ�.", vbInformation
      Exit Sub
    Else
      lngChk = MsgBox("������ ������ �����մϴ�.(������� ����, " & _
        "������������ �ƴϿ��� Ŭ��)", vbInformation + vbYesNoCancel)
      If lngChk = vbCancel Then
        Exit Sub
      End If
    End If
    If lngChk = vbYes Then
      Set rngTmp = Union(rngTemp, rngTemp.Offset(0, 1))
      varInputdata = rngTmp.FormulaR1C1
    Else
      varInputdata = rngTemp.FormulaR1C1
    End If
    Application.ScreenUpdating = False
    On Error Resume Next
    rngTemp.Sort Key1:=rngTemp(1), _
      Header:=xlNo, Orientation:=xlTopToBottom
    If Err > 0 Then
      MsgBox "�迭�����̳� ���ռ��� �ִ� ��� ������ �� �����ϴ�.", vbInformation
      Exit Sub
    End If
    On Error GoTo Err_Step
    If lngChk = vbYes Then
      strAddress = rngTmp(1, rngTmp.Columns.Count).Address
      lngRow = UBound(varInputdata, 1)
      lngCol = UBound(varInputdata, 2)
      For lngI = 1 To lngRow
        varInputdata(lngI, lngCol) = Rnd
      Next lngI
      With ThisWorkbook.Sheets(1)
        Application.DisplayAlerts = False
        .Range(rngTmp.Address).Formula = varInputdata
        .Range(rngTmp.Address).Sort Key1:=.Range(strAddress), _
          Header:=xlNo, Orientation:=xlTopToBottom
        varInputdata = .Range(rngTmp.Address).FormulaR1C1
        Application.DisplayAlerts = True
        rngTemp = varInputdata
        .UsedRange.Clear
      End With
    Else
      lngRow = UBound(varInputdata, 1)
      lngCol = UBound(varInputdata, 2)
      lngL = lngRow * lngCol
      ReDim varTemp1(1 To lngL)
      ReDim varTemp2(1 To lngL)
      For lngI = 1 To lngRow
        For lngJ = 1 To lngCol
          lngK = lngK + 1
          varTemp1(lngK) = varInputdata(lngI, lngJ)
        Next lngJ
      Next lngI
      For lngI = lngL To 1 Step -1
        lngJ = Int(Rnd * lngI) + 1
        varTemp2(lngI) = varTemp1(lngJ)
        varTemp1(lngJ) = varTemp1(lngI)
      Next lngI
      lngK = 0
      For lngI = 1 To lngRow
        For lngJ = 1 To lngCol
          lngK = lngK + 1
          varInputdata(lngI, lngJ) = varTemp2(lngK)
        Next lngJ
      Next lngI
      rngTemp = varInputdata
    End If
    Application.ScreenUpdating = True
    Application.OnRepeat "", ""
    Exit Sub
Err_Step:
    ThisWorkbook.Sheets(1).UsedRange.Clear
    Application.OnRepeat "", ""
End Sub

Private Sub Filter_Reverse()
 '����Ű Alt+L
 '���͸��� ���� �����, �ݴ�� ���������� ���̵��� �մϴ�.
Dim strFilter As String
Dim rngVisible As Range, rngHide As Range, lngCount As Long
    On Error GoTo Err_Step
    Application.EnableCancelKey = xlErrorHandler
    If ActiveSheet.FilterMode Then
      On Error Resume Next
      Set rngVisible = ActiveCell.EntireColumn.SpecialCells(xlCellTypeVisible)
      If Err > 0 Then
        MsgBox "���߿����� �ʹ� Ů�ϴ�.(8,192���� �ʰ�)", vbInformation
        Exit Sub
      Else
        On Error GoTo Err_Step
      End If
      Set rngHide = Range(rngVisible.Areas(1)(rngVisible.Areas(1).Cells.Count _
        + 1), rngVisible.Areas(2)(0, 1))
      If rngVisible.Areas.Count > 2 Then
        For lngCount = 2 To rngVisible.Areas.Count - 1
          Set rngHide = Union(rngHide, Range(rngVisible.Areas(lngCount) _
            (rngVisible.Areas(lngCount). _
            Cells.Count + 1), rngVisible.Areas(lngCount + 1)(0, 1)))
        Next lngCount
      End If
      strFilter = "'" & ActiveSheet.Name & "'!_FilterDatabase"
      Set rngHide = Intersect(rngHide.EntireRow, Names(strFilter).RefersToRange)
      Range(Names(strFilter).RefersToRange(2, 1), _
        Names(strFilter).RefersToRange(Names(strFilter).RefersToRange. _
          Cells.Count)).EntireRow.Hidden = True
      rngHide.EntireRow.Hidden = False
    Else
      MsgBox "���͹����� �����Ͱ� ���͵� ���¿����� �۵��˴ϴ�.", vbInformation
    End If
Err_Step:
    If Err.Number = 18 Then
      MsgBox "����ڿ� ���� ��ҵǾ����ϴ�.", vbInformation
    End If
    Application.OnRepeat "", ""
End Sub

Private Sub Color_Filter()
 '����Ű Ctrl+Alt+L
 '���󺰷� ���͸��� �մϴ�.
Dim ftTemp As AutoFilter
Dim rngFilter As Range, rngTemp As Range
Dim rngEach As Range, rngUnion As Range
Dim blnChk As Boolean
Dim lngColor As Long, lngI As Long
    On Error GoTo Err_Step
    Application.EnableCancelKey = xlErrorHandler
    Set ftTemp = ActiveSheet.AutoFilter
    If ftTemp Is Nothing Then
      MsgBox "�� ����� �ڵ����� ���¿��� ������ �� �ֽ��ϴ�.", vbInformation
      Exit Sub
    End If
    Set rngFilter = ftTemp.Range
    Set rngFilter = Range(rngFilter(2, 1), rngFilter(rngFilter.Rows.Count, _
      rngFilter.Columns.Count))
    Set rngTemp = Intersect(rngFilter, ActiveCell)
    If Not rngTemp Is Nothing Then
      Select Case MsgBox("��������� �����ҷ��� ""��""��," & vbCr & _
        "�۲û����� �����ҷ��� ""�ƴϿ�""�� Ŭ���ϼ���.", vbYesNoCancel + vbInformation)
      Case vbYes
        blnChk = True
        lngColor = ActiveCell.Interior.ColorIndex
      Case vbNo
        lngColor = ActiveCell.Font.ColorIndex
      Case Else
        Exit Sub
      End Select
      Set rngTemp = rngFilter.Columns(ActiveCell.Column - _
        rngFilter(1).Column + 1)
      If ActiveSheet.FilterMode Then
        Set rngTemp = rngTemp.SpecialCells(xlCellTypeVisible)
        For Each rngEach In rngTemp
          If blnChk Then
            If rngEach.Interior.ColorIndex <> lngColor Then
              If rngUnion Is Nothing Then
                Set rngUnion = rngEach
              Else
                Set rngUnion = Union(rngUnion, rngEach)
              End If
            End If
          Else
            If rngEach.Font.ColorIndex <> lngColor Then
              If rngUnion Is Nothing Then
                Set rngUnion = rngEach
              Else
                Set rngUnion = Union(rngUnion, rngEach)
              End If
            End If
          End If
        Next rngEach
        rngUnion.EntireRow.Hidden = True
      Else
        Set rngUnion = ActiveCell
        Selection.AutoFilter Field:=ActiveCell.Column - _
          rngFilter(1).Column + 1, Criteria1:=Chr(9)
        For lngI = 1 To rngTemp.Rows.Count
          If blnChk Then
            If rngTemp.Cells(lngI).Interior.ColorIndex = lngColor Then
              Set rngUnion = Union(rngUnion, rngTemp.Cells(lngI))
            End If
          Else
            If rngTemp.Cells(lngI).Font.ColorIndex = lngColor Then
              Set rngUnion = Union(rngUnion, rngTemp.Cells(lngI))
            End If
          End If
        Next lngI
        rngUnion.EntireRow.Hidden = False
      End If
    End If
Err_Step:
    If Err.Number = 18 Then
      MsgBox "����ڿ� ���� ��ҵǾ����ϴ�.", vbInformation
    End If
    Application.OnRepeat "", ""
End Sub

Private Sub Print_SelectPage()
 '����Ű Ctrl+Shift+P
 '���õ� �������� �μ��մϴ�.(���������� �μ�)
Dim shtActive As Worksheet, rngUsed As Range, rngSelection As Range
Dim lngStartpage As Long, lngLastpage As Long
Dim lngR As Long, lngC As Long, lngHeight As Long, lngWidth As Long
Dim lngOutSide As Long, lngDownstart As Long, lngDownstartend As Long
Dim lngAcrossstart As Long, lngAcrossstartend As Long, lngQues As Long
Dim strAddress As String, strTemp As String
    On Error GoTo Err_Step
    Set shtActive = ActiveSheet
    With shtActive.UsedRange
      If .Cells.Count = 1 Then
        If IsEmpty(.Cells(1)) Then GoTo Err_Step
      End If
    End With
    '�μ� ������ ������ ���� �׷��� ���� ��츦 ����
    strTemp = shtActive.PageSetup.PrintArea
    If strTemp = vbNullString Then
      Set rngUsed = Range(Cells(1, 1), shtActive.UsedRange)
    Else
      Set rngUsed = Range(strTemp)
    End If
    'Ȱ������ ��ġ�� �μ� ������ ��ġ�� ������ Ȯ��
    Set rngSelection = Intersect(Selection.Areas(1), rngUsed)
    If rngSelection Is Nothing Then
      MsgBox "���ü��� �μ⿵�� �ȿ� ���� �ʽ��ϴ�!", vbInformation
      Exit Sub
    End If

    With rngUsed
      '�� �˻�
      lngR = .Row
      lngHeight = Get_RowBreaks(lngR + .Rows.Count - 1) + 1
      lngDownstart = Get_RowBreaks(rngSelection(1).Row) + 1
      lngDownstartend = Get_RowBreaks(rngSelection(rngSelection.Cells.Count). _
        Row) + 1
      If lngR > 1 Then
        lngOutSide = Get_RowBreaks(lngR)
        lngHeight = lngHeight - lngOutSide
        lngDownstart = lngDownstart - lngOutSide
        lngDownstartend = lngDownstartend - lngOutSide
      End If
      '���˻�
      lngC = .Column
      lngWidth = Get_ColBreaks(lngC + .Columns.Count - 1) + 1
      lngAcrossstart = Get_ColBreaks(rngSelection(1).Column) + 1
      lngAcrossstartend = Get_ColBreaks(rngSelection(rngSelection.Cells.Count) _
        .Column) + 1
      If lngC > 1 Then
        lngOutSide = Get_ColBreaks(lngC)
        lngWidth = lngWidth - lngOutSide
        lngAcrossstart = lngAcrossstart - lngOutSide
        lngAcrossstartend = lngAcrossstartend - lngOutSide
      End If
      '�μ� ����
      If shtActive.PageSetup.Order = xlDownThenOver Then
        lngStartpage = lngHeight * (lngAcrossstart - 1) + lngDownstart
        lngLastpage = lngHeight * (lngAcrossstartend - 1) + lngDownstartend
      Else
        lngStartpage = lngWidth * (lngDownstart - 1) + lngAcrossstart
        lngLastpage = lngWidth * (lngDownstartend - 1) + lngAcrossstartend
      End If
    End With
    If lngStartpage > 0 Then
      lngQues = MsgBox("���ü��� " & lngStartpage & " �� ~ " & _
        lngLastpage & "���� �μ��ұ��?", vbYesNo + vbInformation)
      If lngQues = vbYes Then
        shtActive.PrintOut From:=lngStartpage, To:=lngLastpage
      End If
    End If
Err_Step:
    Application.OnRepeat "", ""
End Sub

Private Function Get_ColBreaks(ColNum As Long) As Long
Dim strTemp As String
    '���� �Ǵ� �ڵ� �� ������ �ٷ� �Ʒ� ��鿡 �ش��ϴ� �� ��ȣ �迭�� �����ش�.
    On Error Resume Next
    strTemp = "MATCH(" & ColNum & ",GET.DOCUMENT(65),1)"
      Get_ColBreaks = ExecuteExcel4Macro(strTemp)
End Function

Private Function Get_RowBreaks(RowNum As Long) As Long
Dim strTemp As String
    '���� �Ǵ� �ڵ� �� ������ �ٷ� �Ʒ� ��鿡 �ش��ϴ� �� ��ȣ �迭�� �����ش�.
    On Error Resume Next
    strTemp = "MATCH(" & RowNum & ",GET.DOCUMENT(64),1)"
        Get_RowBreaks = ExecuteExcel4Macro(strTemp)
End Function

Private Sub Help_Text()
 'My_Addin2015�� ���Ͽ�
    Help_Msg.show vbModeless
End Sub

Private Sub Safe_Save()
 '����Ű Ctrl+S
 '����Ű�� ���� ��� �ٽ��ѹ� ���忩�θ� ������ �����ν� �ߴ�Ǽ� ����
Dim wkbAct As Workbook
    On Error GoTo Err_Step
    Set wkbAct = ActiveWorkbook
    If wkbAct.path = "" Then '�űԹ����� ���
      Application.Dialogs(xlDialogSaveAs).show
    Else
      If wkbAct.Saved Then '�̹� ����� ������ ���
        wkbAct.Save
      Else
        If MsgBox("'" & wkbAct.Name & "'�� ���� ������ �����Ͻðڽ��ϱ�?", _
          vbYesNo + vbInformation) = vbYes Then
          Application.EnableEvents = False
            wkbAct.Save
          Application.EnableEvents = True
        End If
      End If
    End If
Err_Step:
    Application.OnRepeat "", ""
End Sub

Sub Quick_Sort(varData As Variant, ByVal lngKey As Long, ByVal blnHigh As Boolean, _
  ByVal lngCol As Long, ByVal lngFirst As Long, ByVal lngLast As Long)
Dim lngLow As Long, lngHigh As Long, lngI As Long
Dim MidValue As Variant
    lngLow = lngFirst
    lngHigh = lngLast
    MidValue = varData((lngFirst + lngLast) \ 2, lngKey)
    Do
      If blnHigh Then
        While varData(lngLow, lngKey) < MidValue
          lngLow = lngLow + 1
        Wend
        While varData(lngHigh, lngKey) > MidValue
          lngHigh = lngHigh - 1
        Wend
      Else
        While varData(lngLow, lngKey) > MidValue
          lngLow = lngLow + 1
        Wend
        While varData(lngHigh, lngKey) < MidValue
          lngHigh = lngHigh - 1
        Wend
      End If
      If lngLow <= lngHigh Then
        For lngI = 1 To lngCol
          Swap_Data varData(lngLow, lngI), varData(lngHigh, lngI)
        Next lngI
        lngLow = lngLow + 1
        lngHigh = lngHigh - 1
      End If
    Loop While lngLow <= lngHigh
    If lngFirst < lngHigh Then
      Quick_Sort varData, lngKey, blnHigh, lngCol, lngFirst, lngHigh
    End If
    If lngLow < lngLast Then
      Quick_Sort varData, lngKey, blnHigh, lngCol, lngLow, lngLast
    End If
End Sub

Private Sub Swap_Data(ByRef varA As Variant, ByRef varB As Variant)
Dim varT As Variant
    varT = varA
    varA = varB
    varB = varT
End Sub
