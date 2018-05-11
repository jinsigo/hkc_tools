Attribute VB_Name = "���Ḯ��Ʈ"

'======================================================================
' ���Ḯ��Ʈ ���� �Լ�
'======================================================================
Sub rmSetNames_old()
Attribute rmSetNames_old.VB_Description = "���Ḯ��Ʈ �̸�����"
Attribute rmSetNames_old.VB_ProcData.VB_Invoke_Func = " \n14"
'
' mk_����LIST ��ũ�� 2013.9.30 / ������
' ���Ḯ��Ʈ �̸�����
'
'    Windows("D:\1.���Ἲ��\����LIST.xls").Activate
'    Windows("D:\1.���Ἲ��\��ɼ��Է�.xlsx").Activate
'    ActiveWorkbook.Names.Item("����LIST").Delete
'    ActiveWorkbook.Names.Item("��ɼ�").Delete
    Dim fn(3) As String
    Dim dn(3) As String
    Dim nn(3) As String
    Dim wrk As Workbook
        
    Set wrk = ActiveWorkbook
    
    dn(1) = "D:\RND\���Ἲ��"
    dn(2) = "D:\RND\���Ἲ��"
    fn(1) = "����LIST.xls"
    fn(2) = "����LIST.��ɼ��Է�.xlsx"
    nn(1) = "����LIST"
    nn(2) = "��ɼ�"
    
    ChDir (dn(1))
    If isOpenWrk(fn(1)) = 0 Then Workbooks.Open FileName:=dn(1) & "\" & fn(1)
    If isOpenWrk(fn(2)) = 0 Then Workbooks.Open FileName:=dn(2) & "\" & fn(2)

    With wrk
        .Activate
        If isOpenName(nn(1)) > 0 Then .Names(nn(1)).Delete
        If isOpenName(nn(2)) > 0 Then .Names(nn(2)).Delete

        .Names.Add Name:=nn(1), RefersToR1C1:="='" & fn(1) & "'!Database"
        .Names.Add Name:=nn(2), RefersToR1C1:="='" & fn(2) & "'!��ɼ�"
        .Names(nn(1)).Comment = nn(1)
        .Names(nn(2)).Comment = nn(1)
        MsgBox wrk.Name
    End With

End Sub




Sub rmPutData()
'
' �����ּ� �ӽ����� ��ũ�� 2015.2.4 / ������
'

' OR Range(Range("A1:K1"), Range("A1:K1").End(xlDown)).Select


    Selection.CurrentRegion.Select
    Set sc = Selection
    MsgBox sc.Address()
    Workbooks("personal.xlsb").Sheets("sheet1").Range("A1") = ActiveWorkbook.path  'path
    Workbooks("personal.xlsb").Sheets("sheet1").Range("B1") = ActiveWorkbook.Name  'File name
    Workbooks("personal.xlsb").Sheets("sheet1").Range("C1") = ActiveSheet.Name  'Sheet name
    Workbooks("personal.xlsb").Sheets("sheet1").Range("D1") = sc.Address()  'Selection address

End Sub

Sub rmGetData()
'
' �����ּ� �������� ��ũ�� 2015.2.4 / ������
'
Dim wb As Workbook
    tr = Selection
    Application.ScreenUpdating = False ' turn off the screen updating
    
    With Workbooks("HKC.xlsm").Sheets("P")
        pn = .Range("A1").Value  'path name
        fn = .Range("B1").Value  'File name
        sn = .Range("C1").Value  'Sheet name
        rn = .Range("D1").Formula  'Selection address
    End With
    
    fName = pn & "\" & fn
    Set wb = Workbooks.Open(fName, True, True) ' open the source workbook, read only
    wb.Worksheets(sn).Range(rn).Copy
    ThisWorkbook.Activate
    Range(tr).Paste
        wb.Close False ' close the source workbook without saving any changes
    Set wb = Nothing ' free memory
    Application.ScreenUpdating = True ' turn on the screen updating
      
End Sub

Sub rmGetDataFromClosedWorkbook()
Dim wb As Workbook
    Application.ScreenUpdating = False ' turn off the screen updating
    Set wb = Workbooks.Open("C:\Foldername\Filename.xls", True, True)
    ' open the source workbook, read only
    With ThisWorkbook.Worksheets("TargetSheetName")
        ' read data from the source workbook
        .Range("A10").Formula = wb.Worksheets("SourceSheetName").Range("A10").Formula
        .Range("A11").Formula = wb.Worksheets("SourceSheetName").Range("A20").Formula
        .Range("A12").Formula = wb.Worksheets("SourceSheetName").Range("A30").Formula
        .Range("A13").Formula = wb.Worksheets("SourceSheetName").Range("A40").Formula
    End With
    wb.Close False ' close the source workbook without saving any changes
    Set wb = Nothing ' free memory
    Application.ScreenUpdating = True ' turn on the screen updating
End Sub

Sub rmFindText()
    Dim i As Integer
    'Search criteria
    With Application.FileSearch
        .LookIn = "c:\my documents\logs" 'path to look in
        .FileType = msoFileTypeAllFiles
        .SearchSubFolders = False
        .TextOrProperty = "*Find*" 'Word to find in this line
        .Execute 'start search
     
        'This loop will bring up a message box with the name of
        'each file that meets the search criteria
        For i = 1 To .FoundFiles.Count
            MsgBox .FoundFiles(i)
        Next i
    End With
 
End Sub

Function rmJU(cd As String, opt As Integer)
' ju�ڵ� �ܰ� �������� / ������
'2013.6.13
    fn = "D:\�˰�����\HKC\LST\"
    cd = Replace(cd, "-", "")
    Set rs = Range("����ܰ�")
    If opt = 0 Then opt = 4
    jcost = WorksheetFunction.VLookup(cd, rs, opt, 0)
    If opt = 4 Then jcost = jcost * 1000
    
End Function

Sub rmRead_BOM()
' BOM���� JU�ڵ� ��������
' 2015.12.21 ������ ;
Dim rs, ro, rJUcode, rComp As Range
Dim np As Integer
    
    Set rJUcode = Selection 'Range(rs.Cells(5, 2), rs.End(xlDown)) 'BOM ���� ����
       
    np = InputBox("�ҿ䷮ �����ġ�� �Է��� �ּ���(�⺻���� '5' �Դϴ�).", "ó�� ��������", 5) '�ҿ䷮ �����ġ ��������
    Set rComp = rJUcode.Offset(0, np)
          
    Set rs = Union(rJUcode, rComp)
    
    'Application.DisplayAlerts = False
    rs.Parent.Parent.Activate
    
    'Sheets(ro.Parent.CodeName).Activate
    
    Set ro = Application.InputBox("�̵��� ��ġ�� ������ �ּ���", "ó�� ������:B", Type:=8)
    Application.DisplayAlerts = True
    
    With ActiveSheet
        rs.Copy
        ro.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=True

    End With
    
'    rJUcode = rJUcode.Replace("JU", "JU-")
'    Selection.Replace What:="JU", Replacement:="JU-", LookAt:=xlPart, _
'        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
'        ReplaceFormat:=False
'    Set rs = Range("����ܰ�")
'    If opt = 0 Then opt = 4
'    jcost = WorksheetFunction.VLookup(cd, rs, opt, 0)
'    If opt = 4 Then jcost = jcost * 1000
    
End Sub

Sub InputBoxTest()
    Dim MySelection As Range
     
    On Error Resume Next
    Set MySelection = Application.InputBox(prompt:="Select a range of cells", Type:=8)
    With MySelection
        .Parent.Parent.Activate
        .Parent.Activate
        .Select
    End With
End Sub
