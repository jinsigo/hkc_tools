Attribute VB_Name = "원료리스트"

'======================================================================
' 원료리스트 관련 함수
'======================================================================
Sub rmSetNames_old()
Attribute rmSetNames_old.VB_Description = "원료리스트 이름정의"
Attribute rmSetNames_old.VB_ProcData.VB_Invoke_Func = " \n14"
'
' mk_원료LIST 매크로 2013.9.30 / 이진성
' 원료리스트 이름정의
'
'    Windows("D:\1.원료성분\원료LIST.xls").Activate
'    Windows("D:\1.원료성분\기능성함량.xlsx").Activate
'    ActiveWorkbook.Names.Item("원료LIST").Delete
'    ActiveWorkbook.Names.Item("기능성").Delete
    Dim fn(3) As String
    Dim dn(3) As String
    Dim nn(3) As String
    Dim wrk As Workbook

    Set wrk = ActiveWorkbook

    dn(1) = "D:\RND\원료성분"
    dn(2) = "D:\RND\원료성분"
    fn(1) = "원료LIST.xls"
    fn(2) = "원료LIST.기능성함량.xlsx"
    nn(1) = "원료LIST"
    nn(2) = "기능성"

    ChDir (dn(1))
    If isOpenWrk(fn(1)) = 0 Then Workbooks.Open FileName:=dn(1) & "\" & fn(1)
    If isOpenWrk(fn(2)) = 0 Then Workbooks.Open FileName:=dn(2) & "\" & fn(2)

    With wrk
        .Activate
        If isOpenName(nn(1)) > 0 Then .Names(nn(1)).Delete
        If isOpenName(nn(2)) > 0 Then .Names(nn(2)).Delete

        .Names.Add Name:=nn(1), RefersToR1C1:="='" & fn(1) & "'!Database"
        .Names.Add Name:=nn(2), RefersToR1C1:="='" & fn(2) & "'!기능성"
        .Names(nn(1)).Comment = nn(1)
        .Names(nn(2)).Comment = nn(1)
        MsgBox wrk.Name
    End With

End Sub




Sub rmPutData()
'
' 영역주소 임시저장 매크로 2015.2.4 / 이진성
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
' 영역주소 가져오기 매크로 2015.2.4 / 이진성
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
' ju코드 단가 가져오기 / 이진성
'2013.6.13
    fn = "D:\똥강아지\HKC\LST\"
    cd = Replace(cd, "-", "")
    Set rs = Range("원료단가")
    If opt = 0 Then opt = 4
    jcost = WorksheetFunction.VLookup(cd, rs, opt, 0)
    If opt = 4 Then jcost = jcost * 1000

End Function

Sub rmRead_BOM()
' BOM에서 JU코드 가져오기
' 2015.12.21 이진성 ;
Dim rs, ro, rJUcode, rComp As Range
Dim np As Integer

    Set rJUcode = Selection 'Range(rs.Cells(5, 2), rs.End(xlDown)) 'BOM 자재 영역

    np = InputBox("소요량 상대위치를 입력해 주세요(기본값은 '5' 입니다).", "처방 가져오기", 5) '소요량 상대위치 가져오기
    Set rComp = rJUcode.Offset(0, np)

    Set rs = Union(rJUcode, rComp)

    'Application.DisplayAlerts = False
    rs.Parent.Parent.Activate

    'Sheets(ro.Parent.CodeName).Activate

    Set ro = Application.InputBox("이동할 위치를 선택해 주세요", "처방 보내기:B", Type:=8)
    Application.DisplayAlerts = True

    With ActiveSheet
        rs.Copy
        ro.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=True

    End With

'    rJUcode = rJUcode.Replace("JU", "JU-")
'    Selection.Replace What:="JU", Replacement:="JU-", LookAt:=xlPart, _
'        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
'        ReplaceFormat:=False
'    Set rs = Range("원료단가")
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
