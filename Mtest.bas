Attribute VB_Name = "Mtest"

Sub LoadArray(ByRef oRange, ByRef vArray)
    Dim i As Integer
    For i = 0 To UBound(vArray)
        oRange.Offset(0, i).Value = vArray(i)
    Next
End Sub



Sub RegExpRed()
    Dim objRegex
    Dim RegMC
    Dim RegM
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
        .Global = True
        .Pattern = "\d+"
        If .test(Cells(1, 1).Value) Then
            Set RegMC = .Execute(Cells(1, 1).Value)
            For Each RegM In RegMC
                Cells(1, 1).Characters(RegM.FirstIndex + 1, RegM.Length).Font.Color = vbRed
            Next
        End If
    End With
End Sub


Sub copy2Works()
'
' 매크로3 매크로
'

'
    Dim wd, ws As Object
    Dim rs As Range
    
    Set ws = ActiveWindow
    Set wd = Windows("성분표기작성5.0.xlsm")
    
    'Windows("20150828100214.CDVSK.xls").Activate
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
    
    Windows("성분표기작성5.0.xlsm").Activate
    Range("B6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("A6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="JU", Replacement:="JU-", LookAt:=xlPart, _
        SearchOrder:=xlByColumns, MatchCase:=True, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Range("B6").Select
    Range(Selection, Selection.End(xlDown)).Select
End Sub

Sub copy2form()
'
' 서식에 붙여넣기
'

'
    Dim wd, ws As Object
    Dim rs, rd As Range
    
    Set ws = Windows("성분표기작성5.0.xlsm")
    Set wd = ActiveWindow
    
    ws.Activate
    Range("E6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    wd.Activate
    Range("B9").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'ActiveWindow.SmallScroll Down:=9
    
    ws.Activate
    Range("G6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    wd.Activate
    Range("C9").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ws.Activate
    Range("B6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    wd.Activate
    Range("E9").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    

End Sub

Sub CDO_Mail_Small_Text()
    Dim iMsg As Object
    Dim iConf As Object
    Dim strbody As String
    '    Dim Flds As Variant

    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")

        iConf.Load -1    ' CDO Source Defaults
        Set Flds = iConf.Fields
        With Flds
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "gw.ihkcos.com"
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "jinsi@ihkcos.com"
            .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "dofqls68"
            .Update
        End With

    strbody = "안녕하세요." & _
              vbNewLine & vbNewLine & _
              "기술개발연구원 메이크업 3팀 이진성입니다." & vbNewLine & _
              "제조계획표 보내드리고 첨부 파일을 확인해주시기 바랍니다~" & vbNewLine & _
              "수고하세요" & vbNewLine & _
              "이 메일은 자동으로 발송되는 메일입니다."
    
    strdate = Format(Date, "yy-mm-dd")
    att1 = "D:\RND.분원\시작품제조\SCH\" & strdate & " 기준서.xls"
    att2 = "D:\RND.분원\시작품제조\SCH\" & strdate & " 제조계획표.xls"
    att3 = "D:\RND.분원\시작품제조\SCH\" & strdate & " 중간공정 계획.xlsx"
    
    With iMsg
        Set .Configuration = iConf
        .To = """이진성"" <jinsi@ihkcos.com>"
        .cc = ""
        .BCC = ""
        .From = """이진성"" <jinsi@ihkcos.com>"
        .Subject = "Mail Testing"
        .TextBody = strbody
        .AddAttachment (att1)
        .AddAttachment (att2)
        .AddAttachment (att3)
        .Send
        
    End With

End Sub


