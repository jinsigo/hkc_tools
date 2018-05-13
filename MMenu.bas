Attribute VB_Name = "MMenu"
'
'======================================================================
' 리본메뉴 만들기
'======================================================================
'
'****** esMenu에서 가져옴 15.6.26
' 도구모음 정의

Option Explicit
Const cName As String = "HKC" '도구모음의 이름 Commandbar Name

Sub MakeMenu(cName As String)
'메뉴 만들기
Dim c As CommandBar
Dim strCur As String
    strCur = ThisWorkbook.Name & "!"
    DeleteMenu (cName)
    Set c = Application.CommandBars.Add(Name:=cName, Position:=msoBarTop, MenuBar:=False, Temporary:=False)

    'MakeSubMenu c, msoControlPopup, False, strCur & "SplitTableIntoMultiSheets", "자주쓰는", "자주쓰는2", 2090

    MakeSubMenu c, msoControlButton, False, strCur & "SplitTableIntoMultiSheets", "구분별 시트생성", "구분별시트생성", 2090
    MakeSubMenu c, msoControlButton, False, strCur & "DeleteMenu", "시트생성매크로종료", "시트생성매크로종료", 2087
    MakeSubMenu c, msoControlButton, False, strCur & "CDO_Mail_Small_Text", "메일보내기", "메일보내기", 24
    MakeSubMenu c, msoControlButton, False, strCur & "exSameBColor_HideColumns", "동일컬럼 감추기", "", 352
    MakeSubMenu c, msoControlButton, False, strCur & "exSameBColor_ShowColumns", "동일컬럼 보이기", "", 351

    MakeSubMenu c, msoControlButton, False, strCur & "openFile", "파일열기", "파일열기2", 23
    MakeSubMenu c, msoControlButton, False, strCur & "edMergeTXT2Cell", "셀병합", "파일열기2", 24
    MakeSubMenu c, msoControlButton, False, strCur & "edMergeTXT2Cell2", "셀병합2", "파일열기2", 24
    MakeSubMenu c, msoControlButton, False, strCur & "edMergeRange", "범위병합", "파일열기2", 25

    c.Visible = True
    c.Name = cName
    Set c = Nothing
End Sub

Sub MakeSubMenu(c As CommandBar, lngType As Long, blnBar As Boolean, strOnAction As String, strTip As String, strCaption As String, lngFaceId As Long)
    With c.Controls.Add(Type:=lngType)
        .BeginGroup = blnBar
        .OnAction = strOnAction
        .TooltipText = strTip
        .Caption = strCaption
        .FaceId = lngFaceId

    End With
End Sub

Sub MakeSubMenu2(c As CommandBar, lngType As Long, blnBar As Boolean, strOnAction As String, strTip As String, strCaption As String, lngFaceId As Long)
    With c.Controls.Add(Type:=lngType)
        .BeginGroup = blnBar
        .TooltipText = strTip
        .Caption = strCaption
        .FaceId = lngFaceId
        .ShortcutText = "Alt+PageUp"
        .OnAction = strOnAction
    End With
End Sub


Sub DeleteMenu(cName As String)
On Error Resume Next
    Application.CommandBars(cName).Delete '메뉴 삭제하기
    AddIns(cName).Installed = False
On Error GoTo 0
End Sub

Sub delMenu2()
    DeleteMenu ("MyAddinFirst")
    DeleteMenu ("MyAddinSecond")
End Sub
'
'======================================================================
' 리본메뉴 만들기 2
'======================================================================
'
Sub QryMenu()
    Dim foundFlag As Boolean
    Dim cb As CommandBar
    Dim newItem As CommandBarControl
    Dim cbName As String
    cbName = "AddMenu"
'    foundFlag = False
'    For Each cb In CommandBars
'        If cb.Name = cbName Then
'            cb.Protection = msoBarNoChangeDock
'            cb.Visible = True
'            foundFlag = True
'        End If
'        MsgBox "Command bar name:" & cb.Name
'    Next cb
'    If Not foundFlag Then
'        MsgBox "The collection does not contain " & cbName
'End If
    Set newItem = CommandBars("Tools").Controls.Add(Type:=msoControlButton)
    With newItem
        .BeginGroup = True
        .Caption = "Make Report"
        .FaceId = 0
        .OnAction = "qtrReport"
    End With
End Sub

Sub AddMenu()
  Dim cmdbar As CommandBar
  Dim toolsMenu As CommandBarControl
  Dim myMenu As CommandBarPopup
  Dim subMenu As CommandBarControl


' Point to the Worksheet Menu Bar
  Set cmdbar = Application.CommandBars("HKC")

' Point to the Tools menu on the menu bar
'  Set toolsMenu = cmdbar.Controls("AddMenu")

' Create My Menu
  Set myMenu = cmdbar.Controls("A그룹").Add(Type:=msoControlPopup)

' Create the sub Menu(s)
  Set subMenu = myMenu.Controls.Add

  With myMenu
    .Caption = "My Menu"
    .BeginGroup = True
    With subMenu
      .Caption = "sub Menu"
      .BeginGroup = True
      .OnAction = "'" & ThisWorkbook.Name & "'!myMacro" ' Assign Macro to Menu Item
    End With
  End With


End Sub

 ' How to remove the menu item
Sub RemoveMenu()

  On Error Resume Next
  Dim cmdbar As CommandBar
  Dim CmdBarMenu As CommandBarControl
  Set cmdbar = Application.CommandBars("Worksheet Menu Bar")
  Set CmdBarMenu = cmdbar.Controls("Tools")
  CmdBarMenu.Controls("My Menu").Delete
End Sub

Private Sub rbAddRibbon()
    Dim ribbonXml As String

    ribbonXml = "<mso:customUI xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">"
    ribbonXml = ribbonXml + "  <mso:ribbon>"
    ribbonXml = ribbonXml + "    <mso:qat/>"
    ribbonXml = ribbonXml + "    <mso:tabs>"
    ribbonXml = ribbonXml + "      <mso:tab id=""tool1"" label=""도구1"" insertBeforeQ=""mso:TabFormat"">"
    ribbonXml = ribbonXml + "        <mso:group id=""testGroup"" label=""Test"" autoScale=""true"">"
    ribbonXml = ribbonXml + "          <mso:button id=""highlightManualTasks"" label=""Toggle Manual Task Color"" "
    ribbonXml = ribbonXml + "imageMso=""DiagramTargetInsertClassic"" onAction=""ToggleManualTasksColor""/>"
    ribbonXml = ribbonXml + "        </mso:group>"
    ribbonXml = ribbonXml + "      </mso:tab>"
    ribbonXml = ribbonXml + "    </mso:tabs>"
    ribbonXml = ribbonXml + "  </mso:ribbon>"
    ribbonXml = ribbonXml + "</mso:customUI>"

    ActiveProject.SetCustomUI (ribbonXml)
End Sub



'======================================================================
' 기타
'======================================================================
'

Sub hello()
    MsgBox "Hello!!"
End Sub
Sub DB열기()
  '사용자정의.이름영역(색상 부분) 파일열기 150406
    Dim rs As Range
    Dim wbk As Worksheet
    Dim i As Range
    Set rs = rdefNames
    Set wbk = ActiveSheet
    ioOpenBydNames
    wbk.Activate
 End Sub

Sub DB_Open_by_Activecell()
  '사용자정의.이름영역(색상 부분) 파일열기 150406
    Dim rs As Range
    Dim wbk As Worksheet
    Dim i As Range
    ioOpendByActiveCell
'    Set rs = rdefNames
    Set wbk = ActiveSheet
'    For Each i In rs.Rows
'        If i.Cells(1, 1).Interior.ColorIndex <> -4142 Then
            ioOpendNames
'            MsgBox i.Cells(1, 1).Interior.ColorIndex & i.Cells(1, 1).Value
'        End If
'    Next i
    wbk.Activate
End Sub


Sub Addin_Open_MyAddin(Optional x As Boolean)
Dim cmdCont As CommandBarControl
    On Error Resume Next
    Set xclass.App = Application '클래스모듈 사용
    Call ImmSetConversionStatus(ImmGetContext(FindWindow("XLMAIN", _
      Application.Caption)), &H1, &H0) 'IME_CMODE_HANGEUL, IME_SMODE_NONE
    With Application
      With .CommandBars("Tools")
        Set cmdCont = .FindControl(Tag:="HKC")
        If Not cmdCont Is Nothing Then cmdCont.Delete
        With .Controls.Add(Type:=msoControlPopup, Before:=1, Temporary:=True)
          .Caption = "HKC"
          .Tag = "HKC Excel Macro"
          .OnAction = "menu_enable"

          With .Controls.Add(Type:=msoControlPopup)
            .Caption = "자주쓰는 단축키"
            .Tag = "자주쓰는 단축키"
            With .Controls.Add(Type:=msoControlButton)
              .FaceId = 40
              .Caption = "아래쪽 범위선택"
              .ShortcutText = "Alt+PageDown"
              .OnAction = "End1_Cell"
            End With
            With .Controls.Add(msoControlButton)
              .FaceId = 39
              .Caption = "오른쪽 범위선택"
              .ShortcutText = "Alt+PageUp"
              .OnAction = "End2_Cell"
            End With
            With .Controls.Add(msoControlButton)
              .Caption = "마지막셀 재설정"
              .ShortcutText = "Ctrl+Alt+End"
              .OnAction = "Sheet_Refresh"
            End With
          End With

          With .Controls.Add(Type:=msoControlButton)
            .BeginGroup = True
            .Caption = "양끝공백 지우기"
            .ShortcutText = "Ctrl+Shift+T"
            .OnAction = "Trim_Text"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .BeginGroup = True
            .FaceId = 160
            .Caption = "선택페이지 인쇄"
            .ShortcutText = "Ctrl+Shift+P"
            .OnAction = "Print_SelectPage"
          End With
          With .Controls.Add(Type:=msoControlButton)
            .BeginGroup = True
            .Caption = "Addin2015 도움말"
            .OnAction = "Help_Text"
            .FaceId = 273 '종모양
          End With
        End With
      End With
    End With
    Onkey_Make
    Randomize
End Sub
