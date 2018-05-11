Attribute VB_Name = "MMenu"
'
'======================================================================
' �����޴� �����
'======================================================================
'
'****** esMenu���� ������ 15.6.26
' �������� ����

Option Explicit
Const cName As String = "HKC" '���������� �̸� Commandbar Name

Sub MakeMenu(cName As String)
'�޴� �����
Dim c As CommandBar
Dim strCur As String
    strCur = ThisWorkbook.Name & "!"
    DeleteMenu (cName)
    Set c = Application.CommandBars.Add(Name:=cName, Position:=msoBarTop, MenuBar:=False, Temporary:=False)
    
    'MakeSubMenu c, msoControlPopup, False, strCur & "SplitTableIntoMultiSheets", "���־���", "���־���2", 2090
    
    MakeSubMenu c, msoControlButton, False, strCur & "SplitTableIntoMultiSheets", "���к� ��Ʈ����", "���к���Ʈ����", 2090
    MakeSubMenu c, msoControlButton, False, strCur & "DeleteMenu", "��Ʈ������ũ������", "��Ʈ������ũ������", 2087
    MakeSubMenu c, msoControlButton, False, strCur & "CDO_Mail_Small_Text", "���Ϻ�����", "���Ϻ�����", 24
    MakeSubMenu c, msoControlButton, False, strCur & "exSameBColor_HideColumns", "�����÷� ���߱�", "", 352
    MakeSubMenu c, msoControlButton, False, strCur & "exSameBColor_ShowColumns", "�����÷� ���̱�", "", 351
    
    MakeSubMenu c, msoControlButton, False, strCur & "openFile", "���Ͽ���", "���Ͽ���2", 23
    MakeSubMenu c, msoControlButton, False, strCur & "edMergeTXT2Cell", "������", "���Ͽ���2", 24
    MakeSubMenu c, msoControlButton, False, strCur & "edMergeTXT2Cell2", "������2", "���Ͽ���2", 24
    MakeSubMenu c, msoControlButton, False, strCur & "edMergeRange", "��������", "���Ͽ���2", 25
        
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
    Application.CommandBars(cName).Delete '�޴� �����ϱ�
    AddIns(cName).Installed = False
On Error GoTo 0
End Sub

Sub delMenu2()
    DeleteMenu ("MyAddinFirst")
    DeleteMenu ("MyAddinSecond")
End Sub
'
'======================================================================
' �����޴� ����� 2
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
  Set myMenu = cmdbar.Controls("A�׷�").Add(Type:=msoControlPopup)
  
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
    ribbonXml = ribbonXml + "      <mso:tab id=""tool1"" label=""����1"" insertBeforeQ=""mso:TabFormat"">"
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
' ��Ÿ
'======================================================================
'

Sub hello()
    MsgBox "Hello!!"
End Sub
Sub DB����()
  '���������.�̸�����(���� �κ�) ���Ͽ��� 150406
    Dim rs As Range
    Dim wbk As Worksheet
    Dim i As Range
    Set rs = rdefNames
    Set wbk = ActiveSheet
    ioOpenBydNames
    wbk.Activate
 End Sub

Sub DB_Open_by_Activecell()
  '���������.�̸�����(���� �κ�) ���Ͽ��� 150406
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
    Set xclass.App = Application 'Ŭ������� ���
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
          End With
          
          With .Controls.Add(Type:=msoControlButton)
            .BeginGroup = True
            .Caption = "�糡���� �����"
            .ShortcutText = "Ctrl+Shift+T"
            .OnAction = "Trim_Text"
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

