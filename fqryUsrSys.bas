Attribute VB_Name = "fqryUsrSys"
'
'======================================================================
' Module Name: 사용자 시스템 정보 확인 함수
'======================================================================
'
'
' Reference http://p2p.wrox.com/access-vba/42695-retrieve-system-information-vba.html
' 16/02/04 modified by Jinsi
' 16/06/20 published by Jinsi
'
' >>Function
' OS_ComputerName   '이진성
' OS_ComputerInfo   '연구소
' OS_UserName       '이진성
' OS_Information
' MAC_Address       'a0:ff:ff:ff
' IP_Address        '133.9.176
' VBA_Version       '7.0

Option Explicit

Public Const mName As String = "시트 나누기 매크로 by 진성"
Public Const cMenu As String = "시트 나누기" '도구 모음의 이름
Public Const cUser As String = "이진성"
Public Const cExp  As String = "2017-06-21"
Public Const vbcPath1 As String = "D:\RND.분원\시작품제조\" '루트 경로1
Public Const vbcKey1 As String = "제품코드"
Public Const vbcWRLST As String = "원료LIST.xls"
Public Const hkc_DB1 As String = "C:\Users\jinsigo\OneDrive\HKC\원료LIST.xls"




Sub show()
Attribute show.VB_ProcData.VB_Invoke_Func = "S\n14"
    MsgBox "Results :" & VBA_Version
End Sub


'======================================================================
' 시스템 정보 함수
'======================================================================
Public Function OS_Information() As Variant
  Dim arrOS_Information(1 To 3)   As String
  Dim curWMI As Object, curObj As Object, Itm
    
  Set curWMI = GetObject("winmgmts:\\.\root\cimv2")
  Set curObj = curWMI.ExecQuery("Select * from Win32_OperatingSystem", , 48)
  For Each Itm In curObj
    arrOS_Information(1) = Itm.Caption 'OS
    arrOS_Information(2) = Itm.BuildNumber 'OS Build
    arrOS_Information(3) = Itm.CSDVersion '
    arrOS_Information(4) = Itm.Version '
  Next
  OS_Information = arrOS_Information()
End Function

Function OS_ComputerName() As String
  OS_ComputerName = CreateObject("wscript.network").Computername
End Function

Public Function OS_ComputerInfo() As Variant
  Dim arrOS_ComputerInfo(1 To 3)   As String
  Dim curWMI As Object, curObj As Object, Itm
    
  Set curWMI = GetObject("winmgmts:\\.\root\cimv2")
  Set curObj = curWMI.ExecQuery("Select * from Win32_ComputerSystem", , 48)
  For Each Itm In curObj
    arrOS_ComputerInfo(1) = Itm.Domain
    arrOS_ComputerInfo(2) = Itm.Manufacturer
    arrOS_ComputerInfo(3) = Itm.Model
  Next
  OS_ComputerInfo = arrOS_ComputerInfo()
End Function

Public Function VBA_Version() As String
  VBA_Version = Application.VBE.Version ' less information, but faster
End Function

Function CheckUser() As Integer
    Dim msg, msg4 As String
    If OS_ComputerName = cUser And Date < DateValue(cExp) Then
        CheckUser = 1
    Else
        msg4 = "개발자에게 문의바랍니다." & Chr(10) & "개발자: 이진성 Tel.010-5382-4086 "
        msg = MsgBox(msg4, 0, cMenu)
        CheckUser = 0
    End If
    
End Function

Function CheckComputer()
    com = ThisWorkbook.Sheets("sheet1").Cells(1, 1)
    MsgBox com
    If com = OS_ComputerName Then
      MsgBox "OK"
    Else
        MsgBox "Wrong"
        Stop
    End If
End Function

Function CheckIP(chk) As Integer
    ln = Len(chk)
    
    If Left(IP_Address, ln) = chk Then
       CheckUser = 1
       Exit Function
      Else
       CheckUser = 0
      End If
    Next
    If CheckUser = 0 Then
        MsgBox "개발자에게 문의바랍니다." & Chr(10) & "Tel.010-5382-4086 이진성"
        Stop
    End If
End Function

Public Function IP_Address()
' 현재 IP Address
    Dim curWMI As Object, curObj As Object, Itm
    
    Set curWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set curObj = curWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    For Each Itm In curObj
      IP_Address = Itm.IPAddress(0)
      Exit Function
    Next
End Function

Public Function MAC_Address() As String
    Dim curWMI As Object, curObj As Object, Itm
    
    Set curWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set curObj = curWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    For Each Itm In curObj
      MAC_Address = Itm.MACAddress
      Exit Function
    Next
End Function

Public Function OS_UserName() As String
  OS_UserName = CreateObject("wscript.network").UserName
End Function


'======================================================================
' 위크시트 정보 함수
'======================================================================
'
Function UseAddIns(fn)

    Set myAddIn = AddIns.Add(FileName:=fn, CopyFile:=True)
    MsgBox myAddIn.Title & " has been added to the list"
End Function

Public Function IsSheet(inp) As Integer
'시트 존재 여부
    Dim ws As Worksheet
    IsSheet = 0
    For Each ws In ActiveWorkbook.Sheets
        If ws.Name = inp Then IsSheet = 1
    Next ws
End Function

'======================================================================
' 특수문자 지우기
'======================================================================
' 160624 jinsigo@naver.com
Function RemoveSpecialChars(strOLD)
    Const SpecialCharacters As String = "!,@,#,$,%,^,/,&,*,(,),{,[,],}"  '지울 문자들
    Dim strNEW As String
    Dim c As Variant
    strNEW = strOLD
    For Each c In Split(SpecialCharacters, ",")
        strNEW = Replace(strNEW, c, "")
    Next
    RemoveSpecialChars = strNEW
End Function
