Attribute VB_Name = "시작품제조"
Sub ShowUserForm()
Attribute ShowUserForm.VB_ProcData.VB_Invoke_Func = "i\n14"
    UserForm1.show
End Sub


Sub ioCopybook()
Dim rngMulti As Range, range2 As Range, multiplerange As Range
Set range1 = shtdata.Range("B14:H" & LastRow)

    For i = 1 To rngINPUT.Columns.Count
        If i = ci(i) Then
            Set rngTmp = Range(.Cells(1, i), .Cells(end_row, i))
            Set rngMulti = Union(rngMulti, rngTmp)
        End If
    Next i
    
    rngMulti.Select
    
Sheets("ODM").Range("C14").Value = Sheets("Agreement").multiplerange.Value

End Sub


Public Function ioFileFolderExists(strFullPath) As Boolean
'Author       : Ken Puls (www.excelguru.ca)
'Macro Purpose: Check if a file or folder exists
    On Error GoTo EarlyExit
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then ioFileFolderExists = True
    
EarlyExit:
    On Error GoTo 0
End Function

Sub ioOpenFolder()
Attribute ioOpenFolder.VB_ProcData.VB_Invoke_Func = "m\n14"

  Dim preFolder, theFolder, theFolderA, theFolderB, fullFolder As String
  Dim r As Integer
    r = Selection.Row
    preFolder = "D:\RND.분원\시작품제조\"
    theFolderA = Cells(r, 3).Value
    theFolderB = Mid(Cells(r, 15).Value, 1, 5)
    theFolder = theFolderA
    fullFolder = preFolder & theFolder
    
    If FileFolderExists(fullFolder) Then
        'MsgBox "File exists!"
        Call Shell("explorer.exe " & fullFolder, vbNormalFocus)
    Else
        MsgBox theFolder & " / Folder does not exist!"
    End If
        
End Sub

Sub ioCheckFile()
  Dim prePath, thePath, sPath, strFile As String
  Dim r, rs As Range
  Dim c  As Integer
    prePath = "D:\RND.분원\시작품제조\"
    
    rs = Range("list")
    c = rs.Column
    
    
    For Each r In rs.Rows
        thePath = r.Cells(1, 8).Value '업체명 경로
        sPath = prePath & thePath
        
        For i = 17 To 20
            r.Cells(1, i) = FileDate()
        Next i
               
    Next r
    

    
    strFile = Cells(1, c) 'Your Variable here
    strFile = Left(strFile, 3)
    If FileFolderExists(sPath) Then
        'MsgBox "File exists!"
        sFil = Dir(sPath & "\*" & strFile & "*.*")
        Do While sFil <> ""
            checkFile = sFil
            sFil = Dir

        Loop
    Else
        sFil = "x"
    End If

End Sub



Sub ioOpenFile()
Attribute ioOpenFile.VB_ProcData.VB_Invoke_Func = "F\n14"
' ctrl+shift+F
  Dim prePath, posPath, sPath, sKey, sFil, sFile As String
  Dim c, r, cCode As Integer

    r = Selection.Row
    c = Selection.Column
    cCode = WorksheetFunction.Match(vbcKey1, Rows(1), 0)
    sKey = Left(Cells(r, cCode), 5)    '18 코드값 열번호, 코드 5자리
    
    posPath = Cells(1, c).Value & "\"
    prePath = vbcPath1   '상위 폴더
    sPath = prePath & posPath    '검색 폴더
 
    If ioFileFolderExists(sPath) Then
        sFil = Dir(sPath & sKey & "*.*")
        MsgBox sFil '"File exists!"
        Do While sFil <> ""
            sFile = sPath & sFil
            ioOpenAnyFile (sFile)
            sFil = Dir
             'Your Code here.
        Loop
    Else
        MsgBox sFil & " / File does not exist!"
    End If
End Sub

Function ioOpenAnyFile(strPath As String)
  Set objShell = CreateObject("Shell.Application")
    If ioFileThere(strPath) Then
        objShell.Open (strPath)
    Else
        MsgBox ("File not found")
    End If
End Function

Function ioFileThere(FileName As String) As Boolean
     If (Dir(FileName) = "") Then
        ioFileThere = False
     Else:
        ioFileThere = True
     End If
End Function

Sub makeCode()
Attribute makeCode.VB_ProcData.VB_Invoke_Func = "H\n14"
    Dim str, strA As Variant
    str = Selection.Value
    strA = code6(str)
    strB = code3(str)
    strC = Format(Date, "YYMMDD")
   
    MsgBox strA & "-" & strB & "-" & strC
End Sub

Function code6(str As Variant)
'JU1234 -> 001234
    Dim st As String
    Dim ln As Integer
    st = Left(str, 9)
    st = edGetNums(st)
    If st = "" Then st = "000000"
    code6 = Format(st, "000000")
End Function

Function code3(str)
    Dim lang, CLS As String
    If InStr(1, str, "영문", vbTextCompare) Then lang = "(EN)"
    If InStr(1, str, "국문", vbTextCompare) Then lang = "(KR)"
    If lang = "" Then lang = "(EN)"
    If InStr(1, str, "SPEC", vbTextCompare) Then CLS = "SPEC"
    If InStr(1, str, "MSDS", vbTextCompare) Then CLS = "MSDS"
    If InStr(1, str, "GHS", vbTextCompare) Then CLS = "MSDS(GHS)"
    If InStr(1, str, "거래명세서", vbTextCompare) Then CLS = "거래명세서"
    If InStr(1, str, "거래명세표", vbTextCompare) Then CLS = "거래명세서"
    If InStr(1, str, "원산지", vbTextCompare) Then CLS = "원산지증명서"
    If InStr(1, str, "origin", vbTextCompare) Then CLS = "원산지증명서"
    If InStr(1, str, "animal", vbTextCompare) Then CLS = "비동물실험확인서"
    If InStr(1, str, "동물", vbTextCompare) Then CLS = "비동물실험확인서"
    If InStr(1, str, "composition", vbTextCompare) Then CLS = "Composition"
    If InStr(1, str, "조성비", vbTextCompare) Then CLS = "Composition"
    If InStr(1, str, "수출서류", vbTextCompare) Then CLS = "수출서류"
    If InStr(1, str, "유효", vbTextCompare) Then CLS = "유효기간"
    If CLS = "" Then CLS = "ZZZZ"
    code3 = CLS & lang
End Function

Sub ioAcrobatFindText2()

'variables
Dim Resp 'For message box responses
Dim gPDFPath As String
Dim sText As String 'String to search for
Dim sStr As String 'Message string
Dim foundText As Integer 'Holds return value from "FindText" method

'hard coding for a PDF to open, it can be changed when needed.
gPDFPath = "D:\RND.지원\원료LIST 작성\0321.취합자료\수출서류\남영상사(주)\JU-1197(Citric Acid)-영문MSDS.pdf"

'Initialize Acrobat by creating App object
Set gApp = CreateObject("AcroExch.App", "")
gApp.Hide

'Set AVDoc object
Set gAvDoc = CreateObject("AcroExch.AVDoc")

' open the PDF
If gAvDoc.Open(gPDFPath, "") Then
sText = "revision"
'FindText params: StringToSearchFor, caseSensitive (1 or 0), WholeWords (1 or 0), 'ResetSearchToBeginOfDocument (1 or 0)
foundText = gAvDoc.FindText(sText, 1, 0, 1) 'Returns -1 if found, 0 otherwise

Else ' if failed, show error message
Resp = MsgBox("Cannot open" & gPDFPath, vbOKOnly)

End If

If foundText = -1 Then

'compose a message
sStr = "Found " & sText
Resp = MsgBox(sStr, vbOKOnly)
Else ' if failed, 'show error message
Resp = MsgBox("Cannot find" & sText, vbOKOnly)
End If

gApp.show
gAvDoc.BringToFront

End Sub

