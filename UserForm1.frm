VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Database"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ListBox1_Click()

End Sub

Private Sub UserForm1_Initialize()
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
        With ListBox1
            .ColumnCount = 3
            .ColumnWidths = "150;120;100"
            .ColumnHeads = True
            .AddItem
            .List(i, 0) = rngDB.Cells(i, 1)    'DB ��Ī
            .List(i, 1) = rngDB.Cells(i, 2)    '���
            .List(i, 2) = rngDB.Cells(i, 3)    'cas
        End With
    Next i
End Sub

Private Sub UserForm_Click()

End Sub
