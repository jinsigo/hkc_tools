Attribute VB_Name = "ShortKey"
Private Sub Workbook_Activate()
    'Application.OnKey "+^{RIGHT}", "YourMacroName"
    Application.OnKey "^m", "부자재재고_고객사별시트"
End Sub

Private Sub Workbook_Deactivate()
    Application.OnKey "+^{RIGHT}"
End Sub

Sub setOnKeys()
    Application.OnKey "+^{M}", "copy2Works"
    Application.OnKey "+^{N}", "copy2Form"
    Application.OnKey "+^{G}", "Go2Database2"
'    Application.OnKey "+^{G}", "Go2Database"
End Sub
