' 仅包含代码的窗体导出示例（没有 VERSION/Begin 设计段）。
' 预览会扫描 UserForm_Initialize 中的布局赋值来推断窗体。
Public SelectedSheetName As String

Private Sub UserForm_Initialize()
    With Me
        .Width = 300
        .Height = 150
        .Caption = "Select a worksheet"
        With .lblTitle
            .Caption = "Select a worksheet"
            .Top = 10
            .Left = 10
            .Width = 280
            .Height = 16
            .Font.Bold = False
            .Font.Size = 14
        End With
        With .cmbSheets
            .Top = 30
            .Left = 10
            .Width = 280
            .Height = 20
        End With
        With .btnOK
            .Caption = "OK"
            .Top = 70
            .Left = 70
            .Width = 70
            .Height = 24
            .Default = True
        End With
        With .btnCancel
            .Caption = "Cancel"
            .Top = 70
            .Left = 160
            .Width = 70
            .Height = 24
            .Cancel = True
        End With
    End With
End Sub

Private Sub btnOK_Click()
    If cmbSheets.ListIndex <> -1 Then SelectedSheetName = cmbSheets.Value
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    SelectedSheetName = ""
    Me.Hide
End Sub
