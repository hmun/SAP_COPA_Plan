Private Sub Workbook_BeforeClose(Cancel As Boolean)
  Ret = delSAPCommandbar
End Sub

Private Sub Workbook_Open()
  Ret = addSAPCommandbar
End Sub

