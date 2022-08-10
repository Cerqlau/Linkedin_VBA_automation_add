
Private Sub Workbook_BeforeClose(Cancel As Boolean)

ThisWorkbook.Application.Visible = True

CloseBook_save

End Sub

Private Sub Workbook_Open()

ThisWorkbook.Application.Visible = False

UserForm.Show


End Sub