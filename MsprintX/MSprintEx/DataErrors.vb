Public Class DataErrors
    Private Sub btnErrors_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnErrors.Click
        RaiseEvent ShowError_Click()
    End Sub
    Private Sub dgvErrors_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvErrors.CellContentClick
        RaiseEvent Address_Click(e)
    End Sub
    Public Property DataSource() As Object
        Get
            Return dgvErrors.DataSource
        End Get
        Set(ByVal value As Object)
            If Not value Is Nothing Then
                dgvErrors.DataSource = value
            End If
        End Set
    End Property
    Public Event ShowError_Click()
    Public Event Address_Click(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
    Private Sub ucDataErrors_Address_Click(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Me.Address_Click
        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

        If sheet.Name.Equals("Plan Selection") Then
            'CreateWeekColumns()
            sheet.Select()
            sheet.Range(Globals.Ribbons.MSprintExRibbon.errors.dgvErrors(e.ColumnIndex, e.RowIndex).Value).Select()
        End If

    End Sub
End Class
