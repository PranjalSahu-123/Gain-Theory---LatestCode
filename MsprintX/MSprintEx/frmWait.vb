Public Class frmWait

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Dispose()
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles imgWait.Click

    End Sub

    Private Sub pbUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub frmWait_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Windows.Forms.Application.DoEvents()
    End Sub
End Class