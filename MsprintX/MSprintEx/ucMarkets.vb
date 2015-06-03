Imports System.Diagnostics
Imports System.IO
Public Class ucMarkets
    Dim daMarkets As New METISTableAdapters.MarketsTableAdapter
    Dim WithEvents dtMarkets As METIS.MarketsDataTable
    Friend dtSelectedMarkets As New METIS.SelectedMarketsDataTable
    Dim strSelectedMG As String
    Dim lbClearMG As Boolean = True
    Dim fileList As List(Of String) = New List(Of String)()
    Private Sub ucMarkets_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

     
        dtMarkets = daMarkets.GetMarkets
        dgvMarkets.DataSource = dtMarkets
        dgvSelectedMarkets.DataSource = dtSelectedMarkets
        For index = 0 To Directory.GetFiles(System.IO.Path.GetTempPath() + "\\MGS", "*.xml").Count - 1
            ' fileList.Add(Path.GetFileNameWithoutExtension(Directory.GetFiles(System.IO.Path.GetTempPath() + "\\MGS", "*.xml")(index)))
                lbMarketGroup.Items.Add(Path.GetFileNameWithoutExtension(Directory.GetFiles(System.IO.Path.GetTempPath() + "\\MGS", "*.xml")(index)))
        Next
            ' lbMarketGroup.DataSource = fileList
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnCreateGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateGroup.Click

        Dim arrSelected() As METIS.SelectedMarketsRow
        Dim c As XElement
        arrSelected = dtSelectedMarkets.Select("MarketGroup = '-1'")
        If arrSelected.Length < 2 Then
            MsgBox("Please select at least 2 markets for a group", MsgBoxStyle.Exclamation, "Market selection")
            Exit Sub
        End If
        If Trim(txtGroupName.Text).Length = 0 Then
            MsgBox("Please enter a valid name for the group", MsgBoxStyle.Exclamation, "Market group name")
            Exit Sub
        End If
        If lbMarketGroup.FindStringExact(txtGroupName.Text) <> -1 Then
            MsgBox("The market group is already created", MsgBoxStyle.Exclamation, "Market group name")
            Exit Sub
        End If
        lbMarketGroup.Items.Add(txtGroupName.Text)
        c = <mg Name=<%= txtGroupName.Text.Trim() %> type="Group">

            </mg>
        For Each dr As METIS.SelectedMarketsRow In arrSelected
            dr.MarketGroup = txtGroupName.Text
            Dim markets As XElement = New XElement("market")
            markets.Value = dr.Market_desc.Trim()
            c.Add(markets)
        Next
        Dim temppath As String = System.IO.Path.GetTempPath()

        If Not (Directory.Exists(temppath + "\\MGS")) Then
            Directory.CreateDirectory(temppath + "\\MGS")
        End If


        c.Save(temppath + "\\MGS\\" + txtGroupName.Text + ".xml")
        dtSelectedMarkets.DefaultView.RowFilter = "MarketGroup = '-1'"
        txtGroupName.Clear()
        chkSelectAll.Checked = False
        For Each dr As METIS.MarketsRow In dtMarkets.Select("Selected = true")
            dr.Selected = False
        Next
    End Sub

    Private Sub lbMarketGroup_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Windows.Forms.Keys.Delete Then
            Try
                Dim strDeleteMG As String = lbMarketGroup.SelectedItem
                lbMarketGroup.Items.Remove(lbMarketGroup.SelectedItem)
                For Each item In From p In dtSelectedMarkets Where p.MarketGroup = strDeleteMG
                    dtSelectedMarkets.RemoveSelectedMarketsRow(item)
                Next
            Catch ex As Exception
                Debug.WriteLine(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub lbMarketGroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        strSelectedMG = CType(sender, Windows.Forms.ListBox).SelectedItem
        If Not strSelectedMG Is Nothing Then
            dtSelectedMarkets.DefaultView.RowFilter = "MarketGroup = '" & strSelectedMG & "'"
            For Each drMarket In dtMarkets.Rows
                drMarket.Selected = False
                lbClearMG = False
                chkSelectAll.Checked = False
                lbClearMG = True
            Next
        Else
            dtSelectedMarkets.DefaultView.RowFilter = "MarketGroup = '-1'"
        End If
        txtGroupName.Clear()
    End Sub

    Private Sub dgvMarkets_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvMarkets.CellClick
        If e.ColumnIndex <> 0 Then
            Dim dr As METIS.MarketsRow = dtMarkets(e.RowIndex)
            lbMarketGroup.ClearSelected()
            dtSelectedMarkets.DefaultView.RowFilter = "MarketGroup = '-1'"
            If Not dr.Selected Then
                If dtSelectedMarkets.FindByMarketGroupMarket_id("-1", dr.Market_id) Is Nothing Then
                    dtSelectedMarkets.AddSelectedMarketsRow("-1", dr.Market_id, dr.Market_desc)
                End If
            Else
                dtSelectedMarkets.RemoveSelectedMarketsRow(dtSelectedMarkets.Select("MarketGroup = '-1' and Market_id = " + CStr(dr.Market_id))(0))
            End If
        End If
    End Sub

    Private Sub chkSelectAll_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkSelectAll.CheckedChanged
        If Not lbClearMG Then Exit Sub
        lbMarketGroup.ClearSelected()

        For Each drSelected As METIS.SelectedMarketsRow In dtSelectedMarkets.Select("MarketGroup = '-1'")
            dtSelectedMarkets.RemoveSelectedMarketsRow(drSelected)
        Next
        For Each drMarket In From p In dtMarkets
            drMarket.Selected = chkSelectAll.Checked
            If chkSelectAll.Checked Then
                If dtSelectedMarkets.FindByMarketGroupMarket_id("-1", drMarket.Market_id) Is Nothing Then
                    dtSelectedMarkets.AddSelectedMarketsRow("-1", drMarket.Market_id, drMarket.Market_desc)
                End If
            End If
        Next
        dtSelectedMarkets.DefaultView.RowFilter = "MarketGroup = '-1'"
    End Sub

    Private Sub TableLayoutPanel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles TableLayoutPanel1.Paint

    End Sub
End Class
