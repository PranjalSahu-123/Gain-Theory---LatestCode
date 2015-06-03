Public Class ucCategory
    Dim categorycopy As Data.DataTable
    Private Sub clbSelectCategories_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clbSelectCategories.SelectedIndexChanged

    End Sub

    Private Sub clbSelectCategories_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles clbSelectCategories.MouseDoubleClick
        If Not (lbSelectedCategories.Items.Contains(clbSelectCategories.SelectedItems(0).ToString())) Then
            lbSelectedCategories.Items.Add(clbSelectCategories.SelectedItems(0).ToString())
        End If


        ' lbSelectedChannels.Items.Remove(lbSelectedChannels.SelectedItem)
        lbSelectedCategories.Refresh()
        clbSelectCategories.Refresh()
        clbSelectCategories.Sorted = True
        lbSelectedCategories.Sorted = True
    End Sub

    Private Sub lbSelectedCategories_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbSelectedCategories.SelectedIndexChanged

    End Sub

    Private Sub lbSelectedCategories_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lbSelectedCategories.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Delete Then
            Try
                Dim strDeleteGenre As String = lbSelectedCategories.SelectedItem
                lbSelectedCategories.Items.Remove(lbSelectedCategories.SelectedItem)
                'For Each item In From p In dtSelectedMarkets Where p.MarketGroup = strDeleteMG
                '    dtSelectedMarkets.RemoveSelectedMarketsRow(item)
                'Next
            Catch ex As Exception
                'Debug.WriteLine(ex.ToString)
            End Try
        End If
    End Sub
    'Private Sub ucCategory_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    Try
    '        categorycopy = Globals.Ribbons.MSprintExRibbon.bslMasterTable.Copy
    '        ' categorycopy = categorycopy.DefaultView.ToTable("Category", True)
    '        For Each row As Data.DataRow In categorycopy.Rows

    '            If Not (Globals.Ribbons.MSprintExRibbon.tpSelections.UcCategory1.clbSelectCategories.Items.Contains(row("Category").ToString())) Then
    '                Globals.Ribbons.MSprintExRibbon.tpSelections.UcCategory1.clbSelectCategories.Items.Add(row("Category").ToString())

    '            End If

    '        Next
    '        clbSelectCategories.Sorted = True
    '    Catch ex As Exception
    '        LogMpsrintExException("Exception occured while loading category user control.Message :" + ex.Message)
    '    End Try
    'End Sub

    Private Sub ClbSelectAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClbSelectAll.CheckedChanged
        Try
            If ClbSelectAll.Checked Then
                lbSelectedCategories.Items.Clear()
                For index = 0 To clbSelectCategories.Items.Count - 1
                    'clbSelectChannels.SetSelected(index, True)
                    lbSelectedCategories.Items.Add(clbSelectCategories.Items(index))

                Next
                lbSelectedCategories.Sorted = True
            Else
                lbSelectedCategories.Items.Clear()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearAll.Click
        lbSelectedCategories.Items.Clear()
    End Sub
End Class
