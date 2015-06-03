Public Class ucBrand
    Dim brandcopy As Data.DataTable
    Private Sub clbSelectBrands_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles clbSelectBrands.MouseDoubleClick
        If Not (lbSelectedBrands.Items.Contains(clbSelectBrands.SelectedItem.ToString())) Then
            lbSelectedBrands.Items.Add(clbSelectBrands.SelectedItems(0).ToString())
        End If


        ' lbSelectedChannels.Items.Remove(lbSelectedChannels.SelectedItem)
        lbSelectedBrands.Refresh()
        clbSelectBrands.Refresh()
        clbSelectBrands.Sorted = True
        lbSelectedBrands.Sorted = True
    End Sub

    Private Sub clbSelectBrands_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clbSelectBrands.SelectedIndexChanged

    End Sub

    Private Sub lbSelectedBrands_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lbSelectedBrands.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Delete Then
            Try
                Dim strDeleteGenre As String = lbSelectedBrands.SelectedItem
                lbSelectedBrands.Items.Remove(lbSelectedBrands.SelectedItem)
                'For Each item In From p In dtSelectedMarkets Where p.MarketGroup = strDeleteMG
                '    dtSelectedMarkets.RemoveSelectedMarketsRow(item)
                'Next
            Catch ex As Exception
                'Debug.WriteLine(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub lbSelectedBrands_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbSelectedBrands.SelectedIndexChanged

    End Sub

    Private Sub lbChosenBrands_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbChosenBrands.Click

    End Sub

    Private Sub lbSelectBrands_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbSelectBrands.Click

    End Sub

    Private Sub ClbSelectAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClbSelectAll.CheckedChanged
        Try
            If ClbSelectAll.Checked Then
                lbSelectedBrands.Items.Clear()
                For index = 0 To clbSelectBrands.Items.Count - 1
                    'clbSelectChannels.SetSelected(index, True)
                    lbSelectedBrands.Items.Add(clbSelectBrands.Items(index))

                Next
                '   lbSelectedAdvertisers.Sorted = True
            Else
                lbSelectedBrands.Items.Clear()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    'Private Sub ucBrand_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    Try
    '        brandcopy = Globals.Ribbons.MSprintExRibbon.bslMasterTable.Copy
    '        '  brandcopy = brandcopy.DefaultView.ToTable("Brand", True)
    '        For Each row As Data.DataRow In brandcopy.Rows
    '            Globals.Ribbons.MSprintExRibbon.tpSelections.UcBrand1.clbSelectBrands.Items.Add(row("Brand").ToString())
    '        Next
    '    Catch ex As Exception
    '        LogMpsrintExException("Exception occured while loading Brand User control.Message :" + ex.Message)
    '    End Try
    'End Sub

    Private Sub btnClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearAll.Click
        lbSelectedBrands.Items.Clear()
    End Sub
End Class
