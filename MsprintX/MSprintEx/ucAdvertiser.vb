Public Class ucAdvertiser
    Dim advertisercopy As Data.DataTable
    Private Sub lbSelectedAdvertisers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbSelectedAdvertisers.SelectedIndexChanged

    End Sub

    Private Sub lbSelectedAdvertisers_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lbSelectedAdvertisers.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Delete Then
            Try
                Dim strDeleteGenre As String = lbSelectedAdvertisers.SelectedItem
                lbSelectedAdvertisers.Items.Remove(lbSelectedAdvertisers.SelectedItem)
                'For Each item In From p In dtSelectedMarkets Where p.MarketGroup = strDeleteMG
                '    dtSelectedMarkets.RemoveSelectedMarketsRow(item)
                'Next
            Catch ex As Exception
                'Debug.WriteLine(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub clbSelectAdvertisers_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles clbSelectAdvertisers.MouseDoubleClick
        If Not (lbSelectedAdvertisers.Items.Contains(clbSelectAdvertisers.SelectedItems(0).ToString())) Then
            lbSelectedAdvertisers.Items.Add(clbSelectAdvertisers.SelectedItems(0).ToString())
        End If


        ' lbSelectedChannels.Items.Remove(lbSelectedChannels.SelectedItem)
        lbSelectedAdvertisers.Refresh()
        clbSelectAdvertisers.Refresh()
        clbSelectAdvertisers.Sorted = True
        lbSelectedAdvertisers.Sorted = True
    End Sub

    Private Sub clbSelectAdvertisers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clbSelectAdvertisers.SelectedIndexChanged

    End Sub

    Private Sub btnClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearAll.Click
        lbSelectedAdvertisers.Items.Clear()
    End Sub

    Private Sub ClbSelectAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClbSelectAll.CheckedChanged
        Try
            If ClbSelectAll.Checked Then
                lbSelectedAdvertisers.Items.Clear()
                For index = 0 To clbSelectAdvertisers.Items.Count - 1
                    'clbSelectChannels.SetSelected(index, True)
                    lbSelectedAdvertisers.Items.Add(clbSelectAdvertisers.Items(index))

                Next
                lbSelectedAdvertisers.Sorted = True
            Else
                lbSelectedAdvertisers.Items.Clear()
            End If
        Catch ex As Exception

        End Try
    End Sub

    'Private Sub ucAdvertiser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    Try
    '        advertisercopy = Globals.Ribbons.MSprintExRibbon.bslMasterTable.Copy
    '        '  advertisercopy = advertisercopy.DefaultView.ToTable("Advertiser", True)
    '        'For Each row As Data.DataRow In advertisercopy.Rows
    '        '    Globals.Ribbons.MSprintExRibbon.tpSelections.UcAdvertiser1.clbSelectAdvertisers.Items.Add(row(0).ToString())
    '        'Next
    '        For Each row As Data.DataRow In advertisercopy.Rows

    '            If Not (Globals.Ribbons.MSprintExRibbon.tpSelections.UcAdvertiser1.clbSelectAdvertisers.Items.Contains(row("Advertiser").ToString())) Then
    '                Globals.Ribbons.MSprintExRibbon.tpSelections.UcAdvertiser1.clbSelectAdvertisers.Items.Add(row("Advertiser").ToString())

    '            End If
    '        Next
    '        clbSelectAdvertisers.Sorted = True
    '    Catch ex As Exception
    '        LogMpsrintExException("Exception occured while loading Advertiser details.Message :" + ex.Message)
    '    End Try
    'End Sub

    Private Sub tbFilterMasterGenres_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbFilterMasterGenres.TextChanged
        Try
            If Not (String.IsNullOrEmpty(tbFilterMasterGenres.Text.Trim())) Then
                clbSelectAdvertisers.Items.Clear()
                For Each row As Data.DataRow In advertisercopy.Rows

                    If row("Advertiser").ToString().ToUpper().StartsWith(tbFilterMasterGenres.Text.Trim().ToUpper()) Then
                        clbSelectAdvertisers.Items.Add(row("Advertiser").ToString())
                    End If

                Next
            ElseIf (tbFilterMasterGenres.Text.Trim().Length = 0) Then

                clbSelectAdvertisers.Items.Clear()
                For Each dr As Data.DataRow In advertisercopy.Rows
                    clbSelectAdvertisers.Items.Add(dr("Advertiser").ToString())
                Next
            End If
            ' clbSelectAdvertisers.Refresh()
            clbSelectAdvertisers.Sorted = True
        Catch ex As Exception

        End Try
    End Sub
End Class
