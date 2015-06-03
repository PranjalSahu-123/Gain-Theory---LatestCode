Public Class ucVariant
    Dim variantcopy As Data.DataTable
    Private Sub btnClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearAll.Click
        lbSelectedVariants.Items.Clear()
    End Sub

    Private Sub lbSelectedVariants_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lbSelectedVariants.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Delete Then
            Try
                Dim strDeleteGenre As String = lbSelectedVariants.SelectedItem
                lbSelectedVariants.Items.Remove(lbSelectedVariants.SelectedItem)
                'For Each item In From p In dtSelectedMarkets Where p.MarketGroup = strDeleteMG
                '    dtSelectedMarkets.RemoveSelectedMarketsRow(item)
                'Next
            Catch ex As Exception
                'Debug.WriteLine(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub lbSelectedVariants_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbSelectedVariants.SelectedIndexChanged

    End Sub

    Private Sub clbSelectVariant_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clbSelectVariant.SelectedIndexChanged

    End Sub

    Private Sub clbSelectVariant_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles clbSelectVariant.MouseDoubleClick
        If Not (lbSelectedVariants.Items.Contains(clbSelectVariant.SelectedItem.ToString())) Then
            lbSelectedVariants.Items.Add(clbSelectVariant.SelectedItems(0).ToString())
        End If


        ' lbSelectedChannels.Items.Remove(lbSelectedChannels.SelectedItem)
        lbSelectedVariants.Refresh()
        clbSelectVariant.Refresh()
        clbSelectVariant.Sorted = True
        lbSelectedVariants.Sorted = True
    End Sub

    Private Sub tbFilterMasterGenres_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbFilterMasterGenres.TextChanged
        Try
            If Not (String.IsNullOrEmpty(tbFilterMasterGenres.Text.Trim())) Then
                clbSelectVariant.Items.Clear()
                For Each row As Data.DataRow In variantcopy.Rows

                    If row("Variant").ToString().ToUpper().StartsWith(tbFilterMasterGenres.Text.Trim().ToUpper()) Then
                        clbSelectVariant.Items.Add(row("Variant").ToString())
                    End If

                Next
            ElseIf (tbFilterMasterGenres.Text.Trim().Length = 0) Then

                clbSelectVariant.Items.Clear()
                For Each dr As Data.DataRow In variantcopy.Rows
                    clbSelectVariant.Items.Add(dr("Variant").ToString())
                Next
            End If
            clbSelectVariant.Refresh()
            clbSelectVariant.Sorted = True
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ClbSelectAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClbSelectAll.CheckedChanged
        Try
            If ClbSelectAll.Checked Then
                lbSelectedVariants.Items.Clear()
                For index = 0 To clbSelectVariant.Items.Count - 1
                    'clbSelectChannels.SetSelected(index, True)
                    lbSelectedVariants.Items.Add(clbSelectVariant.Items(index))

                Next
                lbSelectedVariants.Sorted = True
            Else
                lbSelectedVariants.Items.Clear()
            End If
        Catch ex As Exception

        End Try
    End Sub

    'Private Sub ucVariant_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    Try
    '        variantcopy = Globals.Ribbons.MSprintExRibbon.bslMasterTable.Copy()
    '        '  variantcopy = variantcopy.DefaultView.ToTable("Variant", True)
    '        For Each row As Data.DataRow In variantcopy.Rows

    '            If Not (Globals.Ribbons.MSprintExRibbon.tpSelections.UcVariant1.clbSelectVariant.Items.Contains(row("Variant").ToString())) Then
    '                Globals.Ribbons.MSprintExRibbon.tpSelections.UcVariant1.clbSelectVariant.Items.Add(row("Variant").ToString())
    '            End If

    '        Next
    '        clbSelectVariant.Sorted = True

    '    Catch ex As Exception
    '        LogMpsrintExException("Exception occured while loading Variant user control.Message :" + ex.Message)

    '    End Try
    'End Sub
End Class
