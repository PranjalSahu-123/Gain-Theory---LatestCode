Imports System.Data

Public Class ucChannels
    Friend dtccopy As Data.DataTable
    Dim genres As DataTable
    Private Sub ucChannels_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' dtchannels = New Data.DataTable("Channels")
        genres = Globals.Ribbons.MSprintExRibbon.dtGenres.Copy().DefaultView.ToTable(True, "Name")
        'genres.Columns.RemoveAt(0)
        ' genres.AsEnumerable().Distinct().CopyToDataTable()
        For Each dr As Data.DataRow In genres.Rows
            Dim genre As String = dr(0).ToString().Trim()
            If Not (cbGenre.Items.Contains(genre)) Then
                cbGenre.Items.Add(genre)
            End If
        Next
        cbGenre.Sorted = True
        dtccopy = New Data.DataTable()
        ' ''dtchannels.ReadXmlSchema("C:\\6Jan\\MSprintEx\\MSprintEx\\bin\\Debug\\channelsschema.xsd")
        ''dtchannels.ReadXmlSchema(AppDomain.CurrentDomain.BaseDirectory + "\\channelsschema.xsd")
        ''dtchannels.ReadXml(AppDomain.CurrentDomain.BaseDirectory + "\\channels.xml")
        dtccopy = Globals.Ribbons.MSprintExRibbon.dtchannels.Copy()
        dtccopy.Columns.RemoveAt(0)
        '  Dim channels As String() = New String(dtccopy.Rows.Count) {}
        'Dim dr As Data.DataRow() = New Data.DataRow(dtccopy.Rows.Count) {}
        'dtccopy.Rows.CopyTo(dr, 0)
        '  channels = Array.ConvertAll(dr, New Converter(Of Data.DataRow, String)(AddressOf DataRowToString))
        ' Dim list As List(Of String) = dtccopy.Rows.Cast(Of List(Of String))()
        'clbSelectChannels.Items.AddRange(Array.ConvertAll(dr, New Converter(Of Data.DataRow, String)(AddressOf DataRowToString)))
        For Each dr As Data.DataRow In dtccopy.Rows
            clbSelectChannels.Items.Add(dr(0).ToString())
        Next
        clbSelectChannels.Sorted = True
    End Sub

    Private Sub clbSelectChannels_ItemCheck(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs)
        'clbSelectChannels.Refresh()
        'lbSelectedChannels.Items.Add(clbSelectChannels.CheckedItems(0).ToString())
    End Sub
    Public Shared Function DataRowToString(ByVal drr As Data.DataRow) As String
        Return drr(0).ToString()
    End Function

    Private Sub clbSelectChannels_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles clbSelectChannels.MouseDoubleClick


        If Not (lbSelectedChannels.Items.Contains(clbSelectChannels.SelectedItems(0).ToString())) Then
            lbSelectedChannels.Items.Add(clbSelectChannels.SelectedItems(0).ToString())
        End If


        ' lbSelectedChannels.Items.Remove(lbSelectedChannels.SelectedItem)
        lbSelectedChannels.Refresh()
        clbSelectChannels.Refresh()

    End Sub

    Private Sub chbClearallchannels_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        lbSelectedChannels.Items.Clear()
        lbSelectedChannels.Refresh()
    End Sub

    Private Sub tbFilterMasterGenres_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbFilterMasterGenres.TextChanged
        If Not (String.IsNullOrEmpty(tbFilterMasterGenres.Text.Trim())) Then
            clbSelectChannels.Items.Clear()
            For Each row As Data.DataRow In dtccopy.Rows

                If row(0).ToString().ToUpper().StartsWith(tbFilterMasterGenres.Text.Trim().ToUpper()) Then
                    clbSelectChannels.Items.Add(row(0).ToString())
                End If

            Next
        ElseIf (tbFilterMasterGenres.Text.Trim().Length = 0) Then

            clbSelectChannels.Items.Clear()
            For Each dr As Data.DataRow In dtccopy.Rows
                clbSelectChannels.Items.Add(dr(0).ToString())
            Next
        End If
        clbSelectChannels.Refresh()
    End Sub

    Private Sub lbSelectedChannels_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lbSelectedChannels.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Delete Then
            Try
                Dim strDeleteGenre As String = lbSelectedChannels.SelectedItem
                lbSelectedChannels.Items.Remove(lbSelectedChannels.SelectedItem)
                'For Each item In From p In dtSelectedMarkets Where p.MarketGroup = strDeleteMG
                '    dtSelectedMarkets.RemoveSelectedMarketsRow(item)
                'Next
            Catch ex As Exception
                'Debug.WriteLine(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub btnClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearAll.Click
        lbSelectedChannels.Items.Clear()
        lbSelectedChannels.Refresh()
    End Sub

    Private Sub ClbSelectAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClbSelectAll.CheckedChanged

        If ClbSelectAll.Checked Then
            lbSelectedChannels.Items.Clear()
            For index = 0 To clbSelectChannels.Items.Count - 1
                'clbSelectChannels.SetSelected(index, True)
                lbSelectedChannels.Items.Add(clbSelectChannels.Items(index))

            Next
        Else
            lbSelectedChannels.Items.Clear()
        End If


    End Sub
    Private Sub cbGenre_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbGenre.SelectedIndexChanged
        Try
            clbSelectChannels.Items.Clear()
            PopulateChannelsListBox()
        Catch ex As Exception

        End Try
    End Sub
    Private Function PopulateChannelsListBox()
        '   Dim channels As DataTable = New DataTable
        '  channels.Columns.Add("Channels")
        Try

            If Not (cbGenre.SelectedItem Is Nothing) Then
                Dim genre As String = cbGenre.SelectedItem.ToString()
                Dim rows As DataRow() = Globals.Ribbons.MSprintExRibbon.dtGenres.Select("Name='" + genre + "'")

                If rows.Length > 0 Then
                    For Each row As DataRow In rows
                        Dim channel As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("ID='" + row("ChannelCode") + "'")(0)("Name")
                        ' channels.Rows.Add(channel)
                        clbSelectChannels.Items.Add(channel)
                    Next
                End If

            End If

        Catch ex As Exception
            LogMpsrintExException("Exception occured while populating Channels List Box." + ex.Message)
        End Try
        ' Return channels
    End Function

    Private Sub lbChosenChannels_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbChosenChannels.Click

    End Sub

    Private Sub lbSelectedChannels_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbSelectedChannels.SelectedIndexChanged

    End Sub

    Private Sub lbSelectChannels_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbSelectChannels.Click

    End Sub

    Private Sub clbSelectChannels_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clbSelectChannels.SelectedIndexChanged

    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class
