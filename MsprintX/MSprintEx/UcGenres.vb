Imports System
Imports System.Threading.Tasks
Imports System.Threading
Imports System.Data
Imports System.Linq
Public Class UcGenres
    '  Dim dtGenres As Data.DataTable
    Dim genres As DataTable
    Private Sub clbSelectGenres_ItemCheck(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs)
        'lbSelectedGenres.Refresh()
        'lbSelectedGenres.Items.Add(clbSelectGenres.CheckedItems(0).ToString)
    End Sub

    Private Sub tbFilterMasterGenres_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbFilterMasterGenres.TextChanged
        '  if (String.IsNullOrEmpty(t1.Text.Trim()) == false)
        '{
        '    lb1.Items.Clear();
        '    foreach (string str in list)
        '    {
        '    If (Str.StartsWith(t1.Text.Trim())) Then

        '        {
        '            lb1.Items.Add(str);
        '        }
        '    } 
        '}

        'else if(t1.Text.Trim() == "")
        '{
        '    lb1.Items.Clear();

        '    foreach (string str in list)
        '        {
        '            lb1.Items.Add(str);
        '        }
        '    }                         
        '}       

        If Not (String.IsNullOrEmpty(tbFilterMasterGenres.Text.Trim())) Then
            clbSelectGenres.Items.Clear()
            For Each row As Data.DataRow In genres.Rows

                If row(0).ToString().StartsWith(tbFilterMasterGenres.Text.Trim().ToUpper()) Then
                    clbSelectGenres.Items.Add(row(0).ToString())
                End If

            Next
        ElseIf (tbFilterMasterGenres.Text.Trim().Length = 0) Then

            clbSelectGenres.Items.Clear()
            For Each dr As Data.DataRow In genres.Rows
                clbSelectGenres.Items.Add(dr(0).ToString())
            Next
        End If
        clbSelectGenres.Refresh()
    End Sub

    Private Sub UcGenres_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            'dtGenres = New Data.DataTable("Genres")
            'dtGenres.ReadXmlSchema(AppDomain.CurrentDomain.BaseDirectory + "\\genresschema.xsd")
            'dtGenres.ReadXml(AppDomain.CurrentDomain.BaseDirectory + "\\genres.xml")
            'Dim list As List(Of String) = dtGenres.Rows.CopyTo(.Cast(Of List(Of String))()
            'Dim genres As String() = New String(dtGenres.Rows.Count) {}
            'Dim dr As Data.DataRow() = New Data.DataRow(dtGenres.Rows.Count) {}
            'dtGenres.Rows.CopyTo(dr, 0)
            'genres = Array.ConvertAll(dr, New Converter(Of Data.DataRow, String)(AddressOf DataRowToString))
            genres = Globals.Ribbons.MSprintExRibbon.dtGenres.Copy().DefaultView.ToTable(True, "Name")
            'genres.Columns.RemoveAt(0)
            ' genres.AsEnumerable().Distinct().CopyToDataTable()
            For Each dr As Data.DataRow In genres.Rows
                Dim genre As String = dr(0).ToString().Trim()
                If Not (clbSelectGenres.Items.Contains(genre)) Then
                    clbSelectGenres.Items.Add(genre)
                End If


            Next
            'Parallel.ForEach(Globals.Ribbons.MSprintExRibbon.dtGenres.AsEnumerable(), Sub(dr As DataRow)
            '                                                                              clbSelectGenres.Items.Add(dr(0).ToString())
            '                                                                              ' The more computational work you do here, the greater  
            '                                                                              ' the speedup compared to a sequential foreach loop. 
            '                                                                              'Dim filename As String = System.IO.Path.GetFileName(currentFile)
            '                                                                              'Dim bitmap As New System.Drawing.Bitmap(currentFile)

            '                                                                              'bitmap.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
            '                                                                              'bitmap.Save(System.IO.Path.Combine(newDir, filename))

            '                                                                              '' Peek behind the scenes to see how work is parallelized. 
            '                                                                              '' But be aware: Thread contention for the Console slows down parallel loops!!!

            '                                                                              'Console.WriteLine("Processing {0} on thread {1}", filename, Thread.CurrentThread.ManagedThreadId)
            '                                                                              'close lambda expression and method invocation 
            '                                                                          End Sub)
            clbSelectGenres.Sorted = True
            'DataRow[] dr = new DataRow[dtSource.Rows.Count];
            'dtSource.Rows.CopyTo(dr, 0);
            'double[] dblPrice= Array.ConvertAll(dr, new Converter<DataRow , Double>(DataRowToDouble));
        Catch ex As Exception

        End Try
    End Sub
    Public Shared Function DataRowToString(ByVal drr As Data.DataRow) As String
        Return drr(0).ToString()
    End Function



    Private Sub clbSelectGenres_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles clbSelectGenres.MouseDoubleClick

        If Not (lbSelectedGenres.Items.Contains(clbSelectGenres.SelectedItems(0).ToString())) Then
            lbSelectedGenres.Items.Add(clbSelectGenres.SelectedItems(0).ToString())
        End If


        'lbSelectedGenres.Items.Remove(lbSelectedGenres.SelectedItem)
        lbSelectedGenres.Refresh()
        clbSelectGenres.Refresh()

    End Sub

    Private Sub chbclearall_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbclearall.CheckedChanged

        If chbclearall.Checked Then
            lbSelectedGenres.Items.Clear()
            lbSelectedGenres.Refresh()
        End If

    End Sub

    Private Sub chbSelectAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbSelectAll.CheckedChanged
        lbSelectedGenres.Items.Clear()

        If chbSelectAll.Checked Then
            For index = 0 To clbSelectGenres.Items.Count - 1
                ' clbSelectGenres.SetSelected(index, True)
                lbSelectedGenres.Items.Add(clbSelectGenres.Items(index).ToString())

            Next
        Else
            lbSelectedGenres.Items.Clear()
        End If
        dgshowchannels.DataSource = Nothing
        

    End Sub

    Private Sub lbSelectedGenres_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lbSelectedGenres.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Delete Then
            Try
                Dim strDeleteGenre As String = lbSelectedGenres.SelectedItem
                lbSelectedGenres.Items.Remove(lbSelectedGenres.SelectedItem)
                'For Each item In From p In dtSelectedMarkets Where p.MarketGroup = strDeleteMG
                '    dtSelectedMarkets.RemoveSelectedMarketsRow(item)
                'Next
            Catch ex As Exception
                'Debug.WriteLine(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub clbSelectGenres_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles clbSelectGenres.MouseClick
        Try
            dgshowchannels.DataSource = PopulateChannelsGrid()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub clbSelectGenres_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clbSelectGenres.SelectedIndexChanged
        Try
            dgshowchannels.DataSource = PopulateChannelsGrid()
        Catch ex As Exception

        End Try
    End Sub
    Private Function PopulateChannelsGrid() As DataTable
        Dim channels As DataTable = New DataTable
        channels.Columns.Add("Channels")
        Try

            If Not (clbSelectGenres.SelectedItem Is Nothing) Then
                Dim genre As String = clbSelectGenres.SelectedItem.ToString()
                Dim rows As DataRow() = Globals.Ribbons.MSprintExRibbon.dtGenres.Select("Name='" + genre + "'")

                If rows.Length > 0 Then
                    For Each row As DataRow In rows
                        Dim channel As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("ID='" + row("ChannelCode") + "'")(0)("Name")
                        channels.Rows.Add(channel)
                    Next
                End If

            End If

        Catch ex As Exception
            LogMpsrintExException("Exception occured while populating Channels Grid." + ex.Message)
        End Try
        Return channels
    End Function
End Class
