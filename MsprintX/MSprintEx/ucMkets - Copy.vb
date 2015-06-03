Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Data
'Imports System.Data.DataSetExtensions
Public Class ucMkets
    Dim daMarkets As New METISTableAdapters.MarketsTableAdapter
    Dim WithEvents dtMarkets As METIS.MarketsDataTable
    ' Dim Globals.Ribbons.MSprintExRibbon.dtMarkets1 As System.Data.DataTable
    Friend dtSelectedMarkets As New METIS.SelectedMarketsDataTable
    Dim strSelectedMG As String
    Dim lbClearMG As Boolean = True
    Dim fileList As List(Of String) = New List(Of String)()
    Dim tgDirectoryPath As String = AppDomain.CurrentDomain.BaseDirectory & "\\Masters\TGS"
    Dim mgDirectoryPath As String = AppDomain.CurrentDomain.BaseDirectory & "\\Masters\MGS"
    Private Sub ucMkets_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            If Directory.Exists(mgDirectoryPath) Then
                For index = 0 To Directory.GetFiles(mgDirectoryPath, "*.xml").Count - 1
                    ' fileList.Add(Path.GetFileNameWithoutExtension(Directory.GetFiles(System.IO.Path.GetTempPath() + "\\MGS", "*.xml")(index)))
                    Dim c As XElement = XElement.Load(Directory.GetFiles(mgDirectoryPath, "*.xml")(index))
                    If String.Compare(c.Attribute("type").Value, "Group", True) = 0 Then
                        lbMarketGroup.Items.Add(Path.GetFileNameWithoutExtension(Directory.GetFiles(mgDirectoryPath, "*.xml")(index)))
                    End If

                Next
                lbMarketGroup.Sorted = True
            End If


            ' dtMarkets = daMarkets.GetMarkets
            '' Dim dtMarkets As Data.DataTable = New Data.DataTable("Markets")
            ''  dtMarkets.WriteXml("dtmarks.xml")
            ' 
            '  dtccopy = New Data.DataTable()
            'dtchannels.ReadXmlSchema("C:\\6Jan\\MSprintEx\\MSprintEx\\bin\\Debug\\channelsschema.xsd")
            '  Globals.Ribbons.MSprintExRibbon.dtMarkets1.ReadXmlSchema(AppDomain.CurrentDomain.BaseDirectory + "\\marketschema.xsd")
            ' Globals.Ribbons.MSprintExRibbon.dtMarkets1.ReadXml(AppDomain.CurrentDomain.BaseDirectory + "\\market.xml")
            Globals.Ribbons.MSprintExRibbon.dtMarkets1.Columns.Add("Selected", Type.GetType("System.Boolean"))
            Globals.Ribbons.MSprintExRibbon.dtMarkets1.DefaultView.Sort = "Name ASC"
            ' Globals.Ribbons.MSprintExRibbon.dtMarkets1=Globals.Ribbons.MSprintExRibbon.dtMarkets1.DefaultView.ToTable()
            Globals.Ribbons.MSprintExRibbon.dtMarkets1 = Globals.Ribbons.MSprintExRibbon.dtMarkets1.DefaultView.ToTable()
            dgvMarkets.DataSource = Globals.Ribbons.MSprintExRibbon.dtMarkets1
            dgvMarkets.Columns(0).Visible = False
            dgvMarkets.Columns(2).Visible = False
            dgvSelectedMarkets.DataSource = dtSelectedMarkets
            dgvSelectedMarkets.Columns(0).Visible = False
            dgvSelectedMarkets.Columns(1).Visible = False

            dgvForNewMG.DataSource = Globals.Ribbons.MSprintExRibbon.dtMarkets1
            dgvForNewMG.Columns(0).Visible = False
            dgvForNewMG.Columns(1).ReadOnly = True
            dgvForNewMG.Columns(2).Visible = False

            dgvForNewMG.Columns(1).Width = 250

            flpNewGroup.BringToFront()
            'For index = 0 To Directory.GetFiles(System.IO.Path.GetTempPath() + "\\MGS", "*.xml").Count - 1
            '    ' fileList.Add(Path.GetFileNameWithoutExtension(Directory.GetFiles(System.IO.Path.GetTempPath() + "\\MGS", "*.xml")(index)))
            '    lbMarketGroup.Items.Add(Path.GetFileNameWithoutExtension(Directory.GetFiles(System.IO.Path.GetTempPath() + "\\MGS", "*.xml")(index)))
            'Next
            ' lbMarketGroup.DataSource = fileList
        Catch ex As Exception
            LogMpsrintExException("Exception occured while loading Markets tab.Message" + ex.Message)
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
        c = <mg name=<%= txtGroupName.Text.Trim() %> type="Group">

            </mg>
        For Each dr As METIS.SelectedMarketsRow In arrSelected
            dr.MarketGroup = txtGroupName.Text
            Dim markets As XElement = New XElement("market")
            markets.Value = dr.Market_id
            c.Add(markets)
        Next
        'Dim temppath As String = System.IO.Path.GetTempPath()

        'If Not (Directory.Exists(temppath + "\\MGS")) Then
        '    Directory.CreateDirectory(temppath + "\\MGS")
        'End If


        c.Save(mgDirectoryPath + "\\" + txtGroupName.Text + ".xml")
        dtSelectedMarkets.DefaultView.RowFilter = "MarketGroup = '-1'"
        txtGroupName.Clear()
        chkSelectAll.Checked = False
        'For Each dr As METIS.MarketsRow In dtMarkets.Select("Selected = true")
        '    dr.Selected = False
        'Next
    End Sub

    Private Sub btnSetPlanMG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetPlanMG.Click
        If TabControl1.SelectedTab Is tpMarkets Then
            For Each item As System.Windows.Forms.DataGridViewRow In dgvMarkets.SelectedRows
                If Not (lbPlan.Items.Contains(item.Cells(1).Value)) Then
                    lbPlan.Items.Add(item.Cells(1).Value)
                End If
            Next
        Else
            For Each item As Object In lbMarketGroup.SelectedItems
                If Not (lbPlan.Items.Contains(item.ToString())) Then
                    lbPlan.Items.Add(item.ToString())
                End If
            Next
        End If
    End Sub

    Private Sub btnSetRefMG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetRefMG.Click
        If TabControl1.SelectedTab Is tpMarkets Then
            For Each item As System.Windows.Forms.DataGridViewRow In dgvMarkets.SelectedRows
                If Not (lbRef.Items.Contains(item.Cells(1).Value)) Then
                    lbRef.Items.Add(item.Cells(1).Value)
                End If
            Next
        Else
            For Each item As Object In lbMarketGroup.SelectedItems
                If Not (lbRef.Items.Contains(item.ToString())) Then
                    lbRef.Items.Add(item.ToString())
                End If
            Next
        End If
        ' ClbRefMG.Items.Add(lbMarketGroup.SelectedItem.ToString())
    End Sub

    Private Sub lbMarketGroup_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lbMarketGroup.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Delete Then
            Try
                Dim strDeleteMG As String = lbMarketGroup.SelectedItem
                File.Delete(mgDirectoryPath + "\\" + strDeleteMG + ".xml")
                lbMarketGroup.Items.Remove(lbMarketGroup.SelectedItem)
                For Each item In From p In dtSelectedMarkets Where p.MarketGroup = strDeleteMG
                    dtSelectedMarkets.RemoveSelectedMarketsRow(item)
                Next
            Catch ex As Exception
                Debug.WriteLine(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub chkSelectAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSelectAll.CheckedChanged
        If Not lbClearMG Then Exit Sub
        lbMarketGroup.ClearSelected()

        If chkSelectAll.Checked Then

            For Each row As Windows.Forms.DataGridViewRow In dgvMarkets.Rows
                row.Selected = True
                Try
                    Dim c As XElement
                    ' If e.ColumnIndex <> 0 Then

                    Dim dr As Data.DataRow = Globals.Ribbons.MSprintExRibbon.dtMarkets1.Rows(row.Index)
                    c = <mg name=<%= dr("Name").ToString().Trim() %> type="Single">

                        </mg>
                    ' For Each dr As METIS.SelectedMarketsRow In arrSelected
                    '  dr.MarketGroup = txtGroupName.Text
                    Dim markets As XElement = New XElement("market")
                    markets.Value = dr("ID").ToString()
                    c.Add(markets)
                    ' Next
                    '  Dim temppath As String = System.IO.Path.GetTempPath()

                    'If Not (Directory.Exists(mgDirectoryPath + "\\MGS")) Then
                    '    Directory.CreateDirectory(temppath + "\\MGS")
                    'End If
                    c.Save(mgDirectoryPath + "\\" + dr("Name").ToString().Trim() + ".xml")
                    ' End If
                Catch ex As Exception
                    LogMpsrintExException("Exception occured while selecting all markets." + ex.Message)
                End Try
            Next
        Else
            For Each row As Windows.Forms.DataGridViewRow In dgvMarkets.Rows
                row.Selected = False
            Next
        End If

        'For Each drSelected As METIS.SelectedMarketsRow In dtSelectedMarkets.Select("MarketGroup = '-1'")
        '    dtSelectedMarkets.RemoveSelectedMarketsRow(drSelected)
        'Next
        'For Each drMarket In From p In dtMarkets
        '    drMarket.Selected = chkSelectAll.Checked
        '    If chkSelectAll.Checked Then
        '        If dtSelectedMarkets.FindByMarketGroupMarket_id("-1", drMarket.Market_id) Is Nothing Then
        '            dtSelectedMarkets.AddSelectedMarketsRow("-1", drMarket.Market_id, drMarket.Market_desc)
        '        End If
        '    End If
        'Next
        'dtSelectedMarkets.DefaultView.RowFilter = "MarketGroup = '-1'"
    End Sub

    Private Sub dgvMarkets_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvMarkets.CellClick
        Try
            Dim c As XElement
            If e.ColumnIndex <> 0 Then

                Dim dr As Data.DataRow = Globals.Ribbons.MSprintExRibbon.dtMarkets1.Rows(e.RowIndex)
                c = <mg name=<%= dr("Name").ToString().Trim() %> type="Single">

                    </mg>
                ' For Each dr As METIS.SelectedMarketsRow In arrSelected
                '  dr.MarketGroup = txtGroupName.Text
                Dim markets As XElement = New XElement("market")
                markets.Value = dr("ID").ToString()
                c.Add(markets)
                ' Next
                '  Dim temppath As String = System.IO.Path.GetTempPath()

                'If Not (Directory.Exists(mgDirectoryPath + "\\MGS")) Then
                '    Directory.CreateDirectory(temppath + "\\MGS")
                'End If
                c.Save(mgDirectoryPath + "\\" + dr("Name").ToString().Trim() + ".xml")
            End If
        Catch ex As Exception
            LogMpsrintExException("Exception occured while selecting market and saving market xml." + ex.Message)
        End Try
    End Sub

    Private Sub lbMarketGroup_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbMarketGroup.SelectedIndexChanged
        Try
            Dim mgelement As XElement
            strSelectedMG = CType(sender, Windows.Forms.ListBox).SelectedItem
            If Not strSelectedMG Is Nothing Then
                mgelement = XElement.Load(mgDirectoryPath + "\\" + strSelectedMG + ".xml")
                dtSelectedMarkets.Clear()
                For Each market As XElement In mgelement.Elements
                    dtSelectedMarkets.AddSelectedMarketsRow(strSelectedMG, market.Value, Globals.Ribbons.MSprintExRibbon.dtMarkets1.Select("ID = '" + market.Value + "'")(0)("Name"))
                Next
                dtSelectedMarkets.DefaultView.RowFilter = ""
                dgvSelectedMarkets.DataSource = dtSelectedMarkets

                '    dtSelectedMarkets.DefaultView.RowFilter = "MarketGroup = '" & strSelectedMG & "'"
                '    For Each drMarket In dtMarkets.Rows
                '        drMarket.Selected = False
                '        '  dtSelectedMarkets.AddSelectedMarketsRow(strSelectedMG,
                '        lbClearMG = False
                '        chkSelectAll.Checked = False
                '        lbClearMG = True
                '    Next
                'Else
                '    dtSelectedMarkets.DefaultView.RowFilter = "MarketGroup = '-1'"
            End If
        Catch ex As Exception
            LogMpsrintExException("Exception occured while viewing markets for chosen market group." + ex.Message)
        End Try
        'txtGroupName.Clear()
    End Sub

    Private Sub lbPlan_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbPlan.SelectedIndexChanged

    End Sub

    Private Sub lbRef_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbRef.SelectedIndexChanged

    End Sub

    Private Sub lbPlan_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lbPlan.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Delete Then
            Try
                'Dim strDeleteMG As String = lbMarketGroup.SelectedItem
                For Each item As Object In lbPlan.SelectedItems
                    lbPlan.Items.Remove(item)
                Next

                'For Each item In From p In dtSelectedMarkets Where p.MarketGroup = strDeleteMG
                '    dtSelectedMarkets.RemoveSelectedMarketsRow(item)
                'Next
            Catch ex As Exception
                '  Debug.WriteLine(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub lbRef_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lbRef.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Delete Then
            Try
                ' Dim strDeleteMG As String = lbRef.SelectedItem

                '  lbRef.Items.Remove(lbRef.SelectedItem)
                For Each item As Object In lbRef.SelectedItems
                    lbRef.Items.Remove(item)
                Next
                'For Each item In From p In dtSelectedMarkets Where p.MarketGroup = strDeleteMG
                '    dtSelectedMarkets.RemoveSelectedMarketsRow(item)
                'Next
            Catch ex As Exception
                ' Debug.WriteLine(ex.ToString)
            End Try
        End If
    End Sub

    'Private Sub btnsinglePlanmarket_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsinglePlanmarket.Click
    '    For Each item As Object In lbSingleMarkets.SelectedItems
    '        If Not (lbPlan.Items.Contains(lbSingleMarkets.SelectedItem.ToString())) Then
    '            lbPlan.Items.Add(lbSingleMarkets.SelectedItem.ToString())
    '        End If
    '    Next
    'End Sub

    'Private Sub lbSingleMarkets_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If e.KeyCode = Windows.Forms.Keys.Delete Then
    '        Try
    '            ' Dim strDeleteMG As String = lbRef.SelectedItem
    '            lbSingleMarkets.Items.Remove(lbSingleMarkets.SelectedItem)
    '            'For Each item In From p In dtSelectedMarkets Where p.MarketGroup = strDeleteMG
    '            '    dtSelectedMarkets.RemoveSelectedMarketsRow(item)
    '            'Next
    '        Catch ex As Exception
    '            ' Debug.WriteLine(ex.ToString)
    '        End Try
    '    End If
    'End Sub

    'Private Sub btnSingleRefMarket_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSingleRefMarket.Click
    '    For Each item As Object In lbSingleMarkets.SelectedItems
    '        If Not (lbRef.Items.Contains(lbSingleMarkets.SelectedItem.ToString())) Then
    '            lbRef.Items.Add(lbSingleMarkets.SelectedItem.ToString())
    '        End If
    '    Next
    'End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSelectedMarkets.CellContentClick
        Dim w As String = String.Empty
    End Sub

    Private Sub DataGridView1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvSelectedMarkets.KeyDown
        Dim grid As Windows.Forms.DataGridView = CType(sender, Windows.Forms.DataGridView)
        Dim rows As Windows.Forms.DataGridViewSelectedCellCollection = grid.SelectedCells
        Dim s As String = String.Empty
    End Sub

    Private Sub btnNewGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewGroup.Click
        'flpNewGroup.Visible = Not flpNewGroup.Visible
        If flpNewGroup.Height = 29 Then
            btnNewGroup.Text = "Hide Market Group creation"
            TabControl1.SelectTab(1)
            flpNewGroup.Height = 334
            flpNewGroup.Top = 176
        Else
            btnNewGroup.Text = "Click here to create additional Market Group"
            flpNewGroup.Height = 29
        End If
        flpNewGroup.BringToFront()
    End Sub


    Private Sub dgvForNewMG_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvForNewMG.CellClick
      
    End Sub


    Private Sub TabControl1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 0 Then
            flpNewGroup.Visible = False
        Else
            flpNewGroup.Visible = True
        End If
    End Sub

    Private Sub txtGroupName_Enter(sender As System.Object, e As System.EventArgs) Handles txtGroupName.Enter
        If txtGroupName.Text = "Enter group name here" Then
            txtGroupName.Text = ""
        End If
    End Sub

    Private Sub txtGroupName_Leave(sender As System.Object, e As System.EventArgs) Handles txtGroupName.Leave
        If txtGroupName.Text.Trim = "" Then
            txtGroupName.Text = "Enter group name here"
        End If
    End Sub

    'Private Sub btnAddToGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddToGroup.Click

    'End Sub

    'Private Sub chkAddallMktsToGroup_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAddallMktsToGroup.CheckedChanged

    'End Sub

    Private Sub tlpNewGroup_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles tlpNewGroup.Paint

    End Sub

    Private Sub dgvForNewMG_CellDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvForNewMG.CellDoubleClick
        Try
            Dim c As XElement
            If e.ColumnIndex <> 0 Then
                Dim dr As Data.DataRow = Globals.Ribbons.MSprintExRibbon.dtMarkets1.Rows(e.RowIndex)
                lbMarketGroup.ClearSelected()
                dtSelectedMarkets.DefaultView.RowFilter = "MarketGroup = '-1'"
                Dim s As Boolean = Convert.IsDBNull(dr("Selected"))
                If Convert.IsDBNull(dr("Selected")) Then
                    '  If dtSelectedMarkets.FindByMarketGroupMarket_id("-1", dr("ID").ToString()) Is Nothing Then
                    dtSelectedMarkets.AddSelectedMarketsRow("-1", dr("ID").ToString(), dr("Name").ToString())
                    'Single market changes starts
                    c = <mg name=<%= dr("Name").ToString().Trim() %> type="Single">

                        </mg>
                    ' For Each dr As METIS.SelectedMarketsRow In arrSelected
                    '  dr.MarketGroup = txtGroupName.Text
                    Dim markets As XElement = New XElement("market")
                    markets.Value = dr("ID").ToString()
                    c.Add(markets)
                    ' Next
                    '  Dim temppath As String = System.IO.Path.GetTempPath()

                    'If Not (Directory.Exists(mgDirectoryPath + "\\MGS")) Then
                    '    Directory.CreateDirectory(temppath + "\\MGS")
                    'End If


                    c.Save(mgDirectoryPath + "\\" + dr("Name").ToString().Trim() + ".xml")

                    'If Not (lbSingleMarkets.Items.Contains(dr("Name").ToString().Trim())) Then
                    '    lbSingleMarkets.Items.Add(dr("Name").ToString().Trim())
                    'End If


                    'lbSingleMarkets.Refresh()
                    'Single market changes - ends
                End If
                ' dgvSelectedMarkets.FirstDisplayedScrollingRowIndex = dgvSelectedMarkets.RowCount - 1
            Else
                'Dim filer As String = "MarketGroup = '-1' and Market_id = '" + dr("ID").ToString() + "'"
                'dtSelectedMarkets.RemoveSelectedMarketsRow(dtSelectedMarkets.Select(filer)(0))
                ''   End If
              End If

        Catch ex As Exception
            LogMpsrintExException("Exception occured while adding selected market to group." + ex.Message)
        End Try
    End Sub

    Private Sub btnAddtoGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddtoGroup.Click
        Try
            For Each row As Windows.Forms.DataGridViewRow In dgvForNewMG.SelectedRows
                ' row.Selected = True
                Dim c As XElement
                Dim dr As Data.DataRow = Globals.Ribbons.MSprintExRibbon.dtMarkets1.Rows(row.Index)
                lbMarketGroup.ClearSelected()
                dtSelectedMarkets.DefaultView.RowFilter = "MarketGroup = '-1'"
                Dim s As Boolean = Convert.IsDBNull(dr("Selected"))
                If Convert.IsDBNull(dr("Selected")) Then
                    '  If dtSelectedMarkets.FindByMarketGroupMarket_id("-1", dr("ID").ToString()) Is Nothing Then
                    dtSelectedMarkets.AddSelectedMarketsRow("-1", dr("ID").ToString(), dr("Name").ToString())
                    'Single market changes starts
                    c = <mg name=<%= dr("Name").ToString().Trim() %> type="Single">

                        </mg>
                    ' For Each dr As METIS.SelectedMarketsRow In arrSelected
                    '  dr.MarketGroup = txtGroupName.Text
                    Dim markets As XElement = New XElement("market")
                    markets.Value = dr("ID").ToString()
                    c.Add(markets)
                    ' Next
                    '  Dim temppath As String = System.IO.Path.GetTempPath()

                    'If Not (Directory.Exists(mgDirectoryPath + "\\MGS")) Then
                    '    Directory.CreateDirectory(temppath + "\\MGS")
                    'End If


                    c.Save(mgDirectoryPath + "\\" + dr("Name").ToString().Trim() + ".xml")

                    'If Not (lbSingleMarkets.Items.Contains(dr("Name").ToString().Trim())) Then
                    '    lbSingleMarkets.Items.Add(dr("Name").ToString().Trim())
                    'End If


                    'lbSingleMarkets.Refresh()
                    'Single market changes - ends
                End If
                'Else
                'Dim filer As String = "MarketGroup = '-1' and Market_id = '" + dr("ID").ToString() + "'"
                'dtSelectedMarkets.RemoveSelectedMarketsRow(dtSelectedMarkets.Select(filer)(0))
                'End If
            Next
            ' dgvSelectedMarkets.FirstDisplayedScrollingRowIndex = dgvSelectedMarkets.RowCount - 1
        Catch ex As Exception
            LogMpsrintExException("Exception occured while adding selected markets to group" + ex.Message)
        End Try
    End Sub

    Private Sub chkSelectallMarkets_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSelectallMarkets.CheckedChanged
        Try
            If chkSelectallMarkets.Checked Then

                For Each row As Windows.Forms.DataGridViewRow In dgvForNewMG.Rows
                    row.Selected = True
                    Try
                        Dim c As XElement
                        ' If e.ColumnIndex <> 0 Then

                        Dim dr As Data.DataRow = Globals.Ribbons.MSprintExRibbon.dtMarkets1.Rows(row.Index)
                        c = <mg name=<%= dr("Name").ToString().Trim() %> type="Single">

                            </mg>
                        ' For Each dr As METIS.SelectedMarketsRow In arrSelected
                        '  dr.MarketGroup = txtGroupName.Text
                        Dim markets As XElement = New XElement("market")
                        markets.Value = dr("ID").ToString()
                        c.Add(markets)
                        ' Next
                        '  Dim temppath As String = System.IO.Path.GetTempPath()

                        'If Not (Directory.Exists(mgDirectoryPath + "\\MGS")) Then
                        '    Directory.CreateDirectory(temppath + "\\MGS")
                        'End If
                        c.Save(mgDirectoryPath + "\\" + dr("Name").ToString().Trim() + ".xml")
                        ' End If
                    Catch ex As Exception
                        LogMpsrintExException("Exception occured while selecting all markets." + ex.Message)
                    End Try
                Next
                '  dgvSelectedMarkets.FirstDisplayedScrollingRowIndex = dgvSelectedMarkets.RowCount - 1
            Else
                For Each row As Windows.Forms.DataGridViewRow In dgvForNewMG.Rows
                    row.Selected = False
                Next
                ' dgvSelectedMarkets.DataSource = Nothing
            End If
        Catch ex As Exception

        End Try
    End Sub
End Class
