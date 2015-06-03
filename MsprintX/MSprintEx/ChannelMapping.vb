Public Class ChannelMapping
    Dim ChannelCells As Excel.Range
    Dim ChannelColumn As Excel.ListColumn
    Dim planchannels As Data.DataTable
    Friend masterchannels As Data.DataTable
    Private Sub btnChannels_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChannels.Click
        RaiseEvent ShowChannels_Click()
    End Sub
    Public Event ShowChannels_Click()
    Public Event ShowMoreChannels()

    Private Sub dgvChannels_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvChannels.CellClick
        If e.RowIndex < 0 OrElse Not e.ColumnIndex = _
            dgvChannels.Columns("More").Index Then Return
        RaiseEvent ShowMoreChannels()
    End Sub
    Friend Sub LoadTAMChannels()
        If dtChannelMaster.Count = 0 Then

            Dim drChannelMaster As Plandata.ChannelMasterRow
            'Dim ChannelCode, ChannelName As String
            'Dim daChannelsMETIS As New METISTableAdapters.CHANNEL_MASTERTableAdapter
            'Dim dtChannelsMETIS As METIS.CHANNEL_MASTERDataTable = daChannelsMETIS.GetChannels
            For Each drChannelMETIS As Data.DataRow In Globals.Ribbons.MSprintExRibbon.dtchannels.Rows
                drChannelMaster = dtChannelMaster.NewChannelMasterRow
                drChannelMaster.ChannelCode = drChannelMETIS("ID")
                drChannelMaster.ChannelName = drChannelMETIS("Name")
                dtChannelMaster.AddChannelMasterRow(drChannelMaster)
            Next
            '    'Dim appPath As String = System.AppDomain.CurrentDomain.BaseDirectory() 'Globals.ThisWorkbook.appPath
            '    ''Using r As StreamReader = New StreamReader(appPath & "STATION.TV")
            '    'Using r As StreamReader = New StreamReader(appPath & "\TAMChannelList.log")

            '    '    Dim line As String

            '    '    line = r.ReadLine

            '    '    Do While (Not line Is Nothing)
            '    '        ChannelCode = line.Substring(0, 3)
            '    '        ChannelName = line.Substring(3).Trim
            '    '        If ChannelName.Length > 0 Then
            '    '            drChannelMaster = dtChannelMaster.NewChannelMasterRow
            '    '            drChannelMaster.ChannelCode = ChannelCode
            '    '            drChannelMaster.ChannelName = ChannelName
            '    '            dtChannelMaster.AddChannelMasterRow(drChannelMaster)
            '    '        End If
            '    '        line = r.ReadLine
            '    '    Loop
            '    'End Using

            drChannelMaster = dtChannelMaster.NewChannelMasterRow
            drChannelMaster.ChannelCode = "000"
            drChannelMaster.ChannelName = " - - Select - - "
            dtChannelMaster.AddChannelMasterRow(drChannelMaster)
            dtChannelMaster.AcceptChanges()
            dtChannelMaster.DefaultView.Sort = "ChannelName Asc"
            '   masterchannels = New Data.DataTable()
            '  masterchannels = Globals.Ribbons.MSprintExRibbon.dtchannels
            ' masterchannels.Columns(0).ColumnName = "ChannelCode"
            ' masterchannels.Columns(1).ColumnName = "ChannelName"
            'masterchannels.DefaultView.Sort = "ChannelName Asc"
            'With ChannelMasterBindingSource
            '    .DataSource = dtChannelMaster.DefaultView
            'End With
        End If
        With ChannelMasterBindingSource
            .DataSource = dtChannelMaster.DefaultView
        End With
        'logTaskPane.scMain.Panel2Collapsed = False
        'If Not logTaskPane.showingChannels Then logTaskPane.showChannelMapping(True)

    End Sub
    Friend Sub LoadPlanChannels()
        dtPlanChannels = New Plandata.PlanChannelsDataTable
        Dim drPlanChannel As Plandata.PlanChannelsRow
        '  dtPlanChannels.ChannelCodeColumn.AllowDBNull = True
        planchannels = New Data.DataTable()
        planchannels.Columns.Add("ChannelCode")
        planchannels.Columns.Add("ChannelName")
        Dim currChannelCode As String = "000"
        ChannelColumn = loSpotSelection.ListColumns("Channel")
        ChannelCells = ChannelColumn.DataBodyRange
        For Each ChannelCell As Excel.Range In ChannelCells
            If Not SubtotalRows Is Nothing Then

                If Not loPlanData Is Nothing Then
                    If Not loPlanData.Application.Intersect(ChannelCell, SubtotalRows) Is Nothing Then Continue For
                End If

            End If
            'Dim dr As Data.DataRow = planchannels.NewRow()
            'dr("ChannelName") = ChannelCell.Value
            If dtPlanChannels.Select("ChannelName = '" & ChannelCell.Value & "'").Length > 0 Then Continue For
            drPlanChannel = dtPlanChannels.NewPlanChannelsRow
            drPlanChannel.ChannelName = ChannelCell.Value
            currChannelCode = GetChannelCodeFromMaster(ChannelCell.Value)
            If currChannelCode = "000" Then
                currChannelCode = GetChannelCodeFromMapping(ChannelCell.Value)
            End If
            drPlanChannel.ChannelCode = currChannelCode
            ' drPlanChannel.ChannelName = "Choose"
            dtPlanChannels.AddPlanChannelsRow(drPlanChannel)
            ' planchannels.Rows.Add(dr)
        Next
        dtPlanChannels.AcceptChanges()
        dtPlanChannels.DefaultView.Sort = "ChannelName Asc"
        planchannels.DefaultView.Sort = "ChannelName ASC"
        With PlanChannelsBindingSource
            .DataSource = dtPlanChannels.DefaultView
        End With
    End Sub

    Private Function GetChannelCodeFromMapping(ByVal ChannelName As String) As String
        dtChannelMap = daChannelMap.GetChannels(ChannelName)
        If dtChannelMap.Count > 0 Then
            GetChannelCodeFromMapping = dtChannelMap(0).TAMChannelCode
        Else
            GetChannelCodeFromMapping = "000"
        End If
    End Function

    Private Function GetChannelCodeFromMaster(ByVal ChannelName As String) As String
        ' Dim drChannelMaster() As Plandata.ChannelMasterRow
        Dim drChannelMaster() As Data.DataRow
        drChannelMaster = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name = '" & ChannelName & "'")
        If drChannelMaster.Length > 0 Then
            GetChannelCodeFromMaster = drChannelMaster(0)(0).ToString()
        Else
            GetChannelCodeFromMaster = "000"
        End If
    End Function

    Private Sub dgvChannels_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvChannels.CellContentClick

    End Sub

    Private Sub ChannelMapping_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadTAMChannels()
        LoadPlanChannels()
    End Sub

    Private Sub dgvChannels_DataError(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvChannels.DataError

    End Sub

    Private Sub PlanChannelsBindingSource_CurrentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlanChannelsBindingSource.CurrentChanged

    End Sub
End Class
