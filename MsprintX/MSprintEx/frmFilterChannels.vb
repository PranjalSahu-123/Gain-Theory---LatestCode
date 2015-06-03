Imports System
Imports System.Data
Public Class frmFilterChannels
    Protected Friend CurrentChannelCode As String
    Protected Friend CurrentChannelName As String
    Dim CurrentChannel As System.Data.DataRowView

    Private Sub frmFilterChannels_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        SaveSelection()
    End Sub
    Private Sub frmFilterChannels_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim dt As Data.DataTable = Globals.Ribbons.MSprintExRibbon.dtchannels
        'dt.Columns(0).ColumnName = "ChannelCode"
        'dt.Columns(1).ColumnName = "ChannelName"
        'Me.ChannelMasterDDBindingSource.DataSource = dt
        'dgvChannels.DataSource = Me.PlanChannelsBindingSource
        'Dim dtChannelList As Plandata.ChannelMasterDataTable
        'dtChannelList = dt.Copy()
        'dtChannelList.DefaultView.Sort = "ChannelName Asc"
        'Me.ChannelMasterBindingSource.DataSource = dtChannelList
        'Me.Text = "Select master channel for """ & CurrentChannelName & """"
        Me.ChannelMasterDDBindingSource.DataSource = dtChannelMaster.DefaultView
        dgvChannels.DataSource = Me.PlanChannelsBindingSource
        Dim dtChannelList As Plandata.ChannelMasterDataTable
        dtChannelList = ExcelPlan.dtChannelMaster.Copy()
        dtChannelList.DefaultView.Sort = "ChannelName Asc"
        Me.ChannelMasterBindingSource.DataSource = dtChannelList
        Me.Text = "Select master channel for """ & CurrentChannelName & """"
    End Sub

    Private Sub lbChannelMaster_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbChannelMaster.DoubleClick
        SetSelectedChannel()
    End Sub
    Private Sub Search()
        If txtSearch.Text.Length = 0 Then
            ChannelMasterBindingSource.Filter = ""
        Else
            ChannelMasterBindingSource.Filter = "ChannelName like '%" & txtSearch.Text & "%'"
        End If
    End Sub

    Private Sub dgvChannels_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvChannels.SelectionChanged
        CurrentChannel = PlanChannelsBindingSource.Current
        CurrentChannelCode = CurrentChannel.Row.Item("ChannelCode")
        CurrentChannelName = CurrentChannel.Row.Item("ChannelName")
        txtSearch.Text = CurrentChannelName
        Me.Text = "Select master channel for """ & CurrentChannelName & """"
        Search()
        lbChannelMaster.SelectedValue = CurrentChannelCode
        txtSearch.SelectAll()
    End Sub

    Private Sub txtSearch_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSearch.KeyDown
        On Error Resume Next
        If e.KeyCode = Windows.Forms.Keys.Down Then
            Me.lbChannelMaster.SetSelected(Me.lbChannelMaster.SelectedIndex + 1, True)
        ElseIf e.KeyCode = Windows.Forms.Keys.Up Then
            Me.lbChannelMaster.SetSelected(Me.lbChannelMaster.SelectedIndex - 1, True)
        ElseIf e.KeyCode = Windows.Forms.Keys.Enter Then
            SetSelectedChannel()
        ElseIf e.KeyCode = Windows.Forms.Keys.End Then
            Me.lbChannelMaster.SetSelected(Me.lbChannelMaster.Items.Count - 1, True)
        ElseIf e.KeyCode = Windows.Forms.Keys.Home Then
            Me.lbChannelMaster.SetSelected(0, True)
        End If
    End Sub

    Private Sub txtSearch_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged
        Search()
    End Sub

    Private Sub lbChannelMaster_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lbChannelMaster.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            SetSelectedChannel()
        End If
    End Sub
    Private Sub SetSelectedChannel()
        Dim SelectedValue As String
        SelectedValue = lbChannelMaster.SelectedValue
        If Not SelectedValue Is Nothing Then
            CurrentChannel.Row.Item("ChannelCode") = SelectedValue
            dgvChannels.Select()
        End If
    End Sub

    Private Sub SaveSelection()
        Dim dtChanges As Plandata.PlanChannelsDataTable = dtPlanChannels.GetChanges(DataRowState.Modified)
        If Not dtChanges Is Nothing Then
            For Each dr As Plandata.PlanChannelsRow In dtChanges
                If dr.ChannelCode <> "000" Then
                    daChannelMap.Insert(dr.ChannelCode, dr.ChannelName)
                    dr.AcceptChanges()
                End If
            Next
        End If
    End Sub

    Private Sub lbChannelMaster_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbChannelMaster.SelectedIndexChanged

    End Sub
End Class