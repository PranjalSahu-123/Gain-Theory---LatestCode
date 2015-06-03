Public Class frmRearrangePlanChannels

    Private Sub btnRearrange_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRearrange.Click
        Try

            ' If Globals.Ribbons.MSprintExRibbon.reorderedChannels Is Nothing Then
            Globals.Ribbons.MSprintExRibbon.reorderedChannels = New List(Of String)()
            '  End If
            ' Globals.Ribbons.MSprintExRibbon.reorderedChannels.
            For index = 1 To lstChannelSortOrder.Items.Count
                Globals.Ribbons.MSprintExRibbon.reorderedChannels.Add(lstChannelSortOrder.Items(index - 1).ToString())
            Next
            Me.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnMvChnlUP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMvChnlUP.Click
        Try
            Dim itemToMove As Integer = 0
            Dim swapCounter As Integer = 0
            Dim valToMove As String = String.Empty
            Dim swapStr As String = String.Empty
            ' int itemToMove = 0; int swapCounter = 0
            'string valToMove = ""; string swapStr = ""
            itemToMove = Me.lstChannelSortOrder.SelectedIndex

            If itemToMove = 0 Or itemToMove = -1 Then
                Return
            End If

            'if (itemToMove == 0 || itemToMove == -1 )
            '    return;
            swapCounter = itemToMove - 1
            swapStr = Me.lstChannelSortOrder.Items(swapCounter).ToString()
            valToMove = Me.lstChannelSortOrder.SelectedItem.ToString()
            Me.lstChannelSortOrder.Items(swapCounter) = valToMove
            Me.lstChannelSortOrder.Items(itemToMove) = swapStr
            Me.lstChannelSortOrder.SetSelected(swapCounter, True)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnMvChDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMvChDown.Click
        Try
            ' int itemToMove = 0; int swapCounter = 0;
            Dim itemToMove As Integer = 0
            Dim swapCounter As Integer = 0
            Dim valToMove As String = String.Empty
            Dim swapStr As String = String.Empty
            '  string valToMove = ""; string swapStr = "";
            itemToMove = Me.lstChannelSortOrder.SelectedIndex

            'if (itemToMove == this.lstChannelSortOrder.Items.Count - 1 || itemToMove == -1 )
            '    return;

            If itemToMove = lstChannelSortOrder.Items.Count - 1 Or itemToMove = -1 Then
                Return
            End If

            swapCounter = itemToMove + 1
            swapStr = Me.lstChannelSortOrder.Items(swapCounter).ToString()
            valToMove = Me.lstChannelSortOrder.SelectedItem.ToString()
            Me.lstChannelSortOrder.Items(swapCounter) = valToMove
            Me.lstChannelSortOrder.Items(itemToMove) = swapStr
            Me.lstChannelSortOrder.SetSelected(swapCounter, True)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub frmRearrangePlanChannels_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            If Globals.Ribbons.MSprintExRibbon.reorderedChannels Is Nothing Then
                Dim table As Data.DataTable = CType(loSpotSelection.DataSource, Data.DataTable)
                table = table.DefaultView.ToTable(True, "Channel")
                For Each row As Data.DataRow In table.Rows
                    lstChannelSortOrder.Items.Add(row(0).ToString())
                Next
            Else
                For Each row As String In Globals.Ribbons.MSprintExRibbon.reorderedChannels
                    lstChannelSortOrder.Items.Add(row)
                Next
            End If

            'Dim table As Data.DataTable = CType(loSpotSelection.DataSource, Data.DataTable)
            'table = table.DefaultView.ToTable(True, "Channel")
           
        Catch ex As Exception

        End Try
    End Sub
End Class