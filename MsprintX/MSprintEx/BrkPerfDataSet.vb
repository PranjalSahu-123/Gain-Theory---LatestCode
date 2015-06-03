Partial Class BrkPerfDataSet
    Partial Class ExistingLogDataTable

        Private Sub ExistingLogDataTable_ColumnChanged(ByVal sender As Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanged
            If (e.Column.ColumnName = "Date") Then
                e.Row("SpotDay") = Convert.ToDateTime(e.ProposedValue).ToString("ddd")
            End If
        End Sub

    End Class

End Class

Namespace BrkPerfDataSetTableAdapters

    Partial Public Class ProgRFMtTableAdapter
        Private inCommand As String = "SELECT        Channel, Date, TimeFrom, TimeTo, CommercialName, Cost, AP, TA, Duration, Program, ChannelName, Market, Product, Brand, Variant, Advertiser, Genre, Language, " & _
                                                    "FullString, IsSelected " & _
                                                    "FROM ProgRFMt " & _
                                                    "WHERE        (Channel = @ChannelCode) AND (Date BETWEEN @DateFrom AND @DateTo) " & _
                                                    "AND  (SUBSTRING(DATENAME(dw, Date), 1, 3) IN (@Days)) AND ((TimeFrom BETWEEN @SearchFrom AND @SearchTo) OR (TimeTo BETWEEN @SearchFrom AND @SearchTo))"
        Private Sub _adapter_RowUpdated(ByVal sender As Object, ByVal e As System.Data.SqlServerCe.SqlCeRowUpdatedEventArgs) Handles _adapter.RowUpdated
            'BreakPerformance.SetUpdateProgress()
        End Sub

        'Public Function GetBreaksForChannelAndDays(ByVal ChannelCode As String, ByVal inDays As String, ByVal SearchFrom As Date, ByVal SearchTo As Date) As BrkPerfDataSet.ProgRFMtDataTable
        '    Dim command As New System.Data.SqlServerCe.SqlCeCommand(inCommand.Replace("@Days", inDays) _
        '                                                            .Replace("@ChannelCode", "'" & ChannelCode & "'") _
        '                                                            .Replace("@SearchFrom", "'" & SearchFrom.ToString() & "'") _
        '                                                            .Replace("@SearchTo", "'" & SearchTo.ToString() & "'") _
        '                                                            , Me.Connection)
        '    Me.Adapter.SelectCommand = command
        '    Dim dataTable As BrkPerfDataSet.ProgRFMtDataTable = New BrkPerfDataSet.ProgRFMtDataTable
        '    Me.Adapter.Fill(dataTable)
        '    Return dataTable
        'End Function

        Public Overloads Function GetBreaksForChannelAndDays(ByVal ChannelCode As String, ByVal inDays As String, ByVal SearchFrom As Date, ByVal SearchTo As Date, ByVal DateFrom As Date, ByVal DateTo As Date) As BrkPerfDataSet.ProgRFMtDataTable
            Me.Adapter.SelectCommand = Me.CommandCollection(2)
            Dim cmdString As String
            cmdString = inCommand.Replace("@Days", inDays)
            Dim command As String = inCommand.Replace("@Days", inDays) _
                                                                    .Replace("@ChannelCode", "'" & ChannelCode & "'") _
                                                                    .Replace("@SearchFrom", "'" & SearchFrom.ToString() & "'") _
                                                                    .Replace("@SearchTo", "'" & SearchTo.ToString() & "'") _
                                                                    .Replace("@DateFrom", "'" & DateFrom.ToString() & "'") _
                                                                    .Replace("@DateTo", "'" & DateTo.ToString() & "'")
            System.Diagnostics.Debug.Print(command)
            Me.Adapter.SelectCommand.CommandText = cmdString
            If (ChannelCode Is Nothing) Then
                Throw New Global.System.ArgumentNullException("ChannelCode")
            Else
                Me.Adapter.SelectCommand.Parameters(0).Value = CType(ChannelCode, String)
            End If
            Me.Adapter.SelectCommand.Parameters(1).Value = CType(SearchFrom, Date)
            Me.Adapter.SelectCommand.Parameters(2).Value = CType(SearchTo, Date)
            Me.Adapter.SelectCommand.Parameters(3).Value = CType(DateFrom, Date)
            Me.Adapter.SelectCommand.Parameters(4).Value = CType(DateTo, Date)
            Dim dataTable As BrkPerfDataSet.ProgRFMtDataTable = New BrkPerfDataSet.ProgRFMtDataTable
            Me.Adapter.Fill(dataTable)
            Return dataTable
        End Function

    End Class
End Namespace
