Public Class ucAvgTVRMGSelection

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Try

            If rbBreakCount.Checked Then
                If Globals.Ribbons.MSprintExRibbon.xecelTable.Columns.Contains("Total available Ads") Then
                    Dim col As Data.DataColumn = Globals.Ribbons.MSprintExRibbon.xecelTable.Columns("Total available Ads")
                    col.ColumnName = "Total available breaks" 'Total available breaks
                    Globals.Ribbons.MSprintExRibbon.xecelTable.AcceptChanges()
                End If
                For Each row As Data.DataRow In Globals.Ribbons.MSprintExRibbon.xecelTable.Rows
                    Dim filter As String = String.Format("GUID='{0}' and MG='{1}'", row("GUID").ToString(), cbSelectMG.Text.Trim())
                    Dim rows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Select(filter)
                    row("MG") = rows(0)("MG").ToString()
                    row("0TVR Spots") = rows(0)("0TVR Spots").ToString()
                    row("Avg TVR") = rows(0)("Avg TVR").ToString()
                    row("Std Deviation") = rows(0)("Std Deviation").ToString()
                    row("Total available breaks") = rows(0)("Total available breaks").ToString()
                    '  row("0 to m - 2s") = rows(0)("Break 0 to m - 2s").ToString() 'Break 0 to m - 2s
                    '  row("(m - 2s) to (m - s)") = rows(0)("Break (m - 2s) to (m - s)").ToString() 'Break (m - 2s) to (m - s)
                    row("(Avg TVR- SD) to  Avg TVR") = rows(0)("Break (m - s) to m").ToString() 'Break (m - s) to m 
                    row("Avg TVR to (Avg TVR + SD)") = rows(0)("Break m to (m + s)").ToString() 'Break m to (m + s) 
                    ' row("(m + s) to (m + 2s)") = rows(0)("Break (m + s) to (m + 2s)").ToString() 'Break (m + s) to (m + 2s)
                    'row("> (m + 2s)") = rows(0)("Break > (m + 2s)").ToString() 'Break > (m + 2s) 
                Next

            Else
                If Globals.Ribbons.MSprintExRibbon.xecelTable.Columns.Contains("Total available breaks") Then
                    Dim col As Data.DataColumn = Globals.Ribbons.MSprintExRibbon.xecelTable.Columns("Total available breaks")
                    col.ColumnName = "Total available Ads"
                    Globals.Ribbons.MSprintExRibbon.xecelTable.AcceptChanges()
                End If
                For Each row As Data.DataRow In Globals.Ribbons.MSprintExRibbon.xecelTable.Rows
                    Dim filter As String = String.Format("GUID='{0}' and MG='{1}'", row("GUID").ToString(), cbSelectMG.Text.Trim())
                    Dim rows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Select(filter)
                    row("MG") = rows(0)("MG").ToString()
                    row("0TVR Spots") = rows(0)("0TVR Spots").ToString()
                    row("Avg TVR") = rows(0)("Avg TVR").ToString()
                    row("Std Deviation") = rows(0)("Std Deviation").ToString()
                    row("Total available Ads") = rows(0)("Total available Ads").ToString()
                    ' row("0 to m - 2s") = rows(0)("Ad 0 to m - 2s").ToString() 'Break 0 to m - 2s
                    ' row("(m - 2s) to (m - s)") = rows(0)("Ad (m - 2s) to (m - s)").ToString() 'Break (m - 2s) to (m - s)
                    row("(Avg TVR- SD) to  Avg TVR") = rows(0)("Ad (m - s) to m").ToString() 'Break (m - s) to m 
                    row("Avg TVR to (Avg TVR + SD)") = rows(0)("Ad m to (m + s)").ToString() 'Break m to (m + s) 
                    '  row("(m + s) to (m + 2s)") = rows(0)("Ad (m + s) to (m + 2s)").ToString() 'Break (m + s) to (m + 2s)
                    ' row("> (m + 2s)") = rows(0)("Ad > (m + 2s)").ToString() 'Break > (m + 2s) 
                Next

              

            End If

            'For Each row As Data.DataRow In Globals.Ribbons.MSprintExRibbon.xecelTable.Rows
            '    Dim filter As String = String.Format("GUID='{0}' and MG='{1}'", row("GUID").ToString(), cbSelectMG.Text.Trim())
            '    Dim rows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Select(filter)
            '    row("MG") = rows(0)("MG").ToString()
            '    row("Avg TVR") = rows(0)("Avg TVR").ToString()
            '    row("Std Deviation") = rows(0)("Std Deviation").ToString()
            '    row("Total available breaks") = rows(0)("Total available breaks").ToString()
            '    row("0 to m - 2s") = rows(0)("0 to m - 2s").ToString()
            '    row("(m - 2s) to (m - s)") = rows(0)("(m - 2s) to (m - s)").ToString()
            '    row("(m - s) to m") = rows(0)("(m - s) to m").ToString()
            '    row("m to (m + s)") = rows(0)("m to (m + s)").ToString()
            '    row("(m + s) to (m + 2s)") = rows(0)("(m + s) to (m + 2s)").ToString()
            '    row("> (m + 2s)") = rows(0)("> (m + 2s)").ToString()
            'Next
            loSpotSelection.SetDataBinding(Globals.Ribbons.MSprintExRibbon.xecelTable)

        Catch ex As Exception
            LogMpsrintExException("Exception occured while displaying Average TVR details.Message :" + ex.Message)
        End Try
    End Sub

    Private Sub cbSelectMG_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSelectMG.SelectedIndexChanged
        Try

        Catch ex As Exception

        End Try
    End Sub

    Private Sub ucAvgTVRMGSelection_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                ' Dim doc2 As XmlDocument = New XmlDocument()
                '  doc2.Load()
                '  TG_MGElement.Add(XElement.Load(Path.GetTempPath() + "\\TGS\\" + plantgname + ".xml"))
                cbSelectMG.Items.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim())
                ' TG_MGElement.Add(mg)
                '  allMGElement.Add(mg.Elements())
            Next
            cbSelectMG.Sorted = True
        Catch ex As Exception

        End Try
    End Sub

End Class
