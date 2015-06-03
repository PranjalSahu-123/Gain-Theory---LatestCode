Imports System
Imports System.Windows.Forms
Imports System.Data
Imports System.Threading.Tasks

Public Class ucSpotSelection


    Private Sub btnPushOne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPushOne.Click
        'for DgvToFor Each r As DataGridViewRow In DGVFrom.SelectedRows
        '   Dim Drto As DataRow = DsTo.Tables(0).NewRow
        '   For c As Integer = 0 To DGVFrom.ColumnCount - 1
        '       Drto(c) = r.Cells(c).Value
        '   Next
        '   DsTo.Tables(0).Rows.Add(Drto)
        '   DGVFrom.Rows.Remove(r)
        '  Dim avaipots As DataTable = New DataTable()
        Try
            HideErrorLabel()
            RemoveSelectedSpots()
           
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnPushAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub btnPushOneToSelected_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPushOneToSelected.Click
        Try
            HideErrorLabel()
            If cbWeeks.Text.Equals("All") And cbWeeks.Enabled Then
                ' MessageBox.Show("Please choose a week to select spot(s)")
                ShowErrorLabel("Please choose a week to select spot(s)")
            Else

                Dim selecpots As DataTable = New DataTable()
                Dim fil As String = fil = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
                If cbWeeks.Text.Equals("All") Then
                    fil = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
                Else

                    fil = String.Format("GUID='{0}' and WeekNum={1}", Globals.Ribbons.MSprintExRibbon.currentLineItem, Int32.Parse(cbWeeks.Text).ToString())
                End If


                Dim rows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Select(fil)
                selecpots = rows.CopyToDataTable()
                'If rows.Count > 0 Then

                '    selecpots = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Clone()
                '    Parallel.For(0, rows.Count - 1, Sub(i)
                '                                        selecpots.ImportRow(rows(i))
                '                                    End Sub)
                'End If
                ' = DirectCast(dgSelectedSpotsGrid.DataSource, DataTable)

                Dim dt As DataTable = CType(loSpotSelection.DataSource, DataTable)
                Dim count As Integer
                Try
                    If dt.Columns.Contains("Total Spots") Then
                        Dim filter As String = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
                        count = Convert.ToInt32(dt.Select(filter)(0)("Total Spots").ToString())
                    Else
                        Dim filter As String = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
                        count = Convert.ToInt32(dt.Select(filter)(0)("Week " & Int32.Parse(cbWeeks.Text)).ToString())

                    End If
                Catch ex As Exception
                    Dim filter As String = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
                    If Globals.Ribbons.MSprintExRibbon.xecelTable.Columns.Contains("Total Spots") Then
                        count = Convert.ToInt32(Globals.Ribbons.MSprintExRibbon.xecelTable.Select(filter)(0)("Total Spots").ToString())
                    Else
                        count = Convert.ToInt32(Globals.Ribbons.MSprintExRibbon.xecelTable.Select(filter)(0)("Week " & Int32.Parse(cbWeeks.Text)).ToString())

                    End If
                End Try

                If count > selecpots.Rows.Count And (count - selecpots.Rows.Count) >= dgvAvailableSpotsGrid.SelectedRows.Count Then
                    For Each row As DataGridViewRow In dgvAvailableSpotsGrid.SelectedRows
                        Dim dr As DataRow = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.NewRow()
                        Dim dr1 As DataRow = selecpots.NewRow()
                        dr("GUID") = row.Cells.Item("GUID").Value
                        dr1("GUID") = row.Cells.Item("GUID").Value
                        dr("Spot") = row.Cells.Item("Spot").Value
                        dr1("Spot") = row.Cells.Item("Spot").Value
                        dr("Start Date") = row.Cells.Item("Start Date").Value
                        dr1("Start Date") = row.Cells.Item("Start Date").Value
                        dr("End Date") = row.Cells.Item("End Date").Value
                        dr1("End Date") = row.Cells.Item("End Date").Value
                        dr("WeekNum") = row.Cells.Item("WeekNum").Value
                        dr1("WeekNum") = row.Cells.Item("WeekNum").Value
                        dr("Channel") = row.Cells.Item("Channel").Value
                        dr1("Channel") = row.Cells.Item("Channel").Value
                        '  Dim spotrow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.GetSpotRow(spots(index).Spot)
                        dr("Date") = row.Cells.Item("Date").Value
                        dr1("Date") = row.Cells.Item("Date").Value
                        dr("Start Time") = row.Cells.Item("Start Time").Value
                        dr1("Start Time") = row.Cells.Item("Start Time").Value
                        'dRow("StartTime") = spotrow("StartTime")
                        'dRow("EndTime") = spotrow("EndTime")
                        'dr("End Time") = row.Cells.Item("End Time").Value
                        'dr1("End Time") = row.Cells.Item("End Time").Value
                        dr("Duration(Sec)") = row.Cells.Item("Duration(Sec)").Value
                        dr1("Duration(Sec)") = row.Cells.Item("Duration(Sec)").Value
                        'dRow("Duration(Sec)") = TimeSpan.Parse(dRow("EndTime").ToString()).Subtract(TimeSpan.Parse(dRow("StartTime").ToString())).TotalSeconds
                        'dRow("PA") = spotrow("PA")
                        dr("PA") = row.Cells.Item("PA").Value
                        dr1("PA") = row.Cells.Item("PA").Value
                        '  dr(spots(index).MG + "TVR") = spots(index).TVRVal.Split({","c}, StringSplitOptions.None)(0)
                        ' dRow("TA") = spotrow("TA")
                        dr("TA") = row.Cells.Item("TA").Value
                        dr1("TA") = row.Cells.Item("TA").Value
                        '  dr("Commercial") = row.Cells("Commercial").Value
                        dr("Cost") = row.Cells.Item("Cost").Value
                        dr1("Cost") = row.Cells.Item("Cost").Value
                        For Each market As String In Globals.Ribbons.MSprintExRibbon.markets
                            dr(market) = row.Cells.Item(market).Value
                            dr1(market) = row.Cells.Item(market).Value
                        Next
                        Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Add(dr)
                        Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.AcceptChanges()
                        selecpots.Rows.Add(dr1)
                        dgvAvailableSpotsGrid.Rows.Remove(row)
                    Next
                ElseIf ((count - selecpots.Rows.Count) > 0 And ((count - selecpots.Rows.Count) < dgvAvailableSpotsGrid.SelectedRows.Count)) Then
                    ' Windows.Forms.MessageBox.Show(String.Format("Maximum of {0} spot(s) can be selected.Please reselect {0} spot(s)", count - selecpots.Rows.Count))
                    ShowErrorLabel(String.Format("Maximum of {0} spot(s) can be selected.Please reselect {0} spot(s)", count - selecpots.Rows.Count))
                Else
                    ' Windows.Forms.MessageBox.Show("Maximum number of spot(s) has been selected.Please remove unwanted and/or increase number of required spot(s) count and Try again")
                    ShowErrorLabel("Maximum number of spot(s) has been selected.Please remove unwanted and/or increase number of required spot(s) count and Try again")
                End If

                dgSelectedSpotsGrid.DataSource = Nothing
                dgSelectedSpotsGrid.DataSource = selecpots
                Globals.Ribbons.MSprintExRibbon.HideSelectedSpotsGrid(dgSelectedSpotsGrid)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnPushAllToSelected_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub dgSelectedSpotsGrid_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgSelectedSpotsGrid.CellContentDoubleClick

    End Sub

    Private Sub dgSelectedSpotsGrid_RowHeaderMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgSelectedSpotsGrid.RowHeaderMouseClick

    End Sub

    Private Sub DataGridView2_RowHeaderMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvAvailableSpotsGrid.RowHeaderMouseClick

    End Sub
    Public Function ConstructInputRnFXMLForAvailableSpots() As XElement
        Dim input As XElement
        Dim pchannels As Data.DataTable = Globals.Ribbons.MSprintExRibbon.GetGridTable()
        If pchannels.Rows.Count = 0 Then

            If Globals.Ribbons.MSprintExRibbon.mappedchannels Is Nothing Then
                pchannels = New Data.DataTable()
            Else
                pchannels = Globals.Ribbons.MSprintExRibbon.mappedchannels
            End If


        End If
        Try
            Dim month, month1 As String
            Dim day, day1 As String
            Dim plantgname As String = String.Empty
            Dim reftgname As String = String.Empty
            Dim dtable As System.Data.DataTable = DirectCast(Globals.Ribbons.MSprintExRibbon.tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            plantgname = dtable.Rows(0)(1).ToString().Trim()
            '  Button1.Enabled = False
            ' lbGetting.Text = "Getting Genre Share for chosen TG-MGs..."
            '   lbGetting.Refresh()
            If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Month < 10 Then
                month = "0" + Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Month.ToString()
            Else
                month = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Month.ToString()
            End If
            If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Day < 10 Then
                day = "0" + Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Day.ToString()
            Else
                day = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Day.ToString()
            End If
            If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Month < 10 Then
                month1 = "0" + Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Month.ToString()
            Else
                month1 = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Month.ToString()
            End If
            If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Day < 10 Then
                day1 = "0" + Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Day.ToString()
            Else
                day1 = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Day.ToString()
            End If
            input = <mediaplan>
                        <PreEvalPeriod>
                            <StartDate><%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Year.ToString() + month + day %></StartDate>
                            <EndDate><%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Year.ToString() + month1 + day1 %></EndDate>
                        </PreEvalPeriod>
                    </mediaplan>
            If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.LbDayParts.Items.Count > 0 Then
                Dim dayparts As XElement =
            <DayParts></DayParts>
                For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.LbDayParts.Items.Count - 1
                    '<day_part>0200-0200</day_part>
                    '  <day_part>0200-0200</day_part>
                    Dim dpart As XElement = New XElement("DayPart")
                    dpart.Value = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.LbDayParts.Items(index)
                    dayparts.Add(dpart)
                Next
                input.Add(dayparts)
            End If
            Dim tg As XElement = XElement.Load(Globals.Ribbons.MSprintExRibbon.tgDirectoryPath + plantgname + ".xml")
            Dim TG_MGElement As XElement =
              <tg name=<%= plantgname %> cs=<%= tg.Element("cs").Value %> sec=<%= tg.Element("sec").Value %> sex=<%= tg.Element("sex").Value %> age=<%= tg.Element("age").Value %>>
              </tg>
            'For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
            '    ' Dim doc2 As XmlDocument = New XmlDocument()
            '    '  doc2.Load()
            '    '  TG_MGElement.Add(XElement.Load(Path.GetTempPath() + "\\TGS\\" + plantgname + ".xml"))
            '    TG_MGElement.Add(XElement.Load(Globals.Ribbons.MSprintExRibbon.mgDirectoryPath + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim() + ".xml"))
            'Next
            Dim allMGElement As XElement = <mg name="TotalMarkets" type="group"></mg>
            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                ' Dim doc2 As XmlDocument = New XmlDocument()
                '  doc2.Load()
                '  TG_MGElement.Add(XElement.Load(Path.GetTempPath() + "\\TGS\\" + plantgname + ".xml"))
                Dim mg As XElement = XElement.Load(Globals.Ribbons.MSprintExRibbon.mgDirectoryPath + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim() + ".xml")
                TG_MGElement.Add(mg)
                allMGElement.Add(mg.Elements())
            Next
            TG_MGElement.Add(allMGElement)
           
            input.Add(TG_MGElement)
            Dim planType As String = String.Empty

            If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                planType = "clubbed"
                Dim plan As XElement =
              <plan type=<%= planType %>></plan>
                Dim period As XElement =
                <period StartDate=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Year.ToString() + month + day %> EndDate=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Year.ToString() + month1 + day1 %> year=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Year.ToString() %> WeekNum=<%= String.Empty %>></period>
                '  For index1 = 0 To Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows.Count - 1
                Dim drows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.xecelTable.Select(String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem))

                If drows.Count > 0 Then
                    ' mpTpSpotSelection.dgSelectedSpotsGrid.DataSource = drows.CopyToDataTable()

                    Dim program As XElement
                    Dim channelname As String = pchannels.Select("PCName='" + drows(0)("Channel").ToString() + "'")(0)("MCName")
                    ' Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                    Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + channelname + "'")(0)("ID")
                    'Dim guid As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID").ToString()

                    'If guid.Trim().Length = 0 Then
                    '    Dim guidval As String = GetGUIDFromCopy(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1))
                    '    If guidval.Length = 0 Then
                    '        guid = System.Guid.NewGuid.ToString()
                    '        Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID") = guid
                    '    Else
                    '        guid = guidval
                    '    End If

                    'End If

                    If planType.Equals("clubbed") Then
                        program =
                  <programme guid=<%= drows(0)("GUID").ToString() %> SeqNumber='1' ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= drows(0)("Programme").ToString() %> days=<%= drows(0)("Day").ToString() %> StartTime=<%= GetStartTime(drows(0)("Start Time").ToString()) %> EndTime=<%= GetEndTime(drows(0)("End Time").ToString()) %> CostPer10s=<%= drows(0)("RatePer10Sec").ToString() %> caption=<%= drows(0)("Creative").ToString() %> AdDuration=<%= drows(0)("Duration").ToString() %> NumberOfSpots=<%= drows(0)("Total Spots").ToString() %>>
                  </programme>
                        ' Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID") = (index1 + 1).ToString()
                        'ElseIf planType.Equals("weekwise") Then
                        '    If Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Columns.Contains("Week " & dtweeks.Rows(index)("WeekNumber").ToString()) Then
                        '        Dim col As String = "Week " & dtweeks.Rows(index)("WeekNumber").ToString()
                        '        program =
                        '       <programme guid=<%= (index1 + 1).ToString() %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)(col).ToString() %>>
                        '       </programme>
                        '        Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID") = (index1 + 1).ToString()
                        '    End If
                    End If
                    Dim selectedspts As XElement =
                           <selected_spots></selected_spots>
                    If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then

                        If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count > 0 Then
                            Dim filter As String = String.Format("GUID = '{0}'", drows(0)("GUID").ToString())

                            Dim srows As DataRow() = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Select(filter)
                            If srows.Count > 0 Then

                                For index2 = 0 To srows.Count - 1
                                    Dim spot As XElement =
                                        <spot log=<%= srows(index2)("Spot").ToString() %>></spot>
                                    selectedspts.Add(spot)
                                Next

                            End If
                        End If
                    End If
                    program.Add(selectedspts)
                    period.Add(program)

                    '  Next

                    plan.Add(period)
                    input.Add(plan)
                End If
            ElseIf Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                planType = "weekwise"

                Dim plan As XElement =
                <plan type=<%= planType %>></plan>
                Dim dtweeks As Data.DataTable = CType(Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dgvWeeks.DataSource, Data.DataTable)

                '  If Not (Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks Is Nothing) Then
                For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks.Rows.Count - 1
                    'Wk#,Year,StartDate,EndDate
                    Dim monthh1, month11 As String
                    Dim dayy1, day11 As String
                    '  Button1.Enabled = False
                    ' lbGetting.Text = "Getting Genre Share for chosen TG-MGs..."
                    '   lbGetting.Refresh()
                    Dim startdate As Date = Convert.ToDateTime(dtweeks.Rows(index)("StartDate").ToString()).ToShortDateString()
                    Dim enddate As Date = Convert.ToDateTime(dtweeks.Rows(index)("EndDate").ToString()).ToShortDateString()
                    '  Dim day As String()
                    If startdate.Month < 10 Then
                        monthh1 = "0" + startdate.Month.ToString()
                    Else
                        monthh1 = startdate.Month.ToString()
                    End If
                    If startdate.Day < 10 Then
                        dayy1 = "0" + startdate.Day.ToString()
                    Else
                        dayy1 = startdate.Day.ToString()
                    End If
                    If enddate.Month < 10 Then
                        month11 = "0" + enddate.Month.ToString()
                    Else
                        month11 = enddate.Month.ToString()
                    End If
                    If enddate.Day < 10 Then
                        day11 = "0" + enddate.Day.ToString()
                    Else
                        day11 = enddate.Day.ToString()
                    End If
                    Dim drows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.xecelTable.Select(String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem))
                    Dim period As XElement =
                <period StartDate=<%= startdate.Year.ToString() + monthh1 + dayy1 %> EndDate=<%= enddate.Year.ToString() + month11 + day11 %> year=<%= dtweeks.Rows(index)("Year").ToString() %> WeekNum=<%= dtweeks.Rows(index)("WeekNumber").ToString() %>></period>
                    If drows.Count > 0 Then


                        '<programme guid="" SeqNumber="1" ChannelCode="004" ChannelName="Star Plus" ProgName="Yeh Rishta Kya Kehlata Hai"  days="Thu" StartTime="21:30" EndTime="22:00" CostPer10s="150" caption="Colgate Kids Jumping" AdDuration="30" NumberOfSpots="10">
                        'inpSpotTable.Columns.Add("GUID")
                        'inpSpotTable.Columns(0).AutoIncrement = True
                        'inpSpotTable.Columns(0).AutoIncrementSeed = 1
                        'inpSpotTable.Columns.Add("Channel")
                        'inpSpotTable.Columns.Add("Programme")
                        'inpSpotTable.Columns.Add("Day")
                        'inpSpotTable.Columns.Add("Start Time")
                        'inpSpotTable.Columns.Add("End Time")
                        'inpSpotTable.Columns.Add("RatePer10Sec")
                        'inpSpotTable.Columns.Add("Duration")
                        'inpSpotTable.Columns.Add("Creative")
                        'inpSpotTable.Columns.Add("Total Spots", Type.GetType("System.Int32"))
                        '   For index1 = 0 To Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows.Count - 1


                        Dim program As XElement
                        '  Dim channelname As String = pchannels.Select("PCName='" + Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() + "'")(0)("MCName")
                        '  Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                        'Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() + "'")(0)("ID")
                        Dim channelname As String = pchannels.Select("PCName='" + drows(0)("Channel").ToString() + "'")(0)("MCName")
                        ' Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                        Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + channelname + "'")(0)("ID")
                        '  If planType.Equals("clubbed") Then
                        '      program =
                        '<programme guid=<%= (index1 + 1).ToString() %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Total Spots").ToString() %>>
                        '</programme>
                        '      Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID") = (index1 + 1).ToString()
                        '  ElseIf planType.Equals("weekwise") Then
                        'Dim guid As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID").ToString()

                        'If guid.Trim().Length = 0 Then
                        '    Dim guidval As String = GetGUIDFromCopy(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1))
                        '    If guidval.Length = 0 Then
                        '        guid = System.Guid.NewGuid.ToString()
                        '        Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID") = guid
                        '    Else
                        '        guid = guidval
                        '    End If

                        'End If
                        If Globals.Ribbons.MSprintExRibbon.xecelTable.Columns.Contains("Week " & dtweeks.Rows(index)("WeekNumber").ToString()) Then
                            Dim col As String = "Week " & dtweeks.Rows(index)("WeekNumber").ToString()
                            program =
                           <programme guid=<%= drows(0)("GUID").ToString() %> SeqNumber='1' ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= drows(0)("Programme").ToString() %> days=<%= drows(0)("Day").ToString() %> StartTime=<%= GetStartTime(drows(0)("Start Time").ToString()) %> EndTime=<%= GetEndTime(drows(0)("End Time").ToString()) %> CostPer10s=<%= drows(0)("RatePer10Sec").ToString() %> caption=<%= drows(0)("Creative").ToString() %> AdDuration=<%= drows(0)("Duration").ToString() %> NumberOfSpots=<%= drows(0)(col).ToString() %>>
                           </programme>
                            ' Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID") = (index1 + 1).ToString()
                        End If
                        'End If
                        Dim selectedspts As XElement =
                               <selected_spots></selected_spots>
                        If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then

                            If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count > 0 Then
                                Dim filter As String = String.Format("GUID = '{0}' and WeekNum= {1} ", drows(0)("GUID").ToString(), Convert.ToInt32(dtweeks.Rows(index)("WeekNumber").ToString()))

                                Dim srows As DataRow() = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Select(filter)
                                If srows.Count > 0 Then

                                    For index2 = 0 To srows.Count - 1
                                        Dim spot As XElement =
                                            <spot log=<%= srows(index2)("Spot").ToString() %>></spot>
                                        selectedspts.Add(spot)
                                    Next

                                End If
                            End If
                        End If
                        program.Add(selectedspts)
                        period.Add(program)
                        plan.Add(period)
                    End If
                    Next

                '  Next
                input.Add(plan)
            End If

                      
            '  End If
            ' End If
            '  End If

        Catch ex As Exception
            LogMpsrintExException("Exception occured while constructing input XML for RnF" + ex.Message)
            Throw ex
        End Try
        Return input
    End Function
    Public Function ConstructOpRnFTableForAvailableSpots(ByVal opXmL As XElement) As Data.DataTable
        ' Dim output As Data.DataTable = New Data.DataTable()
        Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots = New Data.DataTable()
        Try
            ' RnFOutputTable = New Data.DataTable()
            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("GUID")
            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("Channel")
            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("Spot")
            ' output.Columns.Add("AvaiSpotString")
            'Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("TG")
            'Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("MG")
            'output.Columns.Add("ReachVal")
            ' Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("TVRVal")
            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("Start Date", System.Type.GetType("System.DateTime"))
            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("End Date", System.Type.GetType("System.DateTime"))
            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("WeekNum", System.Type.GetType("System.Int32"))
            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("Date", System.Type.GetType("System.DateTime"))
            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("Start Time")
            ' RnFSelectedSpots.Columns.Add("End Time")
            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
            ' RnFSelectedSpots.Columns.Add("Commercial")
            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("Cost")
            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("PA")
            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("TA")
            'output.Columns.Add("Day")
            'output.Columns.Add("Programme")
            'output.Columns.Add("TVR000s")
            'output.Columns.Add("TVR")
            'output.Columns.Add("GRP000s")
            'output.Columns.Add("GRP")
            'output.Columns.Add("AvgFreq")
            'output.Columns.Add("CummCost")
            'output.Columns.Add("SpotCPRP")
            'output.Columns.Add("CummCPRP")
            'output.Columns.Add("Reach000s")
            'output.Columns.Add("1+")
            'output.Columns.Add("2+")
            'output.Columns.Add("3+")
            'output.Columns.Add("4+")
            'output.Columns.Add("5+")
            'output.Columns.Add("6+")
            'output.Columns.Add("7+")
            'output.Columns.Add("8+")
            'output.Columns.Add("9+")
            'output.Columns.Add("10+")
            '  Dim rnfoutput As XElement = XElement.Load(AppDomain.CurrentDomain.BaseDirectory + "sampleoutput.xml")
            For Each period As XElement In opXmL.Element("plan").Elements
                Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns("Start Date").DefaultValue = New Date(period.Attribute("StartDate").Value.Substring(0, 4), period.Attribute("StartDate").Value.Substring(4, 2), period.Attribute("StartDate").Value.Substring(6, 2)).ToShortDateString()
                Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns("End Date").DefaultValue = New Date(period.Attribute("EndDate").Value.Substring(0, 4), period.Attribute("EndDate").Value.Substring(4, 2), period.Attribute("EndDate").Value.Substring(6, 2)).ToShortDateString()

                If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                    Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns("WeekNum").DefaultValue = 0
                Else
                    Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns("WeekNum").DefaultValue = Convert.ToInt32(period.Attribute("WeekNum").Value)
                End If


                For Each programme As XElement In period.Elements
                    Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns("Channel").DefaultValue = programme.Attribute("ChannelName").Value
                    'output.Columns("Day").DefaultValue = programme.Attribute("days").Value
                    'output.Columns("Programme").DefaultValue = programme.Attribute("ProgName").Value
                    Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns("GUID").DefaultValue = programme.Attribute("guid").Value
                    'For Each spot As XElement In programme.Element("selected_spots").Elements
                    '    output.Columns("Spot").DefaultValue = spot.Attribute("log").Value
                    '    'RnFSelectedSpots.Columns.Add("Date", System.Type.GetType("System.DateTime"))
                    '    'RnFSelectedSpots.Columns.Add("Start Time")
                    '    '' RnFSelectedSpots.Columns.Add("End Time")
                    '    'RnFSelectedSpots.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
                    '    '' RnFSelectedSpots.Columns.Add("Commercial")
                    '    'RnFSelectedSpots.Columns.Add("Cost")
                    '    'RnFSelectedSpots.Columns.Add("PA")
                    '    'RnFSelectedSpots.Columns.Add("TA")
                    '    'output.Columns.Add("Date", System.Type.GetType("System.DateTime"))
                    '    'output.Columns.Add("Start Time")
                    '    '' RnFSelectedSpots.Columns.Add("End Time")
                    '    'output.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
                    '    '' RnFSelectedSpots.Columns.Add("Commercial")
                    '    'output.Columns.Add("Cost")
                    '    'output.Columns.Add("PA")
                    '    'output.Columns.Add("TA")
                    '    Dim values As String() = spot.Attribute("log").Value.Split({","c}, StringSplitOptions.None)
                    '    '   Dim dr As Data.DataRow = spot.NewRow()
                    '    'For index = 0 To values.Length - 1

                    '    ' If index = 1 Then
                    '    'dr("Date") = New Date(Convert.ToInt32(values(1).Substring(0, 4)), Convert.ToInt32(values(1).Substring(4, 2)), Convert.ToInt32(values(1).Substring(6, 2))).ToShortDateString()
                    '    '' ElseIf index = 2 Then
                    '    '' dr("StartTime") = New TimeSpan(Convert.ToInt32(values(2).Substring(0, 2)), Convert.ToInt32(values(2).Substring(2, 2)), Convert.ToInt32(values(2).Substring(4, 2)))
                    '    ''  dr("StartTime") = values(2).Substring(0, 5)
                    '    'dr("StartTime") = String.Format("{0}:{1}", values(2).Substring(0, 2), values(2).Substring(2, 2))
                    '    'dr("Cost") = values(4)
                    '    'dr("PA") = values(5)
                    '    'dr("TA") = values(6)
                    '    'dr("Duration(Sec)") = values(7)
                    '    output.Columns("Date").DefaultValue = New Date(Convert.ToInt32(values(1).Substring(0, 4)), Convert.ToInt32(values(1).Substring(4, 2)), Convert.ToInt32(values(1).Substring(6, 2))).ToShortDateString()
                    '    output.Columns("Start Time").DefaultValue = String.Format("{0}:{1}", values(2).Substring(0, 2), values(2).Substring(2, 2))
                    '    output.Columns("Cost").DefaultValue = values(4)
                    '    output.Columns("PA").DefaultValue = values(5)
                    '    output.Columns("TA").DefaultValue = values(6)
                    '    output.Columns("Duration(Sec)").DefaultValue = values(7)
                    '    For Each reach As XElement In spot.Elements
                    '        Dim dr As Data.DataRow = output.NewRow()
                    '        dr("TG") = reach.Attribute("tm").Value.Split({"~"c}, StringSplitOptions.None)(0)
                    '        dr("MG") = reach.Attribute("tm").Value.Split({"~"c}, StringSplitOptions.None)(1)
                    '        dr("ReachVal") = reach.Attribute("val").Value
                    '        Dim vals As String() = reach.Attribute("val").Value.Split({","c}, StringSplitOptions.None)
                    '        '  Dim drow As Data.DataRow = reach.NewRow()
                    '        dr("TVR000s") = vals(0)
                    '        dr("TVR") = vals(1)
                    '        dr("GRP000s") = vals(2)
                    '        dr("GRP") = vals(3)
                    '        dr("AvgFreq") = vals(4)
                    '        dr("CummCost") = vals(5)
                    '        dr("SpotCPRP") = vals(6)
                    '        dr("CummCPRP") = vals(7)
                    '        dr("Reach000s") = vals(8)
                    '        dr("1+") = vals(9)
                    '        dr("2+") = vals(10)
                    '        dr("3+") = vals(11)
                    '        dr("4+") = vals(12)
                    '        dr("5+") = vals(13)
                    '        dr("6+") = vals(14)
                    '        dr("7+") = vals(15)
                    '        dr("8+") = vals(16)
                    '        dr("9+") = vals(17)
                    '        dr("10+") = vals(18)
                    '        output.Rows.Add(dr)
                    '    Next
                    'Next
                    'Available Spots

                    If programme.Elements("available_spots").Any() Then
                        For Each spot As XElement In programme.Element("available_spots").Elements
                            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns("Spot").DefaultValue = spot.Attribute("log").Value
                            Dim values As String() = spot.Attribute("log").Value.Split({","c}, StringSplitOptions.None)
                            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns("Date").DefaultValue = New Date(Convert.ToInt32(values(1).Substring(0, 4)), Convert.ToInt32(values(1).Substring(4, 2)), Convert.ToInt32(values(1).Substring(6, 2))).ToShortDateString()
                            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns("Start Time").DefaultValue = String.Format("{0}:{1}", values(2).Substring(0, 2), values(2).Substring(2, 2))
                            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns("Cost").DefaultValue = values(4)
                            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns("PA").DefaultValue = values(5)
                            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns("TA").DefaultValue = values(6)
                            Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns("Duration(Sec)").DefaultValue = values(7)
                            Dim mgcount As Integer = 0
                            For Each reach As XElement In spot.Elements
                                ' Dim dr As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.NewRow()
                                '  dr("TG") = reach.Attribute("tm").Value.Split({"~"c}, StringSplitOptions.None)(0)
                                'dr("MG") = reach.Attribute("tm").Value.Split({"~"c}, StringSplitOptions.None)(1)
                                Dim mgvalue As String = reach.Attribute("tm").Value.Split({"~"c}, StringSplitOptions.None)(1)

                                If Not (Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Contains(mgvalue + "TVR")) Then
                                    Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add(mgvalue + "TVR")
                                End If
                                If Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Contains(mgvalue + "TVR") Then
                                    Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns(mgvalue + "TVR").DefaultValue = reach.Attribute("val").Value.Split({","c}, StringSplitOptions.None)(0)
                                End If

                                If mgcount = spot.Elements.Count - 1 Then
                                    Dim dr1 As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.NewRow()
                                    dr1(mgvalue + "TVR") = reach.Attribute("val").Value.Split({","c}, StringSplitOptions.None)(0)
                                    Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Rows.Add(dr1)
                                End If
                                ' dr("TVRVal") = reach.Attribute("val").Value
                                'Dim mg As String = dr("MG").ToString()
                                'If Not (output.Columns.Contains(mg + "TVR")) Then
                                '    output.Columns.Add(mg + "TVR")
                                'End If

                                'If output.Columns.Contains(mg + "TVR") Then
                                '    dr(mg + "TVR") = reach.Attribute("val").Value.Split({","c}, StringSplitOptions.None)(0)
                                'End If

                                ' Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Rows.Add(dr)
                                mgcount += 1
                            Next
                        Next
                    End If



                Next
            Next
        Catch ex As Exception
            LogMpsrintExException("Exception occured while constructing dataset from output XML." + ex.Message)
            Throw ex
        End Try
        Return Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetAvailableSpots.Click
        Try
            System.Windows.Forms.Application.DoEvents()
            HideErrorLabel()
            btnGetAvailableSpots.Enabled = False
            Globals.ThisAddIn.Application.StatusBar = "Getting Available Spots..."
            System.Windows.Forms.Application.DoEvents()
            Dim input As XElement = ConstructInputRnFXMLForAvailableSpots()
            Dim name As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
            input.Save(Globals.Ribbons.MSprintExRibbon.LogDirectoryPath + "AvailableSpots_Inp_" + name)
            Globals.Ribbons.MSprintExRibbon.rnfoutputXml = GetOpXMLFromWS(input, "http://ec2-54-255-217-55.ap-southeast-1.compute.amazonaws.com:8080/GroupM/spotselectionnew/getavailablespot")

            If Not (Globals.Ribbons.MSprintExRibbon.rnfoutputXml Is Nothing) Then

                If Globals.Ribbons.MSprintExRibbon.rnfoutputXml.Elements.Count > 0 Then
                    Dim name1 As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                    Globals.Ribbons.MSprintExRibbon.rnfoutputXml.Save(Globals.Ribbons.MSprintExRibbon.LogDirectoryPath + "AvailableSpots_Op_" + name1)

                    'output.Columns.Add("GUID", System.Type.GetType("System.Int32"))
                    'output.Columns.Add("ChannelName")
                    'output.Columns.Add("SpotString")
                    'output.Columns.Add("AvaiSpotString")
                    'output.Columns.Add("TG")
                    'output.Columns.Add("MG")
                    'output.Columns.Add("ReachVal")
                    'output.Columns.Add("TVRVal")
                    'output.Columns.Add("Start Date", System.Type.GetType("System.DateTime"))
                    'output.Columns.Add("End Date", System.Type.GetType("System.DateTime"))
                    'output.Columns.Add("WeekNum", System.Type.GetType("System.Int32"))
                    ConstructOpRnFTableForAvailableSpots(Globals.Ribbons.MSprintExRibbon.rnfoutputXml)
                    'Dim spots = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.AsEnumerable().Join(Globals.Ribbons.MSprintExRibbon.RnFOutputTable.AsEnumerable(), Function(o) o.Field(Of String)("GUID"), _
                    '                          Function(c) c.Field(Of String)("GUID"), _
                    '                          Function(c, o) _
                    '                              New With {.GUID = o.Field(Of String)("GUID"), _
                    '                                       .Spot = o.Field(Of String)("AvaiSpotString"), _
                    '                                       .StartDate = o.Field(Of DateTime)("Start Date"),
                    '                                       .EndDate = o.Field(Of DateTime)("End Date"),
                    '                                       .ChannelName = o.Field(Of String)("ChannelName"),
                    '                                        .TG = o.Field(Of String)("TG"),
                    '                                        .MG = o.Field(Of String)("MG"),
                    '                                         .TVRVal = o.Field(Of String)("TVRVal"),
                    '                                        .WeekNum = o.Field(Of Int32)("WeekNum")})

                    '   mpTpSpotSelection = New ucSpotSelection()
                    'Dim avaiSpots As DataTable = New DataTable()
                    'avaiSpots = spots.CopyToDataTable()
                    'Dim temp As DataTable = New DataTable()

                    'Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots = New Data.DataTable()
                    '' spot.Columns.Add("ChannelCode")
                    'temp.Columns.Add("Channel")
                    'temp.Columns.Add("Date", System.Type.GetType("System.DateTime"))
                    'temp.Columns.Add("Start Time")
                    '' temp.Columns.Add("End Time", System.Type.GetType("System.TimeSpan"))
                    '' spot.Columns.Add(
                    'temp.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
                    'temp.Columns.Add("Cost")
                    'temp.Columns.Add("PA")
                    'temp.Columns.Add("TA")

                    ''For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                    ''    Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index) + "TVR", System.Type.GetType("System.Decimal"))

                    ''Next
                    'temp.Columns.Add("GUID", System.Type.GetType("System.Int32"))
                    'temp.Columns.Add("Spot")
                    'temp.Columns.Add("StartDate", System.Type.GetType("System.DateTime"))
                    'temp.Columns.Add("EndDate", System.Type.GetType("System.DateTime"))
                    'temp.Columns.Add("WeekNum", System.Type.GetType("System.Int32"))
                    'temp.Columns.Add("MG")
                    'temp.Columns.Add("TVRVAL")
                    '' Dim groupList = From Group1 In  avaiSpots.AsEnumerable() Group g By g.Field(Of String)("MG") 
                    ''Dim numberGroups = From n In numbers _
                    ''     Group n By num = n.Field(Of Integer)("number")
                    ''Dim spots = From spot In avaiSpots.AsEnumerable() _
                    ''            Group spot By spot.Field(Of String)("MG")
                    ''For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                    ''    ' Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index) + "TVR", System.Type.GetType("System.Decimal"))
                    ''    Dim availablerows As DataRow() = avaiSpots.Select("MG= '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index) + "'")

                    ''    If index = 0 And availablerows.Count > 0 Then
                    ''        Dim dr As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.NewRow()
                    ''        dr("GUID") = spots(index).GUID
                    ''        dr("Spot") = spots(index).Spot
                    ''        dr("StartDate") = spots(index).StartDate
                    ''        dr("EndDate") = spots(index).EndDate
                    ''        dr("WeekNum") = spots(index).WeekNum
                    ''        dr("ChannelName") = spots(index).ChannelName
                    ''        Dim spotrow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.GetSpotRow(spots(index).Spot)
                    ''        dr("Date") = spotrow("Date")
                    ''        dr("StartTime") = spotrow("StartTime")
                    ''        'dRow("StartTime") = spotrow("StartTime")
                    ''        'dRow("EndTime") = spotrow("EndTime")
                    ''        dr("EndTime") = spotrow("EndTime")
                    ''        dr("Duration(Sec)") = spotrow("Duration(Sec)")
                    ''        'dRow("PA") = spotrow("PA")
                    ''        dr("PA") = spotrow("PA")
                    ''        dr("TA") = spotrow("TA")
                    ''        '  dr("Commercial") = spotrow("Commercial")
                    ''        dr("Cost") = spotrow("Cost")
                    ''        Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Rows.Add(dr)
                    ''    End If

                    ''Next
                    'For index = 0 To spots.Count - 1
                    '    Dim dr As Data.DataRow = temp.NewRow()
                    '    dr("GUID") = spots(index).GUID
                    '    dr("Spot") = spots(index).Spot
                    '    dr("StartDate") = spots(index).StartDate
                    '    dr("EndDate") = spots(index).EndDate
                    '    dr("WeekNum") = spots(index).WeekNum
                    '    dr("Channel") = spots(index).ChannelName
                    '    Dim spotrow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.GetSpotRow(spots(index).Spot)
                    '    dr("Date") = spotrow("Date")
                    '    dr("Start Time") = spotrow("StartTime")
                    '    'dRow("StartTime") = spotrow("StartTime")
                    '    'dRow("EndTime") = spotrow("EndTime")
                    '    '  dr("End Time") = spotrow("EndTime")
                    '    dr("Duration(Sec)") = spotrow("Duration(Sec)")
                    '    'dRow("PA") = spotrow("PA")
                    '    dr("PA") = spotrow("PA")
                    '    dr("MG") = spots(index).MG
                    '    dr("TVRVAL") = spots(index).TVRVal
                    '    '  dr(spots(index).MG + "TVR") = spots(index).TVRVal.Split({","c}, StringSplitOptions.None)(0)
                    '    ' dRow("TA") = spotrow("TA")
                    '    dr("TA") = spotrow("TA")
                    '    '  dr("Commercial") = spotrow("Commercial")
                    '    dr("Cost") = spotrow("Cost")
                    '    temp.Rows.Add(dr)
                    'Next
                    'Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots = temp.DefaultView.ToTable(True, New String() {"GUID", "Spot", "StartDate", "EndDate", "WeekNum", "Channel", "Date", "Start Time", "Duration(Sec)", "PA", "TA", "Cost"})
                    'For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                    '    Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index) + "TVR", System.Type.GetType("System.Decimal"))

                    'Next
                    ''Dim id As DataColumn = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("ID", System.Type.GetType("System.Int32"))
                    ''id.AutoIncrement = True
                    'For Each row As DataRow In Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Rows
                    '    For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                    '        Dim tvrval As String = temp.Select(String.Format("MG='{0}' and WeekNum={1} and GUID='{2}'", Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString(), Convert.ToInt32(row("WeekNum").ToString()), Convert.ToInt32(row("GUID").ToString())))(0)("TVRVAL").ToString()
                    '        row(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index) + "TVR") = tvrval.Split({","c}, StringSplitOptions.None)(0)

                    '    Next
                    'Next
                    'Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots = Globals.Ribbons.MSprintExRibbon.RnFOutputTable.DefaultView.ToTable(True, New String() {"GUID", "Spot", "Start Date", "End Date", "WeekNum", "Channel", "Date", "Start Time", "Duration(Sec)", "PA", "TA", "Cost"})
                    'For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                    '    Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index) + "TVR", System.Type.GetType("System.Decimal"))

                    'Next
                    'For Each row As DataRow In Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Rows
                    '    For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                    '        Dim tvrval As String = Globals.Ribbons.MSprintExRibbon.RnFOutputTable.Select(String.Format("MG='{0}' and WeekNum={1} and GUID='{2}'", Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString(), Convert.ToInt32(row("WeekNum").ToString()), row("GUID").ToString()))(0)("TVRVAL").ToString()
                    '        row(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index) + "TVR") = tvrval.Split({","c}, StringSplitOptions.None)(0)

                    '    Next
                    'Next
                    Dim filter As String = String.Format("GUID ='{0}' ", Globals.Ribbons.MSprintExRibbon.currentLineItem)

                    If cbWeeks.Text = "All" Then
                        'filter = String.Format("GUID ='{0}' ", Globals.Ribbons.MSprintExRibbon.currentLineItem)
                        dgvAvailableSpotsGrid.DataSource = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots
                    Else
                        filter = String.Format("GUID ='{0}' and WeekNum={1}", Globals.Ribbons.MSprintExRibbon.currentLineItem, Convert.ToInt32(cbWeeks.Text))
                        Dim rows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Select(filter)

                        If rows.Count > 0 Then
                            'Dim aspots As DataTable = New DataTable()
                            'aspots = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Clone()
                            ''Parallel.For(0, rows.Count - 1, Sub(i)
                            ''                                    aspots.ImportRow(rows(i))
                            ''                                End Sub)
                            'For Each row As DataRow In rows
                            '    aspots.ImportRow(row)
                            'Next
                            dgvAvailableSpotsGrid.DataSource = rows.CopyToDataTable()
                        End If
                    End If

                   
                    Globals.Ribbons.MSprintExRibbon.HideSelectedSpotsGrid(dgvAvailableSpotsGrid)
                Else
                    MessageBox.Show("Unable to retreive Available Spots from Server.")
                End If
              
            Else
                MessageBox.Show("Unable to retreive Available Spots from Server.")
            End If
            btnGetAvailableSpots.Enabled = True
            Globals.ThisAddIn.Application.StatusBar = String.Empty
        Catch ex As Exception
            btnGetAvailableSpots.Enabled = True
            Globals.ThisAddIn.Application.StatusBar = String.Empty
            LogMpsrintExException("Exception occured while getting available spots" + ex.Message)
        End Try
    End Sub

    Private Sub SplitContainer2_Panel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles SplitContainer2.Panel2.Paint

    End Sub

    Private Sub cbWeeks_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbWeeks.SelectedIndexChanged
        Try
            HideErrorLabel()
            Dim filter As String = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
            If cbWeeks.Text = "All" Then
                filter = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
            Else
                filter = String.Format("GUID='{0}' and WeekNum={1}", Globals.Ribbons.MSprintExRibbon.currentLineItem, Int32.Parse(cbWeeks.Text))
            End If


            If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then

                Dim drows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Select(filter)

                If drows.Count > 0 Then
                    mpTpSpotSelection.dgSelectedSpotsGrid.DataSource = drows.CopyToDataTable()
                    Globals.Ribbons.MSprintExRibbon.HideSelectedSpotsGrid(dgSelectedSpotsGrid)
                End If
            End If
            If Not (Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots Is Nothing) Then

                Dim rows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Select(filter)

                If rows.Count > 0 Then
                    mpTpSpotSelection.dgvAvailableSpotsGrid.DataSource = rows.CopyToDataTable()
                    Globals.Ribbons.MSprintExRibbon.HideSelectedSpotsGrid(dgvAvailableSpotsGrid)
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ucSpotSelection_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            HideErrorLabel()
            If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                Dim weeks As DataTable = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks
                For Each row As DataRow In weeks.Rows
                    cbWeeks.Items.Add(row("WeekNumber").ToString())
                Next
                cbWeeks.Enabled = True
            Else
                cbWeeks.Enabled = False
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub dgSelectedSpotsGrid_DataSourceChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgSelectedSpotsGrid.DataSourceChanged
        '  Dim str As String = String.Empty
        Try
            ' Globals.Ribbons.MSprintExRibbon.DisplayCurrentPlanItem()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SplitContainer2_SplitterMoved(ByVal sender As System.Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles SplitContainer2.SplitterMoved

    End Sub

    Private Sub dgSelectedSpotsGrid_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgSelectedSpotsGrid.CellContentClick

    End Sub

    Private Sub SplitContainer2_Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles SplitContainer2.Panel1.Paint

    End Sub

    Private Sub gbSelectedSpots_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub dgSelectedSpotsGrid_CellMouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgSelectedSpotsGrid.CellMouseDoubleClick
        Try
            HideErrorLabel()
            If e.RowIndex > -1 Then

                Dim selecpots As DataTable = DirectCast(dgSelectedSpotsGrid.DataSource, DataTable)
                ' For Each row As DataGridViewRow In dgSelectedSpotsGrid.SelectedRows
                Dim row As DataGridViewRow = dgSelectedSpotsGrid.Rows(e.RowIndex)
                Dim filter As String = String.Format("Spot ='{0}'", row.Cells.Item("Spot").Value)
                Dim row1 As DataRow = selecpots.Select(filter)(0)
                Dim row2 As DataRow = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Select(filter)(0)
             
                If Not (Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots Is Nothing) Then
                    Dim avaispots As Data.DataTable = CType(dgvAvailableSpotsGrid.DataSource, Data.DataTable)
                    Dim dr As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.NewRow()
                    Dim gridrow As DataGridViewRow = dgSelectedSpotsGrid.SelectedRows(0)
                    '  Dim dr As DataRow = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.NewRow()
                    Dim dr1 As DataRow = avaispots.NewRow()
                    dr("GUID") = row2("GUID").ToString()
                    dr1("GUID") = row2("GUID").ToString()
                    dr("Spot") = row2("Spot").ToString()
                    dr1("Spot") = row2("Spot").ToString()
                    dr("Start Date") = row2("Start Date").ToString()
                    dr1("Start Date") = row2("Start Date").ToString()
                    dr("End Date") = row2("End Date").ToString()
                    dr1("End Date") = row2("End Date").ToString()
                    dr("WeekNum") = row2("WeekNum").ToString()
                    dr1("WeekNum") = row2("WeekNum").ToString()
                    dr("Channel") = row2("Channel").ToString()
                    dr1("Channel") = row2("Channel").ToString()
                    '  Dim spotrow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.GetSpotRow(spots(index).Spot)
                    dr("Date") = row2("Date").ToString()
                    dr1("Date") = row2("Date").ToString()
                    dr("Start Time") = row2("Start Time").ToString()
                    dr1("Start Time") = row2("Start Time").ToString()
                    'dRow("StartTime") = spotrow("StartTime")
                    ''dRow("EndTime") = spotrow("EndTime")
                    'dr("End Time") = row.Cells.Item("End Time").ToString()
                    'dr1("End Time") = row.Cells.Item("End Time").ToString()
                    dr("Duration(Sec)") = row2("Duration(Sec)").ToString()
                    dr1("Duration(Sec)") = row2("Duration(Sec)").ToString()
                    'dRow("Duration(Sec)") = TimeSpan.Parse(dRow("EndTime").ToString()).Subtract(TimeSpan.Parse(dRow("StartTime").ToString())).TotalSeconds
                    'dRow("PA") = spotrow("PA")
                    dr("PA") = row2("PA").ToString()
                    dr1("PA") = row2("PA").ToString()
                    '  dr(spots(index).MG + "TVR") = spots(index).TVRVal.Split({","c}, StringSplitOptions.None)(0)
                    ' dRow("TA") = spotrow("TA")
                    dr("TA") = row2("TA").ToString()
                    dr1("TA") = row2("TA").ToString()
                    '  dr("Commercial") = row.Cells("Commercial").ToString()
                    dr("Cost") = row2("Cost").ToString()
                    dr1("Cost") = row2("Cost").ToString()
                    Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Rows.Add(dr)
                    Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.AcceptChanges()
                    avaispots.Rows.Add(dr1)
                    dgvAvailableSpotsGrid.DataSource = avaispots
                    Globals.Ribbons.MSprintExRibbon.HideSelectedSpotsGrid(dgvAvailableSpotsGrid)
                End If
                selecpots.Rows.Remove(row1)
                Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Remove(row2)
                Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.AcceptChanges()
                ' Dim dr As DataRow = avaipots.NewRow()
                'For c As Integer = 0 To dgSelectedSpotsGrid.ColumnCount - 1

                '    If Not (avaipots.Columns.Contains(dgSelectedSpotsGrid.Columns(c).Name)) Then
                '        avaipots.Columns.Add(dgSelectedSpotsGrid.Columns(c).Name)
                '    End If
                '    dr(c) = row.Cells(c).Value
                'Next
                'avaipots.Rows.Add(dr)
                'dgvAvailableSpotsGrid.DataSource = avaipots
                'dgSelectedSpotsGrid.Rows.Remove(row)
                '  Next
                dgSelectedSpotsGrid.DataSource = Nothing
                dgSelectedSpotsGrid.DataSource = selecpots
                Globals.Ribbons.MSprintExRibbon.HideSelectedSpotsGrid(dgSelectedSpotsGrid)

            End If
        Catch ex As Exception
            LogMpsrintExException("Exception occured while removing selected spot." + ex.Message)
        End Try
    End Sub

    Private Sub dgvAvailableSpotsGrid_CellMouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvAvailableSpotsGrid.CellMouseDoubleClick
        Try
            HideErrorLabel()
            If e.RowIndex > -1 Then

                If cbWeeks.Text.Equals("All") And cbWeeks.Enabled Then
                    '  MessageBox.Show("Please choose a week to select spot(s)")
                    ShowErrorLabel("Please choose a week to select spot(s)")
                Else

                    Dim selecpots As DataTable = New DataTable()
                    Dim fil As String = fil = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
                    If cbWeeks.Text.Equals("All") Then
                        fil = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
                    Else

                        fil = String.Format("GUID='{0}' and WeekNum={1}", Globals.Ribbons.MSprintExRibbon.currentLineItem, Int32.Parse(cbWeeks.Text).ToString())
                    End If


                    Dim rows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Select(fil)
                    selecpots = rows.CopyToDataTable()
                    'If rows.Count > 0 Then

                    '    selecpots = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Clone()
                    '    Parallel.For(0, rows.Count - 1, Sub(i)
                    '                                        selecpots.ImportRow(rows(i))
                    '                                    End Sub)
                    'End If
                    ' = DirectCast(dgSelectedSpotsGrid.DataSource, DataTable)

                    Dim dt As DataTable = CType(loSpotSelection.DataSource, DataTable)
                    Dim count As Integer
                    Try
                        If dt.Columns.Contains("Total Spots") Then
                            Dim filter As String = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
                            count = Convert.ToInt32(dt.Select(filter)(0)("Total Spots").ToString())
                        Else
                            Dim filter As String = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
                            count = Convert.ToInt32(dt.Select(filter)(0)("Week " & Int32.Parse(cbWeeks.Text)).ToString())

                        End If
                    Catch ex As Exception
                        Dim filter As String = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
                        If Globals.Ribbons.MSprintExRibbon.xecelTable.Columns.Contains("Total Spots") Then
                            count = Convert.ToInt32(Globals.Ribbons.MSprintExRibbon.xecelTable.Select(filter)(0)("Total Spots").ToString())
                        Else
                            count = Convert.ToInt32(Globals.Ribbons.MSprintExRibbon.xecelTable.Select(filter)(0)("Week " & Int32.Parse(cbWeeks.Text)).ToString())

                        End If
                    End Try

                    If count > selecpots.Rows.Count And (count - selecpots.Rows.Count) >= dgvAvailableSpotsGrid.SelectedRows.Count Then
                        For Each row As DataGridViewRow In dgvAvailableSpotsGrid.SelectedRows
                            Dim dr As DataRow = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.NewRow()
                            Dim dr1 As DataRow = selecpots.NewRow()
                            dr("GUID") = row.Cells.Item("GUID").Value
                            dr1("GUID") = row.Cells.Item("GUID").Value
                            dr("Spot") = row.Cells.Item("Spot").Value
                            dr1("Spot") = row.Cells.Item("Spot").Value
                            dr("Start Date") = row.Cells.Item("Start Date").Value
                            dr1("Start Date") = row.Cells.Item("Start Date").Value
                            dr("End Date") = row.Cells.Item("End Date").Value
                            dr1("End Date") = row.Cells.Item("End Date").Value
                            dr("WeekNum") = row.Cells.Item("WeekNum").Value
                            dr1("WeekNum") = row.Cells.Item("WeekNum").Value
                            dr("Channel") = row.Cells.Item("Channel").Value
                            dr1("Channel") = row.Cells.Item("Channel").Value
                            '  Dim spotrow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.GetSpotRow(spots(index).Spot)
                            dr("Date") = row.Cells.Item("Date").Value
                            dr1("Date") = row.Cells.Item("Date").Value
                            dr("Start Time") = row.Cells.Item("Start Time").Value
                            dr1("Start Time") = row.Cells.Item("Start Time").Value
                            'dRow("StartTime") = spotrow("StartTime")
                            ''dRow("EndTime") = spotrow("EndTime")
                            'dr("End Time") = row.Cells.Item("End Time").Value
                            'dr1("End Time") = row.Cells.Item("End Time").Value
                            dr("Duration(Sec)") = row.Cells.Item("Duration(Sec)").Value
                            dr1("Duration(Sec)") = row.Cells.Item("Duration(Sec)").Value
                            'dRow("Duration(Sec)") = TimeSpan.Parse(dRow("EndTime").ToString()).Subtract(TimeSpan.Parse(dRow("StartTime").ToString())).TotalSeconds
                            'dRow("PA") = spotrow("PA")
                            dr("PA") = row.Cells.Item("PA").Value
                            dr1("PA") = row.Cells.Item("PA").Value
                            '  dr(spots(index).MG + "TVR") = spots(index).TVRVal.Split({","c}, StringSplitOptions.None)(0)
                            ' dRow("TA") = spotrow("TA")
                            dr("TA") = row.Cells.Item("TA").Value
                            dr1("TA") = row.Cells.Item("TA").Value
                            '  dr("Commercial") = row.Cells("Commercial").Value
                            dr("Cost") = row.Cells.Item("Cost").Value
                            dr1("Cost") = row.Cells.Item("Cost").Value
                            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Add(dr)
                            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.AcceptChanges()
                            selecpots.Rows.Add(dr1)
                            dgvAvailableSpotsGrid.Rows.Remove(row)
                        Next
                    ElseIf ((count - selecpots.Rows.Count) > 0 And ((count - selecpots.Rows.Count) < dgvAvailableSpotsGrid.SelectedRows.Count)) Then
                        ' Windows.Forms.MessageBox.Show(String.Format("Maximum of {0} spot(s) can be selected.Please reselect {0} spot(s)", count - selecpots.Rows.Count))
                        ShowErrorLabel(String.Format("Maximum of {0} spot(s) can be selected.Please reselect {0} spot(s)", count - selecpots.Rows.Count))
                    Else
                        '  Windows.Forms.MessageBox.Show("Maximum number of spot(s) has been selected.Please remove unwanted and/or increase number of required spot(s) count and Try again")
                        ShowErrorLabel("Maximum number of spot(s) has been selected.Please remove unwanted and/or increase number of required spot(s) count and Try again")
                    End If

                    dgSelectedSpotsGrid.DataSource = Nothing
                    dgSelectedSpotsGrid.DataSource = selecpots
                    Globals.Ribbons.MSprintExRibbon.HideSelectedSpotsGrid(dgSelectedSpotsGrid)
                End If
            End If
        Catch ex As Exception
            LogMpsrintExException("Exception occured while selecting spot" + ex.Message)
        End Try
    End Sub
    Private Function RemoveSelectedSpots()
        Try
            Dim selecpots As DataTable = DirectCast(dgSelectedSpotsGrid.DataSource, DataTable)
            Dim avaispots As Data.DataTable = CType(dgvAvailableSpotsGrid.DataSource, Data.DataTable)
            For Each row As DataGridViewRow In dgSelectedSpotsGrid.SelectedRows
                '   Dim row As DataGridViewRow = dgSelectedSpotsGrid
                Dim filter As String = String.Format("Spot ='{0}'", row.Cells.Item("Spot").Value)
                Dim row1 As DataRow = selecpots.Select(filter)(0)
                Dim row2 As DataRow = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Select(filter)(0)

                If Not (avaispots Is Nothing) Then

                    Dim dr As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.NewRow()
                    '  Dim dr As DataRow = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.NewRow()
                    ' Dim dr1 As DataRow = avaispots.NewRow()
                    dr("GUID") = row2("GUID").ToString()
                    ' dr1("GUID") = row2("GUID").ToString()
                    dr("Spot") = row2("Spot").ToString()
                    '  dr1("Spot") = row2("Spot").ToString()
                    dr("Start Date") = row2("Start Date").ToString()
                    ' dr1("Start Date") = row2("Start Date").ToString()
                    dr("End Date") = row2("End Date").ToString()
                    '  dr1("End Date") = row2("End Date").ToString()
                    dr("WeekNum") = row2("WeekNum").ToString()
                    '  dr1("WeekNum") = row2("WeekNum").ToString()
                    dr("Channel") = row2("Channel").ToString()
                    '  dr1("Channel") = row2("Channel").ToString()
                    '  Dim spotrow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.GetSpotRow(spots(index).Spot)
                    dr("Date") = row2("Date").ToString()
                    ' dr1("Date") = row2("Date").ToString()
                    dr("Start Time") = row2("Start Time").ToString()
                    '  dr1("Start Time") = row2("Start Time").ToString()
                    'dRow("StartTime") = spotrow("StartTime")
                    ''dRow("EndTime") = spotrow("EndTime")
                    'dr("End Time") = row.Cells.Item("End Time").Value
                    'dr1("End Time") = row.Cells.Item("End Time").Value
                    dr("Duration(Sec)") = row2("Duration(Sec)").ToString()
                    ' dr1("Duration(Sec)") = row2("Duration(Sec)").ToString()
                    'dRow("Duration(Sec)") = TimeSpan.Parse(dRow("EndTime").ToString()).Subtract(TimeSpan.Parse(dRow("StartTime").ToString())).TotalSeconds
                    'dRow("PA") = spotrow("PA")
                    dr("PA") = row2("PA").ToString()
                    ' dr1("PA") = row2("PA").ToString()
                    '  dr(spots(index).MG + "TVR") = spots(index).TVRVal.Split({","c}, StringSplitOptions.None)(0)
                    ' dRow("TA") = spotrow("TA")
                    dr("TA") = row2("TA").ToString()
                    ' dr1("TA") = row2("TA").ToString()
                    '  dr("Commercial") = row.Cells("Commercial").Value
                    dr("Cost") = row2("Cost").ToString()
                    ' dr1("Cost") = row2("Cost").ToString()
                    Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Rows.Add(dr)
                    Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.AcceptChanges()
                    ' avaispots.Rows.Add(dr1)

                End If
                selecpots.Rows.Remove(row1)
                Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Remove(row2)
                Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.AcceptChanges()
                ' Dim dr As DataRow = avaipots.NewRow()
                'For c As Integer = 0 To dgSelectedSpotsGrid.ColumnCount - 1

                '    If Not (avaipots.Columns.Contains(dgSelectedSpotsGrid.Columns(c).Name)) Then
                '        avaipots.Columns.Add(dgSelectedSpotsGrid.Columns(c).Name)
                '    End If
                '    dr(c) = row.Cells(c).Value
                'Next
                'avaipots.Rows.Add(dr)
                'dgvAvailableSpotsGrid.DataSource = avaipots
                'dgSelectedSpotsGrid.Rows.Remove(row)
            Next
          
            If Not (avaispots Is Nothing) Then
                Dim fil As String = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
                If cbWeeks.Text.Equals("All") Then
                    fil = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
                Else

                    fil = String.Format("GUID='{0}' and WeekNum={1}", Globals.Ribbons.MSprintExRibbon.currentLineItem, Int32.Parse(cbWeeks.Text).ToString())
                End If


                Dim rows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Select(fil)
                avaispots = rows.CopyToDataTable()
                dgvAvailableSpotsGrid.DataSource = avaispots
                Globals.Ribbons.MSprintExRibbon.HideSelectedSpotsGrid(dgvAvailableSpotsGrid)
            End If

            dgSelectedSpotsGrid.DataSource = Nothing
            dgSelectedSpotsGrid.DataSource = selecpots
            Globals.Ribbons.MSprintExRibbon.HideSelectedSpotsGrid(dgSelectedSpotsGrid)
        Catch ex As Exception
            LogMpsrintExException("Exception occured while Removing selected spots from grid." + ex.Message)
            Throw ex
        End Try
    End Function
    Private Function ShowErrorLabel(ByVal errMessage As String)
        Try
            lbErrorLabel.Text = errMessage
            lbErrorLabel.Visible = True
            lbErrorLabel.Refresh()
        Catch ex As Exception

        End Try
    End Function
    Private Function HideErrorLabel()
        Try
            lbErrorLabel.Visible = False
            lbErrorLabel.Refresh()
        Catch ex As Exception

        End Try
    End Function

    Private Sub dgSelectedSpotsGrid_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgSelectedSpotsGrid.CellEnter

    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub lbErrorLabel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbErrorLabel.Click

    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbProg.Click

    End Sub
End Class
