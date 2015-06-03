Imports System
Imports System.Data
Imports System.IO
Module ReachNFrequency
    Dim tgDirectoryPath As String = AppDomain.CurrentDomain.BaseDirectory & "\\Masters\TGS\\"
    Dim mgDirectoryPath As String = AppDomain.CurrentDomain.BaseDirectory & "\\Masters\MGS\\"
    Public Function ConstructChannelSummaryTable(ByVal opXml As XElement) As Boolean
        Dim generated As Boolean = True
        Globals.Ribbons.MSprintExRibbon.RnFChannelSummary = New Data.DataTable()
        Try
            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("WeekNum", Type.GetType("System.Int32"))
            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("Channel")
            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("Market")
            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("Universe")
            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("GRP")
            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("AOTS")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("1+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("2+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("3+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("4+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("5+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("6+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("7+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("8+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("9+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("10+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("11+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("12+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("13+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("14+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("15+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("16+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("17+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("18+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("19+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("20+")
            If Int32.TryParse(opXml.Element("plan").Attribute("HReachNumber").Value, Globals.Ribbons.MSprintExRibbon.ChannelSummaryHRN) Then
                Globals.Ribbons.MSprintExRibbon.ChannelSummaryHRN += 1
                For index = 1 To Globals.Ribbons.MSprintExRibbon.ChannelSummaryHRN
                    Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add(index.ToString() + "+")
                Next
            Else
                Globals.Ribbons.MSprintExRibbon.ChannelSummaryHRN = 20
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("1+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("2+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("3+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("4+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("5+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("6+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("7+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("8+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("9+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("10+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("11+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("12+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("13+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("14+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("15+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("16+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("17+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("18+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("19+")
                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("20+")

            End If
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("MidDateGRP000s")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("MidDateGRP")
            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("Gross Outlay")
            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("CPRP")
            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("CPT")

            For Each period As XElement In opXml.Element("plan").Elements
                For Each channel As XElement In period.Elements
                    If period.Attribute("WeekNum").Value = String.Empty Then
                        Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns("WeekNum").DefaultValue = 0
                    Else
                        Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns("WeekNum").DefaultValue = Convert.ToInt32(period.Attribute("WeekNum").Value)
                    End If
                    Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns("Channel").DefaultValue = channel.Attribute("name").Value
                    Dim tlmReach As XElement
                    Dim tlmPeriod As XElement
                    For Each reach As XElement In channel.Elements
                        '   For Each reachElement As XElement In channel.Element("reach").Elements
                        If reach.Attribute("tm").Value.Contains("TotalMarkets") Then
                            tlmReach = reach
                            tlmPeriod = period
                        Else
                            ReadReachChannelSummary(reach, period)
                        End If
                        '    newrow("Gross Outlay") = goutlay / 10
                        'Catch ex As Exception
                        '    newrow("Gross Outlay") = String.Empty
                        'End Try
                        'Try
                        '    newrow("CPRP") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(newrow("GRP"))
                        'Catch ex As Exception
                        '    newrow("CPRP") = String.Empty
                        'End Try
                        'Try
                        '    newrow("CPT") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(finalrow("Reach000s").ToString())
                        'Catch ex As Exception
                        '    newrow("CPT") = String.Empty
                        'End Try
                    Next
                    ReadReachChannelSummary(tlmReach, tlmPeriod)
                Next
            Next
            '  Next
            'Dim dtweekss As Data.DataTable = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks
            'Dim planitems As Data.DataTable = CType(loSpotSelection.DataSource, Data.DataTable)
            'Dim planchannels As Data.DataTable = Globals.Ribbons.MSprintExRibbon.GetGridTable()
            'Dim pchannels As Data.DataTable = planitems.Copy()
            ''  RnFSelectedSpots.DefaultView.ToTable(True, New String() {"GUID", "Spot", "Start Date", "End Date", "WeekNum", "Channel", "Date", "Start Time", "Duration(Sec)", "PA", "TA", "Cost"})
            'pchannels = pchannels.DefaultView.ToTable(True, New String() {"Channel"})

            'If planchannels.Rows.Count = 0 Then

            '    If Globals.Ribbons.MSprintExRibbon.mappedchannels Is Nothing Then
            '        planchannels = New Data.DataTable()
            '    Else
            '        planchannels = Globals.Ribbons.MSprintExRibbon.mappedchannels
            '    End If

            'End If
            'Dim mgs As List(Of String) = New List(Of String)()
            'For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
            '    mgs.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString())
            'Next
            'mgs.Add("TotalMarkets")
            ''  If ConstructOpRnFTable(opXml) Then

            'If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbSingle.Checked Then

            '    For Each channel As Data.DataRow In pchannels.Rows
            '        Dim goutlay As Decimal = 0

            '        'dt.Columns.Add("PCName")
            '        'dt.Columns.Add("MCName")
            '        Dim filterexp As String = String.Format("Channel='{0}'", channel("Channel").ToString())
            '        Dim prows As Data.DataRow() = planitems.Select(filterexp)

            '        If prows.Length > 0 Then
            '            For Each row As Data.DataRow In prows
            '                Dim cost, spots, duration, mulvalue As Decimal

            '                If Not (Decimal.TryParse(row("RatePer10Sec").ToString(), cost)) Then
            '                    cost = 0
            '                End If

            '                If Not (Decimal.TryParse(row("Total Spots").ToString(), spots)) Then
            '                    spots = 0
            '                End If

            '                If Not (Decimal.TryParse(row("Duration").ToString(), duration)) Then
            '                    duration = 0
            '                End If
            '                mulvalue = cost * spots * duration
            '                goutlay = goutlay + mulvalue

            '            Next
            '        End If

            '        For index = 0 To mgs.Count - 1
            '            Dim filterstring As String = String.Format("PCName='{0}'", channel("Channel").ToString())
            '            Dim mcname As String = planchannels.Select(filterstring)(0)("MCName").ToString()
            '            Dim filter As String = String.Format("Channel='{0}' and MG='{1}'", mcname, mgs(index))
            '            Dim rows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Select(filter)

            '            If rows.Length > 0 Then
            '                '   Dim temp As Data.DataTable = rows.CopyToDataTable()
            '                Dim finalrow As Data.DataRow = rows.AsEnumerable().Last()
            '                Dim newrow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.NewRow()
            '                newrow("WeekNum") = finalrow("WeekNum").ToString()
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("MG")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Channel")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Date") 'spot date
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Day")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Start Time")
            '                '' RnfShowResultsTable.Columns.Add("End Time")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Programme")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("PA")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("TA")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("TVR000s")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("TVR")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("GRP000s")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("GRP")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("AvgFreq")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("CummCost")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("SpotCPRP")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("CummCPRP")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Reach000s")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("1+")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("2+")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("3+")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("4+")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("5+")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("6+")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("7+")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("8+")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("9+")
            '                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("10+")
            '                newrow("Channel") = finalrow("Channel").ToString()
            '                newrow("Market") = mgs(index)
            '                Dim periodElement As XElement = Globals.Ribbons.MSprintExRibbon.GetPeriodElementForMG(finalrow("MG").ToString(), String.Empty)
            '                Try
            '                    newrow("Universe") = periodElement.Attribute("universe").Value
            '                Catch ex As Exception
            '                    newrow("Universe") = String.Empty
            '                End Try
            '                newrow("GRP") = finalrow("GRP").ToString()
            '                newrow("1+") = finalrow("1+").ToString()
            '                Try
            '                    newrow("AOTS") = Convert.ToDecimal(finalrow("GRP").ToString()) / Convert.ToDecimal(finalrow("1+").ToString())
            '                Catch ex As Exception
            '                    newrow("AOTS") = String.Empty
            '                End Try
            '                newrow("2+") = finalrow("2+").ToString()
            '                newrow("3+") = finalrow("3+").ToString()
            '                newrow("4+") = finalrow("4+").ToString()
            '                newrow("5+") = finalrow("5+").ToString()
            '                newrow("6+") = finalrow("6+").ToString()
            '                newrow("7+") = finalrow("7+").ToString()
            '                newrow("8+") = finalrow("8+").ToString()
            '                newrow("9+") = finalrow("9+").ToString()
            '                newrow("10+") = finalrow("10+").ToString()
            '                Try
            '                    newrow("Gross Outlay") = goutlay / 10
            '                Catch ex As Exception
            '                    newrow("Gross Outlay") = String.Empty
            '                End Try
            '                Try
            '                    newrow("CPRP") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(newrow("GRP"))
            '                Catch ex As Exception
            '                    newrow("CPRP") = String.Empty
            '                End Try
            '                Try
            '                    newrow("CPT") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(finalrow("Reach000s").ToString())
            '                Catch ex As Exception
            '                    newrow("CPT") = String.Empty
            '                End Try

            '                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Rows.Add(newrow)
            '            End If
            '        Next
            '    Next
            'ElseIf Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
            '    '  Dim planitems As Data.DataTable = CType(loSpotSelection.DataSource, Data.DataTable)
            '    For Each channel As Data.DataRow In pchannels.Rows


            '        'dt.Columns.Add("PCName")
            '        'dt.Columns.Add("MCName")
            '        Dim filterexp As String = String.Format("Channel='{0}'", channel("Channel").ToString())
            '        Dim prows As Data.DataRow() = planitems.Select(filterexp)
            '        For index1 = 0 To mgs.Count - 1
            '            Dim filterstring As String = String.Format("PCName='{0}'", channel("Channel").ToString())
            '            Dim mcname As String = planchannels.Select(filterstring)(0)("MCName").ToString()
            '            For index = 0 To dtweekss.Rows.Count - 1

            '                Dim filter As String = String.Format("Channel='{0}' and WeekNum={1} and MG='{2}'", mcname, Convert.ToInt32(dtweekss.Rows(index)("WeekNumber").ToString()), mgs(index1))
            '                Dim rows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Select(filter)

            '                If rows.Length > 0 Then
            '                    '   Dim temp As Data.DataTable = rows.CopyToDataTable()
            '                    Dim finalrow As Data.DataRow = rows.AsEnumerable().Last()
            '                    Dim newrow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.NewRow()
            '                    newrow("WeekNum") = finalrow("WeekNum").ToString()
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("MG")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Channel")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Date") 'spot date
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Day")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Start Time")
            '                    '' RnfShowResultsTable.Columns.Add("End Time")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Programme")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("PA")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("TA")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("TVR000s")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("TVR")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("GRP000s")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("GRP")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("AvgFreq")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("CummCost")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("SpotCPRP")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("CummCPRP")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Reach000s")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("1+")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("2+")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("3+")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("4+")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("5+")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("6+")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("7+")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("8+")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("9+")
            '                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("10+")
            '                    newrow("Channel") = finalrow("Channel").ToString()
            '                    Dim goutlay As Decimal = 0
            '                    If prows.Length > 0 Then
            '                        For Each row As Data.DataRow In prows
            '                            Dim cost, spots, duration, mulvalue As Decimal

            '                            If Not (Decimal.TryParse(row("RatePer10Sec").ToString(), cost)) Then
            '                                cost = 0
            '                            End If
            '                            Dim col As String = "Week " & finalrow("WeekNum").ToString()

            '                            If Not (Decimal.TryParse(row(col).ToString(), spots)) Then
            '                                spots = 0
            '                            End If

            '                            If Not (Decimal.TryParse(row("Duration").ToString(), duration)) Then
            '                                duration = 0
            '                            End If
            '                            mulvalue = cost * spots * duration
            '                            goutlay = goutlay + mulvalue

            '                        Next
            '                    End If
            '                    newrow("Market") = finalrow("MG").ToString()
            '                    Dim periodElement As XElement = Globals.Ribbons.MSprintExRibbon.GetPeriodElementForMG(finalrow("MG").ToString(), String.Empty)
            '                    Try
            '                        newrow("Universe") = periodElement.Attribute("universe").Value
            '                    Catch ex As Exception
            '                        newrow("Universe") = String.Empty
            '                    End Try
            '                    newrow("GRP") = finalrow("GRP").ToString()
            '                    newrow("1+") = finalrow("1+").ToString()
            '                    Try
            '                        newrow("AOTS") = Convert.ToDecimal(finalrow("GRP").ToString()) / Convert.ToDecimal(finalrow("1+").ToString())
            '                    Catch ex As Exception
            '                        newrow("AOTS") = String.Empty
            '                    End Try
            '                    newrow("2+") = finalrow("2+").ToString()
            '                    newrow("3+") = finalrow("3+").ToString()
            '                    newrow("4+") = finalrow("4+").ToString()
            '                    newrow("5+") = finalrow("5+").ToString()
            '                    newrow("6+") = finalrow("6+").ToString()
            '                    newrow("7+") = finalrow("7+").ToString()
            '                    newrow("8+") = finalrow("8+").ToString()
            '                    newrow("9+") = finalrow("9+").ToString()
            '                    newrow("10+") = finalrow("10+").ToString()
            '                    'newrow("Gross Outlay") = String.Empty
            '                    'newrow("CPRP") = String.Empty
            '                    'newrow("CPT") = String.Empty
            '                    Try
            '                        newrow("Gross Outlay") = goutlay / 10
            '                    Catch ex As Exception
            '                        newrow("Gross Outlay") = String.Empty
            '                    End Try
            '                    Try
            '                        newrow("CPRP") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(newrow("GRP"))
            '                    Catch ex As Exception
            '                        newrow("CPRP") = String.Empty
            '                    End Try
            '                    Try
            '                        newrow("CPT") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(finalrow("Reach000s").ToString())
            '                    Catch ex As Exception
            '                        newrow("CPT") = String.Empty
            '                    End Try

            '                    Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Rows.Add(newrow)
            '                End If
            '            Next
            '        Next
            '    Next
            'End If
            '   End If
        Catch ex As Exception
            generated = False
            LogMpsrintExException("Exception occured while constructing channel summary table." + ex.Message)
            Throw ex
        End Try
        Return generated
    End Function
    Public Function ReadReachChannelSummary(ByVal reach As XElement, ByVal period As XElement)
        Try
            Dim newrow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.NewRow()
            'reach = reach.Element("reach")
            'newrow("Market") = reach.Attribute("market").Value.Split("~")(1)

            Try
                newrow("Market") = reach.Attribute("market").Value.Split("~")(1)
            Catch ex As Exception
                newrow("Market") = reach.Attribute("tm").Value.Split("~")(1)
            End Try
            '  Dim periodElement As XElement = Globals.Ribbons.MSprintExRibbon.GetPeriodElementForMG(finalrow("MG").ToString(), String.Empty)
            Try
                newrow("Universe") = reach.Attribute("universe").Value
            Catch ex As Exception
                newrow("Universe") = String.Empty
            End Try
            newrow("GRP") = reach.Attribute("GRP").Value
            ' newrow("1+") = reach.Attribute("R1").Value
           
            'newrow("2+") = reach.Attribute("R2").Value
            'newrow("3+") = reach.Attribute("R3").Value
            'newrow("4+") = reach.Attribute("R4").Value
            'newrow("5+") = reach.Attribute("R5").Value
            'newrow("6+") = reach.Attribute("R6").Value
            'newrow("7+") = reach.Attribute("R7").Value
            'newrow("8+") = reach.Attribute("R8").Value
            'newrow("9+") = reach.Attribute("R9").Value
            'newrow("10+") = reach.Attribute("R10").Value
            'newrow("11+") = reach.Attribute("R11").Value
            'newrow("12+") = reach.Attribute("R12").Value
            'newrow("13+") = reach.Attribute("R13").Value
            'newrow("14+") = reach.Attribute("R14").Value
            'newrow("15+") = reach.Attribute("R15").Value
            'newrow("16+") = reach.Attribute("R16").Value
            'newrow("17+") = reach.Attribute("R17").Value
            'newrow("18+") = reach.Attribute("R18").Value
            'newrow("19+") = reach.Attribute("R19").Value
            'newrow("20+") = reach.Attribute("R20").Value
            For index = 1 To Globals.Ribbons.MSprintExRibbon.ChannelSummaryHRN
                Dim colname As String = index.ToString() + "+"
                Dim attrName As String = "R" + index.ToString()
                '  If index <= vals.Length - 5 Then
                If reach.Attributes(attrName).Count > 0 Then
                    newrow(colname) = reach.Attribute(attrName).Value
                Else
                    newrow(colname) = 0
                End If


            Next
            Try
                newrow("AOTS") = Convert.ToDecimal(reach.Attribute("GRP").Value) / Convert.ToDecimal(reach.Attribute("R1").Value)
            Catch ex As Exception
                newrow("AOTS") = String.Empty
            End Try
            'newrow("MidDateGRP000s") = reach.Attribute("midDateGRP000s").Value
            'newrow("MidDateGRP") = reach.Attribute("midDateGRP").Value
            'newrow("") = reach.Attribute("R20").Value
            'newrow("") = reach.Attribute("R20").Value
            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Rows.Add(newrow)
            'Try
        Catch ex As Exception

        End Try
    End Function
    Public Function ConstructDurationSummaryTable(ByVal opXML As XElement) As Boolean
        Dim generated As Boolean = True
        Globals.Ribbons.MSprintExRibbon.RnFDurationSummary = New Data.DataTable()
        Try
            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("WeekNum", Type.GetType("System.Int32"))
            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("Duration")
            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("Market")
            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("Universe")
            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("GRP")
            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("AOTS")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("1+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("2+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("3+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("4+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("5+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("6+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("7+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("8+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("9+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("10+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("11+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("12+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("13+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("14+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("15+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("16+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("17+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("18+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("19+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("20+")
            If Int32.TryParse(opXML.Element("plan").Attribute("HReachNumber").Value, Globals.Ribbons.MSprintExRibbon.DSHRN) Then
                Globals.Ribbons.MSprintExRibbon.DSHRN += 1
                For index = 1 To Globals.Ribbons.MSprintExRibbon.DSHRN
                    Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add(index.ToString() + "+")
                Next
            Else
                Globals.Ribbons.MSprintExRibbon.DSHRN = 20
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("1+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("2+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("3+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("4+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("5+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("6+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("7+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("8+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("9+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("10+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("11+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("12+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("13+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("14+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("15+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("16+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("17+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("18+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("19+")
                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("20+")

            End If
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("MidDateGRP000s")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("MidDateGRP")
            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("Gross Outlay")
            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("CPRP")
            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("CPT")
            'duration 6-25-30
            For Each period As XElement In opXML.Element("plan").Elements
                For Each duration As XElement In period.Elements
                    If period.Attribute("WeekNum").Value = String.Empty Then
                        Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns("WeekNum").DefaultValue = 0
                    Else
                        Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns("WeekNum").DefaultValue = Convert.ToInt32(period.Attribute("WeekNum").Value)
                    End If
                    Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns("Duration").DefaultValue = duration.Attribute("value").Value
                    Dim tlmreach As XElement
                    Dim tlmperiod As XElement

                    For Each reach As XElement In duration.Elements
                        '   For Each reachElement As XElement In channel.Element("reach").Elements
                        If reach.Attribute("tm").Value.Contains("TotalMarkets") Then
                            tlmreach = reach
                            tlmperiod = period
                        Else
                            ReachDurationSummary(reach, period)
                        End If

                        '    newrow("Gross Outlay") = goutlay / 10
                        'Catch ex As Exception
                        '    newrow("Gross Outlay") = String.Empty
                        'End Try
                        'Try
                        '    newrow("CPRP") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(newrow("GRP"))
                        'Catch ex As Exception
                        '    newrow("CPRP") = String.Empty
                        'End Try
                        'Try
                        '    newrow("CPT") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(finalrow("Reach000s").ToString())
                        'Catch ex As Exception
                        '    newrow("CPT") = String.Empty
                        'End Try
                    Next
                    ReachDurationSummary(tlmreach, tlmperiod)
                Next
            Next
        Catch ex As Exception
            generated = False
            LogMpsrintExException("Exception occured while constructing duration summary table." + ex.Message)
            Throw ex
        End Try
        Return generated
    End Function
    Public Function ReachDurationSummary(ByVal reach As XElement, ByVal period As XElement)
        Try
            Dim newrow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.NewRow()
            ' reach = reach.Element("reach")
            '  newrow("Market") = reach.Attribute("market").Value.Split("~")(1)

            Try
                newrow("Market") = reach.Attribute("market").Value.Split("~")(1)
            Catch ex As Exception
                newrow("Market") = reach.Attribute("tm").Value.Split("~")(1)
            End Try
            '  Dim periodElement As XElement = Globals.Ribbons.MSprintExRibbon.GetPeriodElementForMG(finalrow("MG").ToString(), String.Empty)
            Try
                newrow("Universe") = reach.Attribute("universe").Value
            Catch ex As Exception
                newrow("Universe") = String.Empty
            End Try
            newrow("GRP") = reach.Attribute("GRP").Value
            '  newrow("1+") = reach.Attribute("R1").Value
           
            'newrow("2+") = reach.Attribute("R2").Value
            'newrow("3+") = reach.Attribute("R3").Value
            'newrow("4+") = reach.Attribute("R4").Value
            'newrow("5+") = reach.Attribute("R5").Value
            'newrow("6+") = reach.Attribute("R6").Value
            'newrow("7+") = reach.Attribute("R7").Value
            'newrow("8+") = reach.Attribute("R8").Value
            'newrow("9+") = reach.Attribute("R9").Value
            'newrow("10+") = reach.Attribute("R10").Value
            'newrow("11+") = reach.Attribute("R10").Value
            'newrow("12+") = reach.Attribute("R10").Value
            'newrow("13+") = reach.Attribute("R10").Value
            'newrow("14+") = reach.Attribute("R10").Value
            'newrow("15+") = reach.Attribute("R10").Value
            'newrow("16+") = reach.Attribute("R10").Value
            'newrow("17+") = reach.Attribute("R17").Value
            'newrow("18+") = reach.Attribute("R18").Value
            'newrow("19+") = reach.Attribute("R19").Value
            'newrow("20+") = reach.Attribute("R20").Value
            For index = 1 To Globals.Ribbons.MSprintExRibbon.DSHRN
                Dim colname As String = index.ToString() + "+"
                Dim attrName As String = "R" + index.ToString()
                '  If index <= vals.Length - 5 Then
                If reach.Attributes(attrName).Count > 0 Then
                    newrow(colname) = reach.Attribute(attrName).Value
                Else
                    newrow(colname) = 0
                End If


            Next
            Try
                newrow("AOTS") = Convert.ToDecimal(reach.Attribute("GRP").Value) / Convert.ToDecimal(newrow("R1").ToString())
            Catch ex As Exception
                newrow("AOTS") = String.Empty
            End Try
            'newrow("MidDateGRP000s") = reach.Attribute("midDateGRP000s").Value
            'newrow("MidDateGRP") = reach.Attribute("midDateGRP").Value
            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Rows.Add(newrow)
            'Try
        Catch ex As Exception

        End Try
    End Function
    Public Function ReachCreativeSummary(ByVal reach As XElement, ByVal period As XElement)
        Try
            Dim newrow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.NewRow()
            ' reach = reach.Element("reach")

            Try
                newrow("Market") = reach.Attribute("market").Value.Split("~")(1)
            Catch ex As Exception
                newrow("Market") = reach.Attribute("tm").Value.Split("~")(1)
            End Try


            '  Dim periodElement As XElement = Globals.Ribbons.MSprintExRibbon.GetPeriodElementForMG(finalrow("MG").ToString(), String.Empty)
            Try
                newrow("Universe") = reach.Attribute("universe").Value
            Catch ex As Exception
                newrow("Universe") = String.Empty
            End Try
            newrow("GRP") = reach.Attribute("GRP").Value
            '  newrow("1+") = reach.Attribute("R1").Value
           
            'newrow("2+") = reach.Attribute("R2").Value
            'newrow("3+") = reach.Attribute("R3").Value
            'newrow("4+") = reach.Attribute("R4").Value
            'newrow("5+") = reach.Attribute("R5").Value
            'newrow("6+") = reach.Attribute("R6").Value
            'newrow("7+") = reach.Attribute("R7").Value
            'newrow("8+") = reach.Attribute("R8").Value
            'newrow("9+") = reach.Attribute("R9").Value
            'newrow("10+") = reach.Attribute("R10").Value
            'newrow("11+") = reach.Attribute("R10").Value
            'newrow("12+") = reach.Attribute("R10").Value
            'newrow("13+") = reach.Attribute("R10").Value
            'newrow("14+") = reach.Attribute("R10").Value
            'newrow("15+") = reach.Attribute("R10").Value
            'newrow("16+") = reach.Attribute("R10").Value
            'newrow("17+") = reach.Attribute("R17").Value
            'newrow("18+") = reach.Attribute("R18").Value
            'newrow("19+") = reach.Attribute("R19").Value
            'newrow("20+") = reach.Attribute("R20").Value
            For index = 1 To Globals.Ribbons.MSprintExRibbon.CreativeSummaryHRN
                Dim colname As String = index.ToString() + "+"
                Dim attrName As String = "R" + index.ToString()
                '  If index <= vals.Length - 5 Then
                If reach.Attributes(attrName).Count > 0 Then
                    newrow(colname) = reach.Attribute(attrName).Value
                Else
                    newrow(colname) = 0
                End If


            Next
            Try
                newrow("AOTS") = Convert.ToDecimal(reach.Attribute("GRP").Value) / Convert.ToDecimal(reach.Attribute("R1").Value)
            Catch ex As Exception
                newrow("AOTS") = String.Empty
            End Try
            'newrow("MidDateGRP000s") = reach.Attribute("midDateGRP000s").Value
            'newrow("MidDateGRP") = reach.Attribute("midDateGRP").Value
            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Rows.Add(newrow)
        Catch ex As Exception
            LogMpsrintExException("Exception occured while adding reach values to creative table.Message: " + ex.Message)
        End Try
    End Function
    Public Function ConstructDurationSummaryTable(ByVal opXML As XElement, ByVal duration As String) As Boolean
        Dim generated As Boolean = True
        Try
            'RnFCreativeSummary.Columns.Add("WeekNum")
            'RnFCreativeSummary.Columns.Add("Creative")
            'RnFCreativeSummary.Columns.Add("Market")
            'RnFCreativeSummary.Columns.Add("Universe")
            'RnFCreativeSummary.Columns.Add("GRP")
            'RnFCreativeSummary.Columns.Add("AOTS")
            'RnFCreativeSummary.Columns.Add("1+")
            'RnFCreativeSummary.Columns.Add("2+")
            'RnFCreativeSummary.Columns.Add("3+")
            'RnFCreativeSummary.Columns.Add("4+")
            'RnFCreativeSummary.Columns.Add("5+")
            'RnFCreativeSummary.Columns.Add("6+")
            'RnFCreativeSummary.Columns.Add("7+")
            'RnFCreativeSummary.Columns.Add("8+")
            'RnFCreativeSummary.Columns.Add("9+")
            'RnFCreativeSummary.Columns.Add("10+")
            Dim dtweekss As Data.DataTable = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks
            Dim mgs As List(Of String) = New List(Of String)()
            For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                mgs.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString())
            Next
            mgs.Add("TotalMarkets")
            If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                Dim weeknum As String = String.Empty

                '   For index1 = 0 To Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows.Count - 1

                For index = 0 To mgs.Count - 1
                    Dim mgroup As String = mgs(index)
                    Dim periodElement As XElement = Globals.Ribbons.MSprintExRibbon.GetPeriodElementForMG(mgroup, String.Empty)
                    Dim row As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.NewRow()
                    row("WeekNum") = weeknum
                    '   row("Duration") = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Creative").ToString()
                    row("Duration") = duration
                    row("Market") = mgroup
                    Try
                        row("Universe") = periodElement.Attribute("universe").Value
                    Catch ex As Exception
                        row("Universe") = String.Empty
                    End Try
                    Try
                        row("GRP") = periodElement.Attribute("GRP").Value
                    Catch ex As Exception
                        row("GRP") = String.Empty
                    End Try

                    Try
                        row("1+") = periodElement.Attribute("R1").Value
                    Catch ex As Exception
                        row("1+") = String.Empty
                    End Try
                    Try
                        row("AOTS") = Convert.ToDecimal(row("GRP").ToString()) / Convert.ToDecimal(row("1+").ToString())
                    Catch ex As Exception
                        row("AOTS") = String.Empty
                    End Try
                    Try
                        row("2+") = periodElement.Attribute("R2").Value
                    Catch ex As Exception
                        row("2+") = String.Empty
                    End Try
                    Try
                        row("3+") = periodElement.Attribute("R3").Value
                    Catch ex As Exception
                        row("3+") = String.Empty
                    End Try
                    Try
                        row("4+") = periodElement.Attribute("R4").Value
                    Catch ex As Exception
                        row("4+") = String.Empty
                    End Try
                    Try
                        row("5+") = periodElement.Attribute("R5").Value
                    Catch ex As Exception
                        row("5+") = String.Empty
                    End Try
                    Try
                        row("6+") = periodElement.Attribute("R6").Value
                    Catch ex As Exception
                        row("6+") = String.Empty
                    End Try
                    Try
                        row("7+") = periodElement.Attribute("R7").Value
                    Catch ex As Exception
                        row("7+") = String.Empty
                    End Try
                    Try
                        row("8+") = periodElement.Attribute("R8").Value
                    Catch ex As Exception
                        row("8+") = String.Empty
                    End Try
                    Try
                        row("9+") = periodElement.Attribute("R9").Value
                    Catch ex As Exception
                        row("9+") = String.Empty
                    End Try
                    Try
                        row("10+") = periodElement.Attribute("R10").Value
                    Catch ex As Exception
                        row("10+") = String.Empty
                    End Try
                    Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Rows.Add(row)
                Next
                ' Next
            ElseIf Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                For index1 = 0 To dtweekss.Rows.Count - 1

                    '  For index2 = 0 To Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows.Count - 1

                    For index = 0 To mgs.Count - 1
                        Dim mgroup As String = mgs(index)
                        Dim periodElement As XElement = Globals.Ribbons.MSprintExRibbon.GetPeriodElementForMG(mgroup, dtweekss.Rows(index1)("WeekNumber").ToString())
                        Dim row As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.NewRow()
                        row("WeekNum") = dtweekss.Rows(index1)("WeekNumber").ToString()
                        '  row("Duration") = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index2)("Creative").ToString()
                        row("Duration") = duration
                        row("Market") = mgroup
                        Try
                            row("Universe") = periodElement.Attribute("universe").Value
                        Catch ex As Exception
                            row("Universe") = String.Empty
                        End Try
                        Try
                            row("GRP") = periodElement.Attribute("GRP").Value
                        Catch ex As Exception
                            row("GRP") = String.Empty
                        End Try

                        Try
                            row("1+") = periodElement.Attribute("R1").Value
                        Catch ex As Exception
                            row("1+") = String.Empty
                        End Try
                        Try
                            row("AOTS") = Convert.ToDecimal(row("GRP").ToString()) / Convert.ToDecimal(row("1+").ToString())
                        Catch ex As Exception
                            row("AOTS") = String.Empty
                        End Try
                        Try
                            row("2+") = periodElement.Attribute("R2").Value
                        Catch ex As Exception
                            row("2+") = String.Empty
                        End Try
                        Try
                            row("3+") = periodElement.Attribute("R3").Value
                        Catch ex As Exception
                            row("3+") = String.Empty
                        End Try
                        Try
                            row("4+") = periodElement.Attribute("R4").Value
                        Catch ex As Exception
                            row("4+") = String.Empty
                        End Try
                        Try
                            row("5+") = periodElement.Attribute("R5").Value
                        Catch ex As Exception
                            row("5+") = String.Empty
                        End Try
                        Try
                            row("6+") = periodElement.Attribute("R6").Value
                        Catch ex As Exception
                            row("6+") = String.Empty
                        End Try
                        Try
                            row("7+") = periodElement.Attribute("R7").Value
                        Catch ex As Exception
                            row("7+") = String.Empty
                        End Try
                        Try
                            row("8+") = periodElement.Attribute("R8").Value
                        Catch ex As Exception
                            row("8+") = String.Empty
                        End Try
                        Try
                            row("9+") = periodElement.Attribute("R9").Value
                        Catch ex As Exception
                            row("9+") = String.Empty
                        End Try
                        Try
                            row("10+") = periodElement.Attribute("R10").Value
                        Catch ex As Exception
                            row("10+") = String.Empty
                        End Try
                        Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Rows.Add(row)
                    Next
                Next
                '   Next
            End If
        Catch ex As Exception
            generated = False
            LogMpsrintExException("Exception occured while constructing duration summary table." + ex.Message)
            Throw ex

        End Try
        Return generated
    End Function
    Public Function ConstructCreativeSummaryTable(ByVal opXML As XElement) As Boolean
        Dim generated As Boolean = True
        Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary = New Data.DataTable()
        Try
            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("WeekNum", Type.GetType("System.Int32"))
            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("Creative")
            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("Market")
            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("Universe")
            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("GRP")
            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("AOTS")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("1+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("2+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("3+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("4+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("5+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("6+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("7+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("8+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("9+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("10+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("11+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("12+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("13+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("14+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("15+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("16+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("17+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("18+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("19+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("20+")
            If Int32.TryParse(opXML.Element("plan").Attribute("HReachNumber").Value, Globals.Ribbons.MSprintExRibbon.CreativeSummaryHRN) Then
                Globals.Ribbons.MSprintExRibbon.CreativeSummaryHRN += 1
                For index = 1 To Globals.Ribbons.MSprintExRibbon.CreativeSummaryHRN
                    Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add(index.ToString() + "+")
                Next
            Else
                Globals.Ribbons.MSprintExRibbon.CreativeSummaryHRN = 20
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("1+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("2+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("3+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("4+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("5+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("6+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("7+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("8+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("9+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("10+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("11+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("12+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("13+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("14+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("15+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("16+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("17+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("18+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("19+")
                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("20+")

            End If
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("MidDateGRP000s")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("MidDateGRP")
            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("Gross Outlay")
            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("CPRP")
            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("CPT")

            For Each period As XElement In opXML.Element("plan").Elements
                For Each creative As XElement In period.Elements
                    If period.Attribute("WeekNum").Value = String.Empty Then
                        Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns("WeekNum").DefaultValue = 0
                    Else
                        Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns("WeekNum").DefaultValue = Convert.ToInt32(period.Attribute("WeekNum").Value)
                    End If
                    Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns("Creative").DefaultValue = creative.Attribute("value").Value
                    Dim tlmReach As XElement
                    Dim tlmPeriod As XElement

                    For Each reach As XElement In creative.Elements
                        '   For Each reachElement As XElement In channel.Element("reach").Elements
                        If reach.Attribute("tm").Value.Contains("TotalMarkets") Then
                            tlmReach = reach
                            tlmPeriod = period
                        Else
                            ReachCreativeSummary(reach, period)
                        End If
                        'Try
                        '    newrow("Gross Outlay") = goutlay / 10
                        'Catch ex As Exception
                        '    newrow("Gross Outlay") = String.Empty
                        'End Try
                        'Try
                        '    newrow("CPRP") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(newrow("GRP"))
                        'Catch ex As Exception
                        '    newrow("CPRP") = String.Empty
                        'End Try
                        'Try
                        '    newrow("CPT") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(finalrow("Reach000s").ToString())
                        'Catch ex As Exception
                        '    newrow("CPT") = String.Empty
                        'End Try
                    Next
                    ReachCreativeSummary(tlmReach, tlmPeriod)
                Next
            Next
        Catch ex As Exception
            generated = False
            LogMpsrintExException("Exception occured while constructing creative summary table." + ex.Message)
            Throw ex
        End Try
        Return generated
    End Function
    Public Function ConstructCreativeSummaryTable(ByVal opXML As XElement, ByVal abc As Integer) As Boolean
        Dim generated As Boolean = True
        Try
            'RnFCreativeSummary.Columns.Add("WeekNum")
            'RnFCreativeSummary.Columns.Add("Creative")
            'RnFCreativeSummary.Columns.Add("Market")
            'RnFCreativeSummary.Columns.Add("Universe")
            'RnFCreativeSummary.Columns.Add("GRP")
            'RnFCreativeSummary.Columns.Add("AOTS")
            'RnFCreativeSummary.Columns.Add("1+")
            'RnFCreativeSummary.Columns.Add("2+")
            'RnFCreativeSummary.Columns.Add("3+")
            'RnFCreativeSummary.Columns.Add("4+")
            'RnFCreativeSummary.Columns.Add("5+")
            'RnFCreativeSummary.Columns.Add("6+")
            'RnFCreativeSummary.Columns.Add("7+")
            'RnFCreativeSummary.Columns.Add("8+")
            'RnFCreativeSummary.Columns.Add("9+")
            'RnFCreativeSummary.Columns.Add("10+")
            Dim dtweekss As Data.DataTable = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks
            Dim mgs As List(Of String) = New List(Of String)()
            For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                mgs.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString())
            Next
            mgs.Add("TotalMarkets")
            If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                Dim weeknum As String = String.Empty

                '  For index1 = 0 To Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows.Count - 1

                For index = 0 To mgs.Count - 1
                    Dim mgroup As String = mgs(index)
                    Dim periodElement As XElement = Globals.Ribbons.MSprintExRibbon.GetPeriodElementForMG(mgroup, String.Empty)
                    Dim row As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.NewRow()
                    row("WeekNum") = weeknum
                    '  row("Creative") = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Creative").ToString()
                    '  row("Creative") = creative
                    row("Market") = mgroup
                    Try
                        row("Universe") = periodElement.Attribute("universe").Value
                    Catch ex As Exception
                        row("Universe") = String.Empty
                    End Try
                    Try
                        row("GRP") = periodElement.Attribute("GRP").Value
                    Catch ex As Exception
                        row("GRP") = String.Empty
                    End Try

                    Try
                        row("1+") = periodElement.Attribute("R1").Value
                    Catch ex As Exception
                        row("1+") = String.Empty
                    End Try
                    Try
                        row("AOTS") = Convert.ToDecimal(row("GRP").ToString()) / Convert.ToDecimal(row("1+").ToString())
                    Catch ex As Exception
                        row("AOTS") = String.Empty
                    End Try
                    Try
                        row("2+") = periodElement.Attribute("R2").Value
                    Catch ex As Exception
                        row("2+") = String.Empty
                    End Try
                    Try
                        row("3+") = periodElement.Attribute("R3").Value
                    Catch ex As Exception
                        row("3+") = String.Empty
                    End Try
                    Try
                        row("4+") = periodElement.Attribute("R4").Value
                    Catch ex As Exception
                        row("4+") = String.Empty
                    End Try
                    Try
                        row("5+") = periodElement.Attribute("R5").Value
                    Catch ex As Exception
                        row("5+") = String.Empty
                    End Try
                    Try
                        row("6+") = periodElement.Attribute("R6").Value
                    Catch ex As Exception
                        row("6+") = String.Empty
                    End Try
                    Try
                        row("7+") = periodElement.Attribute("R7").Value
                    Catch ex As Exception
                        row("7+") = String.Empty
                    End Try
                    Try
                        row("8+") = periodElement.Attribute("R8").Value
                    Catch ex As Exception
                        row("8+") = String.Empty
                    End Try
                    Try
                        row("9+") = periodElement.Attribute("R9").Value
                    Catch ex As Exception
                        row("9+") = String.Empty
                    End Try
                    Try
                        row("10+") = periodElement.Attribute("R10").Value
                    Catch ex As Exception
                        row("10+") = String.Empty
                    End Try
                    Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Rows.Add(row)
                Next
                ' Next
            ElseIf Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                For index1 = 0 To dtweekss.Rows.Count - 1

                    ' For index2 = 0 To Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows.Count - 1

                    For index = 0 To mgs.Count - 1
                        Dim mgroup As String = mgs(index)
                        Dim periodElement As XElement = Globals.Ribbons.MSprintExRibbon.GetPeriodElementForMG(mgroup, dtweekss.Rows(index1)("WeekNumber").ToString())
                        Dim row As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.NewRow()
                        row("WeekNum") = dtweekss.Rows(index1)("WeekNumber").ToString()
                        ' row("Creative") = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index2)("Creative").ToString()
                        '  row("Creative") = creative
                        row("Market") = mgroup
                        Try
                            row("Universe") = periodElement.Attribute("universe").Value
                        Catch ex As Exception
                            row("Universe") = String.Empty
                        End Try
                        Try
                            row("GRP") = periodElement.Attribute("GRP").Value
                        Catch ex As Exception
                            row("GRP") = String.Empty
                        End Try

                        Try
                            row("1+") = periodElement.Attribute("R1").Value
                        Catch ex As Exception
                            row("1+") = String.Empty
                        End Try
                        Try
                            row("AOTS") = Convert.ToDecimal(row("GRP").ToString()) / Convert.ToDecimal(row("1+").ToString())
                        Catch ex As Exception
                            row("AOTS") = String.Empty
                        End Try
                        Try
                            row("2+") = periodElement.Attribute("R2").Value
                        Catch ex As Exception
                            row("2+") = String.Empty
                        End Try
                        Try
                            row("3+") = periodElement.Attribute("R3").Value
                        Catch ex As Exception
                            row("3+") = String.Empty
                        End Try
                        Try
                            row("4+") = periodElement.Attribute("R4").Value
                        Catch ex As Exception
                            row("4+") = String.Empty
                        End Try
                        Try
                            row("5+") = periodElement.Attribute("R5").Value
                        Catch ex As Exception
                            row("5+") = String.Empty
                        End Try
                        Try
                            row("6+") = periodElement.Attribute("R6").Value
                        Catch ex As Exception
                            row("6+") = String.Empty
                        End Try
                        Try
                            row("7+") = periodElement.Attribute("R7").Value
                        Catch ex As Exception
                            row("7+") = String.Empty
                        End Try
                        Try
                            row("8+") = periodElement.Attribute("R8").Value
                        Catch ex As Exception
                            row("8+") = String.Empty
                        End Try
                        Try
                            row("9+") = periodElement.Attribute("R9").Value
                        Catch ex As Exception
                            row("9+") = String.Empty
                        End Try
                        Try
                            row("10+") = periodElement.Attribute("R10").Value
                        Catch ex As Exception
                            row("10+") = String.Empty
                        End Try
                        Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Rows.Add(row)
                    Next
                Next
                ' Next
            End If
        Catch ex As Exception
            generated = False
            LogMpsrintExException("Exception occured while constructing Creative Summary table" + ex.Message)
            Throw ex
        End Try
        Return generated
    End Function
    Public Function ConstructAllSummaryTable(ByVal opXML As XElement) As Boolean
        Dim generated As Boolean = True
        Globals.Ribbons.MSprintExRibbon.RnFMarketSummary = New Data.DataTable
        Globals.Ribbons.MSprintExRibbon.RnFChannelSummary = New Data.DataTable
        Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary = New Data.DataTable
        Globals.Ribbons.MSprintExRibbon.RnFDurationSummary = New Data.DataTable
        Try
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("WeekNum")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("Market")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("Universe")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("GRP")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("AOTS")

            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("WeekNum", Type.GetType("System.Int32"))
            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("Channel")
            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("Market")
            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("Universe")
            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("GRP")
            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("AOTS")

            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("WeekNum", Type.GetType("System.Int32"))
            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("Creative")
            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("Market")
            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("Universe")
            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("GRP")
            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("AOTS")

            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("WeekNum", Type.GetType("System.Int32"))
            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("Duration")
            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("Market")
            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("Universe")
            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("GRP")
            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("AOTS")
            If Int32.TryParse(opXML.Element("plan").Attribute("HReachNumber").Value, Globals.Ribbons.MSprintExRibbon.MSHRN) Then
                Globals.Ribbons.MSprintExRibbon.MSHRN += 1
                Globals.Ribbons.MSprintExRibbon.ChannelSummaryHRN = Globals.Ribbons.MSprintExRibbon.MSHRN
                Globals.Ribbons.MSprintExRibbon.CreativeSummaryHRN = Globals.Ribbons.MSprintExRibbon.MSHRN
                Globals.Ribbons.MSprintExRibbon.DSHRN = Globals.Ribbons.MSprintExRibbon.MSHRN

                For index = 1 To Globals.Ribbons.MSprintExRibbon.ChannelSummaryHRN
                    Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add(index.ToString() + "+")
                    Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add(index.ToString() + "+")
                    Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add(index.ToString() + "+")
                    Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add(index.ToString() + "+")
                Next
            Else
                Globals.Ribbons.MSprintExRibbon.MSHRN = 20
                Globals.Ribbons.MSprintExRibbon.ChannelSummaryHRN = Globals.Ribbons.MSprintExRibbon.MSHRN
                Globals.Ribbons.MSprintExRibbon.CreativeSummaryHRN = Globals.Ribbons.MSprintExRibbon.MSHRN
                Globals.Ribbons.MSprintExRibbon.DSHRN = Globals.Ribbons.MSprintExRibbon.MSHRN
                For index = 1 To 20
                    Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add(index.ToString() + "+")
                    Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add(index.ToString() + "+")
                    Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add(index.ToString() + "+")
                    Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add(index.ToString() + "+")
                Next
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("1+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("2+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("3+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("4+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("5+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("6+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("7+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("8+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("9+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("10+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("11+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("12+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("13+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("14+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("15+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("16+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("17+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("18+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("19+")
                'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("20+")

            End If
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("1+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("2+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("3+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("4+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("5+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("6+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("7+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("8+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("9+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("10+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("11+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("12+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("13+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("14+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("15+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("16+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("17+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("18+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("19+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("20+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateTVR000s")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateTVR")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateGRP000s")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateGRP")

           
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("1+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("2+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("3+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("4+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("5+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("6+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("7+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("8+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("9+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("10+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("11+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("12+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("13+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("14+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("15+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("16+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("17+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("18+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("19+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("20+")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("MidDateGRP000s")
            'Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("MidDateGRP")
            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("Gross Outlay")
            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("CPRP")
            Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns.Add("CPT")

          
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("1+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("2+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("3+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("4+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("5+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("6+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("7+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("8+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("9+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("10+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("11+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("12+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("13+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("14+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("15+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("16+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("17+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("18+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("19+")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("20+")
            ''Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("MidDateGRP000s")
            'Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("MidDateGRP")
            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("Gross Outlay")
            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("CPRP")
            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns.Add("CPT")

         
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("1+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("2+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("3+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("4+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("5+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("6+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("7+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("8+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("9+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("10+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("11+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("12+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("13+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("14+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("15+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("16+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("17+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("18+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("19+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("20+")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("MidDateGRP000s")
            'Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("MidDateGRP")
            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("Gross Outlay")
            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("CPRP")
            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns.Add("CPT")
            For Each period As XElement In opXML.Element("plan").Elements
                Dim tlmReach As XElement
                Dim tlmPeriod As XElement

               

                For Each summary As XElement In period.Elements
                    'If period.Attribute("WeekNum").Value = String.Empty Then
                    '    Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns("WeekNum").DefaultValue = 0
                    'Else
                    '    Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns("WeekNum").DefaultValue = Convert.ToInt32(period.Attribute("WeekNum").Value)
                    'End If
                    ''   Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns("Duration").DefaultValue = duration.Attribute("value").Value

                    ''    For Each reach As XElement In period.Element("duration").Elements
                    ''   For Each reachElement As XElement In channel.Element("reach").Elements
                    'Dim newrow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.NewRow()
                    'newrow("Market") = reach.Attribute("tm").Value.Split("~")(1)
                    ''  Dim periodElement As XElement = Globals.Ribbons.MSprintExRibbon.GetPeriodElementForMG(finalrow("MG").ToString(), String.Empty)
                    'Try
                    '    newrow("Universe") = reach.Attribute("universe").Value
                    'Catch ex As Exception
                    '    newrow("Universe") = String.Empty
                    'End Try
                    'newrow("GRP") = reach.Attribute("GRP").Value
                    'newrow("1+") = reach.Attribute("R1").Value
                    'Try
                    '    newrow("AOTS") = Convert.ToDecimal(reach.Attribute("GRP").Value) / Convert.ToDecimal(reach.Attribute("R1").Value)
                    'Catch ex As Exception
                    '    newrow("AOTS") = String.Empty
                    'End Try
                    'newrow("2+") = reach.Attribute("R2").Value
                    'newrow("3+") = reach.Attribute("R3").Value
                    'newrow("4+") = reach.Attribute("R4").Value
                    'newrow("5+") = reach.Attribute("R5").Value
                    'newrow("6+") = reach.Attribute("R6").Value
                    'newrow("7+") = reach.Attribute("R7").Value
                    'newrow("8+") = reach.Attribute("R8").Value
                    'newrow("9+") = reach.Attribute("R9").Value
                    'newrow("10+") = reach.Attribute("R10").Value
                    'newrow("11+") = reach.Attribute("R11").Value
                    'newrow("12+") = reach.Attribute("R12").Value
                    'newrow("13+") = reach.Attribute("R13").Value
                    'newrow("14+") = reach.Attribute("R14").Value
                    'newrow("15+") = reach.Attribute("R15").Value
                    'newrow("16+") = reach.Attribute("R16").Value
                    'newrow("17+") = reach.Attribute("R17").Value
                    'newrow("18+") = reach.Attribute("R18").Value
                    'newrow("19+") = reach.Attribute("R19").Value
                    'newrow("20+") = reach.Attribute("R20").Value
                    ''Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateTVR000s")
                    ''Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateTVR")
                    ''Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateGRP000s")
                    ''Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateGRP")
                    '' ()
                    ''newrow("MidDateTVR000s") = reach.Attribute("R20").Value
                    ''newrow("MidDateTVR") = vals(30)
                    'newrow("MidDateGRP000s") = reach.Attribute("midDateGRP000s").Value
                    'newrow("MidDateGRP") = reach.Attribute("midDateGRP").Value
                    'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Rows.Add(newrow)
                    Try

                  
                    If summary.Name.LocalName.Equals("marketSummary") Then
                        For Each reach As XElement In summary.Elements
                                If reach.Element("reach").Attribute("market").Value.Contains("TotalMarkets") Then
                                    tlmReach = reach
                                    tlmPeriod = period
                                Else
                                    ReadReachAndAdditToTable(reach, period)
                                End If
                        Next
                        ReadReachAndAdditToTable(tlmReach, tlmPeriod)
                    End If
                    Catch ex As Exception
                        LogMpsrintExException("Exception occured while populating market summary.Message:" + ex.Message)
                    End Try
                    Try

                  
                    If summary.Name.LocalName.Equals("durationSummary") Then
                        For Each duration As XElement In summary.Elements
                            If period.Attribute("WeekNum").Value = String.Empty Then
                                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns("WeekNum").DefaultValue = 0
                            Else
                                Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns("WeekNum").DefaultValue = Convert.ToInt32(period.Attribute("WeekNum").Value)
                            End If
                            Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns("Duration").DefaultValue = duration.Attribute("value").Value
                            Dim tlmreach1 As XElement
                            Dim tlmperiod1 As XElement

                                For Each reach As XElement In duration.Elements
                                    '   For Each reachElement As XElement In channel.Element("reach").Elements
                                    If reach.Attribute("market").Value.Contains("TotalMarkets") Then
                                        tlmreach1 = reach
                                        tlmperiod1 = period
                                    Else
                                        ReachDurationSummary(reach, period)
                                    End If

                                    '    newrow("Gross Outlay") = goutlay / 10
                                    'Catch ex As Exception
                                    '    newrow("Gross Outlay") = String.Empty
                                    'End Try
                                    'Try
                                    '    newrow("CPRP") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(newrow("GRP"))
                                    'Catch ex As Exception
                                    '    newrow("CPRP") = String.Empty
                                    'End Try
                                    'Try
                                    '    newrow("CPT") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(finalrow("Reach000s").ToString())
                                    'Catch ex As Exception
                                    '    newrow("CPT") = String.Empty
                                    'End Try
                                Next
                            ReachDurationSummary(tlmreach1, tlmperiod1)
                        Next
                    End If
                    Catch ex As Exception
                        LogMpsrintExException("Exception occured while populating duration summary.Message:" + ex.Message)
                    End Try
                    Try

                   
                    If summary.Name.LocalName.Equals("creativeSummary") Then
                        For Each creative As XElement In summary.Elements
                            If period.Attribute("WeekNum").Value = String.Empty Then
                                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns("WeekNum").DefaultValue = 0
                            Else
                                Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns("WeekNum").DefaultValue = Convert.ToInt32(period.Attribute("WeekNum").Value)
                            End If
                            Globals.Ribbons.MSprintExRibbon.RnFCreativeSummary.Columns("Creative").DefaultValue = creative.Attribute("value").Value
                            Dim tlmReach2 As XElement
                            Dim tlmPeriod2 As XElement

                                For Each reach As XElement In creative.Elements
                                    '   For Each reachElement As XElement In channel.Element("reach").Elements
                                    If reach.Attribute("market").Value.Contains("TotalMarkets") Then
                                        tlmReach2 = reach
                                        tlmPeriod2 = period
                                    Else
                                        ReachCreativeSummary(reach, period)
                                    End If
                                    'Try
                                    '    newrow("Gross Outlay") = goutlay / 10
                                    'Catch ex As Exception
                                    '    newrow("Gross Outlay") = String.Empty
                                    'End Try
                                    'Try
                                    '    newrow("CPRP") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(newrow("GRP"))
                                    'Catch ex As Exception
                                    '    newrow("CPRP") = String.Empty
                                    'End Try
                                    'Try
                                    '    newrow("CPT") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(finalrow("Reach000s").ToString())
                                    'Catch ex As Exception
                                    '    newrow("CPT") = String.Empty
                                    'End Try
                                Next
                            ReachCreativeSummary(tlmReach2, tlmPeriod2)
                        Next
                        End If
                    Catch ex As Exception
                        LogMpsrintExException("Exception occured while populating creative summary.Message:" + ex.Message)

                    End Try
                    Try

                  
                    If summary.Name.LocalName.Equals("channelSummary") Then
                        For Each channel As XElement In summary.Elements
                            If period.Attribute("WeekNum").Value = String.Empty Then
                                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns("WeekNum").DefaultValue = 0
                            Else
                                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns("WeekNum").DefaultValue = Convert.ToInt32(period.Attribute("WeekNum").Value)
                                End If

                                Globals.Ribbons.MSprintExRibbon.RnFChannelSummary.Columns("Channel").DefaultValue = Globals.Ribbons.MSprintExRibbon.dtchannels.Select(String.Format("ID={0}", Convert.ToInt32(channel.Attribute("code").Value)))(0)("Name").ToString()
                            Dim tlmReach3 As XElement
                            Dim tlmPeriod3 As XElement
                            For Each reach As XElement In channel.Elements
                                '   For Each reachElement As XElement In channel.Element("reach").Elements
                                    If reach.Attribute("market").Value.Contains("TotalMarkets") Then
                                        tlmReach3 = reach
                                        tlmPeriod3 = period
                                    Else
                                        ReadReachChannelSummary(reach, period)
                                    End If
                                '    newrow("Gross Outlay") = goutlay / 10
                                'Catch ex As Exception
                                '    newrow("Gross Outlay") = String.Empty
                                'End Try
                                'Try
                                '    newrow("CPRP") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(newrow("GRP"))
                                'Catch ex As Exception
                                '    newrow("CPRP") = String.Empty
                                'End Try
                                'Try
                                '    newrow("CPT") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(finalrow("Reach000s").ToString())
                                'Catch ex As Exception
                                '    newrow("CPT") = String.Empty
                                'End Try
                            Next
                            ReadReachChannelSummary(tlmReach3, tlmPeriod3)
                        Next
                    End If
                    Catch ex As Exception
                        LogMpsrintExException("Exception occured while populating channel summary.Message:" + ex.Message)

                    End Try



                    'Try
                    '    newrow("Gross Outlay") = goutlay / 10
                    'Catch ex As Exception
                    '    newrow("Gross Outlay") = String.Empty
                    'End Try
                    'Try
                    '    newrow("CPRP") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(newrow("GRP"))
                    'Catch ex As Exception
                    '    newrow("CPRP") = String.Empty
                    'End Try
                    'Try
                    '    newrow("CPT") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(finalrow("Reach000s").ToString())
                    'Catch ex As Exception
                    '    newrow("CPT") = String.Empty
                    'End Try
                Next

            Next
        Catch ex As Exception
            generated = False
            Globals.Ribbons.MSprintExRibbon.SetNormalCursor()
            LogMpsrintExException("Exception occured while constructing ALL summary output table.Message :" + ex.Message)
            Throw ex
        End Try
    End Function
    Public Function ReadReachAndAdditToMarketIndTable(ByVal reach As XElement, ByVal period As XElement)
        Try
            If period.Attribute("WeekNum").Value = String.Empty Then
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns("WeekNum").DefaultValue = 0
            Else
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns("WeekNum").DefaultValue = Convert.ToInt32(period.Attribute("WeekNum").Value)
            End If
            '   Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns("Duration").DefaultValue = duration.Attribute("value").Value

            '    For Each reach As XElement In period.Element("duration").Elements
            '   For Each reachElement As XElement In channel.Element("reach").Elements
            Dim newrow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.NewRow()
            newrow("Market") = reach.Attribute("tm").Value.Split("~")(1)
            '  Dim periodElement As XElement = Globals.Ribbons.MSprintExRibbon.GetPeriodElementForMG(finalrow("MG").ToString(), String.Empty)
            Try
                newrow("Universe") = reach.Attribute("universe").Value
            Catch ex As Exception
                newrow("Universe") = String.Empty
            End Try
            newrow("GRP") = reach.Attribute("GRP").Value
            ' newrow("1+") = reach.Attribute("R1").Value

            'newrow("2+") = reach.Attribute("R2").Value
            'newrow("3+") = reach.Attribute("R3").Value
            'newrow("4+") = reach.Attribute("R4").Value
            'newrow("5+") = reach.Attribute("R5").Value
            'newrow("6+") = reach.Attribute("R6").Value
            'newrow("7+") = reach.Attribute("R7").Value
            'newrow("8+") = reach.Attribute("R8").Value
            'newrow("9+") = reach.Attribute("R9").Value
            'newrow("10+") = reach.Attribute("R10").Value
            'newrow("11+") = reach.Attribute("R11").Value
            'newrow("12+") = reach.Attribute("R12").Value
            'newrow("13+") = reach.Attribute("R13").Value
            'newrow("14+") = reach.Attribute("R14").Value
            'newrow("15+") = reach.Attribute("R15").Value
            'newrow("16+") = reach.Attribute("R16").Value
            'newrow("17+") = reach.Attribute("R17").Value
            'newrow("18+") = reach.Attribute("R18").Value
            'newrow("19+") = reach.Attribute("R19").Value
            'newrow("20+") = reach.Attribute("R20").Value
            For index = 1 To Globals.Ribbons.MSprintExRibbon.MSHRN
                Dim colname As String = index.ToString() + "+"
                Dim attrName As String = "R" + index.ToString()
                '  If index <= vals.Length - 5 Then
                If reach.Attributes(attrName).Count > 0 Then
                    newrow(colname) = reach.Attribute(attrName).Value
                Else
                    newrow(colname) = 0
                End If


            Next
            Try
                '  newrow("AOTS") = Convert.ToDecimal(reach.Attribute("GRP").Value) / Convert.ToDecimal(reach.Attribute("R1").Value)
                newrow("AOTS") = Convert.ToDecimal(reach.Attribute("GRP").Value) / Convert.ToDecimal(newrow("R1").ToString())
            Catch ex As Exception
                newrow("AOTS") = String.Empty
            End Try
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateTVR000s")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateTVR")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateGRP000s")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateGRP")
            ' ()
            'newrow("MidDateTVR000s") = reach.Attribute("R20").Value
            'newrow("MidDateTVR") = vals(30)
            'newrow("MidDateGRP000s") = reach.Attribute("midDateGRP000s").Value
            'newrow("MidDateGRP") = reach.Attribute("midDateGRP").Value
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Rows.Add(newrow)
        Catch ex As Exception
            LogMpsrintExException("Exception occured while reading reach element reach element.Message :" + ex.Message)

        End Try
    End Function
    Public Function ConstructMarketSummaryTable(ByVal opXML As XElement) As Boolean
        Dim generated As Boolean = True
        Globals.Ribbons.MSprintExRibbon.RnFMarketSummary = New Data.DataTable()
        Try

            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("WeekNum")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("Market")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("Universe")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("GRP")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("AOTS")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("1+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("2+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("3+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("4+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("5+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("6+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("7+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("8+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("9+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("10+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("11+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("12+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("13+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("14+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("15+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("16+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("17+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("18+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("19+")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("20+")
            If Int32.TryParse(opXML.Element("plan").Attribute("HReachNumber").Value, Globals.Ribbons.MSprintExRibbon.MSHRN) Then
                Globals.Ribbons.MSprintExRibbon.MSHRN += 1
                For index = 1 To Globals.Ribbons.MSprintExRibbon.MSHRN
                    Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add(index.ToString() + "+")
                Next
            Else
                Globals.Ribbons.MSprintExRibbon.MSHRN = 20
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("1+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("2+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("3+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("4+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("5+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("6+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("7+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("8+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("9+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("10+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("11+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("12+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("13+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("14+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("15+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("16+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("17+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("18+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("19+")
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("20+")

            End If
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateTVR000s")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateTVR")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateGRP000s")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateGRP")
            'market 5-26 
            For Each period As XElement In opXML.Element("plan").Elements
                Dim tlmReach As XElement
                Dim tlmPeriod As XElement
                For Each reach As XElement In period.Elements
                    'If period.Attribute("WeekNum").Value = String.Empty Then
                    '    Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns("WeekNum").DefaultValue = 0
                    'Else
                    '    Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns("WeekNum").DefaultValue = Convert.ToInt32(period.Attribute("WeekNum").Value)
                    'End If
                    ''   Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns("Duration").DefaultValue = duration.Attribute("value").Value

                    ''    For Each reach As XElement In period.Element("duration").Elements
                    ''   For Each reachElement As XElement In channel.Element("reach").Elements
                    'Dim newrow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.NewRow()
                    'newrow("Market") = reach.Attribute("tm").Value.Split("~")(1)
                    ''  Dim periodElement As XElement = Globals.Ribbons.MSprintExRibbon.GetPeriodElementForMG(finalrow("MG").ToString(), String.Empty)
                    'Try
                    '    newrow("Universe") = reach.Attribute("universe").Value
                    'Catch ex As Exception
                    '    newrow("Universe") = String.Empty
                    'End Try
                    'newrow("GRP") = reach.Attribute("GRP").Value
                    'newrow("1+") = reach.Attribute("R1").Value
                    'Try
                    '    newrow("AOTS") = Convert.ToDecimal(reach.Attribute("GRP").Value) / Convert.ToDecimal(reach.Attribute("R1").Value)
                    'Catch ex As Exception
                    '    newrow("AOTS") = String.Empty
                    'End Try
                    'newrow("2+") = reach.Attribute("R2").Value
                    'newrow("3+") = reach.Attribute("R3").Value
                    'newrow("4+") = reach.Attribute("R4").Value
                    'newrow("5+") = reach.Attribute("R5").Value
                    'newrow("6+") = reach.Attribute("R6").Value
                    'newrow("7+") = reach.Attribute("R7").Value
                    'newrow("8+") = reach.Attribute("R8").Value
                    'newrow("9+") = reach.Attribute("R9").Value
                    'newrow("10+") = reach.Attribute("R10").Value
                    'newrow("11+") = reach.Attribute("R11").Value
                    'newrow("12+") = reach.Attribute("R12").Value
                    'newrow("13+") = reach.Attribute("R13").Value
                    'newrow("14+") = reach.Attribute("R14").Value
                    'newrow("15+") = reach.Attribute("R15").Value
                    'newrow("16+") = reach.Attribute("R16").Value
                    'newrow("17+") = reach.Attribute("R17").Value
                    'newrow("18+") = reach.Attribute("R18").Value
                    'newrow("19+") = reach.Attribute("R19").Value
                    'newrow("20+") = reach.Attribute("R20").Value
                    ''Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateTVR000s")
                    ''Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateTVR")
                    ''Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateGRP000s")
                    ''Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateGRP")
                    '' ()
                    ''newrow("MidDateTVR000s") = reach.Attribute("R20").Value
                    ''newrow("MidDateTVR") = vals(30)
                    'newrow("MidDateGRP000s") = reach.Attribute("midDateGRP000s").Value
                    'newrow("MidDateGRP") = reach.Attribute("midDateGRP").Value
                    'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Rows.Add(newrow)

                    If reach.Attribute("tm").Value.Contains("TotalMarkets") Then
                        tlmReach = reach
                        tlmPeriod = period
                    Else
                        ' ReadReachAndAdditToTable(reach, period)
                        ReadReachAndAdditToMarketIndTable(reach, period)
                    End If


                    'Try
                    '    newrow("Gross Outlay") = goutlay / 10
                    'Catch ex As Exception
                    '    newrow("Gross Outlay") = String.Empty
                    'End Try
                    'Try
                    '    newrow("CPRP") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(newrow("GRP"))
                    'Catch ex As Exception
                    '    newrow("CPRP") = String.Empty
                    'End Try
                    'Try
                    '    newrow("CPT") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(finalrow("Reach000s").ToString())
                    'Catch ex As Exception
                    '    newrow("CPT") = String.Empty
                    'End Try
                Next
                '  ReadReachAndAdditToTable(tlmReach, tlmPeriod)
                ReadReachAndAdditToMarketIndTable(tlmReach, tlmPeriod)
            Next
            '  Next
        Catch ex As Exception
            generated = False
            Globals.Ribbons.MSprintExRibbon.SetNormalCursor()
            LogMpsrintExException("Exception occured while constructing Market summary table." + ex.Message)
            Throw ex
        End Try
        Return generated
    End Function
    Public Function ReadReachAndAdditToTable(ByVal reach As XElement, ByVal period As XElement)
        Try
            If period.Attribute("WeekNum").Value = String.Empty Then
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns("WeekNum").DefaultValue = 0
            Else
                Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns("WeekNum").DefaultValue = Convert.ToInt32(period.Attribute("WeekNum").Value)
            End If
            '   Globals.Ribbons.MSprintExRibbon.RnFDurationSummary.Columns("Duration").DefaultValue = duration.Attribute("value").Value

            '    For Each reach As XElement In period.Element("duration").Elements
            '   For Each reachElement As XElement In channel.Element("reach").Elements
            Dim newrow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.NewRow()
            newrow("Market") = reach.Element("reach").Attribute("market").Value.Split("~")(1)
            '  Dim periodElement As XElement = Globals.Ribbons.MSprintExRibbon.GetPeriodElementForMG(finalrow("MG").ToString(), String.Empty)
            Try
                newrow("Universe") = reach.Element("reach").Attribute("universe").Value
            Catch ex As Exception
                newrow("Universe") = String.Empty
            End Try
            newrow("GRP") = reach.Element("reach").Attribute("GRP").Value
            '  newrow("1+") = reach.Element("reach").Attribute("R1").Value
           
            'newrow("2+") = reach.Element("reach").Attribute("R2").Value
            'newrow("3+") = reach.Element("reach").Attribute("R3").Value
            'newrow("4+") = reach.Element("reach").Attribute("R4").Value
            'newrow("5+") = reach.Element("reach").Attribute("R5").Value
            'newrow("6+") = reach.Element("reach").Attribute("R6").Value
            'newrow("7+") = reach.Element("reach").Attribute("R7").Value
            'newrow("8+") = reach.Element("reach").Attribute("R8").Value
            'newrow("9+") = reach.Element("reach").Attribute("R9").Value
            'newrow("10+") = reach.Element("reach").Attribute("R10").Value
            'newrow("11+") = reach.Element("reach").Attribute("R11").Value
            'newrow("12+") = reach.Element("reach").Attribute("R12").Value
            'newrow("13+") = reach.Element("reach").Attribute("R13").Value
            'newrow("14+") = reach.Element("reach").Attribute("R14").Value
            'newrow("15+") = reach.Element("reach").Attribute("R15").Value
            'newrow("16+") = reach.Element("reach").Attribute("R16").Value
            'newrow("17+") = reach.Element("reach").Attribute("R17").Value
            'newrow("18+") = reach.Element("reach").Attribute("R18").Value
            'newrow("19+") = reach.Element("reach").Attribute("R19").Value
            'newrow("20+") = reach.Element("reach").Attribute("R20").Value
            For index = 1 To Globals.Ribbons.MSprintExRibbon.MSHRN
                Dim colname As String = index.ToString() + "+"
                Dim attrName As String = "R" + index.ToString()
                '  If index <= vals.Length - 5 Then
                If reach.Element("reach").Attributes(attrName).Count > 0 Then
                    newrow(colname) = reach.Element("reach").Attribute(attrName).Value
                Else
                    newrow(colname) = 0
                End If


            Next
            Try
                newrow("AOTS") = Convert.ToDecimal(reach.Element("reach").Attribute("GRP").Value) / Convert.ToDecimal(reach.Attribute("R1").Value)
            Catch ex As Exception
                newrow("AOTS") = String.Empty
            End Try
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateTVR000s")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateTVR")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateGRP000s")
            'Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("MidDateGRP")
            ' ()
            'newrow("MidDateTVR000s") = reach.Attribute("R20").Value
            'newrow("MidDateTVR") = vals(30)
            'newrow("MidDateGRP000s") = reach.Element("reach").Attribute("midDateGRP000s").Value
            'newrow("MidDateGRP") = reach.Element("reach").Attribute("midDateGRP").Value
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Rows.Add(newrow)
            'Try
            '    newrow("Gross Outlay") = goutlay / 10
            'Catch ex As Exception
            '    newrow("Gross Outlay") = String.Empty
            'End Try
            'Try
            '    newrow("CPRP") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(newrow("GRP"))
            'Catch ex As Exception
            '    newrow("CPRP") = String.Empty
            'End Try
            'Try
            '    newrow("CPT") = Convert.ToDecimal(newrow("Gross Outlay").ToString()) / Convert.ToDecimal(finalrow("Reach000s").ToString())
            'Catch ex As Exception
            '    newrow("CPT") = String.Empty
            'End Try
            ' Next
        Catch ex As Exception
            LogMpsrintExException("Exception occured while reading reach element reach element.Message :" + ex.Message)
        End Try
    End Function
    Public Function ConstructMarketSummaryTable(ByVal opXML As XElement, ByVal abc As Integer) As Boolean
        Dim generated As Boolean = True
        Globals.Ribbons.MSprintExRibbon.RnFMarketSummary = New Data.DataTable()
        Try
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("WeekNum")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("Market")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("Universe")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("GRP")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("AOTS")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("1+")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("2+")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("3+")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("4+")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("5+")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("6+")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("7+")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("8+")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("9+")
            Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Columns.Add("10+")
            Dim dtweekss As Data.DataTable = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks
            Dim mgs As List(Of String) = New List(Of String)()
            For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                mgs.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString())
            Next
            '  mgs.Add("TotalMarkets")
            If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                Dim weeknum As String = String.Empty

                For index = 0 To mgs.Count - 1
                    Dim mgroup As String = mgs(index)
                    Dim periodElement As XElement = Globals.Ribbons.MSprintExRibbon.GetPeriodElementForMG(mgroup, String.Empty)
                    Dim row As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.NewRow()
                    row("WeekNum") = weeknum
                    row("Market") = mgroup
                    Try
                        row("Universe") = periodElement.Attribute("universe").Value
                    Catch ex As Exception
                        row("Universe") = String.Empty
                    End Try
                    Try
                        row("GRP") = periodElement.Attribute("GRP").Value
                    Catch ex As Exception
                        row("GRP") = String.Empty
                    End Try

                    Try
                        row("1+") = periodElement.Attribute("R1").Value
                    Catch ex As Exception
                        row("1+") = String.Empty
                    End Try
                    Try
                        row("AOTS") = Convert.ToDecimal(row("GRP").ToString()) / Convert.ToDecimal(row("1+").ToString())
                    Catch ex As Exception
                        row("AOTS") = String.Empty
                    End Try
                    Try
                        row("2+") = periodElement.Attribute("R2").Value
                    Catch ex As Exception
                        row("2+") = String.Empty
                    End Try
                    Try
                        row("3+") = periodElement.Attribute("R3").Value
                    Catch ex As Exception
                        row("3+") = String.Empty
                    End Try
                    Try
                        row("4+") = periodElement.Attribute("R4").Value
                    Catch ex As Exception
                        row("4+") = String.Empty
                    End Try
                    Try
                        row("5+") = periodElement.Attribute("R5").Value
                    Catch ex As Exception
                        row("5+") = String.Empty
                    End Try
                    Try
                        row("6+") = periodElement.Attribute("R6").Value
                    Catch ex As Exception
                        row("6+") = String.Empty
                    End Try
                    Try
                        row("7+") = periodElement.Attribute("R7").Value
                    Catch ex As Exception
                        row("7+") = String.Empty
                    End Try
                    Try
                        row("8+") = periodElement.Attribute("R8").Value
                    Catch ex As Exception
                        row("8+") = String.Empty
                    End Try
                    Try
                        row("9+") = periodElement.Attribute("R9").Value
                    Catch ex As Exception
                        row("9+") = String.Empty
                    End Try
                    Try
                        row("10+") = periodElement.Attribute("R10").Value
                    Catch ex As Exception
                        row("10+") = String.Empty
                    End Try
                    Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Rows.Add(row)
                Next
            ElseIf Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                For index1 = 0 To dtweekss.Rows.Count - 1
                    For index = 0 To mgs.Count - 1
                        Dim mgroup As String = mgs(index)
                        Dim periodElement As XElement = Globals.Ribbons.MSprintExRibbon.GetPeriodElementForMG(mgroup, dtweekss.Rows(index1)("WeekNumber").ToString())
                        Dim row As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.NewRow()
                        row("WeekNum") = dtweekss.Rows(index1)("WeekNumber").ToString()
                        row("Market") = mgroup
                        Try
                            row("Universe") = periodElement.Attribute("universe").Value
                        Catch ex As Exception
                            row("Universe") = String.Empty
                        End Try
                        Try
                            row("GRP") = periodElement.Attribute("GRP").Value
                        Catch ex As Exception
                            row("GRP") = String.Empty
                        End Try

                        Try
                            row("1+") = periodElement.Attribute("R1").Value
                        Catch ex As Exception
                            row("1+") = String.Empty
                        End Try
                        Try
                            row("AOTS") = Convert.ToDecimal(row("GRP").ToString()) / Convert.ToDecimal(row("1+").ToString())
                        Catch ex As Exception
                            row("AOTS") = String.Empty
                        End Try
                        Try
                            row("2+") = periodElement.Attribute("R2").Value
                        Catch ex As Exception
                            row("2+") = String.Empty
                        End Try
                        Try
                            row("3+") = periodElement.Attribute("R3").Value
                        Catch ex As Exception
                            row("3+") = String.Empty
                        End Try
                        Try
                            row("4+") = periodElement.Attribute("R4").Value
                        Catch ex As Exception
                            row("4+") = String.Empty
                        End Try
                        Try
                            row("5+") = periodElement.Attribute("R5").Value
                        Catch ex As Exception
                            row("5+") = String.Empty
                        End Try
                        Try
                            row("6+") = periodElement.Attribute("R6").Value
                        Catch ex As Exception
                            row("6+") = String.Empty
                        End Try
                        Try
                            row("7+") = periodElement.Attribute("R7").Value
                        Catch ex As Exception
                            row("7+") = String.Empty
                        End Try
                        Try
                            row("8+") = periodElement.Attribute("R8").Value
                        Catch ex As Exception
                            row("8+") = String.Empty
                        End Try
                        Try
                            row("9+") = periodElement.Attribute("R9").Value
                        Catch ex As Exception
                            row("9+") = String.Empty
                        End Try
                        Try
                            row("10+") = periodElement.Attribute("R10").Value
                        Catch ex As Exception
                            row("10+") = String.Empty
                        End Try
                        Globals.Ribbons.MSprintExRibbon.RnFMarketSummary.Rows.Add(row)
                    Next
                Next
            End If
        Catch ex As Exception
            generated = False
            Globals.Ribbons.MSprintExRibbon.SetNormalCursor()
            LogMpsrintExException("Exception occured while constructing Market summary table." + ex.Message)
            Throw ex
        End Try
        Return generated
    End Function
    Public Function ConstructOpProgAvgTVRTable(ByVal opXml As XElement) As Boolean
        ' Dim output As Data.DataTable = New Data.DataTable()
        Dim generated As Boolean = True
        Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable = New Data.DataTable()
        ' Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable = New Data.DataTable()

        Try
            ' RnFOutputTable = New Data.DataTable()
            'output.Columns.Add("GUID")
            'output.Columns.Add("Channel")
            'output.Columns.Add("Spot")
            'output.Columns.Add("AvaiSpotString")
            'output.Columns.Add("TG")
            'output.Columns.Add("MG")
            'output.Columns.Add("ReachVal")
            'output.Columns.Add("TVRVal")
            'output.Columns.Add("Start Date", System.Type.GetType("System.DateTime"))
            'output.Columns.Add("End Date", System.Type.GetType("System.DateTime"))
            'output.Columns.Add("WeekNum", System.Type.GetType("System.Int32"))
            'output.Columns.Add("Date", System.Type.GetType("System.DateTime"))
            'output.Columns.Add("Start Time")
            '' RnFSelectedSpots.Columns.Add("End Time")
            'output.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
            '' RnFSelectedSpots.Columns.Add("Commercial")
            'output.Columns.Add("Cost")
            'output.Columns.Add("PA")
            'output.Columns.Add("TA")
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

            Globals.Ribbons.MSprintExRibbon.markets.Clear()
            'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Channel")
            'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Programme")
            '' Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Creative")
            'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Day")
            '' Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Date", System.Type.GetType("System.DateTime"))
            'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Start Time")
            'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("End Time")
            '' RnFSelectedSpots.Columns.Add("End Time")
            '' Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
            '' RnFSelectedSpots.Columns.Add("Commercial")
            'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Cost")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("GUID")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("0TVR Spots")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("TG")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("MG")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Avg TVR")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Std Deviation")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Total available breaks")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Break 0 to m - 2s")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Break (m - 2s) to (m - s)")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Break (m - s) to m")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Break m to (m + s)")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Break (m + s) to (m + 2s)")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Break > (m + 2s)")

            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Ad 0 to m - 2s")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Ad (m - 2s) to (m - s)")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Ad (m - s) to m")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Ad m to (m + s)")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Ad (m + s) to (m + 2s)")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Ad > (m + 2s)")
            ' Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Ad > (m + 2s)")
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Total available Ads")
            '  Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("TA")
            ' Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("GUID")
            'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Spot")
            'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Start Date", System.Type.GetType("System.DateTime"))
            'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("End Date", System.Type.GetType("System.DateTime"))
            'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("WeekNum", System.Type.GetType("System.Int32"))


            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Start Date")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("End Date")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("WeekNum", System.Type.GetType("System.Int32"))
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("TG")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("MG")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Channel")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Date") 'spot date
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Day")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Start Time")
            '' RnfShowResultsTable.Columns.Add("End Time")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Programme")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("PA")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("TA")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("TVR000s")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("TVR")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("GRP000s")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("GRP")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("AvgFreq")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("CummCost")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("SpotCPRP")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("CummCPRP")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Reach000s")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("1+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("2+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("3+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("4+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("5+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("6+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("7+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("8+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("9+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("10+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("11+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("12+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("13+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("14+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("15+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("16+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("17+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("18+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("19+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("20+")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("MidDateTVR000s")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("MidDateTVR")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("MidDateGRP000s")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("MidDateGRP")
            '  Dim rnfoutput As XElement = XElement.Load(AppDomain.CurrentDomain.BaseDirectory + "sampleoutput.xml")
            For Each period As XElement In opXml.Element("plan").Elements
                'Dim startDate As Date = New Date(period.Attribute("StartDate").Value.Substring(0, 4), period.Attribute("StartDate").Value.Substring(4, 2), period.Attribute("StartDate").Value.Substring(6, 2)).ToShortDateString()
                'Dim endDate As Date = New Date(period.Attribute("EndDate").Value.Substring(0, 4), period.Attribute("EndDate").Value.Substring(4, 2), period.Attribute("EndDate").Value.Substring(6, 2)).ToShortDateString()
                'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Start Date").DefaultValue = startDate
                'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("End Date").DefaultValue = endDate
                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Start Date").DefaultValue = startDate
                'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("End Date").DefaultValue = endDate
                'If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                '    Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("WeekNum").DefaultValue = 0
                '    Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("WeekNum").DefaultValue = 0
                'Else
                '    Dim weeknum As Integer = Convert.ToInt32(period.Attribute("WeekNum").Value)
                '    Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("WeekNum").DefaultValue = weeknum
                '    Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("WeekNum").DefaultValue = weeknum
                'End If


                For Each programme As XElement In period.Elements
                    'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Channel").DefaultValue = programme.Attribute("ChannelName").Value
                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Channel").DefaultValue = programme.Attribute("ChannelName").Value

                    ''Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Day").DefaultValue = programme.Attribute("days").Value
                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Programme").DefaultValue = programme.Attribute("ProgName").Value
                    'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Programme").DefaultValue = programme.Attribute("ProgName").Value
                    'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Creative").DefaultValue = programme.Attribute("caption").Value
                    ' TotalSpotCountWithZeroTVR

                    Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns("GUID").DefaultValue = programme.Attribute("guid").Value
                    '  0TVR Spots
                    Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns("0TVR Spots").DefaultValue = programme.Attribute("TotalSpotCountWithZeroTVR").Value
                    For Each breaktvr As XElement In programme.Elements
                        Dim dr As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.NewRow
                        dr("TG") = breaktvr.Attribute("TM").Value.Split("~")(0)
                        dr("MG") = breaktvr.Attribute("TM").Value.Split("~")(1)
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Avg TVR")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Std Deviation")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Total available breaks")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("0 to m - 2s")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("(m - 2s) to (m - s)")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("(m - s) to m")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("m to (m + s)")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("(m + s) to (m + 2s)")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("> (m + 2s)")
                        dr("Avg TVR") = breaktvr.Attribute("avgTVR").Value
                        dr("Std Deviation") = breaktvr.Attribute("stdDev").Value
                        dr("Total available breaks") = breaktvr.Attribute("breakCounts").Value.Split(",")(6)
                        dr("Break 0 to m - 2s") = breaktvr.Attribute("breakCounts").Value.Split(",")(0)
                        dr("Break (m - 2s) to (m - s)") = breaktvr.Attribute("breakCounts").Value.Split(",")(1)
                        dr("Break (m - s) to m") = breaktvr.Attribute("breakCounts").Value.Split(",")(2)
                        dr("Break m to (m + s)") = breaktvr.Attribute("breakCounts").Value.Split(",")(3)
                        dr("Break (m + s) to (m + 2s)") = breaktvr.Attribute("breakCounts").Value.Split(",")(4)
                        dr("Break > (m + 2s)") = breaktvr.Attribute("breakCounts").Value.Split(",")(5)

                        dr("Ad 0 to m - 2s") = breaktvr.Attribute("adCounts").Value.Split(",")(0)
                        dr("Ad (m - 2s) to (m - s)") = breaktvr.Attribute("adCounts").Value.Split(",")(1)
                        dr("Ad (m - s) to m") = breaktvr.Attribute("adCounts").Value.Split(",")(2)
                        dr("Ad m to (m + s)") = breaktvr.Attribute("adCounts").Value.Split(",")(3)
                        dr("Ad (m + s) to (m + 2s)") = breaktvr.Attribute("adCounts").Value.Split(",")(4)
                        dr("Ad > (m + 2s)") = breaktvr.Attribute("adCounts").Value.Split(",")(5)
                        dr("Total available Ads") = breaktvr.Attribute("adCounts").Value.Split(",")(6)
                        'dr("Ad 0 to m - 2s") = breaktvr.Attribute("").Value.Split(",")(0)
                        'd("Break 0 to m - 2s")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add(" Break (m - 2s) to (m - s)")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Break (m - s) to m")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Break m to (m + s)")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Break (m + s) to (m + 2s)")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Break > (m + 2s)")

                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Ad 0 to m - 2s")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add(" Ad (m - 2s) to (m - s)")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Ad (m - s) to m")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Ad m to (m + s)")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Ad (m + s) to (m + 2s)")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Ad > (m + 2s)")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Ad > (m + 2s)")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Total available Ads")
                        Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Rows.Add(dr)
                    Next
                    ' Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Spot").DefaultValue = spot.Attribute("log").Value
                    ' Dim spotrow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.NewRow()
                    '  spotrow("Spot") = spot.Attribute("log").Value
                    '   Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Spot").DefaultValue = spot.Attribute("log").Value
                    'RnFSelectedSpots.Columns.Add("Date", System.Type.GetType("System.DateTime"))
                    'RnFSelectedSpots.Columns.Add("Start Time")
                    '' RnFSelectedSpots.Columns.Add("End Time")
                    'RnFSelectedSpots.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
                    '' RnFSelectedSpots.Columns.Add("Commercial")
                    'RnFSelectedSpots.Columns.Add("Cost")
                    'RnFSelectedSpots.Columns.Add("PA")
                    'RnFSelectedSpots.Columns.Add("TA")
                    'output.Columns.Add("Date", System.Type.GetType("System.DateTime"))
                    'output.Columns.Add("Start Time")
                    '' RnFSelectedSpots.Columns.Add("End Time")
                    'output.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
                    '' RnFSelectedSpots.Columns.Add("Commercial")
                    'output.Columns.Add("Cost")
                    'output.Columns.Add("PA")
                    'output.Columns.Add("TA")
                    ' Dim values As String() = spot.Attribute("log").Value.Split({","c}, StringSplitOptions.None)
                    '   Dim dr As Data.DataRow = spot.NewRow()
                    'For index = 0 To values.Length - 1

                    ' If index = 1 Then
                    'dr("Date") = New Date(Convert.ToInt32(values(1).Substring(0, 4)), Convert.ToInt32(values(1).Substring(4, 2)), Convert.ToInt32(values(1).Substring(6, 2))).ToShortDateString()
                    '' ElseIf index = 2 Then
                    '' dr("StartTime") = New TimeSpan(Convert.ToInt32(values(2).Substring(0, 2)), Convert.ToInt32(values(2).Substring(2, 2)), Convert.ToInt32(values(2).Substring(4, 2)))
                    ''  dr("StartTime") = values(2).Substring(0, 5)
                    'dr("StartTime") = String.Format("{0}:{1}", values(2).Substring(0, 2), values(2).Substring(2, 2))
                    'dr("Cost") = values(4)
                    'dr("PA") = values(5)
                    'dr("TA") = values(6)
                    'dr("Duration(Sec)") = values(7)
                    '            Dim dateval As Date = New Date(Convert.ToInt32(values(1).Substring(0, 4)), Convert.ToInt32(values(1).Substring(4, 2)), Convert.ToInt32(values(1).Substring(6, 2))).ToShortDateString()
                    '            Dim starttime As String = String.Format("{0}:{1}", values(2).Substring(0, 2), values(2).Substring(2, 2))
                    '            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Date").DefaultValue = dateval
                    '            '  spotrow("Date") = dateval
                    '            Try
                    '                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Day").DefaultValue = dateval.ToString("ddd")
                    '                Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Day").DefaultValue = dateval.ToString("ddd")
                    '            Catch ex As Exception
                    '                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Day").DefaultValue = programme.Attribute("days").Value
                    '                Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Day").DefaultValue = programme.Attribute("days").Value
                    '            End Try


                    '            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Date").DefaultValue = dateval
                    '            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Start Time").DefaultValue = starttime
                    '            '  spotrow("Start Time") = starttime
                    '            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Start Time").DefaultValue = starttime
                    '            ' Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Cost").DefaultValue = values(4)
                    '            ' spotrow("Cost") = values(4)
                    '            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Cost").DefaultValue = values(4)
                    '            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("PA").DefaultValue = values(5)
                    '            ' spotrow("PA") = values(5)
                    '            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("PA").DefaultValue = values(5)
                    '            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("TA").DefaultValue = values(6)
                    '            ' spotrow("TA") = values(6)
                    '            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("TA").DefaultValue = values(6)
                    '            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Duration(Sec)").DefaultValue = values(7)
                    '            ' spotrow("Duration(Sec)") = values(7)
                    '            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Duration(Sec)").DefaultValue = values(7)
                    '            '  Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Add(spotrow)
                    '            Dim mgcount As Integer = 0
                    '            For Each reach As XElement In spot.Elements
                    '                Dim dr As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.NewRow()
                    '                dr("TG") = reach.Attribute("tm").Value.Split({"~"c}, StringSplitOptions.None)(0)
                    '                dr("MG") = reach.Attribute("tm").Value.Split({"~"c}, StringSplitOptions.None)(1)
                    '                Dim mgvalue As String = reach.Attribute("tm").Value.Split({"~"c}, StringSplitOptions.None)(1)

                    '                If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Contains(mgvalue + "TVR")) Then
                    '                    Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add(mgvalue + "TVR")
                    '                    Globals.Ribbons.MSprintExRibbon.markets.Add(mgvalue + "TVR")
                    '                End If

                    '                '   dr("ReachVal") = reach.Attribute("val").Value
                    '                Dim vals As String() = reach.Attribute("val").Value.Split({","c}, StringSplitOptions.None)
                    '                '  Dim drow As Data.DataRow = reach.NewRow()
                    '                dr("TVR000s") = vals(0)
                    '                dr("TVR") = vals(1)

                    '                If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Contains(mgvalue + "TVR") Then
                    '                    Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns(mgvalue + "TVR").DefaultValue = vals(1)
                    '                End If

                    '                If mgcount = spot.Elements.Count - 1 Then
                    '                    Dim dr1 As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.NewRow()
                    '                    dr1(mgvalue + "TVR") = vals(1)
                    '                    Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Add(dr1)
                    '                End If

                    '                dr("GRP000s") = vals(2)
                    '                dr("GRP") = vals(3)
                    '                dr("AvgFreq") = vals(4)
                    '                dr("CummCost") = vals(5)
                    '                dr("SpotCPRP") = vals(6)
                    '                dr("CummCPRP") = vals(7)
                    '                dr("Reach000s") = vals(8)
                    '                dr("1+") = vals(9)
                    '                dr("2+") = vals(10)
                    '                dr("3+") = vals(11)
                    '                dr("4+") = vals(12)
                    '                dr("5+") = vals(13)
                    '                dr("6+") = vals(14)
                    '                dr("7+") = vals(15)
                    '                dr("8+") = vals(16)
                    '                dr("9+") = vals(17)
                    '                dr("10+") = vals(18)
                    '                dr("11+") = vals(19)
                    '                dr("12+") = vals(20)
                    '                dr("13+") = vals(21)
                    '                dr("14+") = vals(22)
                    '                dr("15+") = vals(23)
                    '                dr("16+") = vals(24)
                    '                dr("17+") = vals(25)
                    '                dr("18+") = vals(26)
                    '                dr("19+") = vals(27)
                    '                dr("20+") = vals(28)
                    '                'MidDateTVR000s,MidDateTVR,MidDateGRP000s,MidDateGRP
                    '                dr("MidDateTVR000s") = vals(29)
                    '                dr("MidDateTVR") = vals(30)
                    '                dr("MidDateGRP000s") = vals(31)
                    '                dr("MidDateGRP") = vals(32)
                    '                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Rows.Add(dr)
                    '                mgcount += 1
                    '            Next
                    '        Next
                    '        'Available Spots

                    '        'If programme.Elements("available_spots").Any() Then
                    '        '    For Each spot As XElement In programme.Element("available_spots").Elements
                    '        '        output.Columns("AvaiSpotString").DefaultValue = spot.Attribute("log").Value
                    '        '        For Each reach As XElement In spot.Elements
                    '        '            Dim dr As Data.DataRow = output.NewRow()
                    '        '            dr("TG") = reach.Attribute("tm").Value.Split({"~"c}, StringSplitOptions.None)(0)
                    '        '            dr("MG") = reach.Attribute("tm").Value.Split({"~"c}, StringSplitOptions.None)(1)
                    '        '            dr("TVRVal") = reach.Attribute("val").Value
                    '        '            output.Rows.Add(dr)
                    '        '        Next
                    '        '    Next
                    '        'End If



                Next
            Next
            Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.AcceptChanges()
            '  Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.AcceptChanges()
        Catch ex As Exception
            LogMpsrintExException("Exception occured while constructing Prog Avg TVRdataset from output XML." + ex.Message)
            Throw ex
            generated = False
        End Try
        Return generated
    End Function
    Private Function EscapeLikeValue(ByVal value As String) As String
        Dim sb As New StringBuilder(value.Length)
        For i As Integer = 0 To value.Length - 1
            Dim c As Char = value(i)
            Select Case c
                Case "]"c, "["c, "%"c, "*"c
                    sb.Append("[").Append(c).Append("]")
                    Exit Select
                Case "'"c
                    sb.Append("''")
                    Exit Select
                Case Else
                    sb.Append(c)
                    Exit Select
            End Select
        Next
        Return sb.ToString()
    End Function
    Public Function ConstructOpRnFTable(ByVal opXmL As XElement) As Boolean
        ' Dim output As Data.DataTable = New Data.DataTable()
        Dim generated As Boolean = True
        Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots = New Data.DataTable()
        Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable = New Data.DataTable()

        Try
            ' RnFOutputTable = New Data.DataTable()
            'output.Columns.Add("GUID")
            'output.Columns.Add("Channel")
            'output.Columns.Add("Spot")
            'output.Columns.Add("AvaiSpotString")
            'output.Columns.Add("TG")
            'output.Columns.Add("MG")
            'output.Columns.Add("ReachVal")
            'output.Columns.Add("TVRVal")
            'output.Columns.Add("Start Date", System.Type.GetType("System.DateTime"))
            'output.Columns.Add("End Date", System.Type.GetType("System.DateTime"))
            'output.Columns.Add("WeekNum", System.Type.GetType("System.Int32"))
            'output.Columns.Add("Date", System.Type.GetType("System.DateTime"))
            'output.Columns.Add("Start Time")
            '' RnFSelectedSpots.Columns.Add("End Time")
            'output.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
            '' RnFSelectedSpots.Columns.Add("Commercial")
            'output.Columns.Add("Cost")
            'output.Columns.Add("PA")
            'output.Columns.Add("TA")
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

            Globals.Ribbons.MSprintExRibbon.markets.Clear()
            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Channel")
            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Programme")
            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Creative")
            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Day")
            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Date", System.Type.GetType("System.DateTime"))
            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Start Time")
            ' RnFSelectedSpots.Columns.Add("End Time")
            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
            ' RnFSelectedSpots.Columns.Add("Commercial")
            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Cost")
            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("PA")
            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("TA")
            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("GUID")
            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Spot")
            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Start Date", System.Type.GetType("System.DateTime"))
            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("End Date", System.Type.GetType("System.DateTime"))
            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("WeekNum", System.Type.GetType("System.Int32"))


            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Start Date")
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("End Date")
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("WeekNum", System.Type.GetType("System.Int32"))
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("TG")
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("MG")
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Channel")
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Date") 'spot date
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Day")
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Start Time")
            ' RnfShowResultsTable.Columns.Add("End Time")
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Programme")
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("PA")
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("TA")
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("TVR000s")
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("TVR")
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("GRP000s")
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("GRP")
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("AvgFreq")
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("CummCost")
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("SpotCPRP")
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("CummCPRP")
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("Reach000s")
            ' Dim HRN As Integer = 0
            If Int32.TryParse(opXmL.Element("plan").Attribute("HReachNumber").Value, Globals.Ribbons.MSprintExRibbon.HRN) Then
                Globals.Ribbons.MSprintExRibbon.HRN += 1
                For index = 1 To Globals.Ribbons.MSprintExRibbon.HRN
                    Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add(index.ToString() + "+")
                Next
            Else
                Globals.Ribbons.MSprintExRibbon.HRN = 21
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("1+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("2+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("3+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("4+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("5+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("6+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("7+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("8+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("9+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("10+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("11+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("12+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("13+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("14+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("15+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("16+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("17+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("18+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("19+")
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("20+")

            End If
            
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("MidDateTVR000s")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("MidDateTVR")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("MidDateGRP000s")
            'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns.Add("MidDateGRP")
            '  Dim rnfoutput As XElement = XElement.Load(AppDomain.CurrentDomain.BaseDirectory + "sampleoutput.xml")
            For Each period As XElement In opXmL.Element("plan").Elements
                Dim startDate As Date = New Date(period.Attribute("StartDate").Value.Substring(0, 4), period.Attribute("StartDate").Value.Substring(4, 2), period.Attribute("StartDate").Value.Substring(6, 2)).ToShortDateString()
                Dim endDate As Date = New Date(period.Attribute("EndDate").Value.Substring(0, 4), period.Attribute("EndDate").Value.Substring(4, 2), period.Attribute("EndDate").Value.Substring(6, 2)).ToShortDateString()
                Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Start Date").DefaultValue = startDate
                Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("End Date").DefaultValue = endDate
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Start Date").DefaultValue = startDate
                Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("End Date").DefaultValue = endDate
                If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                    Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("WeekNum").DefaultValue = 0
                    Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("WeekNum").DefaultValue = 0
                Else
                    Dim weeknum As Integer = Convert.ToInt32(period.Attribute("WeekNum").Value)
                    Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("WeekNum").DefaultValue = weeknum
                    Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("WeekNum").DefaultValue = weeknum
                End If


                For Each programme As XElement In period.Elements
                    Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Channel").DefaultValue = programme.Attribute("ChannelName").Value
                    Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Channel").DefaultValue = programme.Attribute("ChannelName").Value

                    'Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Day").DefaultValue = programme.Attribute("days").Value
                    Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Programme").DefaultValue = programme.Attribute("ProgName").Value
                    Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Programme").DefaultValue = programme.Attribute("ProgName").Value
                    Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Creative").DefaultValue = programme.Attribute("caption").Value

                    Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("GUID").DefaultValue = programme.Attribute("guid").Value
                    For Each spot As XElement In programme.Element("selected_spots").Elements
                        ' Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Spot").DefaultValue = spot.Attribute("log").Value
                        ' Dim spotrow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.NewRow()
                        '  spotrow("Spot") = spot.Attribute("log").Value
                        Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Spot").DefaultValue = spot.Attribute("log").Value
                        'RnFSelectedSpots.Columns.Add("Date", System.Type.GetType("System.DateTime"))
                        'RnFSelectedSpots.Columns.Add("Start Time")
                        '' RnFSelectedSpots.Columns.Add("End Time")
                        'RnFSelectedSpots.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
                        '' RnFSelectedSpots.Columns.Add("Commercial")
                        'RnFSelectedSpots.Columns.Add("Cost")
                        'RnFSelectedSpots.Columns.Add("PA")
                        'RnFSelectedSpots.Columns.Add("TA")
                        'output.Columns.Add("Date", System.Type.GetType("System.DateTime"))
                        'output.Columns.Add("Start Time")
                        '' RnFSelectedSpots.Columns.Add("End Time")
                        'output.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
                        '' RnFSelectedSpots.Columns.Add("Commercial")
                        'output.Columns.Add("Cost")
                        'output.Columns.Add("PA")
                        'output.Columns.Add("TA")
                        Dim values As String() = spot.Attribute("log").Value.Split({","c}, StringSplitOptions.None)
                        '   Dim dr As Data.DataRow = spot.NewRow()
                        'For index = 0 To values.Length - 1

                        ' If index = 1 Then
                        'dr("Date") = New Date(Convert.ToInt32(values(1).Substring(0, 4)), Convert.ToInt32(values(1).Substring(4, 2)), Convert.ToInt32(values(1).Substring(6, 2))).ToShortDateString()
                        '' ElseIf index = 2 Then
                        '' dr("StartTime") = New TimeSpan(Convert.ToInt32(values(2).Substring(0, 2)), Convert.ToInt32(values(2).Substring(2, 2)), Convert.ToInt32(values(2).Substring(4, 2)))
                        ''  dr("StartTime") = values(2).Substring(0, 5)
                        'dr("StartTime") = String.Format("{0}:{1}", values(2).Substring(0, 2), values(2).Substring(2, 2))
                        'dr("Cost") = values(4)
                        'dr("PA") = values(5)
                        'dr("TA") = values(6)
                        'dr("Duration(Sec)") = values(7)
                        Dim dateval As Date = New Date(Convert.ToInt32(values(1).Substring(0, 4)), Convert.ToInt32(values(1).Substring(4, 2)), Convert.ToInt32(values(1).Substring(6, 2))).ToShortDateString()
                        Dim starttime As String = String.Format("{0}:{1}", values(2).Substring(0, 2), values(2).Substring(2, 2))
                        Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Date").DefaultValue = dateval
                        '  spotrow("Date") = dateval
                        Try
                            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Day").DefaultValue = dateval.ToString("ddd")
                            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Day").DefaultValue = dateval.ToString("ddd")
                        Catch ex As Exception
                            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Day").DefaultValue = programme.Attribute("days").Value
                            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Day").DefaultValue = programme.Attribute("days").Value
                        End Try


                        Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Date").DefaultValue = dateval
                        Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Start Time").DefaultValue = starttime
                        '  spotrow("Start Time") = starttime
                        Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Start Time").DefaultValue = starttime
                        ' Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Cost").DefaultValue = values(4)
                        ' spotrow("Cost") = values(4)
                        Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Cost").DefaultValue = values(4)
                        Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("PA").DefaultValue = values(5)
                        ' spotrow("PA") = values(5)
                        Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("PA").DefaultValue = values(5)
                        Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("TA").DefaultValue = values(6)
                        ' spotrow("TA") = values(6)
                        Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("TA").DefaultValue = values(6)
                        Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Columns("Duration(Sec)").DefaultValue = values(7)
                        ' spotrow("Duration(Sec)") = values(7)
                        Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Duration(Sec)").DefaultValue = values(7)
                        '  Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Add(spotrow)
                        Dim mgcount As Integer = 0
                        For Each reach As XElement In spot.Elements
                            Dim dr As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.NewRow()
                            dr("TG") = reach.Attribute("tm").Value.Split({"~"c}, StringSplitOptions.None)(0)
                            dr("MG") = reach.Attribute("tm").Value.Split({"~"c}, StringSplitOptions.None)(1)
                            Dim mgvalue As String = reach.Attribute("tm").Value.Split({"~"c}, StringSplitOptions.None)(1)

                            If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Contains(mgvalue + "TVR")) Then
                                Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add(mgvalue + "TVR")
                                Globals.Ribbons.MSprintExRibbon.markets.Add(mgvalue + "TVR")
                            End If

                            '   dr("ReachVal") = reach.Attribute("val").Value
                            Dim vals As String() = reach.Attribute("val").Value.Split({","c}, StringSplitOptions.None)
                            '  Dim drow As Data.DataRow = reach.NewRow()
                            'dr("TVR000s") = vals(0)


                            If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Contains(mgvalue + "TVR") Then
                                Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns(mgvalue + "TVR").DefaultValue = vals(0)
                            End If

                            If mgcount = spot.Elements.Count - 1 Then
                                Dim dr1 As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.NewRow()
                                dr1(mgvalue + "TVR") = vals(0)
                                Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Add(dr1)
                            End If

                            ' dr("GRP000s") = vals(2)
                            dr("TVR") = vals(0)
                            dr("GRP") = vals(1)
                            '  dr("AvgFreq") = vals(4)
                            dr("CummCost") = vals(2)
                            dr("SpotCPRP") = vals(3)
                            dr("CummCPRP") = vals(4)
                            '  dr("Reach000s") = vals(8)

                            For index = 1 To Globals.Ribbons.MSprintExRibbon.HRN
                                Dim colname As String = index.ToString() + "+"
                                If index <= vals.Length - 5 Then

                                    dr(colname) = vals(index + 4)
                                Else
                                    dr(colname) = 0
                                End If


                            Next

                            'dr("1+") = vals(9)
                            'dr("2+") = vals(10)
                            'dr("3+") = vals(11)
                            'dr("4+") = vals(12)
                            'dr("5+") = vals(13)
                            'dr("6+") = vals(14)
                            'dr("7+") = vals(15)
                            'dr("8+") = vals(16)
                            'dr("9+") = vals(17)
                            'dr("10+") = vals(18)
                            'dr("11+") = vals(19)
                            'dr("12+") = vals(20)
                            'dr("13+") = vals(21)
                            'dr("14+") = vals(22)
                            'dr("15+") = vals(23)
                            'dr("16+") = vals(24)
                            'dr("17+") = vals(25)
                            'dr("18+") = vals(26)
                            'dr("19+") = vals(27)
                            'dr("20+") = vals(28)
                            'MidDateTVR000s,MidDateTVR,MidDateGRP000s,MidDateGRP
                            'dr("MidDateTVR000s") = vals(29)
                            'dr("MidDateTVR") = vals(30)
                            'dr("MidDateGRP000s") = vals(31)
                            'dr("MidDateGRP") = vals(32)
                            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.Rows.Add(dr)
                            mgcount += 1
                        Next
                    Next
                    'Available Spots

                    'If programme.Elements("available_spots").Any() Then
                    '    For Each spot As XElement In programme.Element("available_spots").Elements
                    '        output.Columns("AvaiSpotString").DefaultValue = spot.Attribute("log").Value
                    '        For Each reach As XElement In spot.Elements
                    '            Dim dr As Data.DataRow = output.NewRow()
                    '            dr("TG") = reach.Attribute("tm").Value.Split({"~"c}, StringSplitOptions.None)(0)
                    '            dr("MG") = reach.Attribute("tm").Value.Split({"~"c}, StringSplitOptions.None)(1)
                    '            dr("TVRVal") = reach.Attribute("val").Value
                    '            output.Rows.Add(dr)
                    '        Next
                    '    Next
                    'End If



                Next
            Next
            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.AcceptChanges()
            Globals.Ribbons.MSprintExRibbon.RnFShowResultsTable.AcceptChanges()
        Catch ex As Exception
            LogMpsrintExException("Exception occured while constructing dataset from output XML." + ex.Message)
            Throw ex
            generated = False
        End Try
        Return generated
    End Function
    Public Function ConstructSpotForSelectedRow(ByVal row As Data.DataRow) As String
        Dim spot As String = String.Empty
        Try
            ' spot= 
        Catch ex As Exception

        End Try
    End Function
    'Public Function ConstructBSLSummaryInputXML() As XElement
    '    Dim input As XElement
    '    Dim tp As ucPlanSelections = Globals.Ribbons.MSprintExRibbon.tpSelections
    '    Try
    '        Dim month, month1 As String
    '        Dim day, day1 As String
    '        Dim plantgname As String = String.Empty
    '        Dim reftgname As String = String.Empty
    '        Dim dtable As System.Data.DataTable = DirectCast(Globals.Ribbons.MSprintExRibbon.tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
    '        plantgname = dtable.Rows(0)(1).ToString().Trim()
    '        '  Button1.Enabled = False
    '        ' lbGetting.Text = "Getting Genre Share for chosen TG-MGs..."
    '        '   lbGetting.Refresh()
    '        If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Month < 10 Then
    '            month = "0" + Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Month.ToString()
    '        Else
    '            month = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Month.ToString()
    '        End If
    '        If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Day < 10 Then
    '            day = "0" + Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Day.ToString()
    '        Else
    '            day = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Day.ToString()
    '        End If
    '        If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Month < 10 Then
    '            month1 = "0" + Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Month.ToString()
    '        Else
    '            month1 = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Month.ToString()
    '        End If
    '        If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Day < 10 Then
    '            day1 = "0" + Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Day.ToString()
    '        Else
    '            day1 = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Day.ToString()
    '        End If
    '        input = <mediaplan>
    '                    <PreEvalPeriod>
    '                        <StartDate><%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Year.ToString() + month + day %></StartDate>
    '                        <EndDate><%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Year.ToString() + month1 + day1 %></EndDate>
    '                    </PreEvalPeriod>
    '                </mediaplan>
    '        If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.LbDayParts.Items.Count > 0 Then
    '            Dim dayparts As XElement =
    '        <DayParts></DayParts>
    '            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.LbDayParts.Items.Count - 1
    '                '<day_part>0200-0200</day_part>
    '                '  <day_part>0200-0200</day_part>
    '                Dim dpart As XElement = New XElement("DayPart")
    '                dpart.Value = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.LbDayParts.Items(index)
    '                dayparts.Add(dpart)
    '            Next
    '            input.Add(dayparts)
    '        End If
    '        Dim tg As XElement = XElement.Load(tgDirectoryPath + plantgname + ".xml")
    '        Dim TG_MGElement As XElement =
    '          <tg name=<%= plantgname %> cs=<%= tg.Element("cs").Value %> sec=<%= tg.Element("sec").Value %> sex=<%= tg.Element("sex").Value %> age=<%= tg.Element("age").Value %>>
    '          </tg>
    '        Dim allMGElement As XElement = <mg name="TotalMarkets" type="group"></mg>
    '        For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
    '            ' Dim doc2 As XmlDocument = New XmlDocument()
    '            '  doc2.Load()
    '            '  TG_MGElement.Add(XElement.Load(Path.GetTempPath() + "\\TGS\\" + plantgname + ".xml"))
    '            Dim mg As XElement = XElement.Load(mgDirectoryPath + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim() + ".xml")
    '            TG_MGElement.Add(mg)
    '            allMGElement.Add(mg.Elements())
    '        Next
    '        TG_MGElement.Add(allMGElement)
    '        input.Add(TG_MGElement)
    '        Dim planType As String = String.Empty
    '        '   Dim durations As Data.DataTable = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Copy().DefaultView.ToTable(True, "Duration")
    '        If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
    '            planType = "clubbed"
    '            Dim plan As XElement =
    '          <plan></plan>
    '            Dim period As XElement =
    '          <period StartDate=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Year.ToString() + Month() + Day() %> EndDate=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Year.ToString() + month1 + day1 %> year=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Year.ToString() %> WeekNum=<%= String.Empty %>></period>
    '            Dim brandSelection As XElement = <brandselection></brandselection>

    '            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.lbSelectedChannels.Items.Count - 1
    '                '   brandSelection = <brandselection channel=<%= tp.UcChannels.lbSelectedChannels.Items(index).ToString %>/>

    '                For index1 = 0 To tp.UcAdvertiser1.lbSelectedAdvertisers.Items.Count - 1
    '                    For index2 = 0 To tp.UcCategory1.lbSelectedCategories.Items.Count - 1

    '                        For index3 = 0 To tp.UcBrand1.lbSelectedBrands.Items.Count - 1

    '                            For index4 = 0 To tp.UcVariant1.lbSelectedVariants.Items.Count - 1
    '                                brandSelection = <brandselection channel=<%= tp.UcChannels.lbSelectedChannels.Items(index).ToString %> guid=<%= System.Guid.NewGuid %>/>


    '                            Next


    '                        Next


    '                    Next


    '                Next

    '            Next

    '        Else

    '        End If

    '    Catch ex As Exception

    '    End Try
    'End Function
    Public Function ConstructMarketSummaryInputXML() As XElement
        Dim input As XElement
        'Dim pchannels As Data.DataTable = Globals.Ribbons.MSprintExRibbon.GetGridTable()

        'If pchannels.Rows.Count = 0 Then

        '    If Globals.Ribbons.MSprintExRibbon.mappedchannels Is Nothing Then
        '        pchannels = New Data.DataTable()
        '    Else
        '        pchannels = Globals.Ribbons.MSprintExRibbon.mappedchannels
        '    End If


        'End If

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
            Dim tg As XElement = XElement.Load(tgDirectoryPath + plantgname + ".xml")
            Dim TG_MGElement As XElement =
              <tg name=<%= plantgname %> cs=<%= tg.Element("cs").Value %> sec=<%= tg.Element("sec").Value %> sex=<%= tg.Element("sex").Value %> age=<%= tg.Element("age").Value %>>
              </tg>
            Dim allMGElement As XElement = <mg name="TotalMarkets" type="group"></mg>
            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                ' Dim doc2 As XmlDocument = New XmlDocument()
                '  doc2.Load()
                '  TG_MGElement.Add(XElement.Load(Path.GetTempPath() + "\\TGS\\" + plantgname + ".xml"))
                Dim mg As XElement = XElement.Load(mgDirectoryPath + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim() + ".xml")
                TG_MGElement.Add(mg)
                allMGElement.Add(mg.Elements())
            Next
            TG_MGElement.Add(allMGElement)
            input.Add(TG_MGElement)
            Dim planType As String = String.Empty
            '   Dim durations As Data.DataTable = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Copy().DefaultView.ToTable(True, "Duration")
            If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                planType = "clubbed"
                Dim plan As XElement =
              <plan></plan>
                Dim period As XElement =
                <period StartDate=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Year.ToString() + month + day %> EndDate=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Year.ToString() + month1 + day1 %> year=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Year.ToString() %> WeekNum=<%= String.Empty %>></period>
                'For index1 = 0 To durations.Rows.Count - 1
                'Dim duration As XElement
                '  Dim channelname As String = pchannels.Rows(index1)("MCName").ToString()
                ' Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                'Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + channelname + "'")(0)("ID")
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

                'If planType.Equals("clubbed") Then
                '    '  duration = <duration value=<%= durations.Rows(index1)(0).ToString() %>></duration>
                '    '<programme guid=<%= Guid %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Total Spots").ToString() %>>
                '    '</programme>
                '    ' Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID") = (index1 + 1).ToString()
                '    'ElseIf planType.Equals("weekwise") Then
                '    '    If Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Columns.Contains("Week " & dtweeks.Rows(index)("WeekNumber").ToString()) Then
                '    '        Dim col As String = "Week " & dtweeks.Rows(index)("WeekNumber").ToString()
                '    '        program =
                '    '       <programme guid=<%= (index1 + 1).ToString() %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)(col).ToString() %>>
                '    '       </programme>
                '    '        Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID") = (index1 + 1).ToString()
                '    '    End If
                'End If
                Dim selectedspts As XElement =
                       <spots></spots>
                If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then

                    If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count > 0 Then
                        '  Dim filter As String = String.Format("Duration = '{0}'", durations.Rows(index1)(0).ToString())
                        '  Dim guidrows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Select(filter)
                        ' For Each row In guidrows
                        'Dim filter1 As String = String.Format("GUID = '{0}'", Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID").ToString())
                        '  Dim srows As DataRow() = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Select(filter1)
                        ' If srows.Count > 0 Then

                        For index2 = 0 To Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count - 1
                            Dim spot As XElement =
                                <spot log=<%= Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots(index2)("Spot").ToString() %>></spot>
                            selectedspts.Add(spot)
                        Next

                    End If
                    '   Next

                    '    End If
                End If
                'duration.Add(selectedspts)
                period.Add(selectedspts)

                '    Next
                plan.Add(period)
                input.Add(plan)
            ElseIf Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                planType = "weekwise"

                Dim plan As XElement =
                <plan type=<%= planType %>></plan>
                Dim dtweeks As Data.DataTable = CType(Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dgvWeeks.DataSource, Data.DataTable)

                If Not (Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks Is Nothing) Then
                    Dim period As XElement
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
                        period = <period StartDate=<%= startdate.Year.ToString() + monthh1 + dayy1 %> EndDate=<%= enddate.Year.ToString() + month11 + day11 %> year=<%= dtweeks.Rows(index)("Year").ToString() %> WeekNum=<%= dtweeks.Rows(index)("WeekNumber").ToString() %>></period>

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
                        '  Dim program As XElement
                        '   For index1 = 0 To Globals.Ribbons.MSprintExRibbon.xecelTable.Rows.Count - 1

                        '  Dim channelname As String = pchannels.Select("PCName='" + Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() + "'")(0)("MCName")
                        '  Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                        'Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() + "'")(0)("ID")
                        '   Dim channelname As String = pchannels.Select("PCName='" + Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() + "'")(0)("MCName")
                        ' Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                        '   Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + channelname + "'")(0)("ID")
                        'If planType.Equals("clubbed") Then
                        '    '      program =
                        '    ''<programme guid=<%= (index1 + 1).ToString() %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Total Spots").ToString() %>>
                        '    '</programme>
                        '    ' Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID") = (index1 + 1).ToString()
                        'ElseIf planType.Equals("weekwise") Then
                        '    ''Dim guid As String = Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID").ToString()

                        '    'If guid.Trim().Length = 0 Then
                        '    '    Dim guidval As String = GetGUIDFromCopy(Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1))
                        '    '    If guidval.Length = 0 Then
                        '    '        guid = System.Guid.NewGuid.ToString()
                        '    '        Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID") = guid
                        '    '    Else
                        '    '        guid = guidval
                        '    '    End If

                        '    'End If
                        '    If Globals.Ribbons.MSprintExRibbon.xecelTable.Columns.Contains("Week " & dtweeks.Rows(index)("WeekNumber").ToString()) Then
                        '        Dim col As String = "Week " & dtweeks.Rows(index)("WeekNumber").ToString()
                        '        '  program =
                        '        '<programme guid=<%= guid %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)(col).ToString() %>>
                        '        '</programme>
                        '        '  Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID") = (index1 + 1).ToString()
                        '    End If
                        'End If
                        Dim selectedspts As XElement =
                               <spots></spots>
                        If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then

                            If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count > 0 Then
                                Dim filter As String = String.Format("WeekNum= {0}", Convert.ToInt32(dtweeks.Rows(index)("WeekNumber").ToString()))

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
                        ' program.Add()
                        period.Add(selectedspts)
                        plan.Add(period)
                    Next

                    '  Next
                    input.Add(plan)
                End If
            End If
            Globals.Ribbons.MSprintExRibbon.UpdateUsageReport("MarketSummaryWS", Globals.Ribbons.MSprintExRibbon.tpSelections.tbClientName.Text.Trim, Globals.Ribbons.MSprintExRibbon.tpSelections.tbBrandValue.Text.Trim, input, 0)

        Catch ex As Exception
            LogMpsrintExException("Exception occured while constructing input XML for Duration Summary" + ex.Message)
            Throw ex
        End Try
        Return input
    End Function

    Public Function ConstructDurationSummaryInputXML() As XElement
        Dim input As XElement
        'Dim pchannels As Data.DataTable = Globals.Ribbons.MSprintExRibbon.GetGridTable()

        'If pchannels.Rows.Count = 0 Then

        '    If Globals.Ribbons.MSprintExRibbon.mappedchannels Is Nothing Then
        '        pchannels = New Data.DataTable()
        '    Else
        '        pchannels = Globals.Ribbons.MSprintExRibbon.mappedchannels
        '    End If


        'End If

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
            Dim tg As XElement = XElement.Load(tgDirectoryPath + plantgname + ".xml")
            Dim TG_MGElement As XElement =
              <tg name=<%= plantgname %> cs=<%= tg.Element("cs").Value %> sec=<%= tg.Element("sec").Value %> sex=<%= tg.Element("sex").Value %> age=<%= tg.Element("age").Value %>>
              </tg>
            Dim allMGElement As XElement = <mg name="TotalMarkets" type="group"></mg>
            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                ' Dim doc2 As XmlDocument = New XmlDocument()
                '  doc2.Load()
                '  TG_MGElement.Add(XElement.Load(Path.GetTempPath() + "\\TGS\\" + plantgname + ".xml"))
                Dim mg As XElement = XElement.Load(mgDirectoryPath + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim() + ".xml")
                TG_MGElement.Add(mg)
                allMGElement.Add(mg.Elements())
            Next
            TG_MGElement.Add(allMGElement)
            input.Add(TG_MGElement)
            Dim planType As String = String.Empty
            Dim durations As Data.DataTable = Globals.Ribbons.MSprintExRibbon.xecelTable.Copy().DefaultView.ToTable(True, "Duration")
            If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                planType = "clubbed"
                Dim plan As XElement =
              <plan></plan>
                Dim period As XElement =
                <period StartDate=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Year.ToString() + month + day %> EndDate=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Year.ToString() + month1 + day1 %> year=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Year.ToString() %> WeekNum=<%= String.Empty %>></period>
                For index1 = 0 To durations.Rows.Count - 1
                    Dim duration As XElement
                    '  Dim channelname As String = pchannels.Rows(index1)("MCName").ToString()
                    ' Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                    'Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + channelname + "'")(0)("ID")
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
                        duration = <duration value=<%= durations.Rows(index1)(0).ToString() %>></duration>
                        '<programme guid=<%= Guid %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Total Spots").ToString() %>>
                        '</programme>
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
                           <spots></spots>
                    If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then

                        If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count > 0 Then
                            Dim filter As String = String.Format("Duration = '{0}'", durations.Rows(index1)(0).ToString())
                            Dim guidrows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Select(filter)
                            For Each row In guidrows
                                Dim filter1 As String = String.Format("GUID = '{0}'", row("GUID").ToString())
                                Dim srows As DataRow() = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Select(filter1)
                                If srows.Count > 0 Then

                                    For index2 = 0 To srows.Count - 1
                                        Dim spot As XElement =
                                            <spot log=<%= srows(index2)("Spot").ToString() %>></spot>
                                        selectedspts.Add(spot)
                                    Next

                                End If
                            Next

                        End If
                    End If
                    duration.Add(selectedspts)
                    period.Add(duration)

                Next
                plan.Add(period)
                input.Add(plan)
            ElseIf Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                planType = "weekwise"

                Dim plan As XElement =
                <plan type=<%= planType %>></plan>
                Dim dtweeks As Data.DataTable = CType(Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dgvWeeks.DataSource, Data.DataTable)

                If Not (Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks Is Nothing) Then
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
                        Dim period As XElement =
                       <period StartDate=<%= startdate.Year.ToString() + monthh1 + dayy1 %> EndDate=<%= enddate.Year.ToString() + month11 + day11 %> year=<%= dtweeks.Rows(index)("Year").ToString() %> WeekNum=<%= dtweeks.Rows(index)("WeekNumber").ToString() %>></period>

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
                        For index1 = 0 To durations.Rows.Count - 1
                            Dim duration As XElement
                            duration = <duration value=<%= durations.Rows(index1)(0).ToString() %>></duration>
                            Dim selectedspts As XElement =
                        <spots></spots>
                            If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then
                                '  Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Duration(Sec)").ColumnName = "Duration"
                                If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count > 0 Then
                                    Dim filter As String = String.Format("{0}= {1}  and WeekNum={2}", EscapeLikeValue("Duration"), durations.Rows(index1)(0).ToString(), dtweeks.Rows(index)("WeekNumber").ToString())
                                    'Dim guidrows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Select(filter)
                                    'For Each row In guidrows
                                    ' Dim filter1 As String = String.Format("GUID = '{0}'", Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID").ToString())
                                    Dim srows As DataRow() = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Select(filter)
                                    If srows.Count > 0 Then

                                        For index2 = 0 To srows.Count - 1
                                            Dim spot As XElement =
                                                <spot log=<%= srows(index2)("Spot").ToString() %>></spot>
                                            selectedspts.Add(spot)
                                        Next

                                    End If
                                    ' Next

                                End If
                            End If
                            duration.Add(selectedspts)
                            'For index1 = 0 To Globals.Ribbons.MSprintExRibbon.xecelTable.Rows.Count - 1
                            '    Dim program As XElement
                            '  Dim channelname As String = pchannels.Select("PCName='" + Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() + "'")(0)("MCName")
                            '  Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                            'Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() + "'")(0)("ID")
                            '   Dim channelname As String = pchannels.Select("PCName='" + Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() + "'")(0)("MCName")
                            ' Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                            '   Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + channelname + "'")(0)("ID")
                            'If planType.Equals("clubbed") Then
                            '    '      program =
                            '    ''<programme guid=<%= (index1 + 1).ToString() %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Total Spots").ToString() %>>
                            '    '</programme>
                            '    Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID") = (index1 + 1).ToString()
                            'ElseIf planType.Equals("weekwise") Then
                            '    Dim guid As String = Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID").ToString()

                            '    If guid.Trim().Length = 0 Then
                            '        Dim guidval As String = GetGUIDFromCopy(Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1))
                            '        If guidval.Length = 0 Then
                            '            guid = System.Guid.NewGuid.ToString()
                            '            Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID") = guid
                            '        Else
                            '            guid = guidval
                            '        End If

                            '    End If
                            '    If Globals.Ribbons.MSprintExRibbon.xecelTable.Columns.Contains("Week " & dtweeks.Rows(index)("WeekNumber").ToString()) Then
                            '        Dim col As String = "Week " & dtweeks.Rows(index)("WeekNumber").ToString()
                            '        program =
                            '      <programme guid=<%= guid %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)(col).ToString() %>>
                            '      </programme>
                            '        ' Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID") = (index1 + 1).ToString()
                            '    End If
                            'End If
                            'Dim selectedspts As XElement =
                            '       <selected_spots></selected_spots>
                            'If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then

                            '    If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count > 0 Then
                            '        Dim filter As String = String.Format("GUID = '{0}' and WeekNum= {1} ", Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID").ToString(), Convert.ToInt32(dtweeks.Rows(index)("WeekNumber").ToString()))

                            '        Dim srows As DataRow() = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Select(filter)
                            '        If srows.Count > 0 Then

                            '            For index2 = 0 To srows.Count - 1
                            '                Dim spot As XElement =
                            '                    <spot log=<%= srows(index2)("Spot").ToString() %>></spot>
                            '                selectedspts.Add(spot)
                            '            Next

                            '        End If
                            '    End If
                            'End If
                            '   program.Add(selectedspts)
                            period.Add(duration)

                        Next
                        plan.Add(period)
                    Next
                    input.Add(plan)
                End If

            End If
            Globals.Ribbons.MSprintExRibbon.UpdateUsageReport("DurationSummaryWS", Globals.Ribbons.MSprintExRibbon.tpSelections.tbClientName.Text.Trim, Globals.Ribbons.MSprintExRibbon.tpSelections.tbBrandValue.Text.Trim, input, 0)

            If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Contains("Duration") Then
                Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns("Duration").ColumnName = "Duration(Sec)"
            End If


            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.AcceptChanges()
        Catch ex As Exception
            LogMpsrintExException("Exception occured while constructing input XML for Duration Summary" + ex.Message)
            Throw ex
        End Try
        Return input
    End Function
    Public Function ConstructCreativeSummaryInputXML() As XElement
        Dim input As XElement
        'Dim pchannels As Data.DataTable = Globals.Ribbons.MSprintExRibbon.GetGridTable()

        'If pchannels.Rows.Count = 0 Then

        '    If Globals.Ribbons.MSprintExRibbon.mappedchannels Is Nothing Then
        '        pchannels = New Data.DataTable()
        '    Else
        '        pchannels = Globals.Ribbons.MSprintExRibbon.mappedchannels
        '    End If


        'End If

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
            Dim tg As XElement = XElement.Load(tgDirectoryPath + plantgname + ".xml")
            Dim TG_MGElement As XElement =
              <tg name=<%= plantgname %> cs=<%= tg.Element("cs").Value %> sec=<%= tg.Element("sec").Value %> sex=<%= tg.Element("sex").Value %> age=<%= tg.Element("age").Value %>>
              </tg>
            Dim allMGElement As XElement = <mg name="TotalMarkets" type="group"></mg>
            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                ' Dim doc2 As XmlDocument = New XmlDocument()
                '  doc2.Load()
                '  TG_MGElement.Add(XElement.Load(Path.GetTempPath() + "\\TGS\\" + plantgname + ".xml"))
                Dim mg As XElement = XElement.Load(mgDirectoryPath + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim() + ".xml")
                TG_MGElement.Add(mg)
                allMGElement.Add(mg.Elements())
            Next
            TG_MGElement.Add(allMGElement)
            input.Add(TG_MGElement)
            Dim planType As String = String.Empty
            Dim creatives As Data.DataTable = Globals.Ribbons.MSprintExRibbon.xecelTable.Copy().DefaultView.ToTable(True, "Creative")
            If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                planType = "clubbed"
                Dim plan As XElement =
              <plan></plan>
                Dim period As XElement =
                <period StartDate=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Year.ToString() + month + day %> EndDate=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Year.ToString() + month1 + day1 %> year=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Year.ToString() %> WeekNum=<%= String.Empty %>></period>
                For index1 = 0 To creatives.Rows.Count - 1
                    Dim creative As XElement
                    '  Dim channelname As String = pchannels.Rows(index1)("MCName").ToString()
                    ' Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                    'Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + channelname + "'")(0)("ID")
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
                        creative = <creative value=<%= creatives.Rows(index1)(0).ToString() %>></creative>
                        '<programme guid=<%= Guid %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Total Spots").ToString() %>>
                        '</programme>
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
                           <spots></spots>
                    If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then

                        If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count > 0 Then
                            Dim filter As String = String.Format("Creative = '{0}'", creatives.Rows(index1)(0).ToString())
                            'Dim guidrows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Select(filter)
                            'For Each row In guidrows
                            '    Dim filter1 As String = String.Format("GUID = '{0}'", Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID").ToString())
                            Dim srows As DataRow() = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Select(filter)
                            If srows.Count > 0 Then

                                For index2 = 0 To srows.Count - 1
                                    Dim spot As XElement =
                                        <spot log=<%= srows(index2)("Spot").ToString() %>></spot>
                                    selectedspts.Add(spot)
                                Next

                            End If
                            '  Next

                        End If
                    End If
                    creative.Add(selectedspts)
                    period.Add(creative)

                Next
                plan.Add(period)
                input.Add(plan)
            ElseIf Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                planType = "weekwise"

                Dim plan As XElement =
                <plan type=<%= planType %>></plan>
                Dim dtweeks As Data.DataTable = CType(Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dgvWeeks.DataSource, Data.DataTable)

                If Not (Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks Is Nothing) Then
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
                        Dim period As XElement =
                       <period StartDate=<%= startdate.Year.ToString() + monthh1 + dayy1 %> EndDate=<%= enddate.Year.ToString() + month11 + day11 %> year=<%= dtweeks.Rows(index)("Year").ToString() %> WeekNum=<%= dtweeks.Rows(index)("WeekNumber").ToString() %>></period>
                        For index1 = 0 To creatives.Rows.Count - 1
                            Dim creative As XElement
                            creative = <creative value=<%= creatives.Rows(index1)(0).ToString() %>></creative>
                            Dim selectedspts As XElement =
                         <spots></spots>
                            If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then

                                If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count > 0 Then
                                    Dim filter As String = String.Format("Creative = '{0}' and WeekNum={1}", creatives.Rows(index1)(0).ToString(), dtweeks.Rows(index)("WeekNumber").ToString())
                                    'Dim guidrows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Select(filter)
                                    'For Each row In guidrows
                                    '    Dim filter1 As String = String.Format("GUID = '{0}'", Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID").ToString())
                                    Dim srows As DataRow() = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Select(filter)
                                    If srows.Count > 0 Then

                                        For index2 = 0 To srows.Count - 1
                                            Dim spot As XElement =
                                                <spot log=<%= srows(index2)("Spot").ToString() %>></spot>
                                            selectedspts.Add(spot)
                                        Next

                                    End If
                                    '  Next

                                End If
                            End If
                            creative.Add(selectedspts)
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
                            ''inpSpotTable.Columns.Add("Total Spots", Type.GetType("System.Int32"))
                            'For index1 = 0 To Globals.Ribbons.MSprintExRibbon.xecelTable.Rows.Count - 1
                            '    Dim program As XElement
                            '  Dim channelname As String = pchannels.Select("PCName='" + Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() + "'")(0)("MCName")
                            '  Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                            'Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() + "'")(0)("ID")
                            '   Dim channelname As String = pchannels.Select("PCName='" + Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() + "'")(0)("MCName")
                            ' Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                            ''   Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + channelname + "'")(0)("ID")
                            'If planType.Equals("clubbed") Then
                            '    '      program =
                            '    ''<programme guid=<%= (index1 + 1).ToString() %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Total Spots").ToString() %>>
                            '    '</programme>
                            '    Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID") = (index1 + 1).ToString()
                            'ElseIf planType.Equals("weekwise") Then
                            '    Dim guid As String = Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID").ToString()

                            '    If guid.Trim().Length = 0 Then
                            '        Dim guidval As String = GetGUIDFromCopy(Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1))
                            '        If guidval.Length = 0 Then
                            '            guid = System.Guid.NewGuid.ToString()
                            '            Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID") = guid
                            '        Else
                            '            guid = guidval
                            '        End If

                            '    End If
                            '    If Globals.Ribbons.MSprintExRibbon.xecelTable.Columns.Contains("Week " & dtweeks.Rows(index)("WeekNumber").ToString()) Then
                            '        Dim col As String = "Week " & dtweeks.Rows(index)("WeekNumber").ToString()
                            '        '  program =
                            '        '<programme guid=<%= guid %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)(col).ToString() %>>
                            '        '</programme>
                            '        ' Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID") = (index1 + 1).ToString()
                            '    End If
                            'End If
                            'Dim selectedspts As XElement =
                            '       <selected_spots></selected_spots>
                            'If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then

                            '    If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count > 0 Then
                            '        Dim filter As String = String.Format("GUID = '{0}' and WeekNum= {1} ", Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID").ToString(), Convert.ToInt32(dtweeks.Rows(index)("WeekNumber").ToString()))

                            '        Dim srows As DataRow() = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Select(filter)
                            '        If srows.Count > 0 Then

                            '            For index2 = 0 To srows.Count - 1
                            '                Dim spot As XElement =
                            '                    <spot log=<%= srows(index2)("Spot").ToString() %>></spot>
                            '                selectedspts.Add(spot)
                            '            Next

                            '        End If
                            '    End If
                            'End If
                            '   program.Add(selectedspts)
                            period.Add(creative)

                        Next
                        plan.Add(period)
                    Next
                    input.Add(plan)
                End If
            End If
            Globals.Ribbons.MSprintExRibbon.UpdateUsageReport("CreativeSummaryWS", Globals.Ribbons.MSprintExRibbon.tpSelections.tbClientName.Text.Trim, Globals.Ribbons.MSprintExRibbon.tpSelections.tbBrandValue.Text.Trim, input, 0)

        Catch ex As Exception
            LogMpsrintExException("Exception occured while constructing input XML for CreativeSummary" + ex.Message)
            Throw ex
        End Try
        Return input
    End Function
    Public Function ConstructChannelSummaryInputXML() As XElement

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
            Dim tg As XElement = XElement.Load(tgDirectoryPath + plantgname + ".xml")
            Dim TG_MGElement As XElement =
              <tg name=<%= plantgname %> cs=<%= tg.Element("cs").Value %> sec=<%= tg.Element("sec").Value %> sex=<%= tg.Element("sex").Value %> age=<%= tg.Element("age").Value %>>
              </tg>
            Dim allMGElement As XElement = <mg name="TotalMarkets" type="group"></mg>
            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                ' Dim doc2 As XmlDocument = New XmlDocument()
                '  doc2.Load()
                '  TG_MGElement.Add(XElement.Load(Path.GetTempPath() + "\\TGS\\" + plantgname + ".xml"))
                Dim mg As XElement = XElement.Load(mgDirectoryPath + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim() + ".xml")
                TG_MGElement.Add(mg)
                allMGElement.Add(mg.Elements())
            Next
            TG_MGElement.Add(allMGElement)
            input.Add(TG_MGElement)
            Dim planType As String = String.Empty

            If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                planType = "clubbed"
                Dim plan As XElement =
              <plan></plan>
                Dim period As XElement =
                <period StartDate=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Year.ToString() + month + day %> EndDate=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Year.ToString() + month1 + day1 %> year=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Year.ToString() %> WeekNum=<%= String.Empty %>></period>
                For index1 = 0 To pchannels.Rows.Count - 1
                    Dim Channel As XElement
                    Dim channelname As String = pchannels.Rows(index1)("MCName").ToString()
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
                        Channel = <channel name=<%= channelname %> code=<%= channelcode %>></channel>
                        '<programme guid=<%= Guid %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Total Spots").ToString() %>>
                        '</programme>
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
                           <spots></spots>
                    If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then

                        If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count > 0 Then
                            Dim filter As String = String.Format("Channel = '{0}'", channelname)

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
                    Channel.Add(selectedspts)
                    period.Add(Channel)

                Next
                plan.Add(period)
                input.Add(plan)
            ElseIf Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                '  planType = "weekwise"

                Dim plan As XElement =
                <plan></plan>
                Dim dtweeks As Data.DataTable = CType(Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dgvWeeks.DataSource, Data.DataTable)

                If Not (Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks Is Nothing) Then
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
                        Dim period As XElement =
                       <period StartDate=<%= startdate.Year.ToString() + monthh1 + dayy1 %> EndDate=<%= enddate.Year.ToString() + month11 + day11 %> year=<%= dtweeks.Rows(index)("Year").ToString() %> WeekNum=<%= dtweeks.Rows(index)("WeekNumber").ToString() %>></period>
                        For index1 = 0 To pchannels.Rows.Count - 1
                            Dim Channel As XElement
                            Dim channelname As String = pchannels.Rows(index1)("MCName").ToString()
                            ' Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                            Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + channelname + "'")(0)("ID")
                            Channel = <channel name=<%= channelname %> code=<%= channelcode %>></channel>
                            Dim selectedspts As XElement =
                        <spots></spots>
                            If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then

                                If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count > 0 Then
                                    Dim filter As String = String.Format("Channel = '{0}' and WeekNum={1}", channelname, dtweeks.Rows(index)("WeekNumber"))

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
                            Channel.Add(selectedspts)
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
                            'For index1 = 0 To Globals.Ribbons.MSprintExRibbon.xecelTable.Rows.Count - 1
                            '    Dim program As XElement
                            '  Dim channelname As String = pchannels.Select("PCName='" + Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() + "'")(0)("MCName")
                            '  Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                            'Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() + "'")(0)("ID")
                            '  Dim channelname As String = pchannels.Select("PCName='" + Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Channel").ToString() + "'")(0)("MCName")
                            '  ' Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                            '  Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + channelname + "'")(0)("ID")
                            '  If planType.Equals("clubbed") Then
                            '      program =
                            '<programme guid=<%= (index1 + 1).ToString() %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Total Spots").ToString() %>>
                            '</programme>
                            '      Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID") = (index1 + 1).ToString()
                            '  ElseIf planType.Equals("weekwise") Then
                            '      Dim guid As String = Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID").ToString()

                            '      If guid.Trim().Length = 0 Then
                            '          Dim guidval As String = GetGUIDFromCopy(Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1))
                            '          If guidval.Length = 0 Then
                            '              guid = System.Guid.NewGuid.ToString()
                            '              Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID") = guid
                            '          Else
                            '              guid = guidval
                            '          End If

                            '      End If
                            '      If Globals.Ribbons.MSprintExRibbon.xecelTable.Columns.Contains("Week " & dtweeks.Rows(index)("WeekNumber").ToString()) Then
                            '          Dim col As String = "Week " & dtweeks.Rows(index)("WeekNumber").ToString()
                            '          program =
                            '         <programme guid=<%= guid %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)(col).ToString() %>>
                            '         </programme>
                            '          ' Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID") = (index1 + 1).ToString()
                            '      End If
                            '  End If
                            '  Dim selectedspts As XElement =
                            '         <selected_spots></selected_spots>
                            '  If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then

                            '      If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count > 0 Then
                            '          Dim filter As String = String.Format("GUID = '{0}' and WeekNum= {1} ", Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID").ToString(), Convert.ToInt32(dtweeks.Rows(index)("WeekNumber").ToString()))

                            '          Dim srows As DataRow() = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Select(filter)
                            '          If srows.Count > 0 Then

                            '              For index2 = 0 To srows.Count - 1
                            '                  Dim spot As XElement =
                            '                      <spot log=<%= srows(index2)("Spot").ToString() %>></spot>
                            '                  selectedspts.Add(spot)
                            '              Next

                            '          End If
                            '      End If
                            '  End If
                            '  program.Add(selectedspts)
                            period.Add(Channel)

                        Next
                        plan.Add(period)
                    Next
                    input.Add(plan)
                End If
            End If
            Globals.Ribbons.MSprintExRibbon.UpdateUsageReport("Channel SummaryWS", Globals.Ribbons.MSprintExRibbon.tpSelections.tbClientName.Text.Trim, Globals.Ribbons.MSprintExRibbon.tpSelections.tbBrandValue.Text.Trim, input, 0)

        Catch ex As Exception
            LogMpsrintExException("Exception occured while constructing input XML for ChannelSummary" + ex.Message)
            Throw ex
        End Try
        Return input
    End Function
    Public Function ConstructProgAvgTVR() As XElement
        '  Public Function ConstructInputRnFXML() As XElement
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
            Dim tg As XElement = XElement.Load(tgDirectoryPath + plantgname + ".xml")
            Dim TG_MGElement As XElement =
              <tg name=<%= plantgname %> cs=<%= tg.Element("cs").Value %> sec=<%= tg.Element("sec").Value %> sex=<%= tg.Element("sex").Value %> age=<%= tg.Element("age").Value %>>
              </tg>
            Dim allMGElement As XElement = <mg name="TotalMarkets" type="group"></mg>
            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                ' Dim doc2 As XmlDocument = New XmlDocument()
                '  doc2.Load()
                '  TG_MGElement.Add(XElement.Load(Path.GetTempPath() + "\\TGS\\" + plantgname + ".xml"))
                Dim mg As XElement = XElement.Load(mgDirectoryPath + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim() + ".xml")
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
                Globals.Ribbons.MSprintExRibbon.xecelTable = New Data.DataTable()

                If Globals.Ribbons.MSprintExRibbon.reorderedChannels Is Nothing Then
                    ' Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows.CopyTo(rows, 0)
                    Globals.Ribbons.MSprintExRibbon.xecelTable = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Copy()

                    ' ConstructXMLFromRows(rows, period, planType, pchannels)
                Else
                    Globals.Ribbons.MSprintExRibbon.xecelTable = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Clone()
                    For Each channel As String In Globals.Ribbons.MSprintExRibbon.reorderedChannels
                        ' Dim channelname As String = pchannels.Select("MCName='" + channel + "'")(0)("PCName")
                        Dim rows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Select("Channel='" + channel + "'")
                        For Each row As Data.DataRow In rows
                            Globals.Ribbons.MSprintExRibbon.xecelTable.ImportRow(row)
                        Next
                        ' ConstructXMLFromRows(rows, period, planType, pchannels)
                    Next
                End If
                For index1 = 0 To Globals.Ribbons.MSprintExRibbon.xecelTable.Rows.Count - 1
                    Dim program As XElement
                    Dim channelname As String = pchannels.Select("PCName='" + Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Channel").ToString() + "'")(0)("MCName")
                    ' Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                    Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + channelname + "'")(0)("ID")
                    Dim guid As String = Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID").ToString()

                    If guid.Trim().Length = 0 Then
                        Dim guidval As String = GetGUIDFromCopy(Globals.Ribbons.MSprintExRibbon.xecelTable(index1))
                        If guidval.Length = 0 Then
                            guid = System.Guid.NewGuid.ToString()
                            Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("GUID") = guid
                        Else
                            guid = guidval
                        End If

                    End If

                    If planType.Equals("clubbed") Then

                        'NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("Total Spots").ToString() %>
                        program =
  <programme guid=<%= guid %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("RatePer10Sec").ToString() %>>
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
                    'Dim selectedspts As XElement =
                    '       <selected_spots></selected_spots>
                    'If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then

                    '    If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count > 0 Then
                    '        Dim filter As String = String.Format("GUID = '{0}'", Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("GUID").ToString())

                    '        Dim srows As DataRow() = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Select(filter)
                    '        If srows.Count > 0 Then

                    '            For index2 = 0 To srows.Count - 1
                    '                Dim spot As XElement =
                    '                    <spot log=<%= srows(index2)("Spot").ToString() %>></spot>
                    '                selectedspts.Add(spot)
                    '            Next

                    '        End If
                    '    End If
                    'End If
                    ' program.Add(selectedspts)
                    period.Add(program)

                Next

                plan.Add(period)
                input.Add(plan)
            ElseIf Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                planType = "weekwise"

                Dim plan As XElement =
                <plan type=<%= planType %>></plan>
                Dim dtweeks As Data.DataTable = CType(Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dgvWeeks.DataSource, Data.DataTable)
                Globals.Ribbons.MSprintExRibbon.xecelTable = New Data.DataTable()
                If Not (Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks Is Nothing) Then
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
                        Dim period As XElement =
                       <period StartDate=<%= startdate.Year.ToString() + monthh1 + dayy1 %> EndDate=<%= enddate.Year.ToString() + month11 + day11 %> year=<%= dtweeks.Rows(index)("Year").ToString() %> WeekNum=<%= dtweeks.Rows(index)("WeekNumber").ToString() %>></period>

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

                        If Globals.Ribbons.MSprintExRibbon.reorderedChannels Is Nothing Then
                            ' Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows.CopyTo(rows, 0)
                            Globals.Ribbons.MSprintExRibbon.xecelTable = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable
                            'rows = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows.Cast(Of Data.DataRow)().ToArray()
                            'ConstructXMLFromRows(rows, period, planType, pchannels)
                        Else
                            Globals.Ribbons.MSprintExRibbon.xecelTable = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Clone()
                            For Each channel As String In Globals.Ribbons.MSprintExRibbon.reorderedChannels
                                ' Dim channelname As String = pchannels.Select("MCName='" + channel + "'")(0)("PCName")
                                Dim rows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Select("Channel='" + channel + "'")
                                '  ConstructXMLFromRows(rows, period, planType, pchannels)
                                For Each row As Data.DataRow In rows
                                    Globals.Ribbons.MSprintExRibbon.xecelTable.ImportRow(row)
                                Next
                            Next
                        End If
                        For index1 = 0 To Globals.Ribbons.MSprintExRibbon.xecelTable.Rows.Count - 1
                            Dim program As XElement
                            '  Dim channelname As String = pchannels.Select("PCName='" + Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() + "'")(0)("MCName")
                            '  Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                            'Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() + "'")(0)("ID")
                            Dim channelname As String = pchannels.Select("PCName='" + Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Channel").ToString() + "'")(0)("MCName")
                            ' Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                            Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + channelname + "'")(0)("ID")
                            If planType.Equals("clubbed") Then
                                program =
                          <programme guid=<%= (index1 + 1).ToString() %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Total Spots").ToString() %>>
                          </programme>
                                Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID") = (index1 + 1).ToString()
                            ElseIf planType.Equals("weekwise") Then
                                Dim guid As String = Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID").ToString()

                                If guid.Trim().Length = 0 Then
                                    Dim guidval As String = GetGUIDFromCopy(Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1))
                                    If guidval.Length = 0 Then
                                        guid = System.Guid.NewGuid.ToString()
                                        Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID") = guid
                                    Else
                                        guid = guidval
                                    End If

                                End If
                                If Globals.Ribbons.MSprintExRibbon.xecelTable.Columns.Contains("Week " & dtweeks.Rows(index)("WeekNumber").ToString()) Then
                                    Dim col As String = "Week " & dtweeks.Rows(index)("WeekNumber").ToString()
                                    Dim spotsVal As String = Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)(col).ToString()

                                    If spotsVal.Trim().Length = 0 Then
                                        spotsVal = "0"
                                    End If

                                    program =
                                   <programme guid=<%= guid %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= spotsVal %>>
                                   </programme>
                                    ' Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID") = (index1 + 1).ToString()
                                End If
                            End If
                            'Dim selectedspts As XElement =
                            '       <selected_spots></selected_spots>
                            'If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then

                            '    If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count > 0 Then
                            '        Dim filter As String = String.Format("GUID = '{0}' and WeekNum= {1} ", Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID").ToString(), Convert.ToInt32(dtweeks.Rows(index)("WeekNumber").ToString()))

                            '        Dim srows As DataRow() = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Select(filter)
                            '        If srows.Count > 0 Then

                            '            For index2 = 0 To srows.Count - 1
                            '                Dim spot As XElement =
                            '                    <spot log=<%= srows(index2)("Spot").ToString() %>></spot>
                            '                selectedspts.Add(spot)
                            '            Next

                            '        End If
                            '    End If
                            'End If
                            'program.Add(selectedspts)
                            period.Add(program)

                        Next
                        plan.Add(period)
                    Next
                    input.Add(plan)
                End If
            End If
            Globals.Ribbons.MSprintExRibbon.UpdateUsageReport("AvgProgramTVRWS", Globals.Ribbons.MSprintExRibbon.tpSelections.tbClientName.Text.Trim, Globals.Ribbons.MSprintExRibbon.tpSelections.tbBrandValue.Text.Trim, input, 0)

        Catch ex As Exception
            LogMpsrintExException("Exception occured while constructing input XML for RnF" + ex.Message)
            Throw ex
        End Try
        Return input
    End Function
    Public Function ConstructInputRnFXML(ByVal buttoninvoked As String) As XElement
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
            Dim tg As XElement = XElement.Load(tgDirectoryPath + plantgname + ".xml")
            Dim TG_MGElement As XElement =
              <tg name=<%= plantgname %> cs=<%= tg.Element("cs").Value %> sec=<%= tg.Element("sec").Value %> sex=<%= tg.Element("sex").Value %> age=<%= tg.Element("age").Value %>>
              </tg>
            Dim allMGElement As XElement = <mg name="TotalMarkets" type="group"></mg>
            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                ' Dim doc2 As XmlDocument = New XmlDocument()
                '  doc2.Load()
                '  TG_MGElement.Add(XElement.Load(Path.GetTempPath() + "\\TGS\\" + plantgname + ".xml"))
                Dim mg As XElement = XElement.Load(mgDirectoryPath + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim() + ".xml")
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
                Globals.Ribbons.MSprintExRibbon.xecelTable = New Data.DataTable()

                If Globals.Ribbons.MSprintExRibbon.reorderedChannels Is Nothing Then
                    ' Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows.CopyTo(rows, 0)
                    Globals.Ribbons.MSprintExRibbon.xecelTable = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Copy()

                    ' ConstructXMLFromRows(rows, period, planType, pchannels)
                Else
                    Globals.Ribbons.MSprintExRibbon.xecelTable = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Clone()
                    For Each channel As String In Globals.Ribbons.MSprintExRibbon.reorderedChannels
                        ' Dim channelname As String = pchannels.Select("MCName='" + channel + "'")(0)("PCName")
                        Dim rows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Select("Channel='" + channel + "'")
                        For Each row As Data.DataRow In rows
                            Globals.Ribbons.MSprintExRibbon.xecelTable.ImportRow(row)
                        Next
                        ' ConstructXMLFromRows(rows, period, planType, pchannels)
                    Next
                End If
                For index1 = 0 To Globals.Ribbons.MSprintExRibbon.xecelTable.Rows.Count - 1
                    Dim program As XElement
                    Dim channelname As String = pchannels.Select("PCName='" + Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Channel").ToString() + "'")(0)("MCName")
                    ' Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                    Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + channelname + "'")(0)("ID")
                    Dim guid As String = Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID").ToString()

                    If guid.Trim().Length = 0 Then
                        Dim guidval As String = GetGUIDFromCopy(Globals.Ribbons.MSprintExRibbon.xecelTable(index1))
                        If guidval.Length = 0 Then
                            guid = System.Guid.NewGuid.ToString()
                            Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("GUID") = guid
                        Else
                            guid = guidval
                        End If

                    End If

                    If planType.Equals("clubbed") Then
                        program =
                  <programme guid=<%= guid %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("Duration").ToString() %> MinTVR=<%= Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("Min TVR").ToString() %> MaxTVR=<%= Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("Max TVR").ToString() %> NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("Total Spots").ToString() %>>
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
                    Dim selectedspts As XElement
                    If buttoninvoked.Equals("All Summary") Then
                        selectedspts =
                      <spots></spots>
                    Else
                        selectedspts =
                         <selected_spots></selected_spots>
                    End If

                  
                    If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then

                        If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count > 0 Then
                            Dim filter As String = String.Format("GUID = '{0}'", Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("GUID").ToString())

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

                Next

                plan.Add(period)
                input.Add(plan)
            ElseIf Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                planType = "weekwise"

                Dim plan As XElement =
                <plan type=<%= planType %>></plan>
                Dim dtweeks As Data.DataTable = CType(Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dgvWeeks.DataSource, Data.DataTable)
                Globals.Ribbons.MSprintExRibbon.xecelTable = New Data.DataTable()
                If Not (Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks Is Nothing) Then
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
                        Dim period As XElement =
                       <period StartDate=<%= startdate.Year.ToString() + monthh1 + dayy1 %> EndDate=<%= enddate.Year.ToString() + month11 + day11 %> year=<%= dtweeks.Rows(index)("Year").ToString() %> WeekNum=<%= dtweeks.Rows(index)("WeekNumber").ToString() %>></period>

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

                        If Globals.Ribbons.MSprintExRibbon.reorderedChannels Is Nothing Then
                            ' Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows.CopyTo(rows, 0)
                            Globals.Ribbons.MSprintExRibbon.xecelTable = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable
                            'rows = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows.Cast(Of Data.DataRow)().ToArray()
                            'ConstructXMLFromRows(rows, period, planType, pchannels)
                        Else
                            Globals.Ribbons.MSprintExRibbon.xecelTable = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Clone()
                            For Each channel As String In Globals.Ribbons.MSprintExRibbon.reorderedChannels
                                ' Dim channelname As String = pchannels.Select("MCName='" + channel + "'")(0)("PCName")
                                Dim rows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Select("Channel='" + channel + "'")
                                '  ConstructXMLFromRows(rows, period, planType, pchannels)
                                For Each row As Data.DataRow In rows
                                    Globals.Ribbons.MSprintExRibbon.xecelTable.ImportRow(row)
                                Next
                            Next
                        End If
                        For index1 = 0 To Globals.Ribbons.MSprintExRibbon.xecelTable.Rows.Count - 1
                            Dim program As XElement
                            '  Dim channelname As String = pchannels.Select("PCName='" + Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() + "'")(0)("MCName")
                            '  Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                            'Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() + "'")(0)("ID")
                            Dim channelname As String = pchannels.Select("PCName='" + Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Channel").ToString() + "'")(0)("MCName")
                            ' Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                            Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + channelname + "'")(0)("ID")
                            If planType.Equals("clubbed") Then
                                program =
                          <programme guid=<%= (index1 + 1).ToString() %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Creative").ToString() %> MinTVR=<%= Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("Min TVR").ToString() %> MaxTVR=<%= Globals.Ribbons.MSprintExRibbon.xecelTable(index1)("Max TVR").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Total Spots").ToString() %>>
                          </programme>
                                Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID") = (index1 + 1).ToString()
                            ElseIf planType.Equals("weekwise") Then
                                Dim guid As String = Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID").ToString()

                                If guid.Trim().Length = 0 Then
                                    Dim guidval As String = GetGUIDFromCopy(Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1))
                                    If guidval.Length = 0 Then
                                        guid = System.Guid.NewGuid.ToString()
                                        Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID") = guid
                                    Else
                                        guid = guidval
                                    End If

                                End If
                                If Globals.Ribbons.MSprintExRibbon.xecelTable.Columns.Contains("Week " & dtweeks.Rows(index)("WeekNumber").ToString()) Then
                                    Dim col As String = "Week " & dtweeks.Rows(index)("WeekNumber").ToString()
                                    Dim spotsVal As String = Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)(col).ToString()

                                    If spotsVal.Trim().Length = 0 Then
                                        spotsVal = "0"
                                    End If

                                    program =
                                   <programme guid=<%= guid %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Programme").ToString() %> days=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Day").ToString() %> StartTime=<%= GetStartTime(Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("End Time").ToString()) %> CostPer10s=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("RatePer10Sec").ToString() %> caption=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Creative").ToString() %> AdDuration=<%= Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("Duration").ToString() %> NumberOfSpots=<%= spotsVal %>>
                                   </programme>
                                    ' Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("GUID") = (index1 + 1).ToString()
                                End If
                            End If
                            'Dim selectedspts As XElement =
                            '       <selected_spots></selected_spots>
                            Dim selectedspts As XElement
                            If buttoninvoked.Equals("All Summary") Then
                                selectedspts =
                              <spots></spots>
                            Else
                                selectedspts =
                                 <selected_spots></selected_spots>
                            End If
                            If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then

                                If Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count > 0 Then
                                    Dim filter As String = String.Format("GUID = '{0}' and WeekNum= {1} ", Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(index1)("GUID").ToString(), Convert.ToInt32(dtweeks.Rows(index)("WeekNumber").ToString()))

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

                        Next
                        plan.Add(period)
                    Next
                    input.Add(plan)
                End If
            End If
            Globals.Ribbons.MSprintExRibbon.UpdateUsageReport(buttoninvoked, Globals.Ribbons.MSprintExRibbon.tpSelections.tbClientName.Text.Trim, Globals.Ribbons.MSprintExRibbon.tpSelections.tbBrandValue.Text.Trim, input, 0)

        Catch ex As Exception
            LogMpsrintExException("Exception occured while constructing input XML for RnF" + ex.Message)
            Throw ex
        End Try
        Return input
    End Function
    Public Function ConstructXMLFromRows(ByVal rowsObject As Data.DataRow(), ByVal periodElement As XElement, ByVal planTypeVal As String, ByVal pChannelsObject As Data.DataTable)
        Try
            For index1 = 0 To Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows.Count - 1
                Dim program As XElement
                Dim channelname As String = pChannelsObject.Select("PCName='" + Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString() + "'")(0)("MCName")
                ' Dim channelname As String = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Rows(index1)("Channel").ToString()
                Dim channelcode As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name='" + channelname + "'")(0)("ID")
                Dim guid As String = rowsObject(index1)("GUID").ToString()

                If guid.Trim().Length = 0 Then
                    Dim guidval As String = GetGUIDFromCopy(rowsObject(index1))
                    If guidval.Length = 0 Then
                        guid = System.Guid.NewGuid.ToString()
                        rowsObject(index1)("GUID") = guid
                    Else
                        guid = guidval
                    End If

                End If

                If planTypeVal.Equals("clubbed") Then
                    program =
              <programme guid=<%= guid %> SeqNumber=<%= (index1 + 1).ToString() %> ChannelCode=<%= channelcode %> ChannelName=<%= channelname %> ProgName=<%= rowsObject(index1)("Programme").ToString() %> days=<%= rowsObject(index1)("Day").ToString() %> StartTime=<%= GetStartTime(rowsObject(index1)("Start Time").ToString()) %> EndTime=<%= GetEndTime(rowsObject(index1)("End Time").ToString()) %> CostPer10s=<%= rowsObject(index1)("RatePer10Sec").ToString() %> caption=<%= rowsObject(index1)("Creative").ToString() %> AdDuration=<%= rowsObject(index1)("Duration").ToString() %> NumberOfSpots=<%= rowsObject(index1)("Total Spots").ToString() %>>
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
                        Dim filter As String = String.Format("GUID = '{0}'", rowsObject(index1)("GUID").ToString())

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
                periodElement.Add(program)

            Next
        Catch ex As Exception

        End Try
    End Function

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
            Globals.Ribbons.MSprintExRibbon.UpdateUsageReport("AvailableSpotsWS", Globals.Ribbons.MSprintExRibbon.tpSelections.tbClientName.Text.Trim, Globals.Ribbons.MSprintExRibbon.tpSelections.tbBrandValue.Text.Trim, input, 0)

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
    Public Function GetGUIDFromCopy(ByVal dr As Data.DataRow) As String
        Dim guid As String = String.Empty
        Try
            'inpSpotTable.Columns.Add("Channel")
            'inpSpotTable.Columns.Add("Programme")
            'inpSpotTable.Columns.Add("Day")
            'inpSpotTable.Columns.Add("Start Time")
            'inpSpotTable.Columns.Add("End Time")
            'inpSpotTable.Columns.Add("RatePer10Sec")
            'inpSpotTable.Columns.Add("Creative")
            'inpSpotTable.Columns.Add("Duration")

            If Globals.Ribbons.MSprintExRibbon.xecellTableCopy Is Nothing Then
                Return guid
            Else
                Dim filter As String = String.Format("Channel='{0}' and Programme='{1}' and Day='{2}' and Start Time='{3}' and End Time='{4}' and RatePer10Sec='{5}' and Creative='{6}' and Duration='{7}'", dr("Channel").ToString(), dr("Programme").ToString(), dr("Day").ToString(), GetStartTime(dr("Start Time").ToString()), GetEndTime(dr("End Time").ToString()), dr("RatePer10Sec").ToString(), dr("Creative").ToString(), dr("Duration").ToString())
                Dim rows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.xecellTableCopy.Select(filter)

                If rows.Length > 0 Then
                    guid = rows(0)("GUID").ToString()
                Else
                    guid = String.Empty
                End If
            End If
        Catch ex As Exception
            guid = String.Empty
        End Try
        Return guid
    End Function
    Public Function GetStartTime(ByVal timespanDoubleVal As String) As String
        'Dim starTime As String = String.Empty
        'For Each cell As Microsoft.Office.Interop.Excel.Range In loSpotSelection.ListRows(rIndex + 1).Range.Cells

        '    If cell.Column = 6 Then
        '        starTime = cell.Text

        '    End If

        'Next
        Dim timespanString = String.Empty
        Dim doubleVal As Double = Convert.ToDouble(timespanDoubleVal)
        Dim dateVal As Date = Date.FromOADate(doubleVal)
        Return dateVal.TimeOfDay.ToString("hh\:mm")
        ' Return starTime
    End Function
    Public Function GetEndTime(ByVal timespanDoubleVal As String) As String
        Dim timespanString = String.Empty
        Dim doubleVal As Double = Convert.ToDouble(timespanDoubleVal)
        Dim dateVal As Date = Date.FromOADate(doubleVal)
        Return dateVal.TimeOfDay.ToString("hh\:mm")
    End Function
End Module
