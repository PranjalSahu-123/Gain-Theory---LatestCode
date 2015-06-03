Imports Microsoft.Office.Tools
Imports Microsoft.Office.Interop.Excel
Imports System
Imports System.Data

Module GenreShare
    Friend listObject, listobject1 As Excel.ListObject
    Friend nativeSheet, newSheet, vstoWorkbook As Microsoft.Office.Interop.Excel.Worksheet
    Public Function ConstructGenreShareInputXML(ByVal plantg As String, ByVal reftg As String) As XElement
        Dim input As XElement = New XElement("input")
        Try
            Dim month, month1 As String
            Dim day, day1 As String

            '  Button1.Enabled = False
            ' lbGetting.Text = "Getting Genre Share for chosen TG-MGs..."
            ' lbGetting.Refresh()
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


            input =
                <input>
                    <pre-eval-period>
                        <startdate><%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Year.ToString() + month + day %></startdate>
                        <enddate><%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Year.ToString() + month1 + day1 %></enddate>
                    </pre-eval-period>



                </input>

            '<TG_MG name = "CS 15-44~mg1" type = "Planning">
            '  Dim tgs As XElement = New XElement("targetgroups")
            'Dim doc As XmlDocument = New XmlDocument()
            '  doc.Load()
            ' tgs.Add(XElement.Load(Path.GetTempPath() + "\\TGS\\" + plantg + ".xml"))

            'If reftg <> plantg Then
            '    '  Dim doc1 As XmlDocument = New XmlDocument()
            '    '  doc1.Load()
            '    tgs.Add(XElement.Load(Path.GetTempPath() + "\\TGS\\" + reftg + ".xml"))
            'End If

            '   Dim markets As XElement = New XElement("markets")
            Dim dayparts As XElement =
               <day_parts></day_parts>
            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.LbDayParts.Items.Count - 1
                '<day_part>0200-0200</day_part>
                '  <day_part>0200-0200</day_part>
                Dim dpart As XElement = New XElement("day_part")
                dpart.Value = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.LbDayParts.Items(index)
                dayparts.Add(dpart)
            Next
            input.Add(dayparts)
            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                Dim TG_MGElement As XElement =
                    <TG_MG name=<%= String.Format("{0}~{1}", plantg, Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim()) %> type="Planning">
                    </TG_MG>
                TG_MGElement.Add(XElement.Load(Globals.Ribbons.MSprintExRibbon.tgDirectoryPath + plantg + ".xml"))
                TG_MGElement.Add(XElement.Load(Globals.Ribbons.MSprintExRibbon.mgDirectoryPath + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim() + ".xml"))
                input.Add(TG_MGElement)
                ' Dim doc2 As XmlDocument = New XmlDocument()
                '  doc2.Load()
                ' markets.Add(XElement.Load(Path.GetTempPath() + "\\MGS\\" + tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim() + ".xml"))
            Next


            'If reftg.Length > 0 Then
            '    For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items.Count - 1
            '        Dim TG_MGElement As XElement =
            '         <TG_MG name=<%= String.Format("{0}~{1}", reftg, Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index).ToString().Trim()) %> type="Reference">
            '         </TG_MG>
            '        TG_MGElement.Add(XElement.Load(Globals.Ribbons.MSprintExRibbon.tgDirectoryPath + reftg + ".xml"))
            '        TG_MGElement.Add(XElement.Load(Globals.Ribbons.MSprintExRibbon.mgDirectoryPath + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index).ToString().Trim() + ".xml"))
            '        input.Add(TG_MGElement)
            '    Next
            'End If

            Globals.Ribbons.MSprintExRibbon.UpdateUsageReport("Genre ShareWS", Globals.Ribbons.MSprintExRibbon.tpSelections.tbClientName.Text.Trim, Globals.Ribbons.MSprintExRibbon.tpSelections.tbBrandValue.Text.Trim, input, 0)
        Catch ex As Exception
            LogMpsrintExException("Exception occured while constructing Genre Share XMl")
            Throw ex
        End Try
        Return input
    End Function

    Public Function DisplayGenreShareDetailsOnSheet(ByVal ds As DataSet, Optional ByVal viewVal As MSprintExRibbon.GenreShareView = MSprintExRibbon.GenreShareView.All, Optional ByVal fromPane As Boolean = False)
        Try
            Dim plantg As String = String.Empty
            Dim reftg As String = String.Empty


            Dim dtable As System.Data.DataTable = DirectCast(Globals.Ribbons.MSprintExRibbon.tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            plantg = dtable.Rows(0)(1).ToString().Trim()
            '  reftg = dtable.Rows(1)(1).ToString()
            Dim genres As Data.DataTable = ds.Tables(0).DefaultView.ToTable(True, "Genre")

            'Dim dt1 As System.Data.DataTable = New System.Data.DataTable()
            ''  Dim copyGenreTab As System.Data.DataTable = ds.Tables(0).Copy()

            ''  Dim copyto10 As System.Data.DataTable = ds.Tables(1).Copy()
            'dt1.Columns.Add("Rank", System.Type.GetType("System.Int32"))
            'dt1.Columns.Add("Channel")
            '' dt1.Columns.Add(plantg + "-" + tpSelections.UcMarkets1.lbPlan.Items(0).ToString() + " GRP")
            ''  dt1.Columns.Add(reftg + "-" + tpSelections.UcMarkets1.lbRef.Items(0).ToString() + " GRP")
            'dt.Columns.Add("Genre")
            '  dt.Columns.Add("Plan " + plantg + "-" + tpSelections.UcMarkets1.lbPlan.Items(0).ToString() + " GRP")
            '  dt.Columns.Add("Ref" + reftg + "-" + tpSelections.UcMarkets1.lbRef.Items(0).ToString() + " GRP")
            ' Dim expression As String = "OrderQuantity = 2 and OrderID = 2" 
            ' Sort descending by column named CompanyName. 
            Dim sortOrder As String = "GRPShare in(%) DESC"
            Dim foundRows As DataRow()
            ' Dim exptopten As String = "TG"
            ' Use the Select method to find all rows matching the filter.
            ' Dim dtables As System.Data.DataTable() = New System.Data.DataTable(tpSelections.UcMarkets1.lbPlan.Items.Count) {}
            ' foundtopten = planDs.Tables(1).[Select](expression, sortOrder)

            For index = 0 To genres.Rows.Count - 1

                For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                    Try


                        Dim expression1 As String = "Tgroup = '" + EscapeLikeValue(plantg) + "' and Mgroup = '" + EscapeLikeValue(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString()) + "' and Genre = '" + genres.Rows(index)(0).ToString() + "'"
                        foundRows = ds.Tables(0).[Select](expression1)

                        If foundRows.Count > 0 Then
                            ' dtables(index) = New System.Data.DataTable()
                            'dtables(index) = foundRows.CopyToDataTable()

                            If Not (genres.Columns.Contains(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRPShare in(%)")) Then
                                genres.Columns.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRPShare in(%)")
                            End If

                            'If Not (genres.Columns.Contains(genres.Columns.Contains(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRPShare"))) Then
                            '    genres.Columns.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRPShare")
                            'End If


                            genres.Rows(index)(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRPShare in(%)") = foundRows(0).Item("GRPShare in(%)")
                            ' genres.Rows(index)(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRPShare") = foundRows(0).Item("GRPShare")
                        Else
                            genres.Rows(index)(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRPShare in(%)") = 0
                            'dt1.Columns.Add(tpSelections.UcMarkets1.lbPlan.Items(index).ToString() + "-GRP")
                        End If
                    Catch ex As Exception
                        LogMpsrintExException("Exception occured while filtering data for given mg." + ex.Message)
                    End Try
                Next


            Next
            ' foundRefCopy = copyGenreTab.Select(expcopy, sortOrder)
            'For Each row As DataRow In foundRows
            '    Dim dr As DataRow = dt.NewRow()
            '    dr("Genre") = row("Genre")
            '    For index = 0 To tpSelections.UcMarkets1.lbPlan.Items.Count - 1


            '        Dim expcopy As String = "Tgroup = '" + plantg + "' and Mgroup = '" + tpSelections.UcMarkets1.lbPlan.Items(index).ToString() + "' and Genre = '" + row("Genre").ToString() + "'"

            '        Dim drr As DataRow() = planDs.Tables(0).[Select](expcopy, sortOrder)
            '        If drr Is Nothing Or drr.Count = 0 Then
            '            dr(index + 1) = 0
            '        Else
            '            dr(index + 1) = drr(0).Item("GRP")
            '        End If
            '        ' dr(index + 1) = ds.Tables(0).[Select](expcopy, sortOrder)(0).Item("GRP")
            '        '  foundRefCopy()
            '        ' dr(index + 2) = ds.Tables(0).[Select](expcopy, sortOrder)(0).Item("GRP")

            '    Next
            '    dt.Rows.Add(dr)
            'Next
            'For Each row1 As DataRow In foundtopten
            '    Dim dr As DataRow = dt1.NewRow()
            '    dr("Rank") = row1("Rank")
            '    'Dim i As Integer = Convert.ToInt32(row1("Channel").ToString())
            '    'Dim ss As String = String.Empty
            '    'If i < 10 Then
            '    '    ss = "00" + i.ToString()
            '    'ElseIf i < 100 Then
            '    '    ss = "0" + i.ToString()
            '    'Else
            '    '    ss = row1("Channel").ToString()
            '    'End If
            '    'Dim drow As DataRow = dg.GetChannels().Select("TAM_CHANNEL_CODE = '" + ss + "'")(0)
            '    dr("Channel") = row1("Channel").ToString()
            '    For index = 0 To tpSelections.UcMarkets1.lbPlan.Items.Count - 1
            '        Dim exptop As String = "Tgroup = '" + plantg + "' and Mgroup = '" + tpSelections.UcMarkets1.lbPlan.Items(index).ToString() + "' and Channel = '" + row1("Channel").ToString() + "'"

            '        Dim drr As DataRow() = planDs.Tables(1).[Select](exptop, sortOrder)

            '        If drr Is Nothing Or drr.Count = 0 Then
            '            dr(index + 2) = 0
            '        Else
            '            dr(index + 2) = drr(0).Item("GRP")

            '        End If

            '    Next
            '    'dr(2) = row1(4)
            '    '  foundRefCopy()


            '    dt1.Rows.Add(dr)
            'Next
            If Not CheckSheetExists("Genre Share") Then
                nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                nativeSheet.Name = "Genre Share"
            Else

                If fromPane Then
                    nativeSheet = ReturnActualSheet("Genre Share")
                    '  newSheet.UsedRange.Clear()
                    Globals.Ribbons.MSprintExRibbon.CleanSheet(nativeSheet)
                    nativeSheet.Activate()
                Else
                    Dim sheetcount As Integer = CheckAndReturnSheet("Genre Share")
                    If sheetcount > 0 Then
                        nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                        Dim sname As String = String.Format("Genre Share({0})", sheetcount)
                        nativeSheet.Name = sname
                    End If
                End If

            End If
            'nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            'nativeSheet.UsedRange.Clear()
            'nativeSheet.Name = "Genre Share"

            Dim cell As Microsoft.Office.Interop.Excel.Range = nativeSheet.Range("$A$1", Type.Missing)
            cell.Value2 = "Genre Share"
            cell.ColumnWidth = 15
            cell.Interior.Color = System.Drawing.Color.Yellow

            '  nativeSheet.PageSetup.CenterFooter = "Genre Share and Channel share are calculated based on Program TVR"
            'Dim cell1 As Microsoft.Office.Interop.Excel.Range = nativeSheet.UsedRange


            Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(nativeSheet)
            ' vstoWorkbook.UsedRange.Clear()
            Dim cell1 As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(3, 2), vstoWorkbook.Cells(3, 2 + genres.Columns.Count - 1))
            cell1.Merge(True)
            cell1.Value2 = plantg



            Dim cell2 As Microsoft.Office.Interop.Excel.Range = cell1.Offset(1, 0)
            Dim name As String = "GenreShare" + Date.Now.Minute.ToString() + Date.Now.Second.ToString() + Date.Now.Millisecond.ToString()
            listObject = vstoWorkbook.Controls.AddListObject(cell2, name)
            listObject.AutoSetDataBoundColumnHeaders = True
            ' vstoWorkbook.Controls.AddListObject(
            ' listObject.QueryTable.AdjustColumnWidth = True
            'listObject.Range.Columns
            '  listObject.Range.Columns.AutoFit()

            '  If reftg.Length > 0 Then
            Try

                If viewVal.Equals(MSprintExRibbon.GenreShareView.All) Then
                    listObject.DataSource = genres
                ElseIf viewVal.Equals(MSprintExRibbon.GenreShareView.TopTen) Then
                    listObject.DataSource = genres.AsEnumerable().Take(10).CopyToDataTable()
                End If


            Catch ex As Exception
                LogMpsrintExException("Exception occured while displaying Genre share details on sheet" + ex.Message)
            End Try
            'Else
            'listObject.DataSource = genres
            'End If



            Dim cell3 As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.UsedRange.Next(1, 2)
            Dim chartObjects As ChartObjects = vstoWorkbook.ChartObjects()




            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                Try
                    Dim chart As ChartObject = chartObjects.Add((listObject.ListColumns.Count * 95 * (index + 1)) + (10 * index), listObject.Range.Top, 250, 250)
                    chart.Chart.ChartType = XlChartType.xl3DPie
                    Dim seriesCollection1 As Microsoft.Office.Interop.Excel.SeriesCollection = chart.Chart.SeriesCollection()
                    Dim series11 As Microsoft.Office.Interop.Excel.Series = seriesCollection1.NewSeries()
                    series11.XValues = listObject.ListColumns(1).Range
                    series11.Values = listObject.ListColumns(index + 2).Range
                    series11.Name = listObject.ListColumns(index + 2).Name
                Catch ex As Exception
                    LogMpsrintExException("Exception occured while displaying Genre share details on sheet" + ex.Message)
                End Try
            Next


            'chart.Chart.
            '   chart.Chart.SetSourceData(listObject.Range, Type.Missing)

            Dim celll As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(6 + listObject.ListRows.Count, 1), vstoWorkbook.Cells(6 + listObject.ListRows.Count, 1)), Microsoft.Office.Interop.Excel.Range)
            celll.Value2 = "Top Ten Channels across Genres"
            celll.ColumnWidth = 30
            celll.Interior.Color = System.Drawing.Color.Yellow


            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                Try
                    Dim expression As String = "Tgroup = '" + EscapeLikeValue(plantg) + "' and Mgroup = '" + EscapeLikeValue(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString()) + "'"
                    Dim sOrder As String = "RANK ASC"
                    Dim topTenTable = ds.Tables(1).[Select](expression, sOrder).CopyToDataTable()
                    topTenTable.Columns.RemoveAt(0)
                    topTenTable.Columns.RemoveAt(0)
                    topTenTable.Columns.RemoveAt(0)
                    Dim marketcell As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(7 + listObject.ListRows.Count, 1 + (index * 3)), vstoWorkbook.Cells(7 + listObject.ListRows.Count, 3 + (index * 3))), Microsoft.Office.Interop.Excel.Range)
                    marketcell.Merge(True)
                    marketcell.Value2 = Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString()
                    Dim cell5 As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(8 + listObject.ListRows.Count, 1 + (index * 3)), vstoWorkbook.Cells(8 + listObject.ListRows.Count, 3 + (index * 3))), Microsoft.Office.Interop.Excel.Range)
                    listobject1 = vstoWorkbook.Controls.AddListObject(cell5, "Top10AcrossGenres" + Date.Now.Minute.ToString() + Date.Now.Second.ToString() + Date.Now.Millisecond.ToString())
                    listobject1.AutoSetDataBoundColumnHeaders = True
                    ' listobject1.Range.Columns.AutoFit()
                    listobject1.DataSource = topTenTable
                Catch ex As Exception
                    LogMpsrintExException("Exception occured while displaying Genre share details on sheet" + ex.Message)
                End Try
            Next
        Catch ex As Exception
            LogMpsrintExException("Exception occured while displaying Genre share details on sheet")
            Throw ex
        End Try
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
End Module
