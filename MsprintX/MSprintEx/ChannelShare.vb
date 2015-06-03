Imports Microsoft.Office.Tools
Imports Microsoft.Office.Interop.Excel
Imports System
Imports System.Data
Module ChannelShare
    Friend listObject, listobject1 As Excel.ListObject
    Friend nativeSheet, newSheet, vstoWorkbook As Microsoft.Office.Interop.Excel.Worksheet
    Public Function ConstructChannelShareInputXML(ByVal plantgname As String, ByVal reftgname As String)
        Try


            Dim month, month1 As String
            Dim day, day1 As String
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

            Dim input As XElement =
                <input>
                    <pre-eval-period>
                        <startdate><%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.Year.ToString() + month + day %></startdate>
                        <enddate><%= Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.Year.ToString() + month1 + day1 %></enddate>
                    </pre-eval-period>
                </input>
            ' Dim tgs As XElement = New XElement("targetgroups")
            'Dim doc As XmlDocument = New XmlDocument()
            '  doc.Load()
            '  tgs.Add(XElement.Load(Path.GetTempPath() + "\\TGS\\" + plantgname + ".xml"))

            'If reftgname <> plantgname Then
            '    '  Dim doc1 As XmlDocument = New XmlDocument()
            '    '  doc1.Load()
            '    tgs.Add(XElement.Load(Path.GetTempPath() + "\\TGS\\" + reftgname + ".xml"))
            'End If

            'Dim markets As XElement = New XElement("markets")

            If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.LbDayParts.Items.Count > 0 Then
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
            End If


            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                ' Dim doc2 As XmlDocument = New XmlDocument()
                '  doc2.Load()
                Dim TG_MGElement As XElement =
                  <TG_MG name=<%= String.Format("{0}~{1}", plantgname, Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim()) %> type="Planning">
                  </TG_MG>
                TG_MGElement.Add(XElement.Load(Globals.Ribbons.MSprintExRibbon.tgDirectoryPath + plantgname + ".xml"))
                TG_MGElement.Add(XElement.Load(Globals.Ribbons.MSprintExRibbon.mgDirectoryPath + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim() + ".xml"))

                For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items.Count - 1
                    Dim genre As XElement =
                    <genre name=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index1).ToString() %>/>
                    TG_MGElement.Add(genre)
                Next
                input.Add(TG_MGElement)
                '  markets.Add(XElement.Load(Path.GetTempPath() + "\\MGS\\" + tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim() + ".xml"))
            Next
            '

            'If reftgname.Length > 0 Then
            '    For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items.Count - 1
            '        Dim TG_MGElement As XElement =
            '         <TG_MG name=<%= String.Format("{0}~{1}", reftgname, Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index).ToString().Trim()) %> type="Reference">
            '         </TG_MG>
            '        TG_MGElement.Add(XElement.Load(Globals.Ribbons.MSprintExRibbon.tgDirectoryPath + reftgname + ".xml"))
            '        TG_MGElement.Add(XElement.Load(Globals.Ribbons.MSprintExRibbon.mgDirectoryPath + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index).ToString().Trim() + ".xml"))

            '        For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items.Count - 1
            '            Dim genre As XElement =
            '            <genre name=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index1).ToString() %>/>
            '            TG_MGElement.Add(genre)
            '        Next
            '        input.Add(TG_MGElement)
            '    Next
            'End If
            Globals.Ribbons.MSprintExRibbon.UpdateUsageReport("Channel ShareWS", Globals.Ribbons.MSprintExRibbon.tpSelections.tbClientName.Text.Trim, Globals.Ribbons.MSprintExRibbon.tpSelections.tbBrandValue.Text.Trim, input, 0)

            Return input
        Catch ex As Exception
            LogMpsrintExException("Exception occured while constructing channel share input xml.Message :" + ex.Message)
        End Try
    End Function
    Public Function DisplayChannelShareDetailsonSheet(ByVal ds As DataSet, Optional ByVal fromPane As Boolean = False)
        Dim plantgname As String = String.Empty
        Dim reftg As String = String.Empty
        Try
            Dim dtable As System.Data.DataTable = DirectCast(Globals.Ribbons.MSprintExRibbon.tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            plantgname = dtable.Rows(0)(1).ToString().Trim()
            ' reftg = dtable.Rows(1)(1).ToString()
            Dim genreTables As System.Data.DataSet = New System.Data.DataSet()
            'Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items.Count - 1) {}
            'dt1 = New System.Data.DataTable()
            'dt1.Columns.Add("Rank", System.Type.GetType("System.Int32"))
            'dt1.Columns.Add("Channel")
            'dt1.Columns.Add("Programme")
            'dt1.Columns.Add("Start Hour")
            'dt1.Columns.Add("Plan " + plantgname + "~" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(0).ToString() + " GRP")
            'dt1.Columns.Add("Ref " + reftgname + "~" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(0).ToString() + " GRP")


            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items.Count - 1
                genreTables.Tables.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index))
                genreTables.Tables(index).Columns.Add("Channel")
                For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                    'Dim expression1 As String = "Tgroup = '" + plantgname + "' and Mgroup = '" + tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "' and Genre= '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index).ToString()
                    'Dim sOrder As String = "GRP DESC"
                    'Dim rows As Data.DataRow() = planDs.Tables(0).[Select](expression1, sOrder)

                    genreTables.Tables(index).Columns.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRP(in %)", System.Type.GetType("System.Decimal"))
                    '  genreTables(index).Rows.Add(rows(0).Item("Channel Name").ToString(),
                    '  genreTables(index).Rows(index)(.Columns.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index) + " ~ " + " Ref " + reftgname + "-" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(0).ToString() + " GRP")
                Next
            Next
            'For Each row As Data.DataRow In cGenre.Rows
            '    For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
            '        Dim expression1 As String = "Tgroup = '" + plantgname + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "' and Genre= '" + row("Genre").ToString() + "'"
            '        Dim sOrder As String = "GRP DESC"
            '        Dim rows As Data.DataRow() = planDs.Tables(0).[Select](expression1, sOrder)
            '        genreTables.Tables(row("Genre").ToString()).Rows.Add(row("Channel").ToString(), rows(0).Item("GRP"))
            '    Next
            'Next
            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items.Count - 1

                For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1

                    Dim sortOrder As String = "GRP DESC"
                    Dim expression1 As String = "Tgroup = '" + EscapeLikeValue(plantgname) + "' and Mgroup = '" + EscapeLikeValue(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString()) + "'  and Genre = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index) + "'"
                    Dim foundRows As Data.DataRow() = ds.Tables(0).[Select](expression1, sortOrder)

                    For Each dRow As DataRow In foundRows

                        If index1 = 0 Then
                            Dim dr As DataRow = genreTables.Tables(index).NewRow()
                            dr("Channel") = dRow("Channel Name")

                            dr(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRP(in %)") = dRow("GRP")
                            ' dr(2) = dsRef.Tables(0).[Select](expcopy, sortOrder)(0).Item("GRP")
                            genreTables.Tables(index).Rows.Add(dr)
                        Else
                            For Each rrow As DataRow In genreTables.Tables(index).Rows

                                If rrow("Channel") = dRow("Channel Name") Then
                                    rrow(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRP(in %)") = dRow("GRP")
                                End If
                            Next
                        End If


                        'If genreTables.Tables(index).Rows.Count > 1 Then

                        '    For Each rrow As DataRow In genreTables.Tables(index).Rows

                        '        If rrow("Channel") = dRow("Channel Name") Then
                        '            rrow(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRP") = dRow("GRP")
                        '        Else
                        '            Dim dr As DataRow = genreTables.Tables(index).NewRow()
                        '            dr("Channel") = dRow("Channel Name")

                        '            dr(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRP") = dRow("GRP")
                        '            ' dr(2) = dsRef.Tables(0).[Select](expcopy, sortOrder)(0).Item("GRP")
                        '            genreTables.Tables(index).Rows.Add(dr)
                        '        End If

                        '    Next
                        'Else
                        '    Dim dr As DataRow = genreTables.Tables(index).NewRow()
                        '    dr("Channel") = dRow("Channel Name")

                        '    dr(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRP") = dRow("GRP")
                        '    ' dr(2) = dsRef.Tables(0).[Select](expcopy, sortOrder)(0).Item("GRP")
                        '    genreTables.Tables(index).Rows.Add(dr)
                        'End If
                        'Dim expcopy As String = "Tgroup = '" + reftgname + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(0).ToString() + "' and Genre = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index) + "' and [Channel Name] = '" + dRow("Channel Name") + "'"

                    Next
                Next
            Next
            ' Dim expression As String = "OrderQuantity = 2 and OrderID = 2" 
            ' Sort descending by column named CompanyName. 
            'Dim expression As String = "Tgroup = '" + plantgname + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(0).ToString() + "'"

            'Dim sortOrder As String = "GRP DESC"
            'Dim foundRows, foundtopten As DataRow()
            '' Dim exptopten As String = "TG"
            '' Use the Select method to find all rows matching the filter.
            'foundtopten = ds.Tables(1).[Select](expression, sortOrder)
            ' foundRefCopy = copyGenreTab.Select(expcopy, sortOrder)
            'For Each row As DataRow In foundRows
            '    Dim expcopy As String = "Tgroup = '" + ComboBox2.Text.Trim() + "' and Mgroup = '" + CheckedListBox2.CheckedItems(0).ToString() + "' and Genre = '" + row("Genre").ToString() + "' and Channel Name = '"+row(
            '    Dim dr As DataRow = dt.NewRow()
            '    dr("Genre") = row("Genre")
            '    dr(1) = row("GRP")
            '    '  foundRefCopy()
            '    dr(2) = copyGenreTab.[Select](expcopy, sortOrder)(0).Item("GRP")
            '    dt.Rows.Add(dr)
            'Next

            'For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items.Count - 1
            '    Dim expression1 As String = "Tgroup = '" + plantgname + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(0).ToString() + "'  and Genre = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index) + "'"
            '    foundRows = ds.Tables(0).[Select](expression1, sortOrder)
            '    For Each dRow As DataRow In foundRows
            '        Dim expcopy As String = "Tgroup = '" + reftgname + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(0).ToString() + "' and Genre = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index) + "' and [Channel Name] = '" + dRow("Channel Name") + "'"
            '        Dim dr As DataRow = genreTables(index).NewRow()
            '        dr("Channel") = dRow("Channel Name")
            '        dr(1) = dRow("GRP")
            '        dr(2) = dsRef.Tables(0).[Select](expcopy, sortOrder)(0).Item("GRP")
            '        genreTables(index).Rows.Add(dr)
            '    Next

            'Next

            ' foundRows.CopyToDataTable(,LoadOption.
            'For Each row1 As DataRow In foundtopten
            '    Dim exptop As String = "Tgroup = '" + reftgname + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(0).ToString() + "' and Channel = '" + row1("Channel").ToString() + "' and [Program Name] = '" + row1("Program Name") + "' and [Program Start Time] = '" + row1("Program Start Time") + "'"
            '    Dim dr As DataRow = dt1.NewRow()
            '    dr("Rank") = row1("Rank")
            '    'Dim i As Integer = Convert.ToInt32(row1("ChannelCode").ToString())
            '    'Dim ss As String = String.Empty
            '    'If i < 10 Then
            '    '    ss = "00" + i.ToString()
            '    'ElseIf i < 100 Then
            '    '    ss = "0" + i.ToString()
            '    'Else
            '    '    ss = row1("ChannelCode").ToString()
            '    'End If
            '    'Dim drow As DataRow = dg.GetChannels().Select("TAM_CHANNEL_CODE = '" + ss + "'")(0)
            '    dr("Channel") = row1("Channel").ToString()
            '    dr("Programme") = row1("Program Name").ToString()
            '    dr("Start Hour") = row1("Program Start Time")

            '    dr(4) = row1("GRP")
            '    '  foundRefCopy()
            '    Dim drr As DataRow() = dsRef.Tables(1).[Select](exptop, sortOrder)

            '    If drr Is Nothing Or drr.Count = 0 Then
            '        dr(5) = 0
            '    Else
            '        dr(5) = drr(0).Item("GRP")

            '    End If

            '    dt1.Rows.Add(dr)
            'Next
            'Dim sheets As Microsoft.Office.Interop.Excel.Worksheets = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets
            'sheets.
            'If Not  Then

            'End If

            If Not CheckSheetExists("Channel Share") Then
                newSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                newSheet.Name = "Channel Share"
            Else

                If fromPane Then
                    newSheet = ReturnActualSheet("Channel Share")
                    'newSheet.UsedRange.Clear()
                    Globals.Ribbons.MSprintExRibbon.CleanSheet(newSheet)
                    newSheet.Activate()
                Else
                    Dim sheetcount As Integer = CheckAndReturnSheet("Channel Share")
                    If sheetcount > 0 Then
                        newSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                        Dim name As String = String.Format("Channel Share({0})", sheetcount)
                        newSheet.Name = name
                    End If
                End If

               
              

            End If



            Dim cell As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$A$1", Type.Missing)
            cell.Value2 = "Channel Share"
            cell.Interior.Color = System.Drawing.Color.Yellow
            cell.ColumnWidth = 15
            Dim cell1 As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$C$1", "$C$1")
            cell1.Value2 = "Eval Start Date"
            Dim cell2 As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$D$1", "$D$1")
            cell2.Value2 = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value
            cell2.NumberFormat = "dd/mm/yyyy"
            Dim weekcell1 As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$E$1", "$E$1")
            weekcell1.Value2 = String.Format("Week {0}", Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks.Rows(0)(0).ToString())
            Dim evalendcell As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$C$2", "$C$2")
            evalendcell.Value2 = "Eval End Date"
            Dim cell23 As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$D$2", "$D$2")
            cell23.Value2 = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value
            cell23.NumberFormat = "dd/mm/yyyy"
            Dim weekcell11 As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$E$2", "$E$2")
            weekcell11.Value2 = String.Format("Week {0}", Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks.Rows(0)(0).ToString())
            Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(newSheet)
            ' Dim listObject As Excel.ListObject()
            Dim ii As Integer = 1 + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count
            Dim rocount As Integer = 0
            For index = 0 To genreTables.Tables.Count - 1
                '4,1 - 4,4 ;4,6 - 4,9;4,11- 4,14
                Try

                    Dim genrecell As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(3, 1 + ((ii + 1) * index)), vstoWorkbook.Cells(3, 1 + ((ii + 1) * index))), Microsoft.Office.Interop.Excel.Range)
                    genrecell.Merge(True)
                    genrecell.Value2 = genreTables.Tables(index).TableName
                    '  Dim lrange As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(4, 1 + ((ii + 1) * index)), vstoWorkbook.Cells(4, 1 + (ii * index) + ii)), Microsoft.Office.Interop.Excel.Range)
                    Dim lrange As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(4, 1 + ((ii + 1) * index)), vstoWorkbook.Cells(4, 1 + (ii + 1) * index)), Microsoft.Office.Interop.Excel.Range)
                    Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "cshare" + index.ToString() + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString())
                    listobject.AutoSetDataBoundColumnHeaders = True
                    '  listobject.ListColumns(0).

                    'If reftg.Length > 0 Then
                    listobject.DataSource = genreTables.Tables(index).AsEnumerable().Take(10).CopyToDataTable()
                    'Else
                    'listobject.DataSource = genreTables.Tables(index)
                    'End If

                    '  ii += 5

                    If listobject.ListRows.Count > rocount Then
                        rocount = listobject.ListRows.Count
                    End If
                Catch ex As Exception
                    LogMpsrintExException("Exception occured while displaying channel share details on sheet" + ex.Message)
                    ' ii += 5

                    'If listObject.ListRows.Count > rocount Then
                    '    rocount = listObject.ListRows.Count
                    'End If
                End Try

            Next

            ' Dim row As Integer = Globals.ThisAddIn.Application.ActiveCell.Row + 2
            Dim celll As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(7 + rocount, 1), vstoWorkbook.Cells(7 + rocount, 1)), Microsoft.Office.Interop.Excel.Range)
            celll.Value2 = "Top Ten Programs for period"
            celll.ColumnWidth = 30
            celll.Interior.Color = System.Drawing.Color.Yellow

            ' Dim row11 As Integer = Globals.ThisAddIn.Application.ActiveCell.Row + 2

            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                Try

                    Dim expression As String = "Tgroup = '" + EscapeLikeValue(plantgname) + "' and Mgroup = '" + EscapeLikeValue(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString()) + "'"
                    Dim sortOrder As String = "Rank ASC"
                    Dim foundtopten As DataRow()
                    ' Dim exptopten As String = "TG"
                    ' Use the Select method to find all rows matching the filter.
                    foundtopten = ds.Tables(1).[Select](expression, sortOrder)
                    Dim top10 As Data.DataTable = foundtopten.CopyToDataTable()
                    top10.Columns.RemoveAt(0)
                    top10.Columns.RemoveAt(0)
                    top10.Columns.RemoveAt(0)
                    Dim marketcell As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(9 + rocount, 1 + (index * 5)), vstoWorkbook.Cells(9 + rocount, 1 + (index * 5))), Microsoft.Office.Interop.Excel.Range)
                    marketcell.Merge(True)
                    marketcell.Value2 = Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index)
                    Dim cell5 As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(11 + rocount, 1 + (index * 5)), vstoWorkbook.Cells(11 + rocount, 5 + (index * 5))), Microsoft.Office.Interop.Excel.Range)
                    listobject1 = vstoWorkbook.Controls.AddListObject(cell5, "csharetop10" + Date.Now.Minute.ToString() + Date.Now.Second.ToString() + Date.Now.Millisecond.ToString())
                    listobject1.AutoSetDataBoundColumnHeaders = True
                    ' listobject1.Range.Columns.AutoFit()
                    listobject1.DataSource = top10
                    listobject1.ListColumns(4).DataBodyRange.NumberFormat = "0.00"
                Catch ex As Exception
                    LogMpsrintExException("Exception occured while displaying channel share details on sheet" + ex.Message)
                End Try
            Next
        Catch ex As Exception
            LogMpsrintExException("Exception occured while displaying channel share details." + ex.Message)
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
