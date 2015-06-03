Imports System.Data
Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Tools
Imports Microsoft.Office.Interop.Excel
Public Class ucChannelShare
    Dim dSet, dRefSet As DataSet
    Dim planmgtg As String()
    Dim refmgtg As String()
    Dim ptgname As String
    Dim rtgname, callerVal As String
    Dim cGenre, channels As Data.DataTable
    ' Dim Globals.Ribbons.MSprintExRibbon.tpSelections As ucPlanSelections
    Friend listObject, listobject1 As Excel.ListObject
    Friend nativeSheet, newSheet, vstoWorkbook As Microsoft.Office.Interop.Excel.Worksheet
    Friend dg As New METISTableAdapters.CHANNEL_MASTERTableAdapter
    Public Sub New(ByVal ds As DataSet, ByVal dsReference As DataSet, ByVal plan As String(), ByVal ref As String(), ByVal plantgname As String, ByVal reftgname As String, ByVal tps As ucPlanSelections, ByVal caller As String, ByVal cgen As Data.DataTable, ByVal cnnels As Data.DataTable)
        InitializeComponent()
        dSet = ds
        dRefSet = dsReference
        planmgtg = plan
        refmgtg = ref
        ptgname = plantgname
        rtgname = reftgname
        Globals.Ribbons.MSprintExRibbon.tpSelections = tps
        cGenre = cgen
        channels = cnnels
        callerVal = caller
        '  purposeVal = purpose
    End Sub
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' dSet = 
        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub View_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click

        '   Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlWait
        'dSet = Globals.Ribbons.MSprintExRibbon.planningDataSet
        'dRefSet = Globals.Ribbons.MSprintExRibbon.refDataSet

        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        Dim text As String = String.Empty
        If sheet.Name.Equals("Genre Share") Then
            dSet = Globals.Ribbons.MSprintExRibbon.gshareds
            dRefSet = Globals.Ribbons.MSprintExRibbon.gsharerefds
            DisplayGenreShare()
        ElseIf sheet.Name.Equals("Channel Share") Then
            dSet = Globals.Ribbons.MSprintExRibbon.cshareds
            dRefSet = Globals.Ribbons.MSprintExRibbon.csharerefds
            DisplayChannelShare()
        ElseIf sheet.Name.Equals("Program TVR") Then
            dSet = Globals.Ribbons.MSprintExRibbon.ptvrds
            dRefSet = Globals.Ribbons.MSprintExRibbon.ptvrrefds
            DisplayProgramTVR()
        ElseIf sheet.Name.Equals("Break TVR") Then
            dSet = Globals.Ribbons.MSprintExRibbon.btvrds
            dRefSet = Globals.Ribbons.MSprintExRibbon.btvrrefds
            DisplayBreakTVR()
        End If
        'If callerVal.Equals("Genre") Then

        'ElseIf callerVal.Equals("Channel") Then

        'Else

        'End If


        '    nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        '    nativeSheet.UsedRange.Clear()
        '    Dim dt As System.Data.DataTable = New System.Data.DataTable()
        '    Dim dt1 As System.Data.DataTable = New System.Data.DataTable()
        '    '   Dim copyGenreTab As System.Data.DataTable = dSet.Tables(0).Copy()

        '    ' Dim copyto10 As System.Data.DataTable = dSet.Tables(1).Copy()
        '    dt1.Columns.Add("Rank", System.Type.GetType("System.Int32"))
        '    dt1.Columns.Add("Channel")
        '    ' dt1.Columns.Add("Plan " + cbPlan.Text.Trim() + "-" + cbpmgs.Text.Trim() + " GRP")
        '    ' dt1.Columns.Add("Ref " + CbRef.Text.Trim() + "-" + cbrefmgs.Text.Trim() + " GRP")
        '    dt.Columns.Add("Genre")
        '    ' dt.Columns.Add("Plan " + cbPlan.Text.Trim() + "-" + cbpmgs.Text.Trim() + " GRP")
        '    ' dt.Columns.Add("Ref" + CbRef.Text.Trim() + "-" + cbrefmgs.Text.Trim() + " GRP")
        '    '  Dim expression As String = "Tgroup = '" + cbPlan.Text.Trim() + "' and Mgroup = '" + cbpmgs.Text.Trim() + "'"
        '    ' Dim expression As String = "OrderQuantity = 2 and OrderID = 2" 
        '    ' Sort descending by column named CompanyName. 
        '    Dim sortOrder As String = "GRP DESC"
        '    Dim foundRows, foundtopten As DataRow()
        '    ' Dim exptopten As String = "TG"
        '    ' Use the Select method to find all rows matching the filter.
        '    'foundRows = dSet.Tables(0).[Select](expression, sortOrder)
        '    'foundtopten = dSet.Tables(1).[Select](expression, sortOrder)
        '    ' foundRefCopy = copyGenreTab.Select(expcopy, sortOrder)
        '    'For Each row As DataRow In foundRows
        '    '    Dim expcopy As String = "Tgroup = '" + CbRef.Text.Trim() + "' and Mgroup = '" + cbrefmgs.Text.Trim() + "' and Genre = '" + row("Genre").ToString() + "'"
        '    '    Dim dr As DataRow = dt.NewRow()
        '    '    dr("Genre") = row("Genre")
        '    '    dr(1) = row("GRP")
        '    '    '  foundRefCopy()
        '    '    Dim dRRow As DataRow() = dRefSet.Tables(0).[Select](expcopy, sortOrder)
        '    '    If dRRow Is Nothing Or dRRow.Count = 0 Then
        '    '        dr(2) = 0
        '    '    Else
        '    '        dr(2) = dRRow(0).Item("GRP")

        '    '    End If
        '    '    ' dr(2) = (0).Item("GRP")
        '    '    dt.Rows.Add(dr)
        '    'Next
        '    'For Each row1 As DataRow In foundtopten
        '    '    Dim exptop As String = "Tgroup = '" + CbRef.Text.Trim() + "' and Mgroup = '" + cbrefmgs.Text.Trim() + "' and Channel = '" + row1("Channel").ToString() + "'"
        '    '    Dim dr As DataRow = dt1.NewRow()
        '    '    dr("Rank") = row1("Rank")
        '    '    'Dim i As Integer = Convert.ToInt32(row1("Channel").ToString())
        '    '    'Dim ss As String = String.Empty
        '    '    'If i < 10 Then
        '    '    '    ss = "00" + i.ToString()
        '    '    'ElseIf i < 100 Then
        '    '    '    ss = "0" + i.ToString()
        '    '    'Else
        '    '    '    ss = row1("Channel").ToString()
        '    '    'End If
        '    '    'Dim drow As DataRow = dg.GetChannels().Select("TAM_CHANNEL_CODE = '" + ss + "'")(0)
        '    '    dr("Channel") = row1("Channel").ToString()
        '    '    dr(2) = row1(4)
        '    '    '  foundRefCopy()
        '    '    Dim drr As DataRow() = dRefSet.Tables(1).[Select](exptop, sortOrder)

        '    '    If drr Is Nothing Or drr.Count = 0 Then
        '    '        dr(3) = 0
        '    '    Else
        '    '        dr(3) = drr(0).Item("GRP")

        '    '    End If

        '    '    dt1.Rows.Add(dr)
        '    'Next
        '    Dim cell As Microsoft.Office.Interop.Excel.Range = nativeSheet.Range("$A$1", Type.Missing)
        '    cell.Value2 = "Genre Share"
        '    cell.ColumnWidth = 15
        '    cell.Interior.Color = System.Drawing.Color.Yellow
        '    nativeSheet.Name = "Genre Share"
        '    '  nativeSheet.PageSetup.CenterFooter = "Genre Share and Channel share are calculated based on Program TVR"
        '    Dim cell1 As Microsoft.Office.Interop.Excel.Range = nativeSheet.UsedRange
        '    Dim cell2 As Microsoft.Office.Interop.Excel.Range = cell1.Next(3, 0)

        '    Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(nativeSheet)
        '    listObject = vstoWorkbook.Controls.AddListObject(cell2, "listG" + Date.Now.Hour.ToString() + Date.Now.Minute.ToString() + Date.Now.Millisecond.ToString())
        '    listObject.AutoSetDataBoundColumnHeaders = True
        '    ' listObject.QueryTable.AdjustColumnWidth = True
        '    'listObject.Range.Columns
        '    '  listObject.Range.Columns.AutoFit()
        '    listObject.DataSource = dt
        '    'dt.Columns.RemoveAt(1)
        '    'dt.Columns.RemoveAt(1)
        '    'dt.AcceptChanges()
        '    'dt.WriteXml(IO.Path.GetTempPath() + "\\Genres.xml")
        '    Dim cell3 As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.UsedRange.Next(1, 2)
        '    Dim chartObjects As ChartObjects = vstoWorkbook.ChartObjects()
        '    chartObjects.Delete()
        '    Dim chart As ChartObject = chartObjects.Add(listObject.ListColumns.Count * 85, listObject.Range.Top, 250, 250)
        '    chart.Chart.ChartType = XlChartType.xl3DPie
        '    'chart.Chart.
        '    chart.Chart.SetSourceData(listObject.Range, Type.Missing)

        '    Dim celll As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(6 + listObject.ListRows.Count, 1), vstoWorkbook.Cells(6 + listObject.ListRows.Count, 1)), Microsoft.Office.Interop.Excel.Range)
        '    celll.Value2 = "Top Ten Channels across Genres"
        '    celll.ColumnWidth = 30
        '    celll.Interior.Color = System.Drawing.Color.Yellow

        '    Dim cell5 As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(7 + listObject.ListRows.Count, 1), vstoWorkbook.Cells(7 + listObject.ListRows.Count, 1)), Microsoft.Office.Interop.Excel.Range)
        '    listobject1 = vstoWorkbook.Controls.AddListObject(cell5, "listTT" + Date.Now.Hour.ToString() + Date.Now.Minute.ToString() + Date.Now.Millisecond.ToString())
        '    listobject1.AutoSetDataBoundColumnHeaders = True
        '    ' listobject1.Range.Columns.AutoFit()
        '    listobject1.DataSource = dt1
        '    ' listobject1.QueryTable.AdjustColumnWidth = True
        'End Sub

        'Private Sub ucChannelShare_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '    '  For Each Str As String In planmgtg
        '    'cbpmgs.Items.AddRange(planmgtg)
        '    'cbpmgs.Text = planmgtg(0)
        '    'cbrefmgs.Items.AddRange(refmgtg)
        '    'cbrefmgs.Text = refmgtg(0)
        '    'cbPlan.Items.Add(ptgname)
        '    'cbPlan.Text = ptgname
        '    'CbRef.Items.Add(rtgname)
        '    'CbRef.Text = rtgname
        '    ComboBox1.Items.Add(ptgname)
        '    ComboBox1.Items.Add(rtgname)
        '    ComboBox1.Text = ptgname
        '    lbmgs.Text = String.Empty
        '    For Each Str As String In planmgtg
        '        lbmgs.Text = lbmgs.Text + Environment.NewLine() + Str
        '    Next

        ' Next
        Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault
    End Sub

    Private Sub ucChannelShare_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim dtable As System.Data.DataTable = DirectCast(Globals.Ribbons.MSprintExRibbon.tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)

            'plantgname = dtable.Rows(0)(1).ToString().Trim()
            'reftgname = dtable.Rows(1)(1).ToString()
            ComboBox1.Text = dtable.Rows(0)(1).ToString().Trim()
            ComboBox1.Items.Add(dtable.Rows(0)(1).ToString().Trim())
            ptgname = dtable.Rows(0)(1).ToString().Trim()

            Try
                ComboBox1.Items.Add(dtable.Rows(1)(1).ToString())
                rtgname = dtable.Rows(1)(1).ToString()
            Catch ex As Exception

            End Try


            '  NumericUpDown1.Value = dSet.Tables(0).Rows.Count
            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                lbmgs.Text = lbmgs.Text + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index)

                If index <> Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1 Then
                    lbmgs.Text = lbmgs.Text + ","
                End If


            Next
        Catch ex As Exception

        End Try

    End Sub
    Public Function DisplayGenreShare()
        Try

            ' Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlWait
            If ComboBox1.Text = ptgname Then
                'Dim genres As Data.DataTable = dSet.Tables(0).DefaultView.ToTable(True, "Genre")

                ''Dim dt1 As System.Data.DataTable = New System.Data.DataTable()
                ' ''  Dim copyGenreTab As System.Data.DataTable = ds.Tables(0).Copy()

                ' ''  Dim copyto10 As System.Data.DataTable = ds.Tables(1).Copy()
                ''dt1.Columns.Add("Rank", System.Type.GetType("System.Int32"))
                ''dt1.Columns.Add("Channel")
                ' '' dt1.Columns.Add(plantg + "-" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(0).ToString() + " GRP")
                ' ''  dt1.Columns.Add(reftg + "-" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(0).ToString() + " GRP")
                ''dt.Columns.Add("Genre")
                ''  dt.Columns.Add("Plan " + plantg + "-" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(0).ToString() + " GRP")
                ''  dt.Columns.Add("Ref" + reftg + "-" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(0).ToString() + " GRP")
                '' Dim expression As String = "OrderQuantity = 2 and OrderID = 2" 
                '' Sort descending by column named CompanyName. 
                'Dim sortOrder As String = "GRP DESC"
                'Dim foundRows As DataRow()
                '' Dim exptopten As String = "TG"
                '' Use the Select method to find all rows matching the filter.
                '' Dim dtables As System.Data.DataTable() = New System.Data.DataTable(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count) {}
                '' foundtopten = planDs.Tables(1).[Select](expression, sortOrder)

                'For index = 0 To genres.Rows.Count - 1

                '    For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                '        Dim expression1 As String = "Tgroup = '" + ptgname + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "' and Genre = '" + genres.Rows(index)(0).ToString() + "'"
                '        foundRows = dSet.Tables(0).[Select](expression1, sortOrder)

                '        If foundRows.Count > 0 Then
                '            ' dtables(index) = New System.Data.DataTable()
                '            'dtables(index) = foundRows.CopyToDataTable()

                '            If Not (genres.Columns.Contains(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRP")) Then

                '                genres.Columns.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRP")
                '            End If

                '            genres.Rows(index)(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRP") = foundRows(0).Item("GRP")
                '        Else
                '            genres.Rows(index)(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRP") = 0
                '            'dt1.Columns.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString() + "-GRP")
                '        End If

                '    Next
                'Next
                '' foundRefCopy = copyGenreTab.Select(expcopy, sortOrder)
                ''For Each row As DataRow In foundRows
                ''    Dim dr As DataRow = dt.NewRow()
                ''    dr("Genre") = row("Genre")
                ''    For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1


                ''        Dim expcopy As String = "Tgroup = '" + plantg + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString() + "' and Genre = '" + row("Genre").ToString() + "'"

                ''        Dim drr As DataRow() = planDs.Tables(0).[Select](expcopy, sortOrder)
                ''        If drr Is Nothing Or drr.Count = 0 Then
                ''            dr(index + 1) = 0
                ''        Else
                ''            dr(index + 1) = drr(0).Item("GRP")
                ''        End If
                ''        ' dr(index + 1) = ds.Tables(0).[Select](expcopy, sortOrder)(0).Item("GRP")
                ''        '  foundRefCopy()
                ''        ' dr(index + 2) = ds.Tables(0).[Select](expcopy, sortOrder)(0).Item("GRP")

                ''    Next
                ''    dt.Rows.Add(dr)
                ''Next
                ''For Each row1 As DataRow In foundtopten
                ''    Dim dr As DataRow = dt1.NewRow()
                ''    dr("Rank") = row1("Rank")
                ''    'Dim i As Integer = Convert.ToInt32(row1("Channel").ToString())
                ''    'Dim ss As String = String.Empty
                ''    'If i < 10 Then
                ''    '    ss = "00" + i.ToString()
                ''    'ElseIf i < 100 Then
                ''    '    ss = "0" + i.ToString()
                ''    'Else
                ''    '    ss = row1("Channel").ToString()
                ''    'End If
                ''    'Dim drow As DataRow = dg.GetChannels().Select("TAM_CHANNEL_CODE = '" + ss + "'")(0)
                ''    dr("Channel") = row1("Channel").ToString()
                ''    For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                ''        Dim exptop As String = "Tgroup = '" + plantg + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString() + "' and Channel = '" + row1("Channel").ToString() + "'"

                ''        Dim drr As DataRow() = planDs.Tables(1).[Select](exptop, sortOrder)

                ''        If drr Is Nothing Or drr.Count = 0 Then
                ''            dr(index + 2) = 0
                ''        Else
                ''            dr(index + 2) = drr(0).Item("GRP")

                ''        End If

                ''    Next
                ''    'dr(2) = row1(4)
                ''    '  foundRefCopy()


                ''    dt1.Rows.Add(dr)
                ''Next
                'nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                'nativeSheet.UsedRange.Clear()
                'Dim cell As Microsoft.Office.Interop.Excel.Range = nativeSheet.Range("$A$1", Type.Missing)
                'cell.Value2 = "Genre Share"
                'cell.ColumnWidth = 15
                'cell.Interior.Color = System.Drawing.Color.Yellow
                'nativeSheet.Name = "Genre Share"
                ''  nativeSheet.PageSetup.CenterFooter = "Genre Share and Channel share are calculated based on Program TVR"
                ''Dim cell1 As Microsoft.Office.Interop.Excel.Range = nativeSheet.UsedRange


                'Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(nativeSheet)
                'Dim cell1 As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(3, 2), vstoWorkbook.Cells(3, 2 + genres.Columns.Count - 1))
                'cell1.Merge(True)
                'cell1.Value2 = ptgname



                'Dim cell2 As Microsoft.Office.Interop.Excel.Range = cell1.Offset(1, -(genres.Columns.Count - 2))
                'listObject = vstoWorkbook.Controls.AddListObject(cell2, "GenreShareucp")
                'listObject.AutoSetDataBoundColumnHeaders = True
                '' listObject.QueryTable.AdjustColumnWidth = True
                ''listObject.Range.Columns
                ''  listObject.Range.Columns.AutoFit()
                'Try
                '    listObject.DataSource = genres.AsEnumerable().Take(NumericUpDown1.Value).CopyToDataTable()
                'Catch ex As Exception
                '    listObject.DataSource = genres.AsEnumerable().Take(genres.Rows.Count).CopyToDataTable()
                'End Try


                'Dim cell3 As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.UsedRange.Next(1, 2)
                'Dim chartObjects As ChartObjects = vstoWorkbook.ChartObjects()
                'chartObjects.Delete()
                'For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items.Count - 1
                '    Dim chart As ChartObject = chartObjects.Add((listObject.ListColumns.Count * 85 * (index + 1)) + (10 * index), listObject.Range.Top, 250, 250)
                '    chart.Chart.ChartType = XlChartType.xl3DPie
                '    Dim seriesCollection1 As Microsoft.Office.Interop.Excel.SeriesCollection = chart.Chart.SeriesCollection()
                '    Dim series11 As Microsoft.Office.Interop.Excel.Series = seriesCollection1.NewSeries()
                '    series11.XValues = listObject.ListColumns(1).Range
                '    series11.Values = listObject.ListColumns(index + 2).Range
                '    series11.Name = listObject.ListColumns(index + 2).Name
                'Next
                '' Dim chart As ChartObject = chartObjects.Add(listObject.ListColumns.Count * 85, listObject.Range.Top, 250, 250)
                '' chart.Chart.ChartType = XlChartType.xl3DPie
                ''chart.Chart.
                ''chart.Chart.SetSourceData(listObject.Range, Type.Missing)

                'Dim celll As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(6 + listObject.ListRows.Count, 1), vstoWorkbook.Cells(6 + listObject.ListRows.Count, 1)), Microsoft.Office.Interop.Excel.Range)
                'celll.Value2 = "Top Ten Channels across Genres"
                'celll.ColumnWidth = 30
                'celll.Interior.Color = System.Drawing.Color.Yellow


                'For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                '    Dim expression As String = "Tgroup = '" + ptgname + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString() + "'"
                '    Dim sOrder As String = "RANK ASC"
                '    Dim topTenTable = dSet.Tables(1).[Select](expression, sortOrder).CopyToDataTable()
                '    topTenTable.Columns.RemoveAt(0)
                '    topTenTable.Columns.RemoveAt(0)
                '    topTenTable.Columns.RemoveAt(0)
                '    Dim marketcell As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(6 + listObject.ListRows.Count, 1 + (index * 3)), vstoWorkbook.Cells(6 + listObject.ListRows.Count, 3 + (index * 3))), Microsoft.Office.Interop.Excel.Range)
                '    marketcell.Merge(True)
                '    marketcell.Value2 = Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString()
                '    Dim cell5 As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(7 + listObject.ListRows.Count, 1 + (index * 3)), vstoWorkbook.Cells(7 + listObject.ListRows.Count, 3 + (index * 3))), Microsoft.Office.Interop.Excel.Range)
                '    listobject1 = vstoWorkbook.Controls.AddListObject(cell5, "Top10AcrossGenresucp" + Date.Now.Minute.ToString() + Date.Now.Second.ToString() + Date.Now.Millisecond.ToString())
                '    listobject1.AutoSetDataBoundColumnHeaders = True
                '    ' listobject1.Range.Columns.AutoFit()
                '    listobject1.DataSource = topTenTable
                'Next
                DisplayGenreShareDetailsOnSheet(dSet, , True)
            Else
                Dim genres As Data.DataTable = dRefSet.Tables(0).DefaultView.ToTable(True, "Genre")

                'Dim dt1 As System.Data.DataTable = New System.Data.DataTable()
                ''  Dim copyGenreTab As System.Data.DataTable = ds.Tables(0).Copy()

                ''  Dim copyto10 As System.Data.DataTable = ds.Tables(1).Copy()
                'dt1.Columns.Add("Rank", System.Type.GetType("System.Int32"))
                'dt1.Columns.Add("Channel")
                '' dt1.Columns.Add(plantg + "-" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(0).ToString() + " GRP")
                ''  dt1.Columns.Add(reftg + "-" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(0).ToString() + " GRP")
                'dt.Columns.Add("Genre")
                '  dt.Columns.Add("Plan " + plantg + "-" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(0).ToString() + " GRP")
                '  dt.Columns.Add("Ref" + reftg + "-" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(0).ToString() + " GRP")
                ' Dim expression As String = "OrderQuantity = 2 and OrderID = 2" 
                ' Sort descending by column named CompanyName. 
                Dim sortOrder As String = "GRPShare in(%) DESC"
                Dim foundRows As DataRow()
                ' Dim exptopten As String = "TG"
                ' Use the Select method to find all rows matching the filter.
                ' Dim dtables As System.Data.DataTable() = New System.Data.DataTable(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count) {}
                ' foundtopten = planDs.Tables(1).[Select](expression, sortOrder)

                'For index = 0 To genres.Rows.Count - 1

                '    For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items.Count - 1
                '        Dim expression1 As String = "Tgroup = '" + EscapeLikeValue(rtgname) + "' and Mgroup = '" + EscapeLikeValue(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index1).ToString()) + "' and Genre = '" + genres.Rows(index)(0).ToString() + "'"
                '        foundRows = dRefSet.Tables(0).[Select](expression1)

                '        If foundRows.Count > 0 Then
                '            ' dtables(index) = New System.Data.DataTable()
                '            'dtables(index) = foundRows.CopyToDataTable()
                '            If Not (genres.Columns.Contains(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index1).ToString() + "-GRPShare in(%)")) Then
                '                genres.Columns.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index1).ToString() + "-GRPShare in(%)")
                '            End If
                '            genres.Rows(index)(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index1).ToString() + "-GRPShare in(%)") = foundRows(0).Item("GRPShare in(%)")
                '        Else
                '            genres.Rows(index)(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index1).ToString() + "-GRPShare in(%)") = 0
                '            'dt1.Columns.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString() + "-GRP")
                '        End If

                '    Next
                'Next
                ' foundRefCopy = copyGenreTab.Select(expcopy, sortOrder)
                'For Each row As DataRow In foundRows
                '    Dim dr As DataRow = dt.NewRow()
                '    dr("Genre") = row("Genre")
                '    For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1


                '        Dim expcopy As String = "Tgroup = '" + plantg + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString() + "' and Genre = '" + row("Genre").ToString() + "'"

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
                '    For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                '        Dim exptop As String = "Tgroup = '" + plantg + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString() + "' and Channel = '" + row1("Channel").ToString() + "'"

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
                nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                nativeSheet.UsedRange.Clear()
                For Each lo As ListObject In nativeSheet.ListObjects
                    lo.Delete()
                Next
                Dim cell As Microsoft.Office.Interop.Excel.Range = nativeSheet.Range("$A$1", Type.Missing)
                cell.Value2 = "Genre Share"
                cell.ColumnWidth = 15
                cell.Interior.Color = System.Drawing.Color.Yellow
                nativeSheet.Name = "Genre Share"
                '  nativeSheet.PageSetup.CenterFooter = "Genre Share and Channel share are calculated based on Program TVR"
                'Dim cell1 As Microsoft.Office.Interop.Excel.Range = nativeSheet.UsedRange


                Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(nativeSheet)
                Dim cell1 As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(3, 2), vstoWorkbook.Cells(3, 2 + genres.Columns.Count - 1))
                cell1.Merge(True)
                cell1.Value2 = rtgname



                Dim cell2 As Microsoft.Office.Interop.Excel.Range = cell1.Offset(1, 0)
                listObject = vstoWorkbook.Controls.AddListObject(cell2, "GenreShareuc" + Date.Now.Minute.ToString() + Date.Now.Second.ToString() + Date.Now.Millisecond().ToString())
                listObject.AutoSetDataBoundColumnHeaders = True
                ' listObject.QueryTable.AdjustColumnWidth = True
                'listObject.Range.Columns
                '  listObject.Range.Columns.AutoFit()
                Try
                    listObject.DataSource = genres.AsEnumerable().Take(NumericUpDown1.Value).CopyToDataTable()
                Catch ex As Exception
                    listObject.DataSource = genres.AsEnumerable().Take(genres.Rows.Count).CopyToDataTable()
                End Try

                Dim cell3 As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.UsedRange.Next(1, 2)
                Dim chartObjects As ChartObjects = vstoWorkbook.ChartObjects()

                chartObjects.Delete()

                'Dim chart As ChartObject = chartObjects.Add(listObject.ListColumns.Count * 85, listObject.Range.Top, 250, 250)
                'chart.Chart.ChartType = XlChartType.xl3DPie
                ''chart.Chart.
                'chart.Chart.SetSourceData(listObject.Range, Type.Missing)

                'For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items.Count - 1
                '    Dim chart As ChartObject = chartObjects.Add((listObject.ListColumns.Count * 85 * (index + 1)) + (10 * index), listObject.Range.Top, 250, 250)
                '    chart.Chart.ChartType = XlChartType.xl3DPie
                '    Dim seriesCollection1 As Microsoft.Office.Interop.Excel.SeriesCollection = chart.Chart.SeriesCollection()
                '    Dim series11 As Microsoft.Office.Interop.Excel.Series = seriesCollection1.NewSeries()
                '    series11.XValues = listObject.ListColumns(1).Range
                '    series11.Values = listObject.ListColumns(index + 2).Range
                '    series11.Name = listObject.ListColumns(index + 2).Name
                'Next
                Dim celll As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(6 + listObject.ListRows.Count, 1), vstoWorkbook.Cells(6 + listObject.ListRows.Count, 1)), Microsoft.Office.Interop.Excel.Range)
                celll.Value2 = "Top Ten Channels across Genres"
                celll.ColumnWidth = 30
                celll.Interior.Color = System.Drawing.Color.Yellow


                'For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items.Count - 1
                '    Dim expression As String = "Tgroup = '" + EscapeLikeValue(rtgname) + "' and Mgroup = '" + EscapeLikeValue(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index).ToString()) + "'"
                '    Dim sOrder As String = "RANK ASC"
                '    Dim topTenTable = dRefSet.Tables(1).[Select](expression, sOrder).CopyToDataTable()
                '    topTenTable.Columns.RemoveAt(0)
                '    topTenTable.Columns.RemoveAt(0)
                '    topTenTable.Columns.RemoveAt(0)

                '    Dim marketcell As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(7 + listObject.ListRows.Count, 1 + (index * 3)), vstoWorkbook.Cells(7 + listObject.ListRows.Count, 3 + (index * 3))), Microsoft.Office.Interop.Excel.Range)
                '    marketcell.Merge(True)
                '    marketcell.Value2 = Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString()
                '    Dim cell5 As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(8 + listObject.ListRows.Count, 1 + (index * 3)), vstoWorkbook.Cells(8 + listObject.ListRows.Count, 3 + (index * 3))), Microsoft.Office.Interop.Excel.Range)
                '    listobject1 = vstoWorkbook.Controls.AddListObject(cell5, "Top10AcrossGenresuc" + Date.Now.Minute.ToString() + Date.Now.Second.ToString() + Date.Now.Millisecond.ToString())
                '    listobject1.AutoSetDataBoundColumnHeaders = True
                '    ' listobject1.Range.Columns.AutoFit()
                '    listobject1.DataSource = topTenTable
                'Next
                ' DisplayGenreShareDetailsOnSheet(dRefSet)
            End If
            Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault
        Catch ex As Exception
            LogMpsrintExException("Exception occured while displaying Genre Share details." + ex.Message)
            Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault
            System.Windows.Forms.MessageBox.Show("Exception occured while displying Genre share details.Please refer to error log for more details")
        End Try
    End Function
    Public Function DisplayChannelShare()
        Try

            Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlWait
            If ComboBox1.Text = ptgname Then
                'Dim genreTables As System.Data.DataSet = New System.Data.DataSet()
                ''Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items.Count - 1) {}
                ''dt1 = New System.Data.DataTable()
                ''dt1.Columns.Add("Rank", System.Type.GetType("System.Int32"))
                ''dt1.Columns.Add("Channel")
                ''dt1.Columns.Add("Programme")
                ''dt1.Columns.Add("Start Hour")
                ''dt1.Columns.Add("Plan " + plantgname + "~" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(0).ToString() + " GRP")
                ''dt1.Columns.Add("Ref " + reftgname + "~" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(0).ToString() + " GRP")


                'For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items.Count - 1
                '    genreTables.Tables.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index))
                '    genreTables.Tables(index).Columns.Add("Channel")
                '    For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                '        'Dim expression1 As String = "Tgroup = '" + plantgname + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "' and Genre= '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index).ToString()
                '        'Dim sOrder As String = "GRP DESC"
                '        'Dim rows As Data.DataRow() = planDs.Tables(0).[Select](expression1, sOrder)

                '        genreTables.Tables(index).Columns.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRP")
                '        '  genreTables(index).Rows.Add(rows(0).Item("Channel Name").ToString(),
                '        '  genreTables(index).Rows(index)(.Columns.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index) + " ~ " + " Ref " + reftgname + "-" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(0).ToString() + " GRP")
                '    Next
                'Next
                ''For Each row As Data.DataRow In cGenre.Rows
                ''    For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                ''            Dim expression1 As String = "Tgroup = '" + ptgname + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "' and Genre= '" + row("Genre").ToString() + "'"
                ''        Dim sOrder As String = "GRP DESC"
                ''        Dim rows As Data.DataRow() = dSet.Tables(0).[Select](expression1, sOrder)
                ''        genreTables.Tables(row("Genre").ToString()).Rows.Add(row("Channel").ToString(), rows(0).Item("GRP"))
                ''    Next
                ''Next

                'For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items.Count - 1

                '    For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1

                '        Dim sortOrder As String = "GRP DESC"
                '        Dim expression1 As String = "Tgroup = '" + ptgname + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "'  and Genre = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index) + "'"
                '        Dim foundRows As Data.DataRow() = dSet.Tables(0).[Select](expression1, sortOrder)

                '        For Each dRow As DataRow In foundRows

                '            If index1 = 0 Then
                '                Dim dr As DataRow = genreTables.Tables(index).NewRow()
                '                dr("Channel") = dRow("Channel Name")

                '                dr(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRP") = dRow("GRP")
                '                ' dr(2) = dsRef.Tables(0).[Select](expcopy, sortOrder)(0).Item("GRP")
                '                genreTables.Tables(index).Rows.Add(dr)
                '            Else
                '                For Each rrow As DataRow In genreTables.Tables(index).Rows

                '                    If rrow("Channel") = dRow("Channel Name") Then
                '                        rrow(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "-GRP") = dRow("GRP")
                '                    End If
                '                Next
                '            End If
                '        Next
                '    Next
                'Next
                '' Dim expression As String = "OrderQuantity = 2 and OrderID = 2" 
                '' Sort descending by column named CompanyName. 
                ''Dim expression As String = "Tgroup = '" + plantgname + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(0).ToString() + "'"

                ''Dim sortOrder As String = "GRP DESC"
                ''Dim foundRows, foundtopten As DataRow()
                ' '' Dim exptopten As String = "TG"
                ' '' Use the Select method to find all rows matching the filter.
                ''foundtopten = ds.Tables(1).[Select](expression, sortOrder)
                '' foundRefCopy = copyGenreTab.Select(expcopy, sortOrder)
                ''For Each row As DataRow In foundRows
                ''    Dim expcopy As String = "Tgroup = '" + ComboBox2.Text.Trim() + "' and Mgroup = '" + CheckedListBox2.CheckedItems(0).ToString() + "' and Genre = '" + row("Genre").ToString() + "' and Channel Name = '"+row(
                ''    Dim dr As DataRow = dt.NewRow()
                ''    dr("Genre") = row("Genre")
                ''    dr(1) = row("GRP")
                ''    '  foundRefCopy()
                ''    dr(2) = copyGenreTab.[Select](expcopy, sortOrder)(0).Item("GRP")
                ''    dt.Rows.Add(dr)
                ''Next

                ''For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items.Count - 1
                ''    Dim expression1 As String = "Tgroup = '" + plantgname + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(0).ToString() + "'  and Genre = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index) + "'"
                ''    foundRows = ds.Tables(0).[Select](expression1, sortOrder)
                ''    For Each dRow As DataRow In foundRows
                ''        Dim expcopy As String = "Tgroup = '" + reftgname + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(0).ToString() + "' and Genre = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index) + "' and [Channel Name] = '" + dRow("Channel Name") + "'"
                ''        Dim dr As DataRow = genreTables(index).NewRow()
                ''        dr("Channel") = dRow("Channel Name")
                ''        dr(1) = dRow("GRP")
                ''        dr(2) = dsRef.Tables(0).[Select](expcopy, sortOrder)(0).Item("GRP")
                ''        genreTables(index).Rows.Add(dr)
                ''    Next

                ''Next

                '' foundRows.CopyToDataTable(,LoadOption.
                ''For Each row1 As DataRow In foundtopten
                ''    Dim exptop As String = "Tgroup = '" + reftgname + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(0).ToString() + "' and Channel = '" + row1("Channel").ToString() + "' and [Program Name] = '" + row1("Program Name") + "' and [Program Start Time] = '" + row1("Program Start Time") + "'"
                ''    Dim dr As DataRow = dt1.NewRow()
                ''    dr("Rank") = row1("Rank")
                ''    'Dim i As Integer = Convert.ToInt32(row1("ChannelCode").ToString())
                ''    'Dim ss As String = String.Empty
                ''    'If i < 10 Then
                ''    '    ss = "00" + i.ToString()
                ''    'ElseIf i < 100 Then
                ''    '    ss = "0" + i.ToString()
                ''    'Else
                ''    '    ss = row1("ChannelCode").ToString()
                ''    'End If
                ''    'Dim drow As DataRow = dg.GetChannels().Select("TAM_CHANNEL_CODE = '" + ss + "'")(0)
                ''    dr("Channel") = row1("Channel").ToString()
                ''    dr("Programme") = row1("Program Name").ToString()
                ''    dr("Start Hour") = row1("Program Start Time")

                ''    dr(4) = row1("GRP")
                ''    '  foundRefCopy()
                ''    Dim drr As DataRow() = dsRef.Tables(1).[Select](exptop, sortOrder)

                ''    If drr Is Nothing Or drr.Count = 0 Then
                ''        dr(5) = 0
                ''    Else
                ''        dr(5) = drr(0).Item("GRP")

                ''    End If

                ''    dt1.Rows.Add(dr)
                ''Next

                'newSheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                'newSheet.UsedRange.Clear()
                'newSheet.Name = "Channel Share"
                'Dim cell As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$A$1", Type.Missing)
                'cell.Value2 = "Channel Share"
                'cell.Interior.Color = System.Drawing.Color.Yellow
                'cell.ColumnWidth = 15
                'Dim cell1 As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$C$1", "$C$1")
                'cell1.Value2 = "Eval Start Date"
                'Dim cell2 As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$D$1", "$D$1")
                'cell2.Value2 = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value
                'cell2.NumberFormat = "dd/mm/yyyy"
                'Dim weekcell1 As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$E$1", "$E$1")
                'weekcell1.Value2 = String.Format("Week {0}", Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks.Rows(0)(0).ToString())
                'Dim evalendcell As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$C$2", "$C$2")
                'evalendcell.Value2 = "Eval End Date"
                'Dim cell23 As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$D$2", "$D$2")
                'cell23.Value2 = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value
                'cell23.NumberFormat = "dd/mm/yyyy"
                'Dim weekcell11 As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$E$2", "$E$2")
                'weekcell11.Value2 = String.Format("Week {0}", Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks.Rows(0)(0).ToString())
                'Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(newSheet)
                '' Dim listObject As Excel.ListObject()
                'Dim ii As Integer = 1
                'Dim rocount As Integer = 0
                'For index = 0 To genreTables.Tables.Count - 1
                '    '4,1 - 4,4 ;4,6 - 4,9;4,11- 4,14
                '    Dim genrecell As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(3, ii), vstoWorkbook.Cells(3, ii + 3)), Microsoft.Office.Interop.Excel.Range)
                '    genrecell.Merge(True)
                '    genrecell.Value2 = genreTables.Tables(index).TableName
                '    Dim lrange As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(4, ii), vstoWorkbook.Cells(4, ii + 3)), Microsoft.Office.Interop.Excel.Range)
                '    Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "list1" + index.ToString() + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString())
                '    listobject.AutoSetDataBoundColumnHeaders = True
                '    '  listobject.ListColumns(0).
                '    Try


                '        listobject.DataSource = genreTables.Tables(index).AsEnumerable().Take(NumericUpDown1.Value).CopyToDataTable()
                '    Catch ex As Exception
                '        listobject.DataSource = genreTables.Tables(index).AsEnumerable().Take(genreTables.Tables(index).Rows.Count).CopyToDataTable()

                '    End Try
                '    ii += 5
                '    If listobject.ListRows.Count > rocount Then
                '        rocount = listobject.ListRows.Count
                '    End If
                'Next

                '' Dim row As Integer = Globals.ThisAddIn.Application.ActiveCell.Row + 2
                'Dim celll As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(5 + rocount, 1), vstoWorkbook.Cells(5 + rocount, 1)), Microsoft.Office.Interop.Excel.Range)
                'celll.Value2 = "Top Ten Programs for period"
                'celll.ColumnWidth = 30
                'celll.Interior.Color = System.Drawing.Color.Yellow

                '' Dim row11 As Integer = Globals.ThisAddIn.Application.ActiveCell.Row + 2

                'For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                '    Dim expression As String = "Tgroup = '" + ptgname + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString() + "'"
                '    Dim sortOrder As String = "GRP DESC"
                '    Dim foundtopten As DataRow()
                '    ' Dim exptopten As String = "TG"
                '    ' Use the Select method to find all rows matching the filter.
                '    foundtopten = dSet.Tables(1).[Select](expression, sortOrder)
                '    Dim top10 As Data.DataTable = foundtopten.CopyToDataTable()
                '    top10.Columns.RemoveAt(0)
                '    top10.Columns.RemoveAt(0)
                '    top10.Columns.RemoveAt(0)
                '    top10.Columns.RemoveAt(0)
                '    Dim marketcell As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(6 + rocount, 1 + (index * 5)), vstoWorkbook.Cells(6 + rocount, 5 + (index * 5))), Microsoft.Office.Interop.Excel.Range)
                '    marketcell.Merge(True)
                '    marketcell.Value2 = Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index)
                '    Dim cell5 As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(7 + rocount, 1 + (index * 5)), vstoWorkbook.Cells(7 + rocount, 5 + (index * 5))), Microsoft.Office.Interop.Excel.Range)
                '    listobject1 = vstoWorkbook.Controls.AddListObject(cell5, "csharetop10" + Date.Now.Minute.ToString() + Date.Now.Second.ToString() + Date.Now.Millisecond.ToString())
                '    listobject1.AutoSetDataBoundColumnHeaders = True
                '    ' listobject1.Range.Columns.AutoFit()
                '    listobject1.DataSource = top10
                'Next

                DisplayChannelShareDetailsonSheet(dSet, True)
            Else
                Dim genreTables As System.Data.DataSet = New System.Data.DataSet()
                'Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items.Count - 1) {}
                'dt1 = New System.Data.DataTable()
                'dt1.Columns.Add("Rank", System.Type.GetType("System.Int32"))
                'dt1.Columns.Add("Channel")
                'dt1.Columns.Add("Programme")
                'dt1.Columns.Add("Start Hour")
                'dt1.Columns.Add("Plan " + plantgname + "~" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(0).ToString() + " GRP")
                'dt1.Columns.Add("Ref " + reftgname + "~" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(0).ToString() + " GRP")


                'For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items.Count - 1
                '    genreTables.Tables.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index))
                '    genreTables.Tables(index).Columns.Add("Channel")
                '    For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items.Count - 1
                '        'Dim expression1 As String = "Tgroup = '" + plantgname + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString() + "' and Genre= '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index).ToString()
                '        'Dim sOrder As String = "GRP DESC"
                '        'Dim rows As Data.DataRow() = planDs.Tables(0).[Select](expression1, sOrder)

                '        genreTables.Tables(index).Columns.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index1).ToString() + "-GRP")
                '        '  genreTables(index).Rows.Add(rows(0).Item("Channel Name").ToString(),
                '        '  genreTables(index).Rows(index)(.Columns.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index) + " ~ " + " Ref " + reftgname + "-" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(0).ToString() + " GRP")
                '    Next
                'Next
                'For Each row As Data.DataRow In cGenre.Rows
                '    For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items.Count - 1
                '        Dim expression1 As String = "Tgroup = '" + rtgname + "' and Mgroup = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index1).ToString() + "' and Genre= '" + row("Genre").ToString() + "'"
                '        Dim sOrder As String = "GRP DESC"
                '        Dim rows As Data.DataRow() = dRefSet.Tables(0).[Select](expression1, sOrder)
                '        genreTables.Tables(row("Genre").ToString()).Rows.Add(row("Channel").ToString(), rows(0).Item("GRP"))
                '    Next
                'Next
                'For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items.Count - 1

                '    For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items.Count - 1

                '        Dim sortOrder As String = "GRP DESC"
                '        Dim expression1 As String = "Tgroup = '" + EscapeLikeValue(rtgname) + "' and Mgroup = '" + EscapeLikeValue(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index1).ToString()) + "'  and Genre = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcGenres.lbSelectedGenres.Items(index) + "'"
                '        Dim foundRows As Data.DataRow() = dRefSet.Tables(0).[Select](expression1, sortOrder)

                '        For Each dRow As DataRow In foundRows

                '            If index1 = 0 Then
                '                Dim dr As DataRow = genreTables.Tables(index).NewRow()
                '                dr("Channel") = dRow("Channel Name")

                '                dr(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index1).ToString() + "-GRP") = dRow("GRP")
                '                ' dr(2) = dsRef.Tables(0).[Select](expcopy, sortOrder)(0).Item("GRP")
                '                genreTables.Tables(index).Rows.Add(dr)
                '            Else
                '                For Each rrow As DataRow In genreTables.Tables(index).Rows

                '                    If rrow("Channel") = dRow("Channel Name") Then
                '                        rrow(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index1).ToString() + "-GRP") = dRow("GRP")
                '                    End If
                '                Next
                '            End If
                '        Next
                '    Next
                'Next
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

                newSheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                newSheet.UsedRange.Clear()
                newSheet.Name = "Channel Share"
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
                '   vstoWorkbook.Controls.
                ' Dim listObject As Excel.ListObject()
                'Dim ii As Integer = 1 + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items.Count
                'Dim rocount As Integer = 0
                'For index = 0 To genreTables.Tables.Count - 1
                '    '4,1 - 4,4 ;4,6 - 4,9;4,11- 4,14
                '    ' Dim genrecell As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(3, ii), vstoWorkbook.Cells(3, ii + 3)), Microsoft.Office.Interop.Excel.Range)
                '    Dim genrecell As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(3, 1 + ((ii + 1) * index)), vstoWorkbook.Cells(3, 1 + ((ii + 1) * index))), Microsoft.Office.Interop.Excel.Range)

                '    genrecell.Merge(True)
                '    genrecell.Value2 = genreTables.Tables(index).TableName
                '    ' Dim lrange As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(4, ii), vstoWorkbook.Cells(4, ii + 3)), Microsoft.Office.Interop.Excel.Range)
                '    Dim lrange As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(4, 1 + ((ii + 1) * index)), vstoWorkbook.Cells(4, 1 + (ii + 1) * index)), Microsoft.Office.Interop.Excel.Range)
                '    Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "list1uc" + index.ToString() + Date.Now.Minute.ToString() + Date.Now.Second.ToString() + Date.Now.Millisecond.ToString())
                '    listobject.AutoSetDataBoundColumnHeaders = True
                '    '  listobject.ListColumns(0).
                '    Try

                '        listobject.DataSource = genreTables.Tables(index).AsEnumerable().Take(NumericUpDown1.Value).CopyToDataTable()
                '    Catch ex As Exception
                '        listobject.DataSource = genreTables.Tables(index).AsEnumerable().Take(genreTables.Tables(index).Rows.Count).CopyToDataTable()
                '    End Try
                '    '   ii += 5
                '    If listobject.ListRows.Count > rocount Then
                '        rocount = listobject.ListRows.Count
                '    End If
                'Next

                '' Dim row As Integer = Globals.ThisAddIn.Application.ActiveCell.Row + 2
                'Dim celll As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(5 + rocount, 1), vstoWorkbook.Cells(5 + rocount, 1)), Microsoft.Office.Interop.Excel.Range)
                'celll.Value2 = "Top Ten Programs for period"
                'celll.ColumnWidth = 30
                'celll.Interior.Color = System.Drawing.Color.Yellow

                '' Dim row11 As Integer = Globals.ThisAddIn.Application.ActiveCell.Row + 2
                'Dim tgcell As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(5 + rocount, 1), vstoWorkbook.Cells(5 + rocount, 5 + (Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items.Count * 5))), Microsoft.Office.Interop.Excel.Range)
                'For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items.Count - 1
                '    Dim expression As String = "Tgroup = '" + EscapeLikeValue(rtgname) + "' and Mgroup = '" + EscapeLikeValue(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index).ToString()) + "'"
                '    Dim sortOrder As String = "Rank ASC"
                '    Dim foundtopten As DataRow()
                '    ' Dim exptopten As String = "TG"
                '    ' Use the Select method to find all rows matching the filter.
                '    foundtopten = dRefSet.Tables(1).[Select](expression, sortOrder)
                '    Dim top10 As Data.DataTable = foundtopten.CopyToDataTable()
                '    top10.Columns.RemoveAt(0)
                '    top10.Columns.RemoveAt(0)
                '    top10.Columns.RemoveAt(0)
                '    top10.Columns.RemoveAt(0)
                '    ' Dim marketcell As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(6 + rocount, 1 + (index * 5)), vstoWorkbook.Cells(6 + rocount, 5 + (index * 5))), Microsoft.Office.Interop.Excel.Range)
                '    Dim marketcell As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(9 + rocount, 1 + (index * 5)), vstoWorkbook.Cells(9 + rocount, 1 + (index * 5))), Microsoft.Office.Interop.Excel.Range)

                '    marketcell.Merge(True)
                '    marketcell.Value2 = Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index)
                '    '  Dim cell5 As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(7 + rocount, 1 + (index * 5)), vstoWorkbook.Cells(7 + rocount, 5 + (index * 5))), Microsoft.Office.Interop.Excel.Range)
                '    Dim cell5 As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(11 + rocount, 1 + (index * 5)), vstoWorkbook.Cells(11 + rocount, 5 + (index * 5))), Microsoft.Office.Interop.Excel.Range)
                '    listobject1 = vstoWorkbook.Controls.AddListObject(cell5, "list2uc" + Date.Now.Minute.ToString() + Date.Now.Second.ToString() + Date.Now.Millisecond.ToString())
                '    listobject1.AutoSetDataBoundColumnHeaders = True
                '    ' listobject1.Range.Columns.AutoFit()
                '    listobject1.DataSource = top10
                'Next


            End If
            ' DisplayChannelShareDetailsonSheet(dRefSet)
            '  End If
            Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault
        Catch ex As Exception
            Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault
            LogMpsrintExException("Exception occured while displaying channel share details")
            System.Windows.Forms.MessageBox.Show("Exception occured while displying Channel share details.Please refer to error log for more details")
        End Try
    End Function
    Public Function DisplayBreakTVR()
        Try

            If ComboBox1.Text.Equals(ptgname) Then
                DisplayBreakTVRDetailsOnSheet(dSet)
            Else
                DisplayBreakTVRDetailsOnSheet(dRefSet)
            End If

        Catch ex As Exception
            LogMpsrintExException("Exception occured while displaying Break TVR details" + ex.Message)
            System.Windows.Forms.MessageBox.Show("Exception occured while displaying Break TVR details.Please refer to error log for more details")
        End Try
    End Function
    Public Function DisplayProgramTVR()
        Try

            Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlWait
            If ComboBox1.Text.Equals(ptgname) Then
                'newSheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                'newSheet.UsedRange.Clear()
                'newSheet.Name = "Program TVR"
                'Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(newSheet)
                '' Dim listObject As Excel.ListObject()
                'Dim cell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.Range("$A$4", Type.Missing)
                'cell.Value2 = String.Format("Period : {0} to {1}", Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToShortDateString(), Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.ToShortDateString())
                ''  cell.Interior.Color = System.Drawing.Color.Yellow
                '' cell.ColumnWidth = 15
                'Dim cell1 As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.Range("$A$6", "$A$6")
                'cell1.Value2 = String.Format("TG: {0}", ptgname)
                'Dim ii As Integer = 1
                'Dim rocount As Integer = 0
                'Dim channel = String.Empty
                'Dim lrange As Microsoft.Office.Interop.Excel.Range
                'Dim count As Integer = 8
                'For index = 0 To Globals.Ribbons.MSprintExRibbon.channels.Rows.Count - 1

                '    Dim cel As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count + (24 * index), 1), vstoWorkbook.Cells(count + (24 * index), 7 * dSet.Tables.Count + 3))
                '    cel.Merge(True)
                '    cel.Value2 = Globals.Ribbons.MSprintExRibbon.channels.Rows(index)("CName").ToString()

                '    ' Dim periodCell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count + 1 + (24 * index), 1), vstoWorkbook.Cells(count + 1 + (24 * index), 7 * planningDataSet.Tables.Count + 3))
                '    Dim periodCell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count + 1 + (24 * index), 1), vstoWorkbook.Cells(count + 1 + (24 * index), 7 * dSet.Tables.Count + 3))

                '    periodCell.Merge(True)
                '    periodCell.Value2 = String.Format("Period : {0} to {1}", Globals.Ribbons.MSprintExRibbon.channels.Rows(index)("StartDate").ToString(), Globals.Ribbons.MSprintExRibbon.channels.Rows(index)("EndDate").ToString())

                '    For index1 = 0 To dSet.Tables.Count - 1
                '        ' Dim marketcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count + 2 + (24 * index), 7 * index1 + 1), vstoWorkbook.Cells(count + 1 + (24 * index), (7 * index1 + 1) * index1 + 1))
                '        Dim marketcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count + 2 + (24 * index), 7 * index1 + 1), vstoWorkbook.Cells(count + 2 + (24 * index), (7 * index1 + 1) * index1 + 1))

                '        marketcell.Merge(True)
                '        marketcell.Value2 = dSet.Tables(index1).TableName.Split({"~"c}, StringSplitOptions.None)(1)

                '        Dim lorange As Microsoft.Office.Interop.Excel.Range = marketcell.Offset(1, 0)
                '        Dim listobject = vstoWorkbook.Controls.AddListObject(lorange, "list1C" + index.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString() + Date.Now.Millisecond.ToString())
                '        listobject.AutoSetDataBoundColumnHeaders = True
                '        ' 04/08/2013 ,10/08/2013
                '        '  Dim dt As System.Data.DataTable = planningDataSet.Tables(index1).Select("ChannelName = '" + tvrform.cbChannels.CheckedItems(0).ToString() + "' and PeriodStartDate = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToShortDateString("DD/MM/YYYY") + "' and PeriodEndDate ='" + Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.ToShortDateString("DD/MM/YYYY") + "'").CopyToDataTable()
                '        'channels.Columns.Add("CName")
                '        'channels.Columns.Add("StartDate")
                '        '  channels.Columns.Add("EndDate")
                '        Dim dtrows As System.Data.DataRow() = dSet.Tables(index1).Select("ChannelName = '" + Globals.Ribbons.MSprintExRibbon.channels.Rows(index)("CName").ToString() + "' and PeriodStartDate = '" + Globals.Ribbons.MSprintExRibbon.channels.Rows(index)("StartDate").ToString() + "' and PeriodEndDate = '" + Globals.Ribbons.MSprintExRibbon.channels.Rows(index)("EndDate").ToString() + "' ")
                '        Dim dt As Data.DataTable
                '        If dtrows.Count > 0 Then
                '            dt = dtrows.CopyToDataTable()
                '            dt.Columns.RemoveAt(0)
                '            dt.Columns.RemoveAt(0)
                '            dt.Columns.RemoveAt(0)
                '            dt.Columns.RemoveAt(0)
                '            dt.AcceptChanges()
                '            Try
                '                listobject.DataSource = dt.AsEnumerable().Take(NumericUpDown1.Value).CopyToDataTable()
                '            Catch ex As Exception
                '                listobject.DataSource = dt.AsEnumerable().Take(dt.Rows.Count).CopyToDataTable()
                '            End Try

                '        End If

                '        '.


                '    Next

                '    'If channel.Equals(channels.Rows(index)("CName").ToString()) Or index = 0 Then
                '    '    lrange = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(4, ii), vstoWorkbook.Cells(4, ii + 3)), Microsoft.Office.Interop.Excel.Range)
                '    'Else
                '    '    lrange = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(7 + rocount, 1), vstoWorkbook.Cells(7 + rocount, 1)), Microsoft.Office.Interop.Excel.Range)

                '    'End If


                '    'ii += 10
                '    'rocount = listobject.ListRows.Count
                '    'channel = channels.Rows(index)("CName").ToString()
                'Next
                DisplayProgTVRDetailsOnSheet(dSet, True)
            Else

                'newSheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                'newSheet.UsedRange.Clear()
                'newSheet.Name = "Program TVR"
                'Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(newSheet)
                '' Dim listObject As Excel.ListObject()
                'Dim cell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.Range("$A$4", Type.Missing)
                'cell.Value2 = String.Format("Period : {0} to {1}", Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToShortDateString(), Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.ToShortDateString())
                ''  cell.Interior.Color = System.Drawing.Color.Yellow
                '' cell.ColumnWidth = 15
                'Dim cell1 As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.Range("$A$6", "$A$6")
                'cell1.Value2 = String.Format("TG: {0}", rtgname)
                'Dim ii As Integer = 1
                'Dim rocount As Integer = 0
                'Dim channel = String.Empty
                'Dim lrange As Microsoft.Office.Interop.Excel.Range
                'Dim count As Integer = 8
                'For index = 0 To Globals.Ribbons.MSprintExRibbon.channels.Rows.Count - 1

                '    Dim cel As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count + Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.nudTopPrograms.Value + 4, 1), vstoWorkbook.Cells(count + Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.nudTopPrograms.Value + 4, 7 * dRefSet.Tables.Count + 3))
                '    ' cel.Merge(True)
                '    cel.Value2 = Globals.Ribbons.MSprintExRibbon.channels.Rows(index)("CName").ToString()

                '    ' Dim periodCell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count + 1 + (24 * index), 1), vstoWorkbook.Cells(count + 1 + (24 * index), 7 * planningDataSet.Tables.Count + 3))
                '    Dim periodCell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count + 1 + Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.nudTopPrograms.Value + 4, 1), vstoWorkbook.Cells(count + 1 + Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.nudTopPrograms.Value + 4, 7 * dRefSet.Tables.Count + 3))

                '    'periodCell.Merge(True)
                '    periodCell.Value2 = String.Format("Period : {0} to {1}", Globals.Ribbons.MSprintExRibbon.channels.Rows(index)("StartDate").ToString(), Globals.Ribbons.MSprintExRibbon.channels.Rows(index)("EndDate").ToString())

                '    For index1 = 0 To dRefSet.Tables.Count - 1
                '        ' Dim marketcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count + 2 + (24 * index), 7 * index1 + 1), vstoWorkbook.Cells(count + 1 + (24 * index), (7 * index1 + 1) * index1 + 1))
                '        Dim marketcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count + 2 + Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.nudTopPrograms.Value + 4, 7 * index1 + 1), vstoWorkbook.Cells(count + 2 + Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.nudTopPrograms.Value + 4, (7 * index1 + 1) * index1 + 1))

                '        '  marketcell.Merge(True)
                '        marketcell.Value2 = dRefSet.Tables(index1).TableName.Split({"~"c}, StringSplitOptions.None)(1)

                '        Dim lorange As Microsoft.Office.Interop.Excel.Range = marketcell.Offset(1, 0)
                '        Dim listobject = vstoWorkbook.Controls.AddListObject(lorange, "list1C" + index.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString() + Date.Now.Millisecond.ToString())
                '        listobject.AutoSetDataBoundColumnHeaders = True
                '        ' 04/08/2013 ,10/08/2013
                '        '  Dim dt As System.Data.DataTable = planningDataSet.Tables(index1).Select("ChannelName = '" + tvrform.cbChannels.CheckedItems(0).ToString() + "' and PeriodStartDate = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToShortDateString("DD/MM/YYYY") + "' and PeriodEndDate ='" + Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.ToShortDateString("DD/MM/YYYY") + "'").CopyToDataTable()
                '        'channels.Columns.Add("CName")
                '        'channels.Columns.Add("StartDate")
                '        '  channels.Columns.Add("EndDate")
                '        Dim dtrows As System.Data.DataRow() = dRefSet.Tables(index1).Select("ChannelName = '" + Globals.Ribbons.MSprintExRibbon.channels.Rows(index)("CName").ToString() + "' and PeriodStartDate = '" + Globals.Ribbons.MSprintExRibbon.channels.Rows(index)("StartDate").ToString() + "' and PeriodEndDate = '" + Globals.Ribbons.MSprintExRibbon.channels.Rows(index)("EndDate").ToString() + "' ")
                '        Dim dt As Data.DataTable
                '        If dtrows.Count > 0 Then
                '            dt = dtrows.CopyToDataTable()
                '            dt.Columns.RemoveAt(0)
                '            dt.Columns.RemoveAt(0)
                '            dt.Columns.RemoveAt(0)
                '            dt.Columns.RemoveAt(0)
                '            '  dt.Columns.RemoveAt(0)
                '            dt.AcceptChanges()
                '            Try
                '                listobject.DataSource = dt.AsEnumerable().Take(NumericUpDown1.Value).CopyToDataTable()
                '            Catch ex As Exception
                '                listobject.DataSource = dt.AsEnumerable().Take(dt.Rows.Count).CopyToDataTable()
                '            End Try
                '        End If

                '        '.


                '    Next

                '    'If channel.Equals(channels.Rows(index)("CName").ToString()) Or index = 0 Then
                '    '    lrange = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(4, ii), vstoWorkbook.Cells(4, ii + 3)), Microsoft.Office.Interop.Excel.Range)
                '    'Else
                '    '    lrange = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(7 + rocount, 1), vstoWorkbook.Cells(7 + rocount, 1)), Microsoft.Office.Interop.Excel.Range)

                '    'End If


                '    'ii += 10
                '    'rocount = listobject.ListRows.Count
                '    'channel = channels.Rows(index)("CName").ToString()
                'Next
                DisplayProgTVRDetailsOnSheet(dRefSet, True)
            End If
            Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault
        Catch ex As Exception
            Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault
            LogMpsrintExException("Exception occured while displying Program TVR details" + ex.Message)
            System.Windows.Forms.MessageBox.Show("Exception occured while displying Program TVR details.Please refer to error log for more details")
        End Try
    End Function

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        Dim text As String = String.Empty
        If sheet.Name.Equals("Genre Share") Then
            dSet = Globals.Ribbons.MSprintExRibbon.gshareds
            dRefSet = Globals.Ribbons.MSprintExRibbon.gsharerefds
            '  DisplayGenreShare()
        ElseIf sheet.Name.Equals("Channel Share") Then
            dSet = Globals.Ribbons.MSprintExRibbon.cshareds
            dRefSet = Globals.Ribbons.MSprintExRibbon.csharerefds
            ' DisplayChannelShare()
        ElseIf sheet.Name.Equals("Program TVR") Then
            dSet = Globals.Ribbons.MSprintExRibbon.ptvrds
            dRefSet = Globals.Ribbons.MSprintExRibbon.ptvrrefds
            '  DisplayProgramTVR()
        Else
            dSet = Globals.Ribbons.MSprintExRibbon.btvrds
            dRefSet = Globals.Ribbons.MSprintExRibbon.btvrrefds
            ' DisplayProgramTVR()
        End If
        lbmgs.Text = "Market Group(s) :"
        If ComboBox1.Text = ptgname Then
            NumericUpDown1.Maximum = dSet.Tables(0).Rows.Count
            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                lbmgs.Text = lbmgs.Text + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index)

                If index <> Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1 Then
                    lbmgs.Text = lbmgs.Text + ","
                End If
            Next
        ElseIf ComboBox1.Text = rtgname Then
            'NumericUpDown1.Maximum = dRefSet.Tables(0).Rows.Count
            'For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items.Count - 1
            '    lbmgs.Text = lbmgs.Text + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index)

            '    If index <> Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items.Count - 1 Then
            '        lbmgs.Text = lbmgs.Text + ","
            '    End If
            'Next
        End If
        ' lbmgs.Refresh()




    End Sub
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

    Private Sub btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        RaiseEvent ShowHide_Click()
    End Sub
    Public Event ShowHide_Click()

    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged

    End Sub
End Class
