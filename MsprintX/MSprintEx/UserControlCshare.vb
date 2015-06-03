Imports System.Data
Imports Microsoft.Office.Tools.Ribbon

Imports Microsoft.Office.Tools
Imports Microsoft.Office.Interop.Excel
Public Class UserControlCshare
    Dim dSet, dRefSet As DataSet
    Dim planmgtg As String()
    Dim refmgtg As String()
    Dim ptgname As String
    Dim rtgname As String
    Dim dt1 As System.Data.DataTable
    Friend tpSelections As MSprintEx.ucPlanSelections
    Friend listObject, listobject1 As Excel.ListObject
    Friend nativeSheet, newSheet, vstoWorkbook As Microsoft.Office.Interop.Excel.Worksheet
    Friend dg As New METISTableAdapters.CHANNEL_MASTERTableAdapter
    Public Sub New(ByVal ds As DataSet, ByVal dsReference As DataSet, ByVal plan As String(), ByVal ref As String(), ByVal plantgname As String, ByVal reftgname As String, ByVal tpS As MSprintEx.ucPlanSelections)
        InitializeComponent()
        dSet = ds
        planmgtg = plan
        refmgtg = ref
        ptgname = plantgname
        rtgname = reftgname
        tpSelections = tpS
        dRefSet = dsReference
        '  purposeVal = purpose
    End Sub
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        ' Dim copyGenreTab As System.Data.DataTable = dSet.Tables(0).Copy()
        '   Dim copyto10 As System.Data.DataTable = dSet.Tables(1).Copy()
        Dim genreTables As System.Data.DataTable() = New System.Data.DataTable(3) {}
        'tpSelections.UcGenres.lbSelectedGenres.Items.Count - 1) {}
        dt1 = New System.Data.DataTable()
        dt1.Columns.Add("Rank", System.Type.GetType("System.Int32"))
        dt1.Columns.Add("Channel")
        dt1.Columns.Add("Programme")
        dt1.Columns.Add("Start Hour")
        dt1.Columns.Add("Plan " + ptgname + "~" + tpSelections.UcMarkets1.lbPlan.Items(0).ToString() + " GRP")
        'dt1.Columns.Add("Ref " + rtgname + "~" + tpSelections.UcMarkets1.lbRef.Items(0).ToString() + " GRP")


        'For index = 0 To tpSelections.UcGenres.lbSelectedGenres.Items.Count - 1
        '    genreTables(index) = New System.Data.DataTable(tpSelections.UcGenres.lbSelectedGenres.Items(index))
        '    genreTables(index).Columns.Add("Channel")
        '    genreTables(index).Columns.Add(tpSelections.UcGenres.lbSelectedGenres.Items(index) + " ~ " + " Plan " + ptgname + "-" + tpSelections.UcMarkets1.lbPlan.Items(0).ToString() + " GRP")
        '    genreTables(index).Columns.Add(tpSelections.UcGenres.lbSelectedGenres.Items(index) + " ~ " + " Ref " + rtgname + "-" + tpSelections.UcMarkets1.lbRef.Items(0).ToString() + " GRP")
        'Next
        ' Dim expression As String = "OrderQuantity = 2 and OrderID = 2" 
        ' Sort descending by column named CompanyName. 
        Dim expression As String = "Tgroup = '" + ptgname + "' and Mgroup = '" + tpSelections.UcMarkets1.lbPlan.Items(0).ToString() + "'"

        Dim sortOrder As String = "GRP DESC"
        Dim foundRows, foundtopten As DataRow()
        ' Dim exptopten As String = "TG"
        ' Use the Select method to find all rows matching the filter.
        foundtopten = dSet.Tables(1).[Select](expression, sortOrder)
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

        'For index = 0 To tpSelections.UcGenres.lbSelectedGenres.Items.Count - 1
        '    Dim expression1 As String = "Tgroup = '" + ptgname + "' and Mgroup = '" + tpSelections.UcMarkets1.lbPlan.Items(0).ToString() + "'  and Genre = '" + tpSelections.UcGenres.lbSelectedGenres.Items(index) + "'"
        '    foundRows = dSet.Tables(0).[Select](expression1, sortOrder)
        '    For Each dRow As DataRow In foundRows
        '        Dim expcopy As String = "Tgroup = '" + rtgname + "' and Mgroup = '" + tpSelections.UcMarkets1.lbRef.Items(0).ToString() + "' and Genre = '" + tpSelections.UcGenres.lbSelectedGenres.Items(index) + "' and [Channel Name] = '" + dRow("Channel Name") + "'"
        '        Dim dr As DataRow = genreTables(index).NewRow()
        '        dr("Channel") = dRow("Channel Name")
        '        dr(1) = dRow("GRP")
        '        dr(2) = dRefSet.Tables(0).[Select](expcopy, sortOrder)(0).Item("GRP")
        '        genreTables(index).Rows.Add(dr)
        '    Next

        'Next


        For Each row1 As DataRow In foundtopten
            '  Dim exptop As String = "Tgroup = '" + rtgname + "' and Mgroup = '" + tpSelections.UcMarkets1.lbRef.Items(0).ToString() + "' and Channel = '" + row1("Channel").ToString() + "' and [Program Name] = '" + row1("Program Name") + "' and [Program Start Time] = '" + row1("Program Start Time") + "'"
            Dim exptop As String = String.Empty
            Dim dr As DataRow = dt1.NewRow()
            dr("Rank") = row1("Rank")
            'Dim i As Integer = Convert.ToInt32(row1("ChannelCode").ToString())
            'Dim ss As String = String.Empty
            'If i < 10 Then
            '    ss = "00" + i.ToString()
            'ElseIf i < 100 Then
            '    ss = "0" + i.ToString()
            'Else
            '    ss = row1("ChannelCode").ToString()
            'End If
            'Dim drow As DataRow = dg.GetChannels().Select("TAM_CHANNEL_CODE = '" + ss + "'")(0)
            dr("Channel") = row1("Channel").ToString()
            dr("Programme") = row1("Program Name").ToString()
            dr("Start Hour") = row1("Program Start Time")

            dr(4) = row1("GRP")
            '  foundRefCopy()
            Dim drr As DataRow() = dRefSet.Tables(1).[Select](exptop, sortOrder)

            If drr Is Nothing Or drr.Count = 0 Then
                dr(5) = 0
            Else
                dr(5) = drr(0).Item("GRP")

            End If

            dt1.Rows.Add(dr)
        Next

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
        cell2.Value2 = tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToString()
        Dim weekcell1 As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$E$1", "$E$1")
        weekcell1.Value2 = String.Format("Week {0}", tpSelections.TaskPaneLogFile1.dtWeeks.Rows(0)(0).ToString())
        Dim evalendcell As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$C$2", "$C$2")
        evalendcell.Value2 = "Eval End Date"
        Dim cell23 As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$D$2", "$D$2")
        cell23.Value2 = tpSelections.TaskPaneLogFile1.dtToDate.Value.ToString()
        Dim weekcell11 As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$E$2", "$E$2")
        weekcell11.Value2 = String.Format("Week {0}", tpSelections.TaskPaneLogFile1.dtWeeks.Rows(0)(0).ToString())
        Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(newSheet)
        ' Dim listObject As Excel.ListObject()
        Dim ii As Integer = 1
        Dim rocount As Integer = 0
        For index = 0 To genreTables.Length - 1
            '4,1 - 4,4 ;4,6 - 4,9;4,11- 4,14

            Dim lrange As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(4, ii), vstoWorkbook.Cells(4, ii + 3)), Microsoft.Office.Interop.Excel.Range)
            Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "list1" + index.ToString() + Date.Now.Hour.ToString() + Date.Now.Minute.ToString() + Date.Now.Millisecond().ToString())
            listobject.AutoSetDataBoundColumnHeaders = True
            listobject.DataSource = genreTables(index)
            ii += 5
            rocount = rocount + listobject.ListRows.Count
        Next

        ' Dim row As Integer = Globals.ThisAddIn.Application.ActiveCell.Row + 2
        Dim celll As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(5 + rocount, 1), vstoWorkbook.Cells(5 + rocount, 1)), Microsoft.Office.Interop.Excel.Range)
        celll.Value2 = "Top Ten Programs for period"
        celll.ColumnWidth = 30
        celll.Interior.Color = System.Drawing.Color.Yellow

        ' Dim row11 As Integer = Globals.ThisAddIn.Application.ActiveCell.Row + 2
        Dim cell5 As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(7 + rocount, 1), vstoWorkbook.Cells(7 + rocount, 1)), Microsoft.Office.Interop.Excel.Range)
        listobject1 = vstoWorkbook.Controls.AddListObject(cell5, "list2" + Date.Now.Hour.ToString() + Date.Now.Minute.ToString() + Date.Now.Millisecond().ToString())
        listobject1.AutoSetDataBoundColumnHeaders = True
        ' listobject1.Range.Columns.AutoFit()
        listobject1.DataSource = dt1
    End Sub

    Private Sub UserControlCshare_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cbpmgs.Items.AddRange(planmgtg)
        cbpmgs.Text = planmgtg(0)
        cbrefmgs.Items.AddRange(refmgtg)
        cbrefmgs.Text = refmgtg(0)
        cbPlan.Items.Add(ptgname)
        cbPlan.Text = ptgname
        CbRef.Items.Add(rtgname)
        CbRef.Text = rtgname
    End Sub

    Private Sub cbPlan_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPlan.SelectedIndexChanged

    End Sub

    Private Sub cbpmgs_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbpmgs.SelectedIndexChanged

    End Sub

    Private Sub CbRef_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CbRef.SelectedIndexChanged

    End Sub

    Private Sub cbrefmgs_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbrefmgs.SelectedIndexChanged

    End Sub
End Class
