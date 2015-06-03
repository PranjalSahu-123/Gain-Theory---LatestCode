Imports System.Windows.Forms
Imports System.IO
Imports System.Xml
Imports System.Web
Imports System.Data
Imports System.Net
Imports Microsoft.Office.Tools.Ribbon

Imports Microsoft.Office.Tools
Imports Microsoft.Office.Interop.Excel
Public Class GenreShareForm
    Dim mgs As System.Windows.Forms.ListBox.ObjectCollection
    Dim tgs As System.Windows.Forms.ListBox.ObjectCollection
    Friend tpAudience As ucAudience
    Friend tpSelections, selectionsObject As ucSelections
    Friend mpChannelShare As ucChannelShare
    Friend request As HttpWebRequest
    Friend ws As HttpWebResponse
    Friend stream As Stream
    Friend oStream As Stream
    Friend swriter As StreamWriter
    Friend sreader As StreamReader
    Friend mstream As MemoryStream
    Friend inputstring, postData As String
    Friend data As Byte()
    Friend dg As New METISTableAdapters.CHANNEL_MASTERTableAdapter
    Friend dt, dt1, dt2, dt3 As System.Data.DataTable
    Friend ds, ds1, ds2 As DataSet
    Friend dc As DataColumn
    Friend listObject, listobject1 As Excel.ListObject
    Friend nativeSheet, newSheet, vstoWorkbook As Microsoft.Office.Interop.Excel.Worksheet
    Friend WithEvents MSprintExTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Friend WithEvents MSprintExChannelShare As Microsoft.Office.Tools.CustomTaskPane
    Public Sub New(ByVal markets As System.Windows.Forms.ListBox.ObjectCollection, ByVal tgColl As System.Windows.Forms.ListBox.ObjectCollection)
        InitializeComponent()
        mgs = markets
        tgs = tgColl
    End Sub
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try

       
        Dim month, month1 As String
        Dim day, day1 As String
        Button1.Enabled = False
        lbGetting.Text = "Getting Genre Share for chosen TG-MGs..."
        lbGetting.Refresh()
        If DateTimePicker1.Value.Month < 10 Then
            month = "0" + DateTimePicker1.Value.Month.ToString()
        Else
            month = DateTimePicker1.Value.Month.ToString()
        End If
        If DateTimePicker1.Value.Day < 10 Then
            day = "0" + DateTimePicker1.Value.Day.ToString()
        Else
            day = DateTimePicker1.Value.Day.ToString()
        End If
        If DateTimePicker2.Value.Month < 10 Then
            month1 = "0" + DateTimePicker2.Value.Month.ToString()
        Else
            month1 = DateTimePicker2.Value.Month.ToString()
        End If
        If DateTimePicker2.Value.Day < 10 Then
            day1 = "0" + DateTimePicker2.Value.Day.ToString()
        Else
            day1 = DateTimePicker2.Value.Day.ToString()
        End If

        Dim input As XElement =
            <input>
                <pre-eval-period>
                    <startdate><%= DateTimePicker1.Value.Year.ToString() + month + day %></startdate>
                    <enddate><%= DateTimePicker2.Value.Year.ToString() + month1 + day1 %></enddate>
                </pre-eval-period>
            </input>
        Dim tgs As XElement = New XElement("targetgroups")
        'Dim doc As XmlDocument = New XmlDocument()
        '  doc.Load()
            tgs.Add(XElement.Load(Path.GetTempPath() + "\\TGS\\" + ComboBox1.Text.Trim() + ".xml"))

        If ComboBox2.Text.Trim() <> ComboBox1.Text.Trim() Then
            '  Dim doc1 As XmlDocument = New XmlDocument()
            '  doc1.Load()
            tgs.Add(XElement.Load(Path.GetTempPath() + "\\TGS\\" + ComboBox2.Text.Trim() + ".xml"))
        End If

        Dim markets As XElement = New XElement("markets")

        For index = 0 To CheckedListBox1.CheckedItems.Count - 1
            ' Dim doc2 As XmlDocument = New XmlDocument()
            '  doc2.Load()
            markets.Add(XElement.Load(Path.GetTempPath() + "\\MGS\\" + CheckedListBox1.CheckedItems(index).ToString().Trim() + ".xml"))
        Next
        input.Add(tgs)
        input.Add(markets)
        request = WebRequest.Create("http://ec2-54-254-193-184.ap-southeast-1.compute.amazonaws.com:8080/GroupM/genreshare/1")
        '

        request.Method = "POST"
        request.ContentType = "application/x-www-form-urlencoded"
        request.Timeout = 300000
        request.ServicePoint.MaxIdleTime = 300000
        request.KeepAlive = True
        inputstring = HttpUtility.UrlEncode(MyExtensionClasses.OuterXml(input))
        stream = request.GetRequestStream()
        '  Dim bytes As Byte() = Encoding.Default.GetBytes(input.OuterXml)

        Dim encoding As System.Text.UTF8Encoding = New UTF8Encoding()
        postData = "inputXML=" + inputstring

        data = encoding.GetBytes(postData)
        'input.Save(stream)
        ' request.ContentLength = data.Length
        stream.Write(data, 0, data.Length)


        ws = request.GetResponse()

        oStream = ws.GetResponseStream()
        Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")

        ' Pipe the stream to a higher level stream reader with the required encoding format.
        Dim readStream As New StreamReader(oStream, encode)
        '     Dim separators() As String = {"Genre,Viewership"}
        ' Dim file As [String]() = readStream.ReadToEnd().Split(separators, StringSplitOptions.RemoveEmptyEntries)
        Dim ds As DataSet = New DataSet()
        'Dim ds1 As DataSet = New DataSet()
        ds.Tables.Add("GenreViewership")
        ds.Tables.Add("TopTenChannels")
        ds.Tables(0).Columns.Add("TGroup")
        ds.Tables(0).Columns.Add("MGroup")
        ds.Tables(0).Columns.Add("Genre")
        ds.Tables(0).Columns.Add("GRP", System.Type.GetType("System.Int32"))
        ds.Tables(1).Columns.Add("TGroup")
        ds.Tables(1).Columns.Add("MGroup")
        ds.Tables(1).Columns.Add("Rank")
        ds.Tables(1).Columns.Add("Channel")
        ds.Tables(1).Columns.Add("GRP", System.Type.GetType("System.Int32"))

        While Not (readStream.EndOfStream)
            Dim s As String = readStream.ReadLine()
            Dim values As String()

            If s.Contains("Market:") Then
                Dim v As String() = s.Split({","c}, StringSplitOptions.None)


                ds.Tables(0).Columns("Mgroup").DefaultValue = v(1)
                ds.Tables(1).Columns("Mgroup").DefaultValue = v(1)



            End If
            If s.Contains("TG:") Then
                Dim v As String() = s.Split({","c}, StringSplitOptions.None)

                ds.Tables(0).Columns("TGroup").DefaultValue = v(1)
                ds.Tables(1).Columns("Tgroup").DefaultValue = v(1)


            End If
            If Not (s.Contains("Market:") Or s.Contains("TG:") Or s.Contains("Top 10 Channels") Or s.Equals(" ") Or s.Contains("Rank") Or s.Contains("Genre")) Then
                values = s.Split(New [String]() {","c}, StringSplitOptions.None)

                If Not (values Is Nothing) And values.Length.Equals(2) Then

                    'dr("Genre") = values(0)
                    'dr("GRP") = values(1)
                    ' ds.Tables(0).Rows.

                    ' If ds.Tables(0).Columns(0).DefaultValue.Equals(ComboBox1.Text.Trim()) And CheckedListBox1.Items.Contains(ds.Tables(0).Columns(1).DefaultValue.ToString()) Then
                    Dim dr As DataRow = ds.Tables(0).NewRow()
                    dr("Genre") = values(0)
                    dr("GRP") = values(1)
                    ds.Tables(0).Rows.Add(dr)


                    ' ds.Tables(0).Rows.Add(dr)
                ElseIf Not (values Is Nothing) And values.Length.Equals(3) And Not (values(0).Contains("Rank")) Then
                    ' If ds.Tables(0).Columns(0).DefaultValue.Equals(ComboBox1.Text.Trim()) And CheckedListBox1.Items.Contains(ds.Tables(0).Columns(1).DefaultValue.ToString()) Then
                    Dim dr As DataRow = ds.Tables(1).NewRow()
                    dr("Rank") = values(0)
                    dr("Channel") = values(1)
                    dr("GRP") = values(2)
                    ds.Tables(1).Rows.Add(dr)
                    'Else
                    '    Dim dr As DataRow = ds1.Tables(0).NewRow()
                    '    dr("Rank") = values(0)
                    '    dr("Channel") = values(1)
                    '    dr("GRP") = values(2)
                    '    ds1.Tables(0).Rows.Add(dr)
                    'End If
                End If
            End If

        End While
        ds.WriteXml(Path.GetTempPath() + "\\ds.xml")

        Dim dt As System.Data.DataTable = New System.Data.DataTable()
        Dim dt1 As System.Data.DataTable = New System.Data.DataTable()
        Dim copyGenreTab As System.Data.DataTable = ds.Tables(0).Copy()

        Dim copyto10 As System.Data.DataTable = ds.Tables(1).Copy()
        dt1.Columns.Add("Rank", System.Type.GetType("System.Int32"))
        dt1.Columns.Add("Channel")
        dt1.Columns.Add("Plan " + ComboBox1.Text.Trim() + "-" + CheckedListBox1.CheckedItems(0).ToString() + " GRP")
        dt1.Columns.Add("Ref " + ComboBox2.Text.Trim() + "-" + CheckedListBox2.CheckedItems(0).ToString() + " GRP")
        dt.Columns.Add("Genre")
        dt.Columns.Add("Plan " + ComboBox1.Text.Trim() + "-" + CheckedListBox1.CheckedItems(0).ToString() + " GRP")
        dt.Columns.Add("Ref" + ComboBox2.Text.Trim() + "-" + CheckedListBox2.CheckedItems(0).ToString() + " GRP")
        Dim expression As String = "Tgroup = '" + ComboBox1.Text.Trim() + "' and Mgroup = '" + CheckedListBox1.CheckedItems(0).ToString() + "'"
        ' Dim expression As String = "OrderQuantity = 2 and OrderID = 2" 
        ' Sort descending by column named CompanyName. 
        Dim sortOrder As String = "GRP DESC"
        Dim foundRows, foundtopten As DataRow()
        ' Dim exptopten As String = "TG"
        ' Use the Select method to find all rows matching the filter.
        foundRows = ds.Tables(0).[Select](expression, sortOrder)
        foundtopten = ds.Tables(1).[Select](expression, sortOrder)
        ' foundRefCopy = copyGenreTab.Select(expcopy, sortOrder)
        For Each row As DataRow In foundRows
            Dim expcopy As String = "Tgroup = '" + ComboBox2.Text.Trim() + "' and Mgroup = '" + CheckedListBox2.CheckedItems(0).ToString() + "' and Genre = '" + row("Genre").ToString() + "'"
            Dim dr As DataRow = dt.NewRow()
            dr("Genre") = row("Genre")
            dr(1) = row("GRP")
            '  foundRefCopy()
            dr(2) = copyGenreTab.[Select](expcopy, sortOrder)(0).Item("GRP")
            dt.Rows.Add(dr)
        Next
        For Each row1 As DataRow In foundtopten
            Dim exptop As String = "Tgroup = '" + ComboBox2.Text.Trim() + "' and Mgroup = '" + CheckedListBox2.CheckedItems(0).ToString() + "' and Channel = '" + row1("Channel").ToString() + "'"
            Dim dr As DataRow = dt1.NewRow()
            dr("Rank") = row1("Rank")
            Dim i As Integer = Convert.ToInt32(row1("Channel").ToString())
            Dim ss As String = String.Empty
            If i < 10 Then
                ss = "00" + i.ToString()
            ElseIf i < 100 Then
                ss = "0" + i.ToString()
            Else
                ss = row1("Channel").ToString()
            End If
            Dim drow As DataRow = dg.GetChannels().Select("TAM_CHANNEL_CODE = '" + ss + "'")(0)
            dr("Channel") = drow(1).ToString()
            dr(2) = row1(4)
            '  foundRefCopy()
            Dim drr As DataRow() = copyto10.[Select](exptop, sortOrder)

            If drr Is Nothing Or drr.Count = 0 Then
                dr(3) = 0
            Else
                dr(3) = drr(0).Item("GRP")

            End If

            dt1.Rows.Add(dr)
        Next
        nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        Dim cell As Microsoft.Office.Interop.Excel.Range = nativeSheet.Range("$A$1", Type.Missing)
        cell.Value2 = "Genre Share"
        cell.ColumnWidth = 15
        cell.Interior.Color = System.Drawing.Color.Yellow
        nativeSheet.Name = "Genre Share"
        '  nativeSheet.PageSetup.CenterFooter = "Genre Share and Channel share are calculated based on Program TVR"
        Dim cell1 As Microsoft.Office.Interop.Excel.Range = nativeSheet.UsedRange
        Dim cell2 As Microsoft.Office.Interop.Excel.Range = cell1.Next(3, 0)

        Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(nativeSheet)
        listObject = vstoWorkbook.Controls.AddListObject(cell2, "list1")
        listObject.AutoSetDataBoundColumnHeaders = True
        ' listObject.QueryTable.AdjustColumnWidth = True
        'listObject.Range.Columns
        '  listObject.Range.Columns.AutoFit()
        listObject.DataSource = dt

        Dim cell3 As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.UsedRange.Next(1, 2)
        Dim chartObjects As ChartObjects = vstoWorkbook.ChartObjects()
        Dim chart As ChartObject = chartObjects.Add(listObject.ListColumns.Count * 85, listObject.Range.Top, 250, 250)
        chart.Chart.ChartType = XlChartType.xl3DPie
        'chart.Chart.
        chart.Chart.SetSourceData(listObject.Range, Type.Missing)

        Dim celll As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(6 + listObject.ListRows.Count, 1), vstoWorkbook.Cells(6 + listObject.ListRows.Count, 1)), Microsoft.Office.Interop.Excel.Range)
        celll.Value2 = "Top Ten Channels across Genres"
        celll.ColumnWidth = 30
        celll.Interior.Color = System.Drawing.Color.Yellow

        Dim cell5 As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(7 + listObject.ListRows.Count, 1), vstoWorkbook.Cells(7 + listObject.ListRows.Count, 1)), Microsoft.Office.Interop.Excel.Range)
        listobject1 = vstoWorkbook.Controls.AddListObject(cell5, "list2")
        listobject1.AutoSetDataBoundColumnHeaders = True
        ' listobject1.Range.Columns.AutoFit()
        listobject1.DataSource = dt1
        'listobject1.QueryTable.AdjustColumnWidth = True
        ' vstoWorkbook.Controls.AddControl(System.Windows.Forms.Button,
            Me.Close()
        Catch ex As Exception
            MessageBox.Show("Exception occured while getting requested Genre share details")
            Me.Close()
        End Try
    End Sub

    Private Sub GenreShareForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

      
        ComboBox1.Items.Clear()
        ComboBox2.Items.Clear()
        CheckedListBox1.Items.Clear()
        CheckedListBox2.Items.Clear()
        For index = 0 To tgs.Count - 1
            ComboBox1.Items.Add(tgs(index))
            ComboBox2.Items.Add(tgs(index))
        Next

        'For index = 0 To mgs.Count - 1
        '    CheckedListBox1.Items.Add(mgs(index))
        '    CheckedListBox2.Items.Add(mgs(index))
        'Next
        For index = 0 To Directory.GetFiles(System.IO.Path.GetTempPath() + "\\MGS", "*.xml").Count - 1
            ' fileList.Add(Path.GetFileNameWithoutExtension(Directory.GetFiles(System.IO.Path.GetTempPath() + "\\MGS", "*.xml")(index)))
            CheckedListBox1.Items.Add(Path.GetFileNameWithoutExtension(Directory.GetFiles(System.IO.Path.GetTempPath() + "\\MGS", "*.xml")(index)))
            CheckedListBox2.Items.Add(Path.GetFileNameWithoutExtension(Directory.GetFiles(System.IO.Path.GetTempPath() + "\\MGS", "*.xml")(index)))

            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged
       
    End Sub

    Private Sub CheckedListBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBox2.SelectedIndexChanged
      
    End Sub

    Private Sub CheckedListBox1_ItemCheck(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles CheckedListBox1.ItemCheck
        'Dim text As String = "Planning Market Grps chosen are : "
        'For index = 0 To CheckedListBox1.CheckedItems.Count - 1
        '    text = text + CheckedListBox1.CheckedItems(index).ToString()

        '    If index <> CheckedListBox1.CheckedItems.Count - 1 Then
        '        text = text + ","
        '    End If

        'Next
        'LbplanMgchosen.Text = String.Empty
        'LbplanMgchosen.Text = text
    End Sub

    Private Sub CheckedListBox2_ItemCheck(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles CheckedListBox2.ItemCheck
        'Dim text As String = "Reference Market Grps chosen are :"
        'For index = 0 To CheckedListBox2.CheckedItems.Count - 1
        '    text = text + CheckedListBox2.CheckedItems(index).ToString()
        '    If index <> CheckedListBox2.CheckedItems.Count - 1 Then
        '        text = text + ","
        '    End If
        'Next
        'lbRefMgsCHosen.Text = String.Empty
        'lbRefMgsCHosen.Text = text
    End Sub

    Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged
        If CType(sender, System.Windows.Forms.DateTimePicker).Value < DateTimePicker1.Value Then
            MsgBox("End date cannot be less than start date.  Please correct the entry", MsgBoxStyle.Exclamation, "Resolve date")
            Exit Sub
        End If
    End Sub

    Private Sub DateTimePicker2_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles DateTimePicker2.Validating
        If CType(sender, System.Windows.Forms.DateTimePicker).Value < DateTimePicker1.Value Then
            MsgBox("End date cannot be less than start date.  Please correct the entry", MsgBoxStyle.Exclamation, "Resolve date")
            e.Cancel = True
        End If
    End Sub
End Class