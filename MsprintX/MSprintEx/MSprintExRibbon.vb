Imports Microsoft.Office.Tools.Ribbon
Imports System.Web.HttpUtility
'Imports System.Web.Services
Imports System.Net
Imports System.IO
Imports System.Collections.Generic
Imports System.Xml
Imports System.Text
'Imports System.Web.Http
Imports System.Data
Imports System.Collections
Imports System.Windows.Forms
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.Office.Tools
Imports System.ComponentModel
Imports System.DirectoryServices
Imports Microsoft.Office.Interop.Excel
Imports System.Globalization
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Timers

Public Class MSprintExRibbon
    Friend tpAudience As ucAudience
    Friend frmPrepareServer As frmPrepareServer
    Friend HRN As Integer = 0
    Friend MSHRN As Integer = 0
    Friend DSHRN As Integer = 0
    Friend CreativeSummaryHRN As Integer = 0
    Friend ChannelSummaryHRN As Integer = 0
    Friend tpSelections, selectionsObject As ucPlanSelections
    Friend mpChannelShare As ucChannelShare
    Friend mpGenEndTime As ucGenEndTime
    Dim ChannelCells As Microsoft.Office.Interop.Excel.Range
    Dim ChannelColumn As Microsoft.Office.Interop.Excel.ListColumn
    Dim planchannels As Data.DataTable
    Friend masterchannels As Data.DataTable
    Friend WithEvents channelMapping As ChannelMapping
    Friend errors As DataErrors
    Friend mpUCChannelShare As UserControlCshare
    Friend mpUcTVRScreen As ucTVRScreen
    Friend request As HttpWebRequest
    Friend ws As HttpWebResponse
    Friend stream As Stream
    Friend oStream As Stream
    Friend swriter As StreamWriter
    Friend sreader As StreamReader
    Friend mstream As MemoryStream
    Friend planOpenedSuccessfully As Boolean = False
    Friend WithEvents backGroundWorker As BackgroundWorker
    Friend inputstring, postData As String
    Friend currentLineItem As String = String.Empty
    Friend data As Byte()
    Friend dt, dt1, dt2, dt3 As System.Data.DataTable
    Friend ds, ds1, ds2, dsRef As DataSet
    Friend dc As DataColumn
    Friend dtMarkets1 As Data.DataTable = New Data.DataTable("Markets")
    Friend wsURLS As Data.DataTable = New Data.DataTable()
    Public db_FromDate, db_ToDate As Date
    Public db_WeekNo As Integer
    Friend dtchannels As Data.DataTable = New Data.DataTable("Channels")
    Friend bslMasterTable As Data.DataTable = New Data.DataTable("bsl")
    Friend dtGenres As Data.DataTable = New Data.DataTable("Genres")
    Friend selectedSpotsListObject, listObject, listobject1 As Excel.ListObject
    Dim ptvrRootNode As XElement = New XElement("input")
    Dim genreShareRootNode As XElement = New XElement("input")
    Dim channelShareRootNode As XElement = New XElement("input")
    Dim bslSummary_OP As XElement = New XElement("input")
    Public rnfoutputXml As XElement
    Public mediaplan, testmediaplan As XElement
    Friend gshareds, cshareds, ptvrds, btvrds, planningDataSet, gsharerefds, csharerefds, ptvrrefds, btvrrefds, refDataSet As System.Data.DataSet
    Public channels, mappedchannels, RnFOutputTable, RnFSelectedSpots, RnFMarketSummary, RnFChannelSummary, RnFCreativeSummary, RnFDurationSummary, RnFAvaiSpots, RnFShowResultsTable, RnFProgAvgTVRTable, xecellineItemsTable, xecelProgAvgTVRTable, xecelTable, xecellTableCopy As System.Data.DataTable
    Friend dg As New METISTableAdapters.CHANNEL_MASTERTableAdapter
    Friend nativeSheet, newSheet, vstoWorkbook As Microsoft.Office.Interop.Excel.Worksheet
    Friend rnFWorkBook As Microsoft.Office.Interop.Excel.Workbook
    Friend WithEvents MSprintExTaskPane, SpotSelectionPane As Microsoft.Office.Tools.CustomTaskPane
    Friend WithEvents logFileSavePath, savePlanPath As System.Windows.Forms.SaveFileDialog
    Friend tgDirectoryPath As String = AppDomain.CurrentDomain.BaseDirectory & "\\Masters\TGS\\"
    Friend mgDirectoryPath As String = AppDomain.CurrentDomain.BaseDirectory & "\\Masters\MGS\\"
    Friend ExceptionLogFilePath As String = AppDomain.CurrentDomain.BaseDirectory & "\\Masters\\MspintErrorLog.txt"
    Friend MasterFolderPath As String = AppDomain.CurrentDomain.BaseDirectory & "\\Masters\\"
    '  Friend InputXMLFolderPath As String = AppDomain.CurrentDomain.BaseDirectory & "\\Masters\\InputXMLs\\"
    ' Friend OutputXMLFolderPath As String = AppDomain.CurrentDomain.BaseDirectory & "\\Masters\\OutputXMLs\\"
    Friend LogDirectoryPath As String = String.Empty
    Friend logDirectoryXML As XElement
    Friend loggedInUserName As String
    Friend reorderedChannels As List(Of String)
    Friend markets As List(Of String) = New List(Of String)()
    Friend WithEvents MSprintExChannelShare, AvgTVRMGPane, ChannelPane, ErrorPane, CTSChannelShare, CTPProgTVR, CTPGenreShare, CTPChannelShare, CTPBreakTVR, TVRTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Public Selectedrow As System.Data.DataRow
    Public selectedRowIndex As Integer
    Friend Shared TempTableTimer As System.Timers.Timer
    Public Shared _createdPanes As New Dictionary(Of Workbook, Microsoft.Office.Tools.CustomTaskPane)()
    Enum GenreShareView
        All
        TopTen
    End Enum
    Private Sub MSprintExRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        Try
            wsURLS = GetAllWSURLS()
        Catch ex As Exception
            LogMpsrintExException("Exception occured while loading MsprintXRibbon.Message : " + ex.Message)
        End Try
        ' Console.WriteLine(ControlChars.Cr + "Response stream received")
        '  Dim read(256) As [Char]

        ' Read 256 charcters at a time    .
        ' Dim count As Integer = readStream.Read(read, 0, 256)
        'Console.WriteLine("HTML..." + ControlChars.Lf + ControlChars.Cr)
        'While count > 0

        '    ' Dump the 256 characters on a string and display the string onto the console.
        '    Dim str As New [String](read, 0, count)
        '    'Console.Write(str)
        '    count = readStream.Read(read, 0, 256)

        'End While

    End Sub
    Public Function DisplayAndConfigureDateRange()
        Try

        Catch ex As Exception
            LogMpsrintExException("Exception occured while displaying and configuring server for chosen date range.Message : " + ex.Message)
        End Try
    End Function
    Public Function DisplayTVBuilder()
        Try
            If Not tpSelections Is Nothing Then tpSelections.Dispose()
            tpSelections = New ucPlanSelections()
            ' Dim col As CustomTaskPaneCollection
          
            MSprintExTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(tpSelections, "TV Plan builder", Globals.ThisAddIn.Application.ActiveWindow)
           
            MSprintExTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight

            MSprintExTaskPane.Width = 350
            MSprintExTaskPane.Visible = True
            _createdPanes(Globals.ThisAddIn.Application.ActiveWorkbook) = MSprintExTaskPane ''Added By Alok for Task Pane Show/Hide
        Catch ex As Exception
            LogMpsrintExException("Exception occured while displaying TV Builder pane. Message: " + ex.Message)
            Throw ex
        End Try
    End Function

    'Added By Alok for Task Pane Show/Hide Start
    Public ReadOnly Property TaskPane() As Microsoft.Office.Tools.CustomTaskPane
        Get
            Dim MSprintExTaskPane_temp As Microsoft.Office.Tools.CustomTaskPane = Nothing
            For Each keypair As KeyValuePair(Of Workbook, Microsoft.Office.Tools.CustomTaskPane) In _createdPanes
                If keypair.Key Is Globals.ThisAddIn.Application.ActiveWorkbook Then
                    MSprintExTaskPane_temp = DirectCast(keypair.Value, Microsoft.Office.Tools.CustomTaskPane)
                    Exit For
                End If
            Next
            Return MSprintExTaskPane_temp
        End Get
    End Property

    Private Sub MSprintExTaskPane_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MSprintExTaskPane.VisibleChanged
        Globals.Ribbons.MSprintExRibbon.ToggleButton1.Checked = MSprintExTaskPane.Visible
    End Sub

    Private Sub ToggleButton1_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ToggleButton1.Click
        Me.TaskPane.Visible = TryCast(sender, Microsoft.Office.Tools.Ribbon.RibbonToggleButton).Checked
        'If TryCast(sender, Microsoft.Office.Tools.Ribbon.RibbonToggleButton).Checked Then
        '    Globals.ThisAddIn.CustomTaskPanes.Remove(MSprintExTaskPane)
        '    DisplayTVBuilder()
        '    dtchannels = New Data.DataTable("Channels")
        '    dtGenres = New Data.DataTable("Genres")
        '    dtMarkets1 = New Data.DataTable("Markets")
        '    PopulateGenresChannelsMarkets()
        '    Dim latestStatus As XElement = GetLatestWeekDetails()
        '    PopulateLatestDates(latestStatus)
        'End If
    End Sub
    'Added By Alok for Task Pane Show/Hide End

    Public Sub BackWorker(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles backGroundWorker.DoWork
        Try


            '  Globals.ThisAddIn.Application.StatusBar = "Loading MSprintEx application..."

            ' BackgroundWorker1.RunWorkerAsync()
            If Not (File.Exists(ExceptionLogFilePath)) Then
                File.Create(ExceptionLogFilePath)
            End If

            backGroundWorker.ReportProgress(10)

            If MachineConnectedToInternet() Then
                btnStartMsprint.Enabled = False
                PopulateGenresChannelsMarkets()
                backGroundWorker.ReportProgress(100)
                ' DisplayTVBuilder()
                'backGroundWorker.ReportProgress(90)


            Else
                System.Windows.Forms.MessageBox.Show("MSprintEx communicates with Server over Internet.Please ensure Internet connectivity and Try again.")
            End If
        Catch ex As Exception

        End Try
        ' Globals.ThisAddIn.Application.StatusBar = String.Empty
    End Sub
    Public Sub ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles backGroundWorker.ProgressChanged
        Globals.ThisAddIn.Application.StatusBar = "Loading MSprintEx.." + e.ProgressPercentage.ToString() + "%"
    End Sub
    Public Sub WorkCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles backGroundWorker.RunWorkerCompleted
        Globals.ThisAddIn.Application.StatusBar = String.Empty
        DisplayTVBuilder()
        EnableDisableButtons(True)
    End Sub
    Public Function PopulateBSLTabs()
        Try
            GetVariantMasterDetails()
            PopulateVariantTabs()
        Catch ex As Exception
            LogMpsrintExException("Exception occured while populating VariantMAsterTable .Message: " + ex.Message)
        End Try
    End Function
    Public Function PopulateVariantTabs()
        Try




        Catch ex As Exception

        End Try
    End Function
    Public Function VariantMasterIsUpdated(ByVal db_weekNumber As Integer) As Boolean
        Dim vMasterIsUpdated As Boolean = True
        Try
            Dim bslmaster As XElement = XElement.Load(AppDomain.CurrentDomain.BaseDirectory & "\\Masters\\BSLMaster.xml")
            Dim xmlWeekNumber As Integer = Convert.ToInt32(bslmaster.Attribute("WeekNumber").Value)

            If xmlWeekNumber = db_weekNumber Then
                vMasterIsUpdated = False
            End If

        Catch ex As Exception

        End Try
        Return vMasterIsUpdated
    End Function
    Public Function PopulateGenresChannelsMarkets()
        Try
            PopulateTabsFromMasterWs()
        Catch ex As Exception
            LogMpsrintExException("Exception occured while populating Genres,Channels and Markets Datasource." + ex.Message)
            Throw ex
        End Try
    End Function
    Private Sub btnOpen_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnOpen.Click
        Try
            If Globals.ThisAddIn.Application.Workbooks.Count < 1 Then
                '  Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(,,
                Globals.ThisAddIn.Application.Workbooks.Add(System.Type.Missing)
            End If

            planOpenedSuccessfully = False
            OpenMSprintEx()
            btnCleanupplan.Enabled = True
            btnMapChannels.Enabled = True
            btnRnF.Enabled = True

            btnGetReqSpots.Enabled = True
            btnGenerateEndTime.Enabled = True
            btnReorderPlanChannels.Enabled = True
        Catch ex As Exception
            LogMpsrintExException("Exception occured while creating Plan Selection sheet for entering plan" + ex.Message)
            MessageBox.Show("Exception occured while ")
        End Try


    End Sub

    Private Sub btnGenreShare_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        'Dim waitFrm As frmWait
        'Try
        '    Dim plantg As String = String.Empty
        '    Dim reftg As String = String.Empty


        '    Dim dtable As System.Data.DataTable = DirectCast(tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
        '    plantg = dtable.Rows(0)(1).ToString().Trim()
        '    reftg = dtable.Rows(1)(1).ToString()

        '    If Not (MachineConnectedToInternet()) Then
        '        MessageBox.Show("MSprintEx communicates with Server over Internet.Please ensure Internet connectivity and Try again.")

        '    ElseIf plantg.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
        '        System.Windows.Forms.MessageBox.Show("Please make Planning Target Group and Market(s) and/or Market Group(s) Selections")
        '    ElseIf plantg.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count > 0 Then
        '        System.Windows.Forms.MessageBox.Show("Please choose Planning Target group for Market groups")
        '    ElseIf plantg.Length > 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
        '        System.Windows.Forms.MessageBox.Show("Please choose Market(s) and/or Market Group(s) for chosen Planning Target Group")
        '    ElseIf reftg.Length = 0 And tpSelections.UcMarkets1.lbRef.Items.Count > 0 Then
        '        System.Windows.Forms.MessageBox.Show("Please choose Reference Target group for Market groups")
        '    ElseIf reftg.Length > 0 And tpSelections.UcMarkets1.lbRef.Items.Count < 1 Then
        '        System.Windows.Forms.MessageBox.Show("Please choose  Market(s) and/or Market Group(s) for reference Target Group chosen")
        '    Else

        '        Globals.ThisAddIn.Application.StatusBar = "Getting requested Genre Share details..."
        '        System.Windows.Forms.Application.DoEvents()
        '        ' Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlWait
        '        '  waitFrm = New frmWait()
        '        'waitFrm.Show()
        '        'waitFrm.Refresh()
        '        gshareds = New DataSet()

        '        If reftg.Length > 0 Then
        '            gsharerefds = New DataSet()
        '        End If

        '        GetGenreShare(plantg, reftg, gshareds, gsharerefds)


        '        If reftg.Length > 0 And gshareds.Tables.Count > 0 Then
        '            'Dim plan, ref As List(Of String)
        '            'plan = New List(Of String)()
        '            'ref = New List(Of String)()
        '            'For index = 0 To tpSelections.UcMarkets1.lbPlan.Items.Count - 1
        '            '    plan.Add(tpSelections.UcMarkets1.lbPlan.Items(index))
        '            'Next
        '            'For index = 0 To tpSelections.UcMarkets1.lbRef.Items.Count - 1
        '            '    ref.Add(tpSelections.UcMarkets1.lbRef.Items(index))
        '            'Next
        '            If CTPGenreShare Is Nothing Then
        '                mpChannelShare = New ucChannelShare()
        '                CTPGenreShare = Globals.ThisAddIn.CustomTaskPanes.Add(mpChannelShare, "Genre Share Selections")
        '                CTPGenreShare.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
        '                CTPGenreShare.Height = 226
        '                CTPGenreShare.Width = 300
        '                CTPGenreShare.Visible = True
        '            Else
        '                CTPGenreShare.Visible = True
        '                '  MSprintExChannelShare.Title = "Genre Share Selections"
        '            End If

        '            'tpSelections.TaskPaneLogFile1.scMain.Panel2Collapsed = False
        '            'If Not tpSelections.TaskPaneLogFile1.showingChannels Then
        '            '    tpSelections.TaskPaneLogFile1.ucChannelsMapping.btnShowHide.Text = "Genre Share Selections"
        '            '    tpSelections.TaskPaneLogFile1.ucChannelsMapping.btnShowHide.Refresh()
        '            '    tpSelections.TaskPaneLogFile1.showChannelMapping(True)
        '            'End If
        '            'tpSelections.TaskPaneLogFile1.Show()
        '        End If

        '        'waitFrm.Close()


        '        'waitFrm.Dispose()
        '        Globals.ThisAddIn.Application.StatusBar = String.Empty
        '        Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault

        '    End If
        'Catch ex As Exception

        '    'If Not (waitFrm) Is Nothing Then
        '    '    waitFrm.Close()
        '    '    waitFrm.Dispose()
        '    'End If

        '    Globals.ThisAddIn.Application.StatusBar = String.Empty
        '    Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault
        '    LogMpsrintExException("Exception occured while retreiving requested Genre Share details" + ex.Message)
        '    System.Windows.Forms.MessageBox.Show("Exception occured while getting requested Genre share details.Please refer to Error log for more details")


        'End Try
    End Sub
    Private Sub GetChannelShare(ByVal plantgname As String, ByVal reftgname As String, ByVal planDs As Data.DataSet, ByVal refDs As Data.DataSet, ByVal cGenre As Data.DataTable)
        Try

            'If plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
            '    System.Windows.Forms.MessageBox.Show("Please make Target Group and Market Group(s) Selections")
            'ElseIf reftgname.Length > 0 And tpSelections.UcMarkets1.lbRef.Items.Count < 1 Then
            '    System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for reference Target Group chosen")
            'ElseIf tpSelections.UcGenres.lbSelectedGenres.Items.Count < 1 Then
            '    System.Windows.Forms.MessageBox.Show("Please choose Genre(s) to view their Channel Share details")
            'Else
            Globals.ThisAddIn.Application.StatusBar = "Getting requested Channel Share details..."

            Dim input As XElement = ConstructChannelShareInputXML(plantgname, reftgname)
            'input.Add(tgs)
            'input.Add(markets)
            '  Dim genrelist As XElement = New XElement("genre_list")

            '  input.Add(genrelist)

            'request = WebRequest.Create("http://ec2-54-255-217-55.ap-southeast-1.compute.amazonaws.com:8080/GroupM/channelshare/")
            ''

            'request.Method = "POST"
            'request.ContentType = "application/x-www-form-urlencoded"
            'request.Timeout = 300000
            'request.ServicePoint.MaxIdleTime = 300000
            ''  request.KeepAlive = True
            'inputstring = HttpUtility.UrlEncode(MyExtensionClasses.OuterXml(input))

            'stream = request.GetRequestStream()
            ''  Dim bytes As Byte() = Encoding.Default.GetBytes(input.OuterXml)

            'Dim encoding As System.Text.UTF8Encoding = New UTF8Encoding()
            'postData = "inputXML=" + inputstring

            'data = encoding.GetBytes(postData)
            ''input.Save(stream)
            '' request.ContentLength = data.Length
            'stream.Write(data, 0, data.Length)

            '' request.Proxy = Nothing
            'ws = request.GetResponse()

            'oStream = ws.GetResponseStream()
            'Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")

            '' Pipe the stream to a higher level stream reader with the required encoding format.
            'Dim readStream As New StreamReader(oStream, encode)
            ''     Dim separators() As String = {"Genre,Viewership"}
            '' Dim file As [String]() = readStream.ReadToEnd().Split(separators, StringSplitOptions.RemoveEmptyEntries)
            Dim name As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
            input.Save(LogDirectoryPath + "ChannelShare_Inp_" + name)
            '  channelShareRootNode = GetOpXMLFromWS(input, "http://ec2-54-255-217-55.ap-southeast-1.compute.amazonaws.com:8080/GroupM/channelshare/")
            channelShareRootNode = GetOpXMLFromWS(input, Globals.Ribbons.MSprintExRibbon.GetURLForWS("ChannelShareWSURL_New"))
            ' channelShareRootNode = XElement.Parse(readStream.ReadToEnd, Xml.Linq.LoadOptions.None)
            'Parallel.ForEach(x.Elements("tg"),Sub(

            '    ds.Tables(0).DefaultView.ToTable(
            ' Next
            ' ds = New DataSet()

            If Not (channelShareRootNode Is Nothing) Then
                Dim name1 As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                channelShareRootNode.Save(LogDirectoryPath + "ChannelShare_Op_" + name1)
            End If

            If channelShareRootNode.Elements("TG_MG").Count > 0 Then

                planDs.Tables.Add("ChannelViewership")
                planDs.Tables.Add("TopTenPrograms")
                planDs.Tables(0).Columns.Add("TGroup")
                planDs.Tables(0).Columns.Add("MGroup")
                planDs.Tables(0).Columns.Add("ChannelCode", System.Type.GetType("System.Int32"))
                planDs.Tables(0).Columns.Add("Channel Name")
                planDs.Tables(0).Columns.Add("Genre")
                planDs.Tables(0).Columns.Add("GRP", System.Type.GetType("System.Decimal"))
                ' planDs.Tables(0).Columns.Add("GRPShare")
                planDs.Tables(1).Columns.Add("TGroup")
                planDs.Tables(1).Columns.Add("MGroup")
                planDs.Tables(1).Columns.Add("Rank", System.Type.GetType("System.Int32"))
                planDs.Tables(1).Columns.Add("Channel")
                planDs.Tables(1).Columns.Add("Program Name")
                planDs.Tables(1).Columns.Add("Program Start Time")
                ' ds.Tables(1).Columns.Add("Genre")
                planDs.Tables(1).Columns.Add("AvgTVR", System.Type.GetType("System.Decimal"))
                'dsRef = New DataSet()

                If reftgname.Length > 0 Then
                    refDs.Tables.Add("ChannelViewership")
                    refDs.Tables.Add("TopTenPrograms")
                    refDs.Tables(0).Columns.Add("TGroup")
                    refDs.Tables(0).Columns.Add("MGroup")
                    refDs.Tables(0).Columns.Add("ChannelCode", System.Type.GetType("System.Int32"))
                    refDs.Tables(0).Columns.Add("Channel Name")
                    refDs.Tables(0).Columns.Add("Genre")
                    refDs.Tables(0).Columns.Add("GRP", System.Type.GetType("System.Decimal"))
                    ' planDs.Tables(0).Columns.Add("GRPShare")

                    refDs.Tables(1).Columns.Add("TGroup")
                    refDs.Tables(1).Columns.Add("MGroup")
                    refDs.Tables(1).Columns.Add("Rank", System.Type.GetType("System.Int32"))
                    refDs.Tables(1).Columns.Add("Channel")
                    refDs.Tables(1).Columns.Add("Program Name")
                    refDs.Tables(1).Columns.Add("Program Start Time")
                    ' ds.Tables(1).Columns.Add("Genre")
                    refDs.Tables(1).Columns.Add("AvgTVR", System.Type.GetType("System.Decimal"))
                End If


                For Each tgelement As XElement In channelShareRootNode.Elements("TG_MG")


                    'Parallel.ForEach(x.Elements("TG_MG"), Sub(tgElement)



                    ' Dim tg As XElement = DirectCast(tgelement, XElement)

                    If tgelement.Attribute("type").Value.Equals("Planning") Then


                        planDs.Tables(0).Columns("TGroup").DefaultValue = tgelement.Attribute("name").Value.Split({"~"c})(0)
                        planDs.Tables(1).Columns("TGroup").DefaultValue = tgelement.Attribute("name").Value.Split({"~"c})(0)
                        planDs.Tables(0).Columns("MGroup").DefaultValue = tgelement.Attribute("name").Value.Split({"~"c})(1)
                        planDs.Tables(1).Columns("MGroup").DefaultValue = tgelement.Attribute("name").Value.Split({"~"c})(1)

                        If tgelement.Elements("output").Any Then


                            If tgelement.Element("output").Elements("channel_share").Any Then

                                Dim channelShareElement As XElement = tgelement.Element("output").Element("channel_share")
                                Dim channelsharetotal As Integer = 0
                                '  channelsharetotal = Convert.ToInt32(channelShareElement.Attribute("AllChannelGRP").Value)
                                For Each genreElement As XElement In channelShareElement.Elements
                                    planDs.Tables(0).Columns("Genre").DefaultValue = genreElement.Attribute("name").Value
                                    cGenre.Columns("Genre").DefaultValue = genreElement.Attribute("name").Value

                                    'For Each channel As XElement In genreElement.Elements
                                    '    channelsharetotal = channelsharetotal + Convert.ToInt32(channel.Attribute("GRP").Value)
                                    'Next
                                    For Each channel As XElement In genreElement.Elements
                                        Dim dr As DataRow = planDs.Tables(0).NewRow()
                                        dr("ChannelCode") = channel.Attribute("code").Value
                                        dr("Channel Name") = channel.Attribute("name").Value
                                        dr("GRP") = channel.Attribute("ChannelShare").Value
                                        'Try
                                        '    If channelsharetotal > 0 Then
                                        '        dr("GRP") = Convert.ToInt32(channel.Attribute("GRP").Value) / channelsharetotal * 100
                                        '    End If
                                        'Catch ex As Exception
                                        '    dr("GRP") = channel.Attribute("GRP").Value
                                        'End Try
                                        planDs.Tables(0).Rows.Add(dr)
                                        Dim drow As Data.DataRow = cGenre.NewRow()
                                        drow("Channel") = channel.Attribute("name").Value
                                        cGenre.Rows.Add(drow)
                                    Next
                                Next
                            End If


                            If tgelement.Element("output").Elements("TopTenPrograms").Any Then
                                Dim toptenPrograms As XElement = tgelement.Element("output").Element("TopTenPrograms")
                                For Each program As XElement In toptenPrograms.Elements
                                    Dim dRow As DataRow = planDs.Tables(1).NewRow()
                                    dRow("Rank") = program.Attribute("rank").Value
                                    dRow("Channel") = program.Attribute("channelname").Value
                                    dRow("Program Name") = program.Attribute("name").Value
                                    dRow("Program Start Time") = program.Attribute("start_hour").Value.Substring(0, 5)
                                    dRow("AvgTVR") = program.Attribute("avgTVR").Value
                                    planDs.Tables(1).Rows.Add(dRow)
                                Next
                            End If


                        End If
                    ElseIf tgelement.Attribute("type").Value.Equals("Reference") And reftgname.Length > 0 Then

                        refDs.Tables(0).Columns("TGroup").DefaultValue = tgelement.Attribute("name").Value.Split({"~"c})(0)
                        refDs.Tables(1).Columns("TGroup").DefaultValue = tgelement.Attribute("name").Value.Split({"~"c})(0)
                        refDs.Tables(0).Columns("MGroup").DefaultValue = tgelement.Attribute("name").Value.Split({"~"c})(1)
                        refDs.Tables(1).Columns("MGroup").DefaultValue = tgelement.Attribute("name").Value.Split({"~"c})(1)
                        If tgelement.Elements("output").Any Then


                            If tgelement.Element("output").Elements("channel_share").Any Then

                                Dim channelShareElement As XElement = tgelement.Element("output").Element("channel_share")
                                For Each genreElement As XElement In channelShareElement.Elements
                                    refDs.Tables(0).Columns("Genre").DefaultValue = genreElement.Attribute("name").Value
                                    '   Dim channelsharetotal As Integer = 0
                                    'For Each channel As XElement In genreElement.Elements
                                    '    channelsharetotal = channelsharetotal + Convert.ToInt32(channel.Attribute("GRP").Value)
                                    'Next
                                    For Each channel As XElement In genreElement.Elements
                                        Dim dr As DataRow = refDs.Tables(0).NewRow()
                                        dr("ChannelCode") = channel.Attribute("code").Value
                                        dr("Channel Name") = channel.Attribute("name").Value
                                        dr("GRP") = channel.Attribute("ChannelShare").Value
                                        'Try
                                        '    If channelsharetotal > 0 Then
                                        '        dr("GRP") = Convert.ToInt32(channel.Attribute("GRP").Value) / channelsharetotal * 100
                                        '    End If
                                        'Catch ex As Exception
                                        '    dr("GRP") = channel.Attribute("GRP").Value
                                        'End Try
                                        refDs.Tables(0).Rows.Add(dr)
                                    Next
                                Next
                            End If
                            If tgelement.Element("output").Elements("TopTenPrograms").Any Then
                                Dim toptenPrograms As XElement = tgelement.Element("output").Element("TopTenPrograms")
                                For Each program As XElement In toptenPrograms.Elements
                                    Dim dRow As DataRow = refDs.Tables(1).NewRow()
                                    dRow("Rank") = program.Attribute("rank").Value
                                    dRow("Channel") = program.Attribute("channelname").Value
                                    dRow("Program Name") = program.Attribute("name").Value
                                    dRow("Program Start Time") = program.Attribute("start_hour").Value.Substring(0, 5)
                                    dRow("AvgTVR") = Math.Round(Decimal.Parse(program.Attribute("avgTVR").Value), 2)
                                    refDs.Tables(1).Rows.Add(dRow)
                                Next
                            End If
                        End If
                    End If
                Next
                DisplayChannelShareDetailsonSheet(planDs)
                'ds.WriteXml(Path.GetTempPath() + "\\ds1.xml")
                'dsRef.WriteXml(Path.GetTempPath() + "\\dsRef1.xml")
                ''While Not (readStream.EndOfStream)
                '    Dim s As String = readStream.ReadLine()
                '    Dim values As String()

                '    If s.Contains("Market:") Then
                '        Dim v As String() = s.Split({","c}, StringSplitOptions.None)


                '        ds.Tables(0).Columns("Mgroup").DefaultValue = v(1)
                '        ds.Tables(1).Columns("Mgroup").DefaultValue = v(1)



                '    End If
                '    If s.Contains("TG:") Then
                '        Dim v As String() = s.Split({","c}, StringSplitOptions.None)

                '        ds.Tables(0).Columns("TGroup").DefaultValue = v(1)
                '        ds.Tables(1).Columns("Tgroup").DefaultValue = v(1)


                '    End If
                '    If Not (s.Contains("Market:") Or s.Contains("TG:") Or s.Contains("Top 10 Programs") Or s.Equals(" ") Or s.Contains("Rank") Or s.Contains("Genre")) Then
                '        values = s.Split(New [String]() {","c}, StringSplitOptions.None)

                '        If Not (values Is Nothing) And values.Length.Equals(4) Then

                '            'dr("Genre") = values(0)
                '            'dr("GRP") = values(1)
                '            ' ds.Tables(0).Rows.

                '            ' If ds.Tables(0).Columns(0).DefaultValue.Equals(ComboBox1.Text.Trim()) And CheckedListBox1.Items.Contains(ds.Tables(0).Columns(1).DefaultValue.ToString()) Then
                '            Dim dr As DataRow = ds.Tables(0).NewRow()
                '            dr("ChannelCode") = values(0)
                '            dr("Channel Name") = values(1)
                '            dr("Genre") = values(2)
                '            dr("GRP") = values(3)
                '            ds.Tables(0).Rows.Add(dr)


                '            ' ds.Tables(0).Rows.Add(dr)
                '        ElseIf Not (values Is Nothing) And values.Length.Equals(6) And Not (values(0).Contains("Rank")) Then
                '            ' If ds.Tables(0).Columns(0).DefaultValue.Equals(ComboBox1.Text.Trim()) And CheckedListBox1.Items.Contains(ds.Tables(0).Columns(1).DefaultValue.ToString()) Then
                '            Dim dr As DataRow = ds.Tables(1).NewRow()
                '            dr("Rank") = values(0)
                '            dr("ChannelCode") = values(1)
                '            dr("Program Name") = values(2)
                '            dr("Program Start Time") = values(3)
                '            dr("Genre") = values(4)
                '            dr("GRP") = values(5)

                '            ds.Tables(1).Rows.Add(dr)
                '            'Else
                '            '    Dim dr As DataRow = ds1.Tables(0).NewRow()
                '            '    dr("Rank") = values(0)
                '            '    dr("Channel") = values(1)
                '            '    dr("GRP") = values(2)
                '            '    ds1.Tables(0).Rows.Add(dr)
                '            'End If
                '        End If
                '    End If
                'End While

                '  Dim copyGenreTab As System.Data.DataTable = ds.Tables(0).Copy()
                ' Dim copyto10 As System.Data.DataTable = ds.Tables(1).Copy()

            Else
                System.Windows.Forms.MessageBox.Show("Unable to get requested Channel Share details from Server")
            End If
            '  End If
        Catch ex As Exception
            LogMpsrintExException("Exception occured while retreiving channel share details" + ex.Message)
            Throw ex
        End Try
    End Sub
    Private Sub GetGenreShare(ByVal plantgname As String, ByVal reftgname As String, ByVal planDs As DataSet, ByVal refDs As DataSet, ByVal viewVal As GenreShareView)
        Try

            Dim plantg As String = String.Empty
            Dim reftg As String = String.Empty
            plantg = plantgname
            '  reftg = reftgname

            Dim input As XElement = ConstructGenreShareInputXML(plantg, reftg)

            ' input.Add(tgs)
            ' input.Add(markets)
            'request = WebRequest.Create("http://ec2-54-255-217-55.ap-southeast-1.compute.amazonaws.com:8080/GroupM/genreshare/")
            ''

            'request.Method = "POST"
            'request.ContentType = "application/x-www-form-urlencoded"
            'request.Timeout = 300000
            'request.ServicePoint.MaxIdleTime = 300000
            'request.KeepAlive = True
            'inputstring = HttpUtility.UrlEncode(MyExtensionClasses.OuterXml(input))
            'stream = request.GetRequestStream()
            ''  Dim bytes As Byte() = Encoding.Default.GetBytes(input.OuterXml)

            'Dim encoding As System.Text.UTF8Encoding = New UTF8Encoding()
            'postData = "inputXML=" + inputstring

            'data = encoding.GetBytes(postData)
            ''input.Save(stream)
            '' request.ContentLength = data.Length
            'stream.Write(data, 0, data.Length)


            'ws = request.GetResponse()

            'oStream = ws.GetResponseStream()
            'Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")

            '' Pipe the stream to a higher level stream reader with the required encoding format.
            'Dim readStream As New StreamReader(oStream, encode)
            ''     Dim separators() As String = {"Genre,Viewership"}
            '' Dim file As [String]() = readStream.ReadToEnd().Split(separators, StringSplitOptions.RemoveEmptyEntries)
            Dim name As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
            input.Save(LogDirectoryPath + "GenreShare_Inp_" + name)
            ' genreShareRootNode = GetOpXMLFromWS(input, "http://ec2-54-255-217-55.ap-southeast-1.compute.amazonaws.com:8080/GroupM/genreshare/")
            genreShareRootNode = GetOpXMLFromWS(input, Globals.Ribbons.MSprintExRibbon.GetURLForWS("GenreShareWSURL_New"))
            ' genreShareRootNode = XElement.Parse(readStream.ReadToEnd, Xml.Linq.LoadOptions.None)
            'Parallel.ForEach(x.Elements("tg"),Sub(

            '    ds.Tables(0).DefaultView.ToTable(
            ' Next
            ' ds = New DataSet()
            'Dim ds1 As DataSet = New DataSet()

            If Not (genreShareRootNode Is Nothing) Then
                Dim name1 As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"

                genreShareRootNode.Save(LogDirectoryPath + "GenreShare_Op_" + name1)
            End If

            If genreShareRootNode.Elements("TG_MG").Count > 0 Then

                planDs.Tables.Add("GenreViewership")
                planDs.Tables.Add("TopTenChannels")
                planDs.Tables(0).Columns.Add("TGroup")
                planDs.Tables(0).Columns.Add("MGroup")
                planDs.Tables(0).Columns.Add("Genre")
                planDs.Tables(0).Columns.Add("GRPShare in(%)", System.Type.GetType("System.Int32"))
                '  planDs.Tables(0).Columns.Add("GRPShare")

                planDs.Tables(1).Columns.Add("TGroup")
                planDs.Tables(1).Columns.Add("MGroup")
                planDs.Tables(1).Columns.Add("Rank", System.Type.GetType("System.Int32"))
                planDs.Tables(1).Columns.Add("Channel")
                planDs.Tables(1).Columns.Add("GRP", System.Type.GetType("System.Decimal"))
                '  RefDs = New DataSet()

                If Not (refDs) Is Nothing Then
                    refDs.Tables.Add("GenreViewership")
                    refDs.Tables.Add("TopTenChannels")
                    refDs.Tables(0).Columns.Add("TGroup")
                    refDs.Tables(0).Columns.Add("MGroup")
                    refDs.Tables(0).Columns.Add("Genre")
                    refDs.Tables(0).Columns.Add("GRPShare in(%)", System.Type.GetType("System.Int32"))
                    ' refDs.Tables(0).Columns.Add("GRPShare")

                    refDs.Tables(1).Columns.Add("TGroup")
                    refDs.Tables(1).Columns.Add("MGroup")
                    refDs.Tables(1).Columns.Add("Rank", System.Type.GetType("System.Int32"))
                    refDs.Tables(1).Columns.Add("Channel")
                    refDs.Tables(1).Columns.Add("GRP", System.Type.GetType("System.Decimal"))
                End If
                For Each tgelement As XElement In genreShareRootNode.Elements("TG_MG")


                    'Parallel.ForEach(x.Elements("TG_MG"), Sub(tgElement)



                    ' Dim tg As XElement = DirectCast(tgelement, XElement)

                    If tgelement.Attribute("type").Value.Equals("Planning") Then

                        Dim str As String() = tgelement.Attribute("name").Value.Split({"~"c}, StringSplitOptions.None)
                        planDs.Tables(0).Columns("TGroup").DefaultValue = tgelement.Attribute("name").Value.Split({"~"c})(0)
                        planDs.Tables(1).Columns("TGroup").DefaultValue = tgelement.Attribute("name").Value.Split({"~"c})(0)
                        planDs.Tables(0).Columns("MGroup").DefaultValue = tgelement.Attribute("name").Value.Split({"~"c})(1)
                        planDs.Tables(1).Columns("MGroup").DefaultValue = tgelement.Attribute("name").Value.Split({"~"c})(1)

                        If tgelement.Elements("output").Any Then

                            If tgelement.Element("output").Elements("genre_share").Any Then
                                Dim genreShareElement As XElement = tgelement.Element("output").Element("genre_share")
                                Dim genregrpTotal As Integer = 0
                                For Each genre As XElement In genreShareElement.Elements
                                    genregrpTotal = genregrpTotal + genre.Attribute("GRP").Value
                                Next
                                For Each genre As XElement In genreShareElement.Elements
                                    Dim dr As DataRow = planDs.Tables(0).NewRow()
                                    dr("Genre") = genre.Attribute("name").Value
                                    ' dr("GRPShare in(%)") = genre.Attribute("GRP").Value
                                    Try
                                        If genregrpTotal > 0 Then
                                            dr("GRPShare in(%)") = Convert.ToInt32(genre.Attribute("GRP").Value) / genregrpTotal * 100
                                        End If
                                    Catch ex As Exception

                                    End Try


                                    planDs.Tables(0).Rows.Add(dr)
                                Next
                            End If

                            If tgelement.Element("output").Elements("TopTenChannels").Any Then
                                Dim toptenChannelsElement As XElement = tgelement.Element("output").Element("TopTenChannels")
                                For Each channel As XElement In toptenChannelsElement.Elements
                                    Dim dRow As DataRow = planDs.Tables(1).NewRow()
                                    dRow("Rank") = channel.Attribute("rank").Value
                                    dRow("Channel") = channel.Attribute("name").Value
                                    dRow("GRP") = channel.Attribute("GRP").Value
                                    planDs.Tables(1).Rows.Add(dRow)
                                Next
                            End If


                        End If



                    ElseIf tgelement.Attribute("type").Value.Equals("Reference") And Not (refDs) Is Nothing Then

                        refDs.Tables(0).Columns("TGroup").DefaultValue = tgelement.Attribute("name").Value.Split({"~"c})(0)
                        refDs.Tables(1).Columns("TGroup").DefaultValue = tgelement.Attribute("name").Value.Split({"~"c})(0)
                        refDs.Tables(0).Columns("MGroup").DefaultValue = tgelement.Attribute("name").Value.Split({"~"c})(1)
                        refDs.Tables(1).Columns("MGroup").DefaultValue = tgelement.Attribute("name").Value.Split({"~"c})(1)
                        If tgelement.Elements("output").Any Then

                            If tgelement.Element("output").Elements("genre_share").Any Then

                                Dim genreShareElement As XElement = tgelement.Element("output").Element("genre_share")
                                Dim genregrpTotal As Integer = 0
                                For Each genre As XElement In genreShareElement.Elements
                                    genregrpTotal = genregrpTotal + genre.Attribute("GRP").Value
                                Next
                                For Each genre As XElement In genreShareElement.Elements
                                    Dim dr As DataRow = refDs.Tables(0).NewRow()
                                    dr("Genre") = genre.Attribute("name").Value
                                    '  dr("GRPShare in(%)") = genre.Attribute("GRP").Value
                                    Try
                                        If genregrpTotal > 0 Then
                                            dr("GRPShare in(%)") = Convert.ToInt32(genre.Attribute("GRP").Value) / genregrpTotal * 100
                                        End If
                                    Catch ex As Exception

                                    End Try
                                    refDs.Tables(0).Rows.Add(dr)
                                Next
                            End If
                            If tgelement.Element("output").Elements("TopTenChannels").Any Then
                                Dim toptenChannelsElement As XElement = tgelement.Element("output").Element("TopTenChannels")
                                For Each channel As XElement In toptenChannelsElement.Elements
                                    Dim dRow As DataRow = refDs.Tables(1).NewRow()
                                    dRow("Rank") = channel.Attribute("rank").Value
                                    dRow("Channel") = channel.Attribute("name").Value
                                    dRow("GRP") = channel.Attribute("GRP").Value
                                    refDs.Tables(1).Rows.Add(dRow)
                                Next
                            End If
                        End If
                    End If
                Next
                DisplayGenreShareDetailsOnSheet(planDs, viewVal)
                'planDs.WriteXml(Path.GetTempPath() + "\\planDs.xml")
                'refDs.WriteXml(Path.GetTempPath() + "\\RefDs.xml")
                'While Not (readStream.EndOfStream)
                '    Dim s As String = readStream.ReadLine()
                '    Dim values As String()

                '    If s.Contains("Market:") Then
                '        Dim v As String() = s.Split({","c}, StringSplitOptions.None)


                '        ds.Tables(0).Columns("Mgroup").DefaultValue = v(1)
                '        ds.Tables(1).Columns("Mgroup").DefaultValue = v(1)



                '    End If
                '    If s.Contains("TG:") Then
                '        Dim v As String() = s.Split({","c}, StringSplitOptions.None)

                '        ds.Tables(0).Columns("TGroup").DefaultValue = v(1)
                '        ds.Tables(1).Columns("Tgroup").DefaultValue = v(1)


                '    End If
                '    If Not (s.Contains("Market:") Or s.Contains("TG:") Or s.Contains("Top 10 Channels") Or s.Equals(" ") Or s.Contains("Rank") Or s.Contains("Genre")) Then
                '        values = s.Split(New [String]() {","c}, StringSplitOptions.None)

                '        If Not (values Is Nothing) And values.Length.Equals(2) Then

                '            'dr("Genre") = values(0)
                '            'dr("GRP") = values(1)
                '            ' ds.Tables(0).Rows.

                '            ' If ds.Tables(0).Columns(0).DefaultValue.Equals(ComboBox1.Text.Trim()) And CheckedListBox1.Items.Contains(ds.Tables(0).Columns(1).DefaultValue.ToString()) Then
                '            Dim dr As DataRow = ds.Tables(0).NewRow()
                '            dr("Genre") = values(0)
                '            dr("GRP") = values(1)
                '            ds.Tables(0).Rows.Add(dr)


                '            ' ds.Tables(0).Rows.Add(dr)
                '        ElseIf Not (values Is Nothing) And values.Length.Equals(3) And Not (values(0).Contains("Rank")) Then
                '            ' If ds.Tables(0).Columns(0).DefaultValue.Equals(ComboBox1.Text.Trim()) And CheckedListBox1.Items.Contains(ds.Tables(0).Columns(1).DefaultValue.ToString()) Then
                '            Dim dr As DataRow = ds.Tables(1).NewRow()
                '            dr("Rank") = values(0)
                '            dr("Channel") = values(1)
                '            dr("GRP") = values(2)
                '            ds.Tables(1).Rows.Add(dr)
                '            'Else
                '            '    Dim dr As DataRow = ds1.Tables(0).NewRow()
                '            '    dr("Rank") = values(0)
                '            '    dr("Channel") = values(1)
                '            '    dr("GRP") = values(2)
                '            '    ds1.Tables(0).Rows.Add(dr)
                '            'End If
                '        End If
                '    End If

                'End While


                ' Dim genreTable As System.Data.DataTable = New System.Data.DataTable()

            Else
                System.Windows.Forms.MessageBox.Show("Unable to get requested Genre Share details from Server")

            End If
            'listobject1.QueryTable.AdjustColumnWidth = True
            ' vstoWorkbook.Controls.AddControl(System.Windows.Forms.Button,

        Catch ex As Exception
            LogMpsrintExException("Exception occured while getting requested genre share" + ex.Message)
            ' System.Windows.Forms.MessageBox.Show("Exception occured while getting requested Genre share details")
            Throw ex

        End Try
    End Sub
    Public Function CleanSheet(ByVal sheet As Microsoft.Office.Interop.Excel.Worksheet)
        Try
            sheet.UsedRange.Clear()
            Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(sheet)
            For Each lo As ListObject In vstoWorkbook.ListObjects
                lo.Delete()
            Next
            If vstoWorkbook.ChartObjects.Count > 0 Then
                vstoWorkbook.ChartObjects.Delete()
            End If
        Catch ex As Exception

        End Try
    End Function

    Private Sub btnChannelShare_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnChannelShare.Click
        'newSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
        'newSheet.Name = "Channel Share"
        ''newSheet.Cells(0, 0) = "Channel Share"
        'Dim cell As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$A$1", Type.Missing)
        'cell.Value2 = "Channel Share"
        'cell.Interior.Color = System.Drawing.Color.Yellow
        'cell.ColumnWidth = 15
        'Dim ceel As Microsoft.Office.Interop.Excel.Range = newSheet.UsedRange.Next(3, 0)
        'ceel.Value2 = "Top Ten Programs"
        'Dim cell1 As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$C$1", "$E$2")
        'Dim values As String() = {"Eval Start Date", "8/4/2013", "Week 32", "Eval end Date", "8/10/2013", "Week 32"}
        'If MSprintExChannelShare Is Nothing Then
        '    mpChannelShare = New ucChannelShare()
        '    MSprintExChannelShare = Globals.ThisAddIn.CustomTaskPanes.Add(mpChannelShare, "Channel Programs")
        '    MSprintExChannelShare.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
        'End If
        'MSprintExChannelShare.Height = 226
        'MSprintExChannelShare.Width = 271
        'MSprintExChannelShare.Visible = True

        'Dim i As Int32 = cell1.Range.Range.Rows.Count - 1
        'For index = 0 To cell1.Range.Rows.Count - 1

        'Next

        'foreach (Excel.Range row in myRange.Rows)  
        '{  
        '    Excel.Range cell = (Excel.Range)row.Cells[1, 1];  
        '    if (cell.Value2 != null)  
        '         System.Diagnostics.Debug.WriteLine(cell.Value2.ToString());  
        '} 
        Dim frm As frmWait
        Dim cGenre As Data.DataTable
        Try
            'If plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
            '    System.Windows.Forms.MessageBox.Show("Please make Target Group and Market Group(s) Selections")
            'ElseIf reftgname.Length > 0 And tpSelections.UcMarkets1.lbRef.Items.Count < 1 Then
            '    System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for reference Target Group chosen")
            'ElseIf tpSelections.UcGenres.lbSelectedGenres.Items.Count < 1 Then
            '    System.Windows.Forms.MessageBox.Show("Please choose Genre(s) to view their Channel Share details")
            'Else
            'End If
            Dim plantg As String = String.Empty
            Dim reftg As String = String.Empty
            '    '  Dim tgListBox As Windows.Forms.Control() = tpSelections.tpAudience.Controls.Find("lbTGDefs", True)
            '    '  Dim mglistBox As Windows.Forms.Control() = tpSelections.tpMarkets.Controls.Find("lbMarketGroup", True)
            'Dim tgg As Windows.Forms.ListBox
            'Dim mgg As Windows.Forms.ListBox
            'Dim form As GenreShareForm
            '    ' If TypeOf tgListBox(0) Is Windows.Forms.ListBox And (TypeOf mglistBox(0) Is Windows.Forms.ListBox) Then
            '    'tgg = DirectCast(tgListBox(0), Windows.Forms.ListBox)
            '    'mgg = DirectCast(mglistBox(0), Windows.Forms.ListBox)
            '    'tgg.Items
            '    form = New GenreShareForm(mgg.Items, tgg.Items)

            '    form.ShowDialog()

            Dim dtable As System.Data.DataTable = DirectCast(tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            plantg = dtable.Rows(0)(1).ToString().Trim()
            ' reftg = dtable.Rows(1)(1).ToString()
            If plantg.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please make Target Group and Market Group(s) Selections")
            ElseIf plantg.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count > 0 Then
                System.Windows.Forms.MessageBox.Show("Please choose Planning Target group for Market groups")
            ElseIf plantg.Length > 1 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please choose Market(s) and/or Market Group(s) for chosen Primary target group")
                'ElseIf reftg.Length > 0 And tpSelections.UcMarkets1.lbRef.Items.Count < 1 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for reference Target Group chosen")
                'ElseIf reftg.Length = 0 And tpSelections.UcMarkets1.lbRef.Items.Count > 0 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose Reference Target group for Market groups")
            ElseIf tpSelections.UcGenres.lbSelectedGenres.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please choose Genre(s) to view their Channel Share details")
            Else
                Globals.ThisAddIn.Application.StatusBar = "Getting requested Channel Share details.."
                System.Windows.Forms.Application.DoEvents()
                '   Globals.ThisAddIn.Application.DoEvents()
                '   Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlWait
                'frm = New frmWait()
                'frm.Show()
                'frm.Panel1.Refresh()
                'frm.Refresh()

                ' frm.Refresh()
                cshareds = New DataSet()

                If reftg.Length > 0 Then
                    csharerefds = New DataSet()
                End If
                cGenre = New Data.DataTable()
                cGenre.Columns.Add("Genre")
                cGenre.Columns.Add("Channel")
                GetChannelShare(plantg, reftg, cshareds, csharerefds, cGenre)
                'ds = New DataSet()
                'ds.ReadXml(Path.GetTempPath() + "\\ds1.xml")
                'Dim dsRef As DataSet = New DataSet()
                'dsRef.ReadXml(Path.GetTempPath() + "\\dsRef1.xml")

                If reftg.Length > 0 And cshareds.Tables.Count > 0 Then



                    'Dim plan, ref As List(Of String)
                    'plan = New List(Of String)()
                    'ref = New List(Of String)()
                    'For index = 0 To tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                    '    plan.Add(tpSelections.UcMarkets1.lbPlan.Items(index))
                    'Next
                    'For index = 0 To tpSelections.UcMarkets1.lbRef.Items.Count - 1
                    '    ref.Add(tpSelections.UcMarkets1.lbRef.Items(index))
                    'Next

                    If CTPChannelShare Is Nothing Then
                        mpChannelShare = New ucChannelShare()
                        CTPChannelShare = Globals.ThisAddIn.CustomTaskPanes.Add(mpChannelShare, "Channel Share Selections")
                        CTPChannelShare.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
                        CTPChannelShare.Height = 226
                        CTPChannelShare.Width = 300
                        CTPChannelShare.Visible = True
                    Else
                        CTPChannelShare.Visible = True
                        '  MSprintExChannelShare.Title = "Genre Share Selections"
                    End If
                    'If CTSChannelShare Is Nothing Then
                    '    ' CTSChannelShare.Dispose()
                    '    mpChannelShare = New ucChannelShare(planningDataSet, refDataSet, plan.ToArray(), ref.ToArray(), plantg, reftg, tpSelections, "Channel", cGenre, Nothing)
                    '    CTSChannelShare = Globals.ThisAddIn.CustomTaskPanes.Add(mpChannelShare, "Channel Share Selections")
                    '    CTSChannelShare.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
                    '    CTSChannelShare.Height = 226
                    '    CTSChannelShare.Width = 271
                    '    CTSChannelShare.Visible = True
                    'End If
                    'tpSelections.TaskPaneLogFile1.scMain.Panel2Collapsed = False
                    'If Not tpSelections.TaskPaneLogFile1.showingChannels Then
                    '    ' tpSelections.TaskPaneLogFile1.ucChannelsMapping.btnShowHide.Text = "Genre Share Selections"
                    '    tpSelections.TaskPaneLogFile1.ucChannelsMapping.btnShowHide.Text = "Channel Share Selections"
                    '    tpSelections.TaskPaneLogFile1.showChannelMapping(True)
                    '    tpSelections.TaskPaneLogFile1.ucChannelsMapping.btnShowHide.Refresh()
                    '    tpSelections.TaskPaneLogFile1.Show()
                    'End If
                End If
                Globals.ThisAddIn.Application.StatusBar = String.Empty
                ' frm.Close()
                Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault
            End If
        Catch ex As Exception
            Globals.ThisAddIn.Application.StatusBar = String.Empty

            If Not (frm Is Nothing) Then
                frm.Close()
            End If
            Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault
            LogMpsrintExException("Exception occured while retreiving Channel share details." + ex.Message)
            System.Windows.Forms.MessageBox.Show("Exception occured while retreiving requested Channel Share details.Please refer to ErrorLog for details")
        End Try
    End Sub
    Friend Sub LoadTAMChannels()
        If dtChannelMaster.Count = 0 Then

            Dim drChannelMaster As Plandata.ChannelMasterRow
            'Dim ChannelCode, ChannelName As String
            'Dim daChannelsMETIS As New METISTableAdapters.CHANNEL_MASTERTableAdapter
            'Dim dtChannelsMETIS As METIS.CHANNEL_MASTERDataTable = daChannelsMETIS.GetChannels
            For Each drChannelMETIS As Data.DataRow In Globals.Ribbons.MSprintExRibbon.dtchannels.Rows
                drChannelMaster = dtChannelMaster.NewChannelMasterRow
                drChannelMaster.ChannelCode = drChannelMETIS("ID")
                drChannelMaster.ChannelName = drChannelMETIS("Name")
                dtChannelMaster.AddChannelMasterRow(drChannelMaster)
            Next
            '    'Dim appPath As String = System.AppDomain.CurrentDomain.BaseDirectory() 'Globals.ThisWorkbook.appPath
            '    ''Using r As StreamReader = New StreamReader(appPath & "STATION.TV")
            '    'Using r As StreamReader = New StreamReader(appPath & "\TAMChannelList.log")

            '    '    Dim line As String

            '    '    line = r.ReadLine

            '    '    Do While (Not line Is Nothing)
            '    '        ChannelCode = line.Substring(0, 3)
            '    '        ChannelName = line.Substring(3).Trim
            '    '        If ChannelName.Length > 0 Then
            '    '            drChannelMaster = dtChannelMaster.NewChannelMasterRow
            '    '            drChannelMaster.ChannelCode = ChannelCode
            '    '            drChannelMaster.ChannelName = ChannelName
            '    '            dtChannelMaster.AddChannelMasterRow(drChannelMaster)
            '    '        End If
            '    '        line = r.ReadLine
            '    '    Loop
            '    'End Using

            drChannelMaster = dtChannelMaster.NewChannelMasterRow
            drChannelMaster.ChannelCode = "000"
            drChannelMaster.ChannelName = " - - Select - - "
            dtChannelMaster.AddChannelMasterRow(drChannelMaster)
            dtChannelMaster.AcceptChanges()
            dtChannelMaster.DefaultView.Sort = "ChannelName Asc"
            '   masterchannels = New Data.DataTable()
            '  masterchannels = Globals.Ribbons.MSprintExRibbon.dtchannels
            ' masterchannels.Columns(0).ColumnName = "ChannelCode"
            ' masterchannels.Columns(1).ColumnName = "ChannelName"
            'masterchannels.DefaultView.Sort = "ChannelName Asc"
            'With ChannelMasterBindingSource
            '    .DataSource = dtChannelMaster.DefaultView
            'End With
        End If
        'With ChannelMasterBindingSource
        '    .DataSource = dtChannelMaster.DefaultView
        'End With
        'logTaskPane.scMain.Panel2Collapsed = False
        'If Not logTaskPane.showingChannels Then logTaskPane.showChannelMapping(True)

    End Sub
    Friend Sub LoadPlanChannels()
        dtPlanChannels = New Plandata.PlanChannelsDataTable
        Dim drPlanChannel As Plandata.PlanChannelsRow
        '  dtPlanChannels.ChannelCodeColumn.AllowDBNull = True
        planchannels = New Data.DataTable()
        planchannels.Columns.Add("ChannelCode")
        planchannels.Columns.Add("ChannelName")
        Dim currChannelCode As String = "000"
        ChannelColumn = loSpotSelection.ListColumns("Channel")
        ChannelCells = ChannelColumn.DataBodyRange
        For Each ChannelCell As Microsoft.Office.Interop.Excel.Range In ChannelCells
            If Not SubtotalRows Is Nothing Then
                If Not loPlanData.Application.Intersect(ChannelCell, SubtotalRows) Is Nothing Then Continue For
            End If
            'Dim dr As Data.DataRow = planchannels.NewRow()
            'dr("ChannelName") = ChannelCell.Value
            If dtPlanChannels.Select("ChannelName = '" & ChannelCell.Value & "'").Length > 0 Then Continue For
            drPlanChannel = dtPlanChannels.NewPlanChannelsRow
            drPlanChannel.ChannelName = ChannelCell.Value
            currChannelCode = GetChannelCodeFromMaster(ChannelCell.Value)
            If currChannelCode = "000" Then
                currChannelCode = GetChannelCodeFromMapping(ChannelCell.Value)
            End If
            drPlanChannel.ChannelCode = currChannelCode
            ' drPlanChannel.ChannelName = "Choose"
            dtPlanChannels.AddPlanChannelsRow(drPlanChannel)
            ' planchannels.Rows.Add(dr)
        Next
        dtPlanChannels.AcceptChanges()
        dtPlanChannels.DefaultView.Sort = "ChannelName Asc"
        planchannels.DefaultView.Sort = "ChannelName ASC"
        'With PlanChannelsBindingSource
        '    .DataSource = dtPlanChannels.DefaultView
        'End With
    End Sub

    Private Function GetChannelCodeFromMapping(ByVal ChannelName As String) As String
        dtChannelMap = daChannelMap.GetChannels(ChannelName)
        If dtChannelMap.Count > 0 Then
            GetChannelCodeFromMapping = dtChannelMap(0).TAMChannelCode
        Else
            GetChannelCodeFromMapping = "000"
        End If
    End Function

    Private Function GetChannelCodeFromMaster(ByVal ChannelName As String) As String
        ' Dim drChannelMaster() As Plandata.ChannelMasterRow
        Dim drChannelMaster() As Data.DataRow
        drChannelMaster = Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name = '" & ChannelName & "'")
        If drChannelMaster.Length > 0 Then
            GetChannelCodeFromMaster = drChannelMaster(0)(0).ToString()
        Else
            GetChannelCodeFromMaster = "000"
        End If
    End Function

    Private Sub btnMapChannels_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnMapChannels.Click
        Try
            'LoadTAMChannels()
            'LoadPlanChannels()
            Dim frmSelectChannel As New frmFilterChannels
            channelMapping = New ChannelMapping()
            ChannelPane = Globals.ThisAddIn.CustomTaskPanes.Add(channelMapping, "Channel Mapping")
            'Dim CurrentChannel As System.Data.DataRowView = Globals.Ribbons.MSprintExRibbon.channelMapping.PlanChannelsBindingSource.Current
            'frmSelectChannel.CurrentChannelCode = CurrentChannel.Row.Item("ChannelCode")
            'frmSelectChannel.CurrentChannelName = CurrentChannel.Row.Item("ChannelName")
            frmSelectChannel.PlanChannelsBindingSource = Globals.Ribbons.MSprintExRibbon.channelMapping.PlanChannelsBindingSource
            If frmSelectChannel.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Dim SelectedValue As String
                SelectedValue = frmSelectChannel.lbChannelMaster.SelectedValue
                ' CurrentChannel.Row.Item("ChannelCode") = SelectedValue
            End If
            frmSelectChannel.Dispose()
        Catch ex As Exception

        End Try
        'If Not (ChannelPane Is Nothing) Then ChannelPane.Dispose()

        'ChannelPane = Globals.ThisAddIn.CustomTaskPanes.Add(channelMapping, "Channel Mapping")
        'ChannelPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
        'ChannelPane.Height = 226
        'ChannelPane.Width = 271
        'ChannelPane.Visible = True
        'Else
        'ChannelPane.Visible = True
        ''  MSprintExChannelShare.Title = "Genre Share Selections"
        'End If
        'tpSelections.TaskPaneLogFile1.scMain.Panel2Collapsed = False
        'If Not tpSelections.TaskPaneLogFile1.showingChannels Then
        '    tpSelections.TaskPaneLogFile1.ucChannelsMapping.btnShowHide.Text = "Genre Share Selections"

        '    tpSelections.TaskPaneLogFile1.showChannelMapping(True)
        'End If

        'logTaskPane.scMain.Panel2Collapsed = False
        'If Not logTaskPane.showingChannels Then logTaskPane.showChannelMapping(True)
    End Sub

    Private Sub btnCleanupplan_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnCleanupplan.Click
        CleanUpPlan()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        If Not tpAudience Is Nothing Then tpAudience.Dispose()
        tpAudience = New ucAudience
        MSprintExTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(tpAudience, "TV Plan builder")
        MSprintExTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight

        MSprintExTaskPane.Width = 300
        MSprintExTaskPane.Visible = True
    End Sub

    Private Sub btnMarkets_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)

    End Sub
    Private Sub CallingWS()
        request = WebRequest.Create("http://ec2-54-254-193-184.ap-southeast-1.compute.amazonaws.com:8080/GroupM/genreshare/1")
        '

        request.Method = "POST"
        request.ContentType = "application/x-www-form-urlencoded"
        request.Timeout = 300000
        request.ServicePoint.MaxIdleTime = 300000
        request.KeepAlive = True

        ' request.ContentType = "text/xml"
        stream = request.GetRequestStream()
        ' sreader = File.OpenText("C:\\MSprintEx\\MSprintEx\\input.xml")
        ' inputstring = sreader.ReadToEnd()
        ' sreader.Close()
        Dim input As XmlDocument = New XmlDocument()
        '  input.ImportNode(
        input.Load("C:\\MSprintEx\\MSprintEx\\input.xml")
        'inputstring = HttpUtility.UrlEncode(input.OuterXml)
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

        'Dim columnNames As [String]() = readStream.ReadLine().Split(New [Char]() {","c}, StringSplitOptions.None)
        'dt = New DataTable("Users")

        'dt = New DataTable("Genre Chart")

        'For index = 0 To columnNames.Length - 1
        '    dc = New DataColumn(columnNames.ElementAt(index))
        '    dt.Columns.Add(dc)
        'Next

        'While Not (readStream.EndOfStream)
        ' dt.Rows.Add(readStream.ReadLine().Split(New [Char]() {","c}, StringSplitOptions.None))
        Dim separators() As String = {"Genre,Viewership"}
        Dim file As [String]() = readStream.ReadToEnd().Split(separators, StringSplitOptions.RemoveEmptyEntries)
        Dim planmgtgstring, refmgtgstring As String
        Dim list As List(Of String)
        list = New List(Of String)(file)

        For index = 0 To list.Count - 1

            If list(index).Contains("mg1") And list(index).Contains("CS 15-44") Then
                ds = New DataSet("Plan mg1-CS 15-44")
                ds.Tables.Add("GenreViewerShip")
                ds.Tables(0).Columns.Add("Genre")
                ds.Tables(0).Columns.Add("Viewership", System.Type.GetType("System.Int32"))
                ds.Tables.Add("TopTenPrograms")
                ds.Tables(1).Columns.Add("Rank", System.Type.GetType("System.Int32"))
                ds.Tables(1).Columns.Add("Channel Code", System.Type.GetType("System.Int32"))
                ds.Tables(1).Columns.Add("Viewership", System.Type.GetType("System.Int32"))
                planmgtgstring = list(index)
            End If
            If list(index).Contains("mg2") And list(index).Contains("CS 15-24 A") Then
                ds1 = New DataSet("Ref mg2-CS 15-24 A")
                ds1.Tables.Add("GenreViewerShip")
                ds1.Tables(0).Columns.Add("Genre")
                ds1.Tables(0).Columns.Add("Viewership", System.Type.GetType("System.Int32"))
                ds1.Tables.Add("TopTenPrograms")
                ds1.Tables(1).Columns.Add("Rank", System.Type.GetType("System.Int32"))
                ds1.Tables(1).Columns.Add("Channel Code", System.Type.GetType("System.Int32"))
                ds1.Tables(1).Columns.Add("Viewership", System.Type.GetType("System.Int32"))
                refmgtgstring = list(index)
            End If
        Next
        '  Dim fInfo, finfo1 As FileInfo
        Dim swriter, swriter1 As StreamWriter
        'fInfo = New FileInfo("C:\\MSprintEx\\csv1.csv")
        'finfo1 = New FileInfo("C:\\MSprintEx\\csv2.csv")
        'fInfo.Create()
        'finfo1.Create()
        'fInfo.OpenWrite()
        'finfo1.OpenWrite()
        Dim planmarkets, planrows, plantopten, reftopten, refmarkets, refrows As [String]()
        ' Dim objReader As New StreamReader("c:\test.txt")
        Dim separators1() As String = {"CS 15-44", "Rank,Channel Code, Viewership"}
        '  planmgtgstring = objReader.ReadLine()
        planmarkets = planmgtgstring.Split(separators1, StringSplitOptions.RemoveEmptyEntries)
        ' planrows = planmarkets(1).Split({","c}, StringSplitOptions.RemoveEmptyEntries)
        ' plantopten = planmarkets(2).Split({","c}, StringSplitOptions.RemoveEmptyEntries)
        Dim refseperators() As String = {"CS 15-24 A", "Rank,Channel Code, Viewership"}
        refmarkets = refmgtgstring.Split(refseperators, StringSplitOptions.RemoveEmptyEntries)
        refrows = refmarkets(1).Split({","c}, StringSplitOptions.RemoveEmptyEntries)
        reftopten = refmarkets(2).Split({","c}, StringSplitOptions.RemoveEmptyEntries)
        swriter = New StreamWriter("C:\\MSprintEx\\csv1.csv")
        swriter1 = New StreamWriter("C:\\MSprintEx\\csv2.csv")
        swriter.Write(planmarkets(1))
        swriter1.Write(planmarkets(2))
        swriter.Close()
        swriter1.Close()
        Using reader As New CsvFileReader("C:\\MSprintEx\\csv1.csv")
            Dim row As New CsvRow()
            While reader.ReadRow(row)
                For Each s As String In row
                    Console.Write(s)
                    Console.Write(" ")
                Next
                Console.WriteLine()
            End While
        End Using

        'For index = 0 To planrows.Length - 1
        '    Dim row As DataRow = ds.Tables(0).NewRow()

        '    If Not (planrows(index).Contains("Top 10 Channels")) Then
        '        row(index) = planrows(index)
        '        ds.Tables(0).Rows.Add(row)
        '    End If

        'Next

        'For index = 0 To plantopten.Length - 1
        '    Dim row As DataRow = ds.Tables(1).NewRow()
        '    row(index) = plantopten(index)
        '    ds.Tables(1).Rows.Add(row)

        'Next
        'For index = 0 To refrows.Length - 1
        '    Dim row As DataRow = ds1.Tables(0).NewRow()
        '    row(index) = refrows(index)
        '    ds1.Tables(0).Rows.Add(row)

        'Next
        'For index = 0 To reftopten.Length - 1
        '    Dim row As DataRow = ds1.Tables(1).NewRow()
        '    row(index) = reftopten(index)
        '    ds1.Tables(1).Rows.Add(row)

        'Next
    End Sub

    Private Sub btnProgramTVR_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnProgramTVR.Click
        ' Globals.ThisAddIn.Application.StatusBar = "Getting requested Channel Share details..."
        Dim frm As frmWait
        Try


            Dim plantgname As String = String.Empty
            Dim reftgname As String = String.Empty
            Dim dtable As System.Data.DataTable = DirectCast(tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            plantgname = dtable.Rows(0)(1).ToString().Trim()
            '  reftgname = dtable.Rows(1)(1).ToString()
            'Dim chds As DataSet = New DataSet()

            If plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please make Target Group and Market Group(s) Selections")
            ElseIf plantgname.Length > 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please choose Market(s) and/or Market Group(s) for chosen Primary target group")
            ElseIf plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count > 0 Then
                System.Windows.Forms.MessageBox.Show("Please choose Planning Target group for Market groups")
                'ElseIf reftgname.Length > 0 And tpSelections.UcMarkets1.lbRef.Items.Count < 1 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for reference Target Group chosen")
                'ElseIf reftgname.Length = 0 And tpSelections.UcMarkets1.lbRef.Items.Count > 0 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose Reference Target group for Market groups")
            ElseIf tpSelections.UcChannels.lbSelectedChannels.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please choose Channel(s) to view their Program TVR details")
            Else

                System.Windows.Forms.Application.DoEvents()
                'chds.ReadXml(Path.GetTempPath() + "ds1.xml")
                'Dim tvrform As TVRForm = New TVRForm(chds, tpSelections)
                'tvrform.ShowDialog()
                'If TVRTaskPane Is Nothing Then
                '    mpUcTVRScreen = New ucTVRScreen(chds, tpSelections)
                '    TVRTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(mpUcTVRScreen, "Program TVR Selections")
                '    TVRTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
                'End If
                'TVRTaskPane.Height = 226
                'TVRTaskPane.Width = 271
                'TVRTaskPane.Visible = True
                Globals.ThisAddIn.Application.StatusBar = "Getting requested Program TVR details..."
                'Globals.ThisAddIn.Application.DoEvents()
                'frm = New frmWait()
                'frm.Show()
                'frm.Panel1.Refresh()
                'frm.Refresh()
                '  Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlWait
                Dim input As XElement = ConstructInputXMLForProgTVR()
                'input.Add(tgs)
                'input.Add(markets)
                '  Dim genrelist As XElement = New XElement("genre_list")

                '  input.Add(genrelist)

                'request = WebRequest.Create("http://ec2-54-255-217-55.ap-southeast-1.compute.amazonaws.com:8080/GroupM/programtvr/")
                ''

                'request.Method = "POST"
                'request.ContentType = "application/x-www-form-urlencoded"
                'request.Timeout = 300000
                'request.ServicePoint.MaxIdleTime = 300000
                ''  request.KeepAlive = True
                'inputstring = HttpUtility.UrlEncode(MyExtensionClasses.OuterXml(input))

                'stream = request.GetRequestStream()
                ''  Dim bytes As Byte() = Encoding.Default.GetBytes(input.OuterXml)

                'Dim encoding As System.Text.UTF8Encoding = New UTF8Encoding()
                'postData = "inputXML=" + inputstring

                'data = encoding.GetBytes(postData)
                ''input.Save(stream)
                '' request.ContentLength = data.Length
                'stream.Write(data, 0, data.Length)

                '' request.Proxy = Nothing
                'ws = request.GetResponse()

                'oStream = ws.GetResponseStream()
                'Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")

                '' Pipe the stream to a higher level stream reader with the required encoding format.
                'Dim readStream As New StreamReader(oStream, encode)
                ''   Using reader As StreamReader = New StreamReader("C:\\ws\\Ptvr.xml.xml")
                'Dim strr As String = readStream.ReadToEnd()
                Dim name As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                input.Save(LogDirectoryPath + "ProgramTVR_Inp_" + name)
                ' ptvrRootNode = GetOpXMLFromWS(input, "http://ec2-54-255-217-55.ap-southeast-1.compute.amazonaws.com:8080/GroupM/programtvr/")
                ptvrRootNode = GetOpXMLFromWS(input, Globals.Ribbons.MSprintExRibbon.GetURLForWS("BPTVRWSURL_New"))
                ' ptvrRootNode = XElement.Parse(strr, Xml.Linq.LoadOptions.None)
                'Parallel.ForEach(x.Elements("tg"),Sub(

                If Not (ptvrRootNode Is Nothing) Then
                    Dim name1 As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                    ptvrRootNode.Save(LogDirectoryPath + "ProgramTVR_Op_" + name1)
                End If

                If ptvrRootNode.Elements("TG_MG").Count > 0 Then

                    '    ds.Tables(0).DefaultView.ToTable(
                    ' Next
                    ptvrds = New DataSet("Planning")
                    ptvrrefds = New DataSet("Planning")
                    Dim tabcount As Integer = 0
                    Dim reftabcount As Integer = 0
                    channels = New System.Data.DataTable()
                    For Each tgelement As XElement In ptvrRootNode.Elements("TG_MG")

                        'Parallel.ForEach(x.Elements("TG_MG"), Sub(tgElement)



                        ' Dim tg As XElement = DirectCast(tgelement, XElement)

                        If tgelement.Attribute("type").Value.Equals("Planning") Then


                            ptvrds.Tables.Add(tgelement.Attribute("name").Value)
                            ptvrds.Tables(tabcount).Columns.Add("ChannelName")
                            ptvrds.Tables(tabcount).Columns.Add("PeriodStartDate")
                            ptvrds.Tables(tabcount).Columns.Add("PeriodEndDate")
                            ptvrds.Tables(tabcount).Columns.Add("Rank")
                            ptvrds.Tables(tabcount).Columns.Add("Programme")
                            ptvrds.Tables(tabcount).Columns.Add("Day")
                            ptvrds.Tables(tabcount).Columns.Add("Time")
                            ptvrds.Tables(tabcount).Columns.Add("TVR", System.Type.GetType("System.Decimal"))

                            If tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                                ptvrds.Tables(tabcount).Columns.Add("AverageTVR", System.Type.GetType("System.Decimal"))
                                ptvrds.Tables(tabcount).Columns.Add("Variance(in %)", System.Type.GetType("System.Decimal"))
                            End If


                            '   Dim dr As DataRow = planningDataSet.Tables(0).NewRow()



                            If Not channels.Columns.Contains("CName") Then
                                channels.Columns.Add("CName")
                            End If

                            If Not channels.Columns.Contains("StartDate") Then
                                channels.Columns.Add("StartDate")
                            End If

                            If Not channels.Columns.Contains("EndDate") Then
                                channels.Columns.Add("EndDate")
                            End If

                            'channels.Columns.Add("StartDate")
                            'channels.Columns.Add("EndDate")

                            If tgelement.Elements("output").Any Then

                                Dim outputElement As XElement = tgelement.Element("output")

                                For Each channelElement As XElement In outputElement.Elements("channel")


                                    'Parallel.ForEach(outputElement.Elements("channel"), Sub(channelElement)
                                    Dim chElement As XElement = DirectCast(channelElement, XElement)
                                    ptvrds.Tables(tabcount).Columns("ChannelName").DefaultValue = chElement.Attribute("name").Value


                                    channels.Columns("CName").DefaultValue = chElement.Attribute("name").Value
                                    For Each periodElement As XElement In chElement.Elements("period")


                                        'Parallel.ForEach(chElement.Elements("period"), Sub(periodElement)
                                        Dim prElement As XElement = DirectCast(periodElement, XElement)
                                        ptvrds.Tables(tabcount).Columns("PeriodStartDate").DefaultValue = prElement.Attribute("startdate").Value
                                        ptvrds.Tables(tabcount).Columns("PeriodEndDate").DefaultValue = prElement.Attribute("enddate").Value
                                        Dim drow As DataRow = channels.NewRow()
                                        drow("StartDate") = prElement.Attribute("startdate").Value
                                        drow("EndDate") = prElement.Attribute("enddate").Value
                                        channels.Rows.Add(drow)
                                        channels.AcceptChanges()
                                        For Each programElement As XElement In prElement.Elements("program")

                                            'Parallel.ForEach(prElement.Elements("program"), Sub(programElement)
                                            Dim progElement As XElement = DirectCast(programElement, XElement)
                                            Dim dr As DataRow = ptvrds.Tables(tabcount).NewRow()
                                            dr("Rank") = progElement.Attribute("rank").Value
                                            dr("Programme") = progElement.Attribute("name").Value
                                            dr("Day") = progElement.Attribute("day").Value
                                            dr("Time") = progElement.Attribute("time").Value.Substring(0, 5)
                                            dr("TVR") = progElement.Attribute("tvr").Value

                                            If ptvrds.Tables(tabcount).Columns.Contains("AverageTVR") Then
                                                dr("AverageTVR") = Convert.ToDecimal(progElement.Attribute("averagetvr").Value)
                                            End If
                                            ', "(AverageTVR - TVR)/TVR *100"

                                            'If ptvrds.Tables(tabcount).Columns.Contains("Variance(in %)") Then
                                            '    Try
                                            '        dr("Variance(in %)") = (Convert.ToDecimal(dr("AverageTVR").ToString()) - Convert.ToDecimal(dr("TVR").ToString())) / Convert.ToDecimal(dr("TVR").ToString()) * 100
                                            '    Catch ex As Exception
                                            '        dr("Variance(in %)") = 0
                                            '    End Try

                                            'End If
                                            If ptvrds.Tables(tabcount).Columns.Contains("Variance(in %)") Then
                                                dr("Variance(in %)") = Convert.ToDecimal(progElement.Attribute("variance").Value)
                                            End If
                                            ptvrds.Tables(tabcount).Rows.Add(dr)
                                            ptvrds.Tables(tabcount).AcceptChanges()

                                        Next    '   End Sub)
                                    Next  'End Sub)
                                Next ' End Sub)
                            End If

                            tabcount += 1
                        ElseIf (tgelement.Attribute("type").Value.Equals("Reference")) Then


                            ptvrrefds.Tables.Add(tgelement.Attribute("name").Value)
                            ptvrrefds.Tables(reftabcount).Columns.Add("ChannelName")
                            ptvrrefds.Tables(reftabcount).Columns.Add("PeriodStartDate")
                            ptvrrefds.Tables(reftabcount).Columns.Add("PeriodEndDate")
                            ptvrrefds.Tables(reftabcount).Columns.Add("Rank")
                            ptvrrefds.Tables(reftabcount).Columns.Add("Programme")
                            ptvrrefds.Tables(reftabcount).Columns.Add("Day")
                            ptvrrefds.Tables(reftabcount).Columns.Add("Time")
                            ptvrrefds.Tables(reftabcount).Columns.Add("TVR", System.Type.GetType("System.Decimal"))
                            If tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                                ptvrrefds.Tables(reftabcount).Columns.Add("AverageTVR", System.Type.GetType("System.Decimal"))
                                ptvrrefds.Tables(reftabcount).Columns.Add("Variance(in %)", System.Type.GetType("System.Decimal"))
                            End If
                            '   Dim dr As DataRow = ptvrds.Tables(0).NewRow()


                            If tgelement.Elements("output").Any Then

                                Dim outputElement As XElement = tgelement.Element("output")

                                For Each channelElement As XElement In outputElement.Elements("channel")


                                    'Parallel.ForEach(outputElement.Elements("channel"), Sub(channelElement)
                                    Dim chElement As XElement = DirectCast(channelElement, XElement)
                                    ptvrrefds.Tables(reftabcount).Columns("ChannelName").DefaultValue = chElement.Attribute("name").Value
                                    For Each periodElement As XElement In chElement.Elements("period")


                                        'Parallel.ForEach(chElement.Elements("period"), Sub(periodElement)
                                        Dim prElement As XElement = DirectCast(periodElement, XElement)
                                        ptvrrefds.Tables(reftabcount).Columns("PeriodStartDate").DefaultValue = prElement.Attribute("startdate").Value
                                        ptvrrefds.Tables(reftabcount).Columns("PeriodEndDate").DefaultValue = prElement.Attribute("enddate").Value
                                        For Each programElement As XElement In prElement.Elements("program")

                                            'Parallel.ForEach(prElement.Elements("program"), Sub(programElement)
                                            Dim progElement As XElement = DirectCast(programElement, XElement)
                                            Dim dr As DataRow = ptvrrefds.Tables(reftabcount).NewRow()
                                            dr("Rank") = progElement.Attribute("rank").Value
                                            dr("Programme") = progElement.Attribute("name").Value
                                            dr("Day") = progElement.Attribute("day").Value
                                            dr("Time") = progElement.Attribute("time").Value.Substring(0, 5)
                                            dr("TVR") = progElement.Attribute("tvr").Value
                                            If ptvrrefds.Tables(reftabcount).Columns.Contains("AverageTVR") Then
                                                dr("AverageTVR") = Convert.ToDecimal(progElement.Attribute("averagetvr").Value)
                                            End If
                                            'If ptvrrefds.Tables(reftabcount).Columns.Contains("Variance(in %)") Then
                                            '    Try
                                            '        dr("Variance(in %)") = (Convert.ToDecimal(dr("AverageTVR").ToString()) - Convert.ToDecimal(dr("TVR").ToString())) / Convert.ToDecimal(dr("TVR").ToString()) * 100
                                            '    Catch ex As Exception
                                            '        dr("Variance(in %)") = 0
                                            '    End Try

                                            'End If
                                            If ptvrrefds.Tables(reftabcount).Columns.Contains("Variance(in %)") Then
                                                dr("Variance(in %)") = Convert.ToDecimal(progElement.Attribute("variance").Value)
                                            End If
                                            ptvrrefds.Tables(reftabcount).Rows.Add(dr)
                                            ptvrrefds.Tables(reftabcount).AcceptChanges()

                                        Next    '   End Sub)
                                    Next  'End Sub)
                                Next ' End Sub)
                            End If
                            reftabcount += 1
                        End If
                        ' planningDataSet.Tables.Item(

                        '  tabcount += 1
                    Next ' End Sub)
                    channels = channels.DefaultView.ToTable(True, New String() {"CName", "StartDate", "EndDate"})
                    DisplayProgTVRDetailsOnSheet(ptvrds)
                    ' x.LoadXml(strr)
                    '   End Using
                    ' newSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)


                    'Globals.ThisAddIn.Application.StatusBar = String.Empty
                    'For index = 0 To genreTables.Length - 1
                    '    '4,1 - 4,4 ;4,6 - 4,9;4,11- 4,14

                    '    Dim lrange As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(4, ii), vstoWorkbook.Cells(4, ii + 3)), Microsoft.Office.Interop.Excel.Range)
                    '    Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "list1" + index.ToString())
                    '    listobject.AutoSetDataBoundColumnHeaders = True
                    '    listobject.DataSource = genreTables(index)
                    '    ii += 5
                    '    rocount = rocount + listobject.ListRows.Count
                    'Next
                Else
                    System.Windows.Forms.MessageBox.Show("Unable to get the requested Program TVR details from server")
                End If
                '' Dim row11 As Integer = Globals.ThisAddIn.Application.ActiveCell.Row + 2
                'Dim cell5 As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(7 + rocount, 1), vstoWorkbook.Cells(7 + rocount, 1)), Microsoft.Office.Interop.Excel.Range)
                'listobject1 = vstoWorkbook.Controls.AddListObject(cell5, "list2")
                'listobject1.AutoSetDataBoundColumnHeaders = True
                '' listobject1.Range.Columns.AutoFit()
                'listobject1.DataSource = dt1
                If reftgname.Length > 0 And ptvrds.Tables.Count > 0 Then


                    If CTPProgTVR Is Nothing Then
                        mpChannelShare = New ucChannelShare()
                        CTPProgTVR = Globals.ThisAddIn.CustomTaskPanes.Add(mpChannelShare, "Program TVR Selections")
                        CTPProgTVR.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
                        CTPProgTVR.Height = 226
                        CTPProgTVR.Width = 300
                        CTPProgTVR.Visible = True
                    Else
                        CTPProgTVR.Visible = True
                        '  MSprintExChannelShare.Title = "Genre Share Selections"
                    End If
                    'Dim plan, ref As List(Of String)
                    'plan = New List(Of String)()
                    'ref = New List(Of String)()
                    'For index = 0 To tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                    '    plan.Add(tpSelections.UcMarkets1.lbPlan.Items(index))
                    'Next
                    'For index = 0 To tpSelections.UcMarkets1.lbRef.Items.Count - 1
                    '    ref.Add(tpSelections.UcMarkets1.lbRef.Items(index))
                    'Next


                    'If CTPProgTVR Is Nothing Then
                    '    ' CTSChannelShare.Dispose()
                    '    mpChannelShare = New ucChannelShare(planningDataSet, ptvrrefds, plan.ToArray(), ref.ToArray(), plantgname, reftgname, tpSelections, "Program", Nothing, channels)
                    '    CTPProgTVR = Globals.ThisAddIn.CustomTaskPanes.Add(mpChannelShare, "Program TVR Selections")
                    '    CTPProgTVR.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
                    '    CTPProgTVR.Height = 226
                    '    CTPProgTVR.Width = 271
                    '    CTPProgTVR.Visible = True
                    'End If
                    'tpSelections.TaskPaneLogFile1.scMain.Panel2Collapsed = False
                    'If Not tpSelections.TaskPaneLogFile1.showingChannels Then
                    '    ' tpSelections.TaskPaneLogFile1.ucChannelsMapping.btnShowHide.Text = "Genre Share Selections"
                    '    tpSelections.TaskPaneLogFile1.ucChannelsMapping.btnShowHide.Text = "Program TVR Selections"
                    '    tpSelections.TaskPaneLogFile1.ucChannelsMapping.btnShowHide.Refresh()
                    '    tpSelections.TaskPaneLogFile1.showChannelMapping(True)
                    '    tpSelections.TaskPaneLogFile1.Show()
                    'End If
                End If
                Globals.ThisAddIn.Application.StatusBar = String.Empty
                Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault
            End If
            If Not (frm Is Nothing) Then
                frm.Close()
            End If
        Catch ex As Exception
            Globals.ThisAddIn.Application.StatusBar = String.Empty
            Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault

            If Not (frm Is Nothing) Then
                frm.Close()
            End If
            LogMpsrintExException("Exception occured while retreiving Program TVR details" + ex.Message)
            MessageBox.Show("Exception occured while retreiving Program TVR details.Please view error log for more details")

        End Try
    End Sub
    Private Sub btnBreakTVR_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnBreakTVR.Click
        ' Globals.ThisAddIn.Application.StatusBar = "Getting requested Channel Share details..."
        Dim frm As frmWait
        Try
            Dim plantgname As String = String.Empty
            Dim reftgname As String = String.Empty
            Dim dtable As System.Data.DataTable = DirectCast(tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            plantgname = dtable.Rows(0)(1).ToString().Trim()
            '  reftgname = dtable.Rows(1)(1).ToString()
            'Dim chds As DataSet = New DataSet()

            If plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please make Target Group and Market Group(s) Selections")
            ElseIf plantgname.Length > 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please choose Market(s) and/or Market Group(s) for chosen Primary target group")
            ElseIf plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count > 0 Then
                System.Windows.Forms.MessageBox.Show("Please choose Planning Target group for Market groups")
                'ElseIf reftgname.Length > 0 And tpSelections.UcMarkets1.lbRef.Items.Count < 1 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for reference Target Group chosen")
                'ElseIf reftgname.Length = 0 And tpSelections.UcMarkets1.lbRef.Items.Count > 0 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose Reference Target group for Market groups")
            ElseIf tpSelections.UcChannels.lbSelectedChannels.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please choose Channel(s) to view their Program TVR details")
            Else
                System.Windows.Forms.Application.DoEvents()
                Dim input As XElement = ConstructInputXMLForBreakTVR()
                'Globals.ThisAddIn.Application.DoEvents()
                SetWaitCursor("Getting requested Break TVR details..")
                'input.Add(tgs)
                'input.Add(markets)
                '  Dim genrelist As XElement = New XElement("genre_list")

                '  input.Add(genrelist)

                'request = WebRequest.Create("http://ec2-54-255-217-55.ap-southeast-1.compute.amazonaws.com:8080/GroupM/programtvr/")
                ''

                'request.Method = "POST"
                'request.ContentType = "application/x-www-form-urlencoded"
                'request.Timeout = 300000
                'request.ServicePoint.MaxIdleTime = 300000
                ''  request.KeepAlive = True
                'inputstring = HttpUtility.UrlEncode(MyExtensionClasses.OuterXml(Input))

                'stream = request.GetRequestStream()
                ''  Dim bytes As Byte() = Encoding.Default.GetBytes(input.OuterXml)

                'Dim encoding As System.Text.UTF8Encoding = New UTF8Encoding()
                'postData = "inputXML=" + inputstring

                'data = encoding.GetBytes(postData)
                ''input.Save(stream)
                '' request.ContentLength = data.Length
                'stream.Write(data, 0, data.Length)

                '' request.Proxy = Nothing
                'ws = request.GetResponse()

                'oStream = ws.GetResponseStream()
                'Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")

                '' Pipe the stream to a higher level stream reader with the required encoding format.
                'Dim readStream As New StreamReader(oStream, encode)
                ''   Using reader As StreamReader = New StreamReader("C:\\ws\\Ptvr.xml.xml")
                'Dim strr As String = readStream.ReadToEnd()
                ' ptvrRootNode = XElement.Parse(strr, Xml.Linq.LoadOptions.None)
                Dim name As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                input.Save(LogDirectoryPath + "BreakTVR_Inp_" + name)
                ' ptvrRootNode = GetOpXMLFromWS(input, "http://ec2-54-255-217-55.ap-southeast-1.compute.amazonaws.com:8080/GroupM/programtvr/")
                ptvrRootNode = GetOpXMLFromWS(input, Globals.Ribbons.MSprintExRibbon.GetURLForWS("BPTVRWSURL_New"))
                'Parallel.ForEach(x.Elements("tg"),Sub(

                '    ds.Tables(0).DefaultView.ToTable(
                ' Next

                If ptvrRootNode Is Nothing Then
                    MessageBox.Show("Unable to retreive Break TVR details from Server.")
                ElseIf ptvrRootNode.Elements("TG_MG").Count > 0 Then
                    Dim name1 As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                    ptvrRootNode.Save(LogDirectoryPath + "BreakTVR_Op_" + name1)
                    btvrds = New DataSet("Planning")
                    btvrrefds = New DataSet("Planning")
                    Dim tabcount As Integer = 0
                    Dim reftabcount As Integer = 0
                    channels = New System.Data.DataTable()
                    channels.Columns.Add("CName")
                    channels.Columns.Add("StartDate")
                    channels.Columns.Add("EndDate")
                    For Each tgelement As XElement In ptvrRootNode.Elements("TG_MG")

                        'Parallel.ForEach(x.Elements("TG_MG"), Sub(tgElement)



                        ' Dim tg As XElement = DirectCast(tgelement, XElement)

                        If tgelement.Attribute("type").Value.Equals("Planning") Then


                            btvrds.Tables.Add(tgelement.Attribute("name").Value)
                            btvrds.Tables(tabcount).Columns.Add("ChannelName")
                            btvrds.Tables(tabcount).Columns.Add("PeriodStartDate")
                            btvrds.Tables(tabcount).Columns.Add("PeriodEndDate")
                            btvrds.Tables(tabcount).Columns.Add("Rank")
                            btvrds.Tables(tabcount).Columns.Add("Programme")
                            btvrds.Tables(tabcount).Columns.Add("Day")
                            btvrds.Tables(tabcount).Columns.Add("Time")
                            btvrds.Tables(tabcount).Columns.Add("TVR", System.Type.GetType("System.Decimal"))

                            If tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                                btvrds.Tables(tabcount).Columns.Add("AverageTVR", System.Type.GetType("System.Decimal"))
                                btvrds.Tables(tabcount).Columns.Add("Variance(in %)", System.Type.GetType("System.Decimal"))
                            End If


                            '   Dim dr As DataRow = planningDataSet.Tables(0).NewRow()



                            If tgelement.Elements("output").Any Then

                                Dim outputElement As XElement = tgelement.Element("output")

                                For Each channelElement As XElement In outputElement.Elements("channel")


                                    'Parallel.ForEach(outputElement.Elements("channel"), Sub(channelElement)
                                    Dim chElement As XElement = DirectCast(channelElement, XElement)
                                    btvrds.Tables(tabcount).Columns("ChannelName").DefaultValue = chElement.Attribute("name").Value


                                    channels.Columns("CName").DefaultValue = chElement.Attribute("name").Value
                                    For Each periodElement As XElement In chElement.Elements("period")


                                        'Parallel.ForEach(chElement.Elements("period"), Sub(periodElement)
                                        Dim prElement As XElement = DirectCast(periodElement, XElement)
                                        btvrds.Tables(tabcount).Columns("PeriodStartDate").DefaultValue = prElement.Attribute("startdate").Value
                                        btvrds.Tables(tabcount).Columns("PeriodEndDate").DefaultValue = prElement.Attribute("enddate").Value
                                        Dim drow As DataRow = channels.NewRow()
                                        drow("StartDate") = prElement.Attribute("startdate").Value
                                        drow("EndDate") = prElement.Attribute("enddate").Value
                                        channels.Rows.Add(drow)
                                        For Each programElement As XElement In prElement.Elements("program")

                                            'Parallel.ForEach(prElement.Elements("program"), Sub(programElement)
                                            Dim progElement As XElement = DirectCast(programElement, XElement)
                                            Dim dr As DataRow = btvrds.Tables(tabcount).NewRow()
                                            dr("Rank") = progElement.Attribute("rank").Value
                                            dr("Programme") = progElement.Attribute("name").Value
                                            dr("Day") = progElement.Attribute("day").Value
                                            dr("Time") = progElement.Attribute("time").Value.Substring(0, 5)
                                            dr("TVR") = progElement.Attribute("tvr").Value

                                            If btvrds.Tables(tabcount).Columns.Contains("AverageTVR") Then
                                                dr("AverageTVR") = Convert.ToDecimal(progElement.Attribute("averagetvr").Value)
                                            End If

                                            If btvrds.Tables(tabcount).Columns.Contains("Variance(in %)") Then
                                                dr("Variance(in %)") = Convert.ToDecimal(progElement.Attribute("variance").Value)
                                            End If

                                            btvrds.Tables(tabcount).Rows.Add(dr)
                                            btvrds.Tables(tabcount).AcceptChanges()

                                        Next    '   End Sub)
                                    Next  'End Sub)
                                Next ' End Sub)
                            End If

                            tabcount += 1
                        ElseIf (tgelement.Attribute("type").Value.Equals("Reference")) Then


                            btvrrefds.Tables.Add(tgelement.Attribute("name").Value)
                            btvrrefds.Tables(reftabcount).Columns.Add("ChannelName")
                            btvrrefds.Tables(reftabcount).Columns.Add("PeriodStartDate")
                            btvrrefds.Tables(reftabcount).Columns.Add("PeriodEndDate")
                            btvrrefds.Tables(reftabcount).Columns.Add("Rank")
                            btvrrefds.Tables(reftabcount).Columns.Add("Programme")
                            btvrrefds.Tables(reftabcount).Columns.Add("Day")
                            btvrrefds.Tables(reftabcount).Columns.Add("Time")
                            btvrrefds.Tables(reftabcount).Columns.Add("TVR", System.Type.GetType("System.Decimal"))
                            If tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                                btvrrefds.Tables(reftabcount).Columns.Add("AverageTVR", System.Type.GetType("System.Decimal"))
                                btvrrefds.Tables(reftabcount).Columns.Add("Variance(in %)", System.Type.GetType("System.Decimal"))
                            End If
                            '   Dim dr As DataRow = planningDataSet.Tables(0).NewRow()


                            If tgelement.Elements("output").Any Then
                                Dim outputElement As XElement = tgelement.Element("output")


                                For Each channelElement As XElement In outputElement.Elements("channel")


                                    'Parallel.ForEach(outputElement.Elements("channel"), Sub(channelElement)
                                    Dim chElement As XElement = DirectCast(channelElement, XElement)
                                    btvrrefds.Tables(reftabcount).Columns("ChannelName").DefaultValue = chElement.Attribute("name").Value
                                    For Each periodElement As XElement In chElement.Elements("period")


                                        'Parallel.ForEach(chElement.Elements("period"), Sub(periodElement)
                                        Dim prElement As XElement = DirectCast(periodElement, XElement)
                                        btvrrefds.Tables(reftabcount).Columns("PeriodStartDate").DefaultValue = prElement.Attribute("startdate").Value
                                        btvrrefds.Tables(reftabcount).Columns("PeriodEndDate").DefaultValue = prElement.Attribute("enddate").Value
                                        For Each programElement As XElement In prElement.Elements("program")

                                            'Parallel.ForEach(prElement.Elements("program"), Sub(programElement)
                                            Dim progElement As XElement = DirectCast(programElement, XElement)
                                            Dim dr As DataRow = btvrrefds.Tables(reftabcount).NewRow()
                                            dr("Rank") = progElement.Attribute("rank").Value
                                            dr("Programme") = progElement.Attribute("name").Value
                                            dr("Day") = progElement.Attribute("day").Value
                                            dr("Time") = progElement.Attribute("time").Value.Substring(0, 5)
                                            dr("TVR") = progElement.Attribute("tvr").Value
                                            If btvrrefds.Tables(reftabcount).Columns.Contains("AverageTVR") Then
                                                dr("AverageTVR") = Convert.ToDecimal(progElement.Attribute("averagetvr").Value)
                                            End If
                                            If btvrrefds.Tables(reftabcount).Columns.Contains("Variance(in %)") Then
                                                dr("Variance(in %)") = Convert.ToDecimal(progElement.Attribute("variance").Value)
                                            End If
                                            btvrrefds.Tables(reftabcount).Rows.Add(dr)
                                            btvrrefds.Tables(reftabcount).AcceptChanges()

                                        Next    '   End Sub)
                                    Next  'End Sub)
                                Next
                            End If
                            ' End Sub)
                            reftabcount += 1
                        End If
                        ' btvrds.Tables.Item(

                        '  tabcount += 1
                    Next ' End Sub)


                    ' x.LoadXml(strr)
                    '   End Using
                    channels = channels.DefaultView.ToTable(True, New String() {"CName", "StartDate", "EndDate"})
                    DisplayBreakTVRDetailsOnSheet(btvrds)

                    'Globals.ThisAddIn.Application.StatusBar = String.Empty
                    'For index = 0 To genreTables.Length - 1
                    '    '4,1 - 4,4 ;4,6 - 4,9;4,11- 4,14

                    '    Dim lrange As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(4, ii), vstoWorkbook.Cells(4, ii + 3)), Microsoft.Office.Interop.Excel.Range)
                    '    Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "list1" + index.ToString())
                    '    listobject.AutoSetDataBoundColumnHeaders = True
                    '    listobject.DataSource = genreTables(index)
                    '    ii += 5
                    '    rocount = rocount + listobject.ListRows.Count
                    'Next


                    '' Dim row11 As Integer = Globals.ThisAddIn.Application.ActiveCell.Row + 2
                    'Dim cell5 As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(7 + rocount, 1), vstoWorkbook.Cells(7 + rocount, 1)), Microsoft.Office.Interop.Excel.Range)
                    'listobject1 = vstoWorkbook.Controls.AddListObject(cell5, "list2")
                    'listobject1.AutoSetDataBoundColumnHeaders = True
                    '' listobject1.Range.Columns.AutoFit()
                    'listobject1.DataSource = dt1
                    If reftgname.Length > 1 Then

                        If CTPBreakTVR Is Nothing Then
                            mpChannelShare = New ucChannelShare()
                            CTPBreakTVR = Globals.ThisAddIn.CustomTaskPanes.Add(mpChannelShare, "Break TVR Selections")
                            CTPBreakTVR.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
                            CTPBreakTVR.Height = 226
                            CTPBreakTVR.Width = 300
                            CTPBreakTVR.Visible = True
                        Else
                            CTPBreakTVR.Visible = True
                            '  MSprintExChannelShare.Title = "Genre Share Selections"
                        End If
                        'Dim plan, ref As List(Of String)
                        'plan = New List(Of String)()
                        'ref = New List(Of String)()
                        'For index = 0 To tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                        '    plan.Add(tpSelections.UcMarkets1.lbPlan.Items(index))
                        'Next
                        'For index = 0 To tpSelections.UcMarkets1.lbRef.Items.Count - 1
                        '    ref.Add(tpSelections.UcMarkets1.lbRef.Items(index))
                        'Next


                        'If CTPBreakTVR Is Nothing Then
                        '    ' CTSChannelShare.Dispose()
                        '    mpChannelShare = New ucChannelShare(planningDataSet, refDataSet, plan.ToArray(), ref.ToArray(), plantgname, reftgname, tpSelections, "Program", Nothing, channels)
                        '    CTPBreakTVR = Globals.ThisAddIn.CustomTaskPanes.Add(mpChannelShare, "Break TVR Selections")
                        '    CTPBreakTVR.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
                        '    CTPBreakTVR.Height = 226
                        '    CTPBreakTVR.Width = 271
                        '    CTPBreakTVR.Visible = True
                        'End If

                        'tpSelections.TaskPaneLogFile1.scMain.Panel2Collapsed = False
                        'If Not tpSelections.TaskPaneLogFile1.showingChannels Then
                        '    tpSelections.TaskPaneLogFile1.ucChannelsMapping.btnShowHide.Text = "Break TVR Selections"
                        '    tpSelections.TaskPaneLogFile1.ucChannelsMapping.btnShowHide.Refresh()
                        '    tpSelections.TaskPaneLogFile1.showChannelMapping(True)
                        '    tpSelections.TaskPaneLogFile1.Show()
                        'End If
                    End If
                Else
                    MessageBox.Show("Unable to retreive requested BreakTVR details from Server")

                End If
                Globals.ThisAddIn.Application.StatusBar = String.Empty
                Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault
            End If
            If Not (frm Is Nothing) Then
                frm.Close()
            End If
        Catch ex As Exception
            Globals.ThisAddIn.Application.StatusBar = String.Empty
            Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault

            If Not (frm Is Nothing) Then
                frm.Close()
            End If
            LogMpsrintExException("Exception occured while retreiving BreakTVR details." + ex.Message)
            MessageBox.Show("Exception occured while retreiving BreakTVR details.Please view error log for more details")

        End Try
    End Sub

    Private Sub btnLogFile_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnLogFile.Click
        CreateLog()
    End Sub
    Dim SaveFile As System.Windows.Forms.SaveFileDialog
    Friend Sub CreateLog()
        If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then ' And BreakPerformance.isBrkPerfFile Then
            Try
                TotalSpotsWritten = 0
                TotalPlanSpots = 0
                SaveFile = Globals.Ribbons.MSprintExRibbon.logFileSavePath
                SaveFile = New SaveFileDialog()
                ' SaveFile.
                SaveFile.InitialDirectory = LogDirectoryPath
                'SaveFile.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory
                'SaveFile.
                '   SaveFile.RestoreDirectory = False
                SaveFile.Filter = "MsprintEx Log Files | *.log"
                SaveFile.DefaultExt = "log"
                SaveFile.FileName = Globals.ThisAddIn.Application.ActiveWorkbook.Name + ".log"
                If SaveFile.ShowDialog() = DialogResult.OK Then
                    Globals.ThisAddIn.Application.StatusBar = "Creating log File.."
                    System.Windows.Forms.Application.DoEvents()
                    'Globals.ThisAddIn.Application.DoEvents()
                    ' Globals.ThisAddIn.Application.Cursor = Excel.XlMousePointer.xlWait
                    Dim fl As New FileInfo(SaveFile.FileName)
                    'frmProgress = New frmWait
                    'frmProgress.Show()
                    'frmProgress.imgWait.Visible = True
                    Dim input As XElement = New XElement("req_for_log")
                    '   Dim input As XElement = XElement.Load("C:\ASR\Spots.xml")
                    For index = 0 To Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count - 1
                        Dim spot As XElement =
                            <spot SeqNumber=<%= (index + 1).ToString() %> log=<%= Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows(index)("Spot").ToString() %>></spot>
                        input.Add(spot)
                    Next

                    'Dim inputXml As XElement =
                    '<req_for_log>
                    '    <!-- outputs log file format in text -->
                    '    <spot SeqNumber="1" log="4,20130808,215459,215519,,10,12,20"/>
                    '    <spot SeqNumber="2" log="4,20130808,214544,214559,,13,18,15"/>
                    '    <spot SeqNumber="4" log="4,20130808,214401,214403,,6,18,2"/>
                    '    <spot SeqNumber="5" log="4,20130808,215519,215539,,11,12,20"/>
                    '    <spot SeqNumber="6" log="4,20130808,215539,215559,,12,12,20"/>
                    '    <spot SeqNumber="7" log="4,20130808,215404,215414,,6,12,10"/>
                    '    <spot SeqNumber="8" log="4,20130808,215344,215404,,5,12,20"/>
                    '    <spot SeqNumber="9" log="4,20130808,215429,215449,,8,12,20"/>
                    '    <spot SeqNumber="10" log="4,20130808,215324,215344,,4,12,20"/>
                    '    <spot SeqNumber="11" log="4,20130808,215449,215459,,9,12,10"/>
                    '    <spot SeqNumber="12" log="4,20130808,212233,212303,,8,17,30"/>
                    '    <spot SeqNumber="13" log="4,20130808,212343,212408,,11,17,25"/>
                    '    <spot SeqNumber="14" log="4,20130808,212343,212408,,11,17,25"/>
                    '    <spot SeqNumber="15" log="4,20130808,212448,212458,,14,17,10"/>
                    '    <spot SeqNumber="16" log="4,20130808,212158,212203,,1,17,5"/>
                    '    <spot SeqNumber="17" log="4,20130808,212203,212223,,2,17,20"/>
                    '    <spot SeqNumber="18" log="4,20130808,212408,212438,,12,17,30"/>
                    '    <spot SeqNumber="19" log="4,20130808,212333,212343,,10,17,10"/>
                    '    <spot SeqNumber="3" log="4,20130808,212538,212558,,17,17,20"/>
                    '</req_for_log>
                    ' fl.CreateText()
                    Dim name As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                    ' mediaplan.Save(LogDirectoryPath + "RNF_Inp_" + name)
                    input.Save(LogDirectoryPath + "SpotsLogFile_Inp_" + name)
                    Globals.Ribbons.MSprintExRibbon.UpdateUsageReport("WriteLogWS", Globals.Ribbons.MSprintExRibbon.tpSelections.tbClientName.Text.Trim, Globals.Ribbons.MSprintExRibbon.tpSelections.tbBrandValue.Text.Trim, input, 0)

                    '  Dim logfile As String = GetOpForLogWS(input, "http://ec2-54-255-217-55.ap-southeast-1.compute.amazonaws.com:8080/GroupM/spotselectionnew/getlogformat")
                    Dim logfile As String = GetOpForLogWS(input, Globals.Ribbons.MSprintExRibbon.GetURLForWS("WriteLogWSURL_New"))
                    Using PlanLogFile As StreamWriter = New StreamWriter(SaveFile.FileName)
                        PlanLogFile.WriteLine(logfile)
                    End Using
                    Dim output As XElement = New XElement("output")
                    output.Value = logfile
                    output.Save(LogDirectoryPath + "SpotsLogFile_Op_" + name)
                    MessageBox.Show("Log File Successfully Created")
                    'If logTaskPane.rbSingle.Checked Then
                    '    ' frmProgress.lblTitle.Text = "Writing to file:" & vbCrLf & fl.Name
                    '    ' CreateLog(SaveFile.FileName, 0, loPlanData.ListColumns("Total Spots").Index, logTaskPane.dtFromDate.Value, logTaskPane.dtToDate.Value)
                    'Else
                    '    'For Each kp As KeyValuePair(Of WeekYear, Excel.ListColumn) In WeekColumns
                    '    '    Dim drWeek As Plandata.WeeksRow
                    '    '    drWeek = logTaskPane.dtWeeks.FindByWeekNumberYear(kp.Key.WeekNumber, kp.Key.WeekYear)
                    '    '    frmProgress.lblTitle.Text = "Writing to file:" & vbCrLf & "Week-" & kp.Key.WeekNumber & "-" & fl.Name
                    '    '    CreateLog(fl.DirectoryName & "\Week-" & kp.Key.WeekNumber & "-" & fl.Name, kp.Key.WeekNumber, kp.Value.Index, drWeek.StartDate, drWeek.EndDate)
                    '    'Next
                    'End If
                    'frmProgress.lblPercentage.Text = "Finished writing log. Total spots:" & vbCrLf & TotalSpotsWritten & " of " & TotalPlanSpots & " spots"
                    'frmProgress.pbUpload.Style = ProgressBarStyle.Continuous
                    'frmProgress.pbUpload.Value = 100
                    'frmProgress.imgWait.Visible = False
                    Globals.ThisAddIn.Application.StatusBar = String.Empty
                    '  Globals.ThisAddIn.Application.Cursor = ExXlMousePointer.xlDefault
                End If
            Catch ex As Exception
                Globals.ThisAddIn.Application.StatusBar = String.Empty
                'Globals.ThisAddIn.Application.Cursor = Excel.XlMousePointer.xlDefault
                LogMpsrintExException("Exception occured while saving log file." + ex.Message)
                MessageBox.Show("Exception occured while creating log file.Please refer to Error log for more details")
            End Try
        End If
    End Sub
    Public Function GetAllWSURLS() As Data.DataTable
        Dim wsURLS As Data.DataSet = New Data.DataSet()

        Try
            '  Dim con As SqlConnection = New SqlConnection("Data Source=BANSQLD01101;Initial Catalog=MsprintXTracker;User ID=ctgreport;Password=Meritus123")
            Dim con As SqlConnection = New SqlConnection("Server= MUMSQLP01107\GRMINDSQL01;Database=MsprintXTracker;User Id=MSXAdmin;Password=MSXAdmin@123;")
            con.Open()
            Dim queryString As String = "SELECT *  FROM [MsprintXTracker].[dbo].[WSURLS]"
            Dim adapter As SqlDataAdapter = New SqlDataAdapter( _
              queryString, con)
            adapter.Fill(wsURLS, "WSURLS")
            con.Close()

        Catch ex As Exception
            '  WriteToErrorLog(ex.Message + "Error occured for Executive Summary Details")
            LogMpsrintExException("Exception occured while getting all WS URLS from DB.Message : " + ex.Message)
            ' Throw ex
        End Try
        Return wsURLS.Tables(0)
    End Function
    Public Function GetURLForWS(ByVal wsName As String) As String
        Dim URL As String = String.Empty
        Try
            Dim filter As String = String.Format("Name='{0}'", wsName)
            Dim rows As Data.DataRow() = wsURLS.Select(filter)

            If rows.Count > 0 Then
                URL = rows(0)("URL").ToString()
            End If

        Catch ex As Exception
            LogMpsrintExException("Exception occured while returning URL for given WSName.Message: " + ex.Message)
        End Try
        Return URL
    End Function
    Public Function SetWaitCursor(ByVal waitMessage As String)
        Try
            Globals.ThisAddIn.Application.StatusBar = waitMessage
            Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault
        Catch ex As Exception

        End Try
    End Function
    Public Function SetNormalCursor()
        Try
            Globals.ThisAddIn.Application.StatusBar = String.Empty
            Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault
        Catch ex As Exception

        End Try
    End Function
    Public Function GetGridTable() As Data.DataTable
        Dim dt As Data.DataTable = New Data.DataTable()
        Dim CBCell As DataGridViewComboBoxCell

        dt.Columns.Add("PCName")
        dt.Columns.Add("MCName")
        Try
            Dim view As DataView = CType(channelMapping.PlanChannelsBindingSource.DataSource, Data.DataView)
            Dim table As Data.DataTable = view.ToTable()
            For Each row As Data.DataRow In table.Rows
                Dim dr As Data.DataRow = dt.NewRow()
                'CBCell = row.Cells("ChannelNameDataGridViewTextBoxColumn")

                dr("PCName") = row("ChannelName").ToString()
                Dim filter As String = String.Format("ID={0}", Convert.ToInt32(row("ChannelCode").ToString()))
                Dim channel As String = Globals.Ribbons.MSprintExRibbon.dtchannels.Select(filter)(0)("Name")
                dr("MCName") = channel
                dt.Rows.Add(dr)
            Next
            '    
            'Next()
        Catch ex As Exception

        End Try
        Return dt
    End Function
    Public Function AllChannelsMapped() As Boolean
        Dim allchannlsmapped As Boolean = True
        Try
            'For Each row As System.Windows.Forms.DataGridViewRow In channelMapping.dgvChannels.Rows

            '    If Convert.ToString(row.Cells.Item("ChannelCodeDataGridViewTextBoxColumn").Value).Equals(" - - Select - - ") Then
            '        allchannlsmapped = False
            '        Exit For
            '    End If

            '    'Dim dr As Data.DataRow = dt.NewRow()
            '    'dr("PCName") = row.Cells.Item("ChannelNameDataGridViewTextBoxColumn").Value
            '    'dr("MCName") = row.Cells.Item("ChannelCodeDataGridViewTextBoxColumn").Value
            '    'dt.Rows.Add(dr)
            'Next
            ''    

            If planOpenedSuccessfully Then
                '  allchannlsmapped = True
                Return allchannlsmapped
            End If

            If channelMapping.PlanChannelsBindingSource.DataSource Is Nothing Then
                allchannlsmapped = False
            Else

                Dim view As DataView = CType(channelMapping.PlanChannelsBindingSource.DataSource, Data.DataView)
                Dim table As Data.DataTable = view.ToTable()
                ' Dim filter As String = String.Format("ChannelCode='{0}'", )

                If table.Rows.Count > 0 Then
                    Dim rows As Data.DataRow() = table.Select("ChannelCode='000'")

                    If rows.Count > 0 Then
                        allchannlsmapped = False
                    Else
                        allchannlsmapped = True
                    End If
                Else
                    allchannlsmapped = False
                End If


            End If


            'Next()
        Catch ex As Exception
            allchannlsmapped = False
        End Try
        Return allchannlsmapped
    End Function
    Public Function CreateRnFWorkBook() As Microsoft.Office.Interop.Excel.Workbook

    End Function
    Private Sub btnRnF_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnRnF.Click
        'Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        'Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(sheet)
        'For Each range As Microsoft.Office.Interop.Excel.Range In vstoWorkbook.UsedRange
        '    System.Windows.Forms.MessageBox.Show(Convert.ToString(range.Value2))
        'Next
        Dim frm As frmWait
        Try

            'Dim str As String = String.Empty
            'For Each row As Microsoft.Office.Interop.Excel.ListRow In loSpotSelection.ListRows
            '    '  loSpotSelection.ListRows(0).Range.Cells.
            '    'For Each cell As Microsoft.Office.Interop.Excel.Range In row.Range.Cells
            '    '    ' str = row.Range.Cells(row.Index, col.Index).ToString()
            '    '    str = cell.Text
            '    'Next

            'Next

            If Not (RnFAvaiSpots Is Nothing) Then
                RnFAvaiSpots.Rows.Clear()
            End If

            ' Dim dtt As BindingSource = CType(channelMapping.dgvChannels.DataSource, BindingSource)

            Dim plantgname As String = String.Empty
            Dim reftgname As String = String.Empty
            Dim dtable As System.Data.DataTable = DirectCast(tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            plantgname = dtable.Rows(0)(1).ToString().Trim()
            ' reftgname = dtable.Rows(1)(1).ToString().Trim()
            If plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please make Target Group and Market Group(s) Selections")
            ElseIf plantgname.Length > 1 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for chosen TG")
            ElseIf plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count > 0 Then
                System.Windows.Forms.MessageBox.Show("Please choose Planning Target group for Market groups")
            ElseIf loSpotSelection Is Nothing Then
                System.Windows.Forms.MessageBox.Show("Please enter plan to view Reach and Frequency")
            ElseIf Not (isPlanClean) Then
                System.Windows.Forms.MessageBox.Show("Please clean plan to view Reach and Frequency and ensure no duplication of rows")
            ElseIf Not (AllChannelsMapped()) Then
                System.Windows.Forms.MessageBox.Show("Please map all channels with master channellist")
                'ElseIf reftgname.Length > 0 And tpSelections.UcMarkets1.lbRef.Items.Count < 1 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for reference Target Group chosen")
            Else
                SetWaitCursor("Getting Reach and Frequency details..")

                '  Globals.ThisAddIn.Application.DoEvents()
                System.Windows.Forms.Application.DoEvents()
                'frm = New frmWait()
                'frm.Show()
                ' loSpotSelection.RefreshDataRows()
                xecellineItemsTable = CType(loSpotSelection.DataSource, Data.DataTable)
                mediaplan = ConstructInputRnFXML("RnFWS")
                ' mediaplan = XElement.Load("C:\\ASR\\For Manish\\Logs\\rnf.xml")
                ' mediaplan = XElement.Load("Z:\\asr\\Errors\\RNF_Inp_20032014_150040.xml")


                'testmediaplan =
                '<mediaplan>
                '    <!-- TO BE FROZEN 26 Dec 13 -->
                '    <!-- common section start -->
                '    <PreEvalPeriod>
                '        <StartDate>20130804</StartDate>
                '        <EndDate>20130817</EndDate>
                '    </PreEvalPeriod>

                '    <DayParts>
                '        <DayPart>0800-1200</DayPart>
                '        <DayPart>2100-2200</DayPart>
                '    </DayParts>


                '    <!-- common section ends -->


                '    <tg name="CS 15-44" cs="1" sec="1,2,3,4" sex="1,2" age="3,4,5">

                '        <mg name="mg1" type="group">
                '            <market>1</market>
                '            <market>3</market>
                '        </mg>

                '        <mg name="mg2" type="single">
                '            <market>2</market>
                '        </mg>
                '    </tg>


                '    <!-- TVR000s,TVR,GRP000s,GRP,AvgFreq,CummCost,SpotCPRP,CummCPRP,Reach000s,R1,R2,R3,R4,R5,R6,R7,R8,R9,R10 -->

                '    <!-- all output -->
                '    <plan type="weekwise"><!-- clubbed-->
                '        <period StartDate="20130804" EndDate="20130810" year="2013" WeekNum="32">
                '            <programme guid="1" SeqNumber="1" ChannelCode="004" ChannelName="Star Plus" ProgName="Yeh Rishta Kya Kehlata Hai" days="Thu" StartTime="21:30" EndTime="22:00" CostPer10s="150" caption="Colgate Kids Jumping" AdDuration="30" NumberOfSpots="10">

                '                <selected_spots><!-- this is a demo output - spot count and dates may be incorrect. attributes are correctNO SPACES BETWEEN VALUES IN CSV -->
                '                    <spot log="004,10082013,032238,032308,0,001,005,00030"/>
                '                </selected_spots>
                '            </programme>
                '            <programme guid="2" SeqNumber="2" ChannelCode="004" ChannelName="Star Plus" ProgName="DIYA AUR BAATI HUM" days="Mon,Tue,Wed,Thu,Fri" StartTime="21:00" EndTime="21:30" CostPer10s="120" caption="Colgate Kids Jumping" AdDuration="20" NumberOfSpots="9">

                '                <selected_spots><!-- this is a demo output - spot count and dates may be incorrect. attributes are correctNO SPACES BETWEEN VALUES IN CSV -->
                '                    <spot log="004,10082013,032238,032308,0,001,005,00030"/>
                '                </selected_spots>
                '            </programme>
                '        </period>
                '    </plan>

                '</mediaplan>


                '  request = WebRequest.Create("")
                '

                'request.Method = "POST"
                'request.ContentType = "application/x-www-form-urlencoded"
                'request.Timeout = 300000
                'request.ServicePoint.MaxIdleTime = 300000
                ''  request.KeepAlive = True
                'inputstring = HttpUtility.UrlEncode(MyExtensionClasses.OuterXml(testmediaplan))

                'stream = request.GetRequestStream()
                ''  Dim bytes As Byte() = Encoding.Default.GetBytes(input.OuterXml)

                'Dim encoding As System.Text.UTF8Encoding = New UTF8Encoding()
                'postData = "inputXML=" + inputstring

                'data = encoding.GetBytes(postData)
                ''input.Save(stream)
                '' request.ContentLength = data.Length
                'stream.Write(data, 0, data.Length)

                '' request.Proxy = Nothing
                'ws = request.GetResponse()

                'oStream = ws.GetResponseStream()
                'Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")

                '' Pipe the stream to a higher level stream reader with the required encoding format.
                'Dim readStream As New StreamReader(oStream, encode)
                '     Dim separators() As String = {"Genre,Viewership"}
                ' Dim file As [String]() = readStream.ReadToEnd().Split(separators, StringSplitOptions.RemoveEmptyEntries)

                Dim name As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                mediaplan.Save(LogDirectoryPath + "RNF_Inp_" + name)
                '  rnfoutputXml = GetOpXMLFromWS(mediaplan, "http://ec2-54-255-217-55.ap-southeast-1.compute.amazonaws.com:8080/GroupM-new/spotselectionnew/getselectedspot")
                rnfoutputXml = GetOpXMLFromWS(mediaplan, Globals.Ribbons.MSprintExRibbon.GetURLForWS("RnFTillZeroWSURL_New"))
                ' rnfoutputXml = XElement.Load("C:\ASR\RNF_.xml")
                If Not (rnfoutputXml Is Nothing) Then
                    Dim name1 As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"

                    rnfoutputXml.Save(LogDirectoryPath + "RNF_Op_" + name1)

                    If ConstructOpRnFTable(rnfoutputXml) Then
                        RnFShowResultsTable = nbdmain(RnFShowResultsTable)
                        RnFShowResultsTable = ReverseCalculateTVRReach(RnFShowResultsTable)
                        btnLogFile.Enabled = True
                        ' xecellTableCopy = xecellineItemsTable.Copy()
                        xecellTableCopy = xecelTable.Copy()
                        Try

                            If Not (xecelTable.Columns.Contains("Number of Spots Returned")) Then
                                xecelTable.Columns.Add("Number of Spots Returned")
                            End If
                            For Each row As Data.DataRow In xecelTable.Rows
                                Dim filter As String = String.Format("GUID='{0}'", row("GUID").ToString())
                                Dim rows As Data.DataRow() = RnFSelectedSpots.Select(filter)
                                row("Number of Spots Returned") = rows.Length
                            Next
                            loSpotSelection.SetDataBinding(xecelTable)
                        Catch ex As Exception
                            LogMpsrintExException("Exception occured while calculating number of returned spots")
                        End Try
                        'Dim spots = xecellineItemsTable.AsEnumerable().Join(RnFOutputTable.AsEnumerable(), Function(o) o.Field(Of Int32)("GUID"), _
                        '                          Function(c) c.Field(Of Int32)("GUID"), _
                        '                          Function(c, o) _
                        '                              New With {.GUID = o.Field(Of Int32)("GUID"), _
                        '                                       .Spot = o.Field(Of String)("SpotString"), _
                        '                                       .StartDate = o.Field(Of DateTime)("Start Date"),
                        '                                       .EndDate = o.Field(Of DateTime)("End Date"),
                        '                                       .ChannelName = o.Field(Of String)("ChannelName"),
                        '                                        .Day = c.Field(Of String)("Day"),
                        '                                        .Program = c.Field(Of String)("Programme"),
                        '                                        .TG = o.Field(Of String)("TG"),
                        '                                        .MG = o.Field(Of String)("MG"),
                        '                                         .ReachVal = o.Field(Of String)("ReachVal"),
                        '                                        .WeekNum = o.Field(Of Int32)("WeekNum")})
                        ' mpTpSpotSelection = New ucSpotSelection()
                        'Dim RnfReach As Data.DataTable = New Data.DataTable()
                        'RnfReach = spots.CopyToDataTable()
                        ' Dim RnfSSpots As Data.DataTable = spots.CopyToDataTable()
                        '  RnFSelectedSpots = New Data.DataTable()
                        ' Dim RnfOutPutCopy1 As Data.DataTable = RnFOutputTable.Copy()
                        '  RnFSelectedSpots = RnfOutPutCopy1
                        ' RnFSelectedSpots = RnFSelectedSpots.DefaultView.ToTable(True, New String() {"GUID", "Spot", "Start Date", "End Date", "WeekNum", "Channel", "Date", "Start Time", "Duration(Sec)", "PA", "TA", "Cost"})
                        ' Dim RnfShowResultsTable As Data.DataTable = New Data.DataTable()
                        ' Dim RnfOutputCopy2 As Data.DataTable = RnFOutputTable.Copy
                        ' RnfShowResultsTable = RnfOutputCopy2
                        ' RnfShowResultsTable = RnfShowResultsTable.DefaultView.ToTable(True, New String() {"Start Date", "End Date", "WeekNum", "TG", "MG", "Channel", "Date", "Day", "Start Time", "Duration(Sec)", "Programme", "PA", "TA", "TVR000s", "TVR", "GRP000s", "GRP", "AvgFreq", "CummCost", "SpotCPRP", "CummCPRP", "Reach000s", "1+", "2+", "3+", "4+", "5+", "6+", "7+", "8+", "9+", "10+"})
                        ' spot.Columns.Add("ChannelCode")
                        'RnFSelectedSpots.Columns.Add("Channel")
                        'RnFSelectedSpots.Columns.Add("Date", System.Type.GetType("System.DateTime"))
                        'RnFSelectedSpots.Columns.Add("Start Time")
                        '' RnFSelectedSpots.Columns.Add("End Time")
                        'RnFSelectedSpots.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
                        '' RnFSelectedSpots.Columns.Add("Commercial")
                        'RnFSelectedSpots.Columns.Add("Cost")
                        'RnFSelectedSpots.Columns.Add("PA")
                        'RnFSelectedSpots.Columns.Add("TA")
                        'RnFSelectedSpots.Columns.Add("GUID", System.Type.GetType("System.Int32"))
                        'RnFSelectedSpots.Columns.Add("Spot")
                        'RnFSelectedSpots.Columns.Add("StartDate", System.Type.GetType("System.DateTime"))
                        'RnFSelectedSpots.Columns.Add("EndDate", System.Type.GetType("System.DateTime"))
                        'RnFSelectedSpots.Columns.Add("WeekNum", System.Type.GetType("System.Int32"))


                        'RnfShowResultsTable.Columns.Add("StartDate")
                        'RnfShowResultsTable.Columns.Add("EndDate")
                        'RnfShowResultsTable.Columns.Add("WeekNum", System.Type.GetType("System.Int32"))
                        'RnfShowResultsTable.Columns.Add("TG")
                        'RnfShowResultsTable.Columns.Add("MG")
                        'RnfShowResultsTable.Columns.Add("Channel")
                        'RnfShowResultsTable.Columns.Add("Date") 'spot date
                        'RnfShowResultsTable.Columns.Add("Day")
                        'RnfShowResultsTable.Columns.Add("Start Time")
                        '' RnfShowResultsTable.Columns.Add("End Time")
                        'RnfShowResultsTable.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
                        'RnfShowResultsTable.Columns.Add("Programme")
                        'RnfShowResultsTable.Columns.Add("PA")
                        'RnfShowResultsTable.Columns.Add("TA")
                        'RnfShowResultsTable.Columns.Add("TVR000s")
                        'RnfShowResultsTable.Columns.Add("TVR")
                        'RnfShowResultsTable.Columns.Add("GRP000s")
                        'RnfShowResultsTable.Columns.Add("GRP")
                        'RnfShowResultsTable.Columns.Add("AvgFreq")
                        'RnfShowResultsTable.Columns.Add("CummCost")
                        'RnfShowResultsTable.Columns.Add("SpotCPRP")
                        'RnfShowResultsTable.Columns.Add("CummCPRP")
                        'RnfShowResultsTable.Columns.Add("Reach000s")
                        'RnfShowResultsTable.Columns.Add("1+")
                        'RnfShowResultsTable.Columns.Add("2+")
                        'RnfShowResultsTable.Columns.Add("3+")
                        'RnfShowResultsTable.Columns.Add("4+")
                        'RnfShowResultsTable.Columns.Add("5+")
                        'RnfShowResultsTable.Columns.Add("6+")
                        'RnfShowResultsTable.Columns.Add("7+")
                        'RnfShowResultsTable.Columns.Add("8+")
                        'RnfShowResultsTable.Columns.Add("9+")
                        'RnfShowResultsTable.Columns.Add("10+")
                        'For index = 0 To spots.Count - 1
                        '    Dim spotrow As Data.DataRow = GetSpotRow(spots(index).Spot)
                        '    Dim reachrow As Data.DataRow = GetReachRow(spots(index).ReachVal)
                        '    Dim dr As Data.DataRow = RnFSelectedSpots.NewRow()
                        '    dr("GUID") = spots(index).GUID
                        '    dr("Spot") = spots(index).Spot
                        '    dr("StartDate") = spots(index).StartDate
                        '    dr("EndDate") = spots(index).EndDate
                        '    dr("WeekNum") = spots(index).WeekNum
                        '    dr("Channel") = spots(index).ChannelName
                        '    dr("Date") = spotrow("Date")
                        '    dr("Start Time") = spotrow("StartTime")
                        '    '  dr("End Time") = spotrow("EndTime")
                        '    dr("Duration(Sec)") = spotrow("Duration(Sec)")
                        '    dr("PA") = spotrow("PA")
                        '    dr("TA") = spotrow("TA")
                        '    ' dr("Commercial") = spotrow("Commercial")
                        '    dr("Cost") = spotrow("Cost")
                        '    '  dr("ID") = index + 1
                        '    ' RnFSelectedSpots.AsEnumerable().Contains(
                        '    RnFSelectedSpots.Rows.Add(dr)
                        '    'rnfshowoutputtable
                        '    Dim dRow As Data.DataRow = RnfShowResultsTable.NewRow()
                        '    dRow("StartDate") = spots(index).StartDate
                        '    dRow("EndDate") = spots(index).EndDate
                        '    dRow("WeekNum") = spots(index).WeekNum
                        '    dRow("Channel") = spots(index).ChannelName
                        '    dRow("TG") = spots(index).TG
                        '    dRow("MG") = spots(index).MG
                        '    dRow("Programme") = spots(index).Program
                        '    dRow("Day") = spots(index).Day
                        '    dRow("Date") = spotrow("Date")
                        '    dRow("Start Time") = spotrow("StartTime")
                        '    ' dRow("End Time") = spotrow("EndTime")
                        '    dRow("Duration(Sec)") = spotrow("Duration(Sec)")
                        '    dRow("PA") = spotrow("PA")
                        '    dRow("TA") = spotrow("TA")
                        '    dRow("TVR000s") = reachrow("TVR000s")
                        '    dRow("TVR") = reachrow("TVR")
                        '    dRow("GRP000s") = reachrow("GRP000s")
                        '    dRow("GRP") = reachrow("GRP")
                        '    dRow("AvgFreq") = reachrow("AvgFreq")
                        '    dRow("CummCost") = reachrow("CummCost")
                        '    dRow("SpotCPRP") = reachrow("SpotCPRP")
                        '    dRow("CummCPRP") = reachrow("CummCPRP")
                        '    dRow("Reach000s") = reachrow("Reach000s")
                        '    dRow("1+") = reachrow("R1")
                        '    dRow("2+") = reachrow("R2")
                        '    dRow("3+") = reachrow("R3")
                        '    dRow("4+") = reachrow("R4")
                        '    dRow("5+") = reachrow("R5")
                        '    dRow("6+") = reachrow("R6")
                        '    dRow("7+") = reachrow("R7")
                        '    dRow("8+") = reachrow("R8")
                        '    dRow("9+") = reachrow("R9")
                        '    dRow("10+") = reachrow("R10")
                        '    RnfShowResultsTable.Rows.Add(dRow)
                        'Next

                        'Parallel.For(0, spots.Count - 1, Sub(index)
                        '                                     Dim spotrow As Data.DataRow = GetSpotRow(spots(index).Spot)
                        '                                     Dim reachrow As Data.DataRow = GetReachRow(spots(index).ReachVal)
                        '                                     Dim dr As Data.DataRow = RnFSelectedSpots.NewRow()
                        '                                     dr("GUID") = spots(index).GUID
                        '                                     dr("Spot") = spots(index).Spot
                        '                                     dr("StartDate") = spots(index).StartDate
                        '                                     dr("EndDate") = spots(index).EndDate
                        '                                     dr("WeekNum") = spots(index).WeekNum
                        '                                     dr("Channel") = spots(index).ChannelName
                        '                                     dr("Date") = spotrow("Date")
                        '                                     dr("Start Time") = spotrow("StartTime")
                        '                                     '  dr("End Time") = spotrow("EndTime")
                        '                                     dr("Duration(Sec)") = spotrow("Duration(Sec)")
                        '                                     dr("PA") = spotrow("PA")
                        '                                     dr("TA") = spotrow("TA")
                        '                                     ' dr("Commercial") = spotrow("Commercial")
                        '                                     dr("Cost") = spotrow("Cost")
                        '                                     '  dr("ID") = index + 1
                        '                                     ' RnFSelectedSpots.AsEnumerable().Contains(
                        '                                     RnFSelectedSpots.Rows.Add(dr)
                        '                                     'rnfshowoutputtable
                        '                                     Dim dRow As Data.DataRow = RnfShowResultsTable.NewRow()
                        '                                     dRow("StartDate") = spots(index).StartDate
                        '                                     dRow("EndDate") = spots(index).EndDate
                        '                                     dRow("WeekNum") = spots(index).WeekNum
                        '                                     dRow("Channel") = spots(index).ChannelName
                        '                                     dRow("TG") = spots(index).TG
                        '                                     dRow("MG") = spots(index).MG
                        '                                     dRow("Programme") = spots(index).Program
                        '                                     dRow("Day") = spots(index).Day
                        '                                     dRow("Date") = spotrow("Date")
                        '                                     dRow("Start Time") = spotrow("StartTime")
                        '                                     ' dRow("End Time") = spotrow("EndTime")
                        '                                     dRow("Duration(Sec)") = spotrow("Duration(Sec)")
                        '                                     dRow("PA") = spotrow("PA")
                        '                                     dRow("TA") = spotrow("TA")
                        '                                     dRow("TVR000s") = reachrow("TVR000s")
                        '                                     dRow("TVR") = reachrow("TVR")
                        '                                     dRow("GRP000s") = reachrow("GRP000s")
                        '                                     dRow("GRP") = reachrow("GRP")
                        '                                     dRow("AvgFreq") = reachrow("AvgFreq")
                        '                                     dRow("CummCost") = reachrow("CummCost")
                        '                                     dRow("SpotCPRP") = reachrow("SpotCPRP")
                        '                                     dRow("CummCPRP") = reachrow("CummCPRP")
                        '                                     dRow("Reach000s") = reachrow("Reach000s")
                        '                                     dRow("1+") = reachrow("R1")
                        '                                     dRow("2+") = reachrow("R2")
                        '                                     dRow("3+") = reachrow("R3")
                        '                                     dRow("4+") = reachrow("R4")
                        '                                     dRow("5+") = reachrow("R5")
                        '                                     dRow("6+") = reachrow("R6")
                        '                                     dRow("7+") = reachrow("R7")
                        '                                     dRow("8+") = reachrow("R8")
                        '                                     dRow("9+") = reachrow("R9")
                        '                                     dRow("10+") = reachrow("R10")
                        '                                     RnfShowResultsTable.Rows.Add(dRow)
                        '                                 End Sub)
                        ' RnFSelectedSpots = RnFSelectedSpots.DefaultView.ToTable(True, New String() {"GUID", "Spot", "StartDate", "EndDate", "WeekNum", "Channel", "Date", "Start Time", "Duration(Sec)", "PA", "TA", "Cost"})
                        'Dim id As DataColumn = RnFSelectedSpots.Columns.Add("ID", System.Type.GetType("System.Int32"))
                        'id.AutoIncrement = True

                        'End Sub)

                        rnFWorkBook = Globals.ThisAddIn.Application.Workbooks.Add(Type.Missing)

                        'If Not CheckSheetExists("ReachNFrequency") Then

                        '    newSheet.Name = "ReachNFrequency"
                        'Else
                        '    newSheet = CheckAndReturnSheet("ReachNFrequency")
                        '    ' newSheet.UsedRange.Clear()
                        '    CleanSheet(newSheet)
                        '    newSheet.Activate()
                        'End If
                        Dim mgs As List(Of String) = New List(Of String)()
                        For index1 = 0 To tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                            mgs.Add(tpSelections.UcMarkets1.lbPlan.Items(index1).ToString())
                        Next
                        mgs.Add("TotalMarkets")
                        Try


                            Dim periodcount As Integer = 1
                            For index1 = 0 To mgs.Count - 1
                                Dim rowcount As Integer = 3
                                newSheet = rnFWorkBook.Worksheets.Add(Type.Missing, Type.Missing, 1, Type.Missing)

                                Dim dtweekss As Data.DataTable = tpSelections.TaskPaneLogFile1.dtWeeks
                                Dim tgcell As Microsoft.Office.Interop.Excel.Range = newSheet.Range("$A$2", Type.Missing)
                                tgcell.Value2 = plantgname
                                Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(newSheet)
                                'Dim mgcell As Microsoft.Office.Interop.Excel.Range
                                'If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                                '    mgcell = vstoWorkbook.get_Range(vstoWorkbook.Cells(1 * rowcount + 1 + (8 * index1), 1), vstoWorkbook.Cells(1 * rowcount + 1 + (8 * index1), 1))
                                'Else
                                '    mgcell = vstoWorkbook.get_Range(vstoWorkbook.Cells((1 * rowcount + 1 + (8 * index1)) * periodcount, 1), vstoWorkbook.Cells((1 * rowcount + 1 + (8 * index1)) * periodcount, 1))

                                'End If

                                Dim mgroup As String = mgs(index1)
                                vstoWorkbook.Name = mgroup
                                ' Dim periodElement As XElement = GetPeriodElementForMG(mgroup)
                                '  mgcell.Value2 = "MG:" + mgroup
                                'Dim lrange As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(4, ii), vstoWorkbook.Cells(4, ii + 3)), Microsoft.Office.Interop.Excel.Range)
                                'listobject.AutoSetDataBoundColumnHeaders = True

                                If tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                                    Dim periodElement As XElement = GetPeriodElementForMG(mgroup, String.Empty)
                                    Dim startdate As String = tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToString("dd/MM/yyyy")
                                    Dim enddate As String = tpSelections.TaskPaneLogFile1.dtToDate.Value.ToString("dd/MM/yyyy")
                                    Dim periodrange As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(1 + 3 + 2, 1), vstoWorkbook.Cells(1 + 3 + 2, 1))
                                    periodrange.Value2 = "Period : " + startdate + "To " + enddate
                                    Dim universeRange As Microsoft.Office.Interop.Excel.Range = periodrange.Next(2, 0)
                                    Try
                                        universeRange.Value2 = "Univserse: " + periodElement.Attribute("universe").Value
                                    Catch ex As Exception
                                        universeRange.Value2 = "Univserse: "
                                    End Try
                                    Dim scountrange As Microsoft.Office.Interop.Excel.Range = universeRange.Next(1, 2)
                                    Try
                                        scountrange.Value2 = "Sample Count: " + periodElement.Attribute("SampleCount").Value
                                    Catch ex As Exception
                                        scountrange.Value2 = "Sample Count: "
                                    End Try

                                    Dim middate As Microsoft.Office.Interop.Excel.Range = scountrange.Next(1, 2)
                                    Try
                                        middate.Value2 = "Mid Date: " + Convert.ToDateTime(periodElement.Attribute("MidDate").Value).ToString("dd/MM/yyyy")
                                    Catch ex As Exception
                                        middate.Value2 = "Mid Date: "
                                    End Try
                                    Dim rnfrows As Data.DataRow() = RnFShowResultsTable.Select("TG='" + plantgname + "' and MG='" + mgroup + "'")


                                    ' dttt.Rows.
                                    If rnfrows.Count > 0 Then
                                        'Dim table As Data.DataTable = New Data.DataTable()
                                        'table = RnfShowResultsTable.Clone()
                                        'For Each row1 As Data.DataRow In rnfrows
                                        '    table.ImportRow(row1)
                                        'Next
                                        Dim lrange As Microsoft.Office.Interop.Excel.Range = middate.Next(2, -4)
                                        Dim table As Data.DataTable = rnfrows.CopyToDataTable()

                                        '  table.Columns.Remove("End Time")
                                        Dim blankrow As Data.DataRow = table.NewRow()
                                        table.Rows.InsertAt(blankrow, table.Rows.Count)
                                        Dim finalrow As Data.DataRow = table.NewRow()
                                        '                ' AvgFreq="1.03" CummCPRP="0" CummCost="0"
                                        'EndDate="20130810" GRP="64.97" GRP000s="13502.34"
                                        'MidDate="2013-08-07" NumberOfSpots="9" R1="62.85"
                                        'R2="25.07" R3="18.21" R4="14.98" R5="11.34" R6="8.18"
                                        'R7="5.58" R8="3.49" R9="0" Reach000s="9833.79"
                                        'SampleCount="2165.0" SpotCPRP="0"
                                        Try


                                            '  finalrow("GRP000s") = periodElement.Attribute("GRP000s").Value
                                            finalrow("GRP000s") = table.Rows(table.Rows.Count - 2)("GRP000s").ToString()
                                            ' finalrow("GRP") = periodElement.Attribute("GRP").Value
                                            '  finalrow("AvgFreq") = periodElement.Attribute("AvgFreq").Value
                                            finalrow("CummCost") = periodElement.Attribute("CummCost").Value
                                            finalrow("SpotCPRP") = periodElement.Attribute("SpotCPRP").Value
                                            finalrow("CummCPRP") = periodElement.Attribute("CummCPRP").Value
                                            '  finalrow("Reach000s") = periodElement.Attribute("Reach000s").Value
                                            finalrow("Reach000s") = table.Rows(table.Rows.Count - 2)("Reach000s").ToString()
                                            '  Dim reachValues As Decimal() = New Decimal() {periodElement.Attribute("R1").Value, periodElement.Attribute("R2").Value, periodElement.Attribute("R3").Value, periodElement.Attribute("R4").Value, periodElement.Attribute("R5").Value, periodElement.Attribute("R6").Value, periodElement.Attribute("R7").Value, periodElement.Attribute("R8").Value, periodElement.Attribute("R9").Value, periodElement.Attribute("R10").Value, periodElement.Attribute("R11").Value, periodElement.Attribute("R12").Value, periodElement.Attribute("R13").Value, periodElement.Attribute("R14").Value, periodElement.Attribute("R15").Value, periodElement.Attribute("R16").Value, periodElement.Attribute("R17").Value, periodElement.Attribute("R18").Value, periodElement.Attribute("R19").Value, periodElement.Attribute("R20").Value}
                                            '  Dim rValues As Decimal() = nbd(periodElement.Attribute("GRP").Value, reachValues)
                                            'finalrow("1+") = periodElement.Attribute("R1").Value
                                            'finalrow("2+") = periodElement.Attribute("R2").Value
                                            'finalrow("3+") = periodElement.Attribute("R3").Value
                                            'finalrow("4+") = periodElement.Attribute("R4").Value
                                            'finalrow("5+") = periodElement.Attribute("R5").Value
                                            'finalrow("6+") = periodElement.Attribute("R6").Value
                                            'finalrow("7+") = periodElement.Attribute("R7").Value
                                            'finalrow("8+") = periodElement.Attribute("R8").Value
                                            'finalrow("9+") = periodElement.Attribute("R9").Value
                                            'finalrow("10+") = periodElement.Attribute("R10").Value
                                            'finalrow("11+") = periodElement.Attribute("R11").Value
                                            'finalrow("12+") = periodElement.Attribute("R12").Value
                                            'finalrow("13+") = periodElement.Attribute("R13").Value
                                            'finalrow("14+") = periodElement.Attribute("R14").Value
                                            'finalrow("15+") = periodElement.Attribute("R15").Value
                                            'finalrow("16+") = periodElement.Attribute("R16").Value
                                            'finalrow("17+") = periodElement.Attribute("R17").Value
                                            'finalrow("18+") = periodElement.Attribute("R18").Value
                                            'finalrow("19+") = periodElement.Attribute("R19").Value
                                            '  finalrow("20+") = periodElement.Attribute("R20").Value
                                            finalrow("GRP") = table.Rows(table.Rows.Count - 2)("GRP").ToString()
                                            finalrow("AvgFreq") = table.Rows(table.Rows.Count - 2)("AvgFreq").ToString()
                                            If HRN > 20 Then
                                                finalrow("1+") = table.Rows(table.Rows.Count - 2)("1+").ToString()
                                                '  finalrow("GRP") = Convert.ToDecimal(finalrow("AvgFreq").ToString) * Convert.ToDecimal(finalrow("1+").ToString())
                                                finalrow("2+") = table.Rows(table.Rows.Count - 2)("2+").ToString()
                                                finalrow("3+") = table.Rows(table.Rows.Count - 2)("3+").ToString()
                                                finalrow("4+") = table.Rows(table.Rows.Count - 2)("4+").ToString()
                                                finalrow("5+") = table.Rows(table.Rows.Count - 2)("5+").ToString()
                                                finalrow("6+") = table.Rows(table.Rows.Count - 2)("6+").ToString()
                                                finalrow("7+") = table.Rows(table.Rows.Count - 2)("7+").ToString()
                                                finalrow("8+") = table.Rows(table.Rows.Count - 2)("8+").ToString()
                                                finalrow("9+") = table.Rows(table.Rows.Count - 2)("9+").ToString()
                                                finalrow("10+") = table.Rows(table.Rows.Count - 2)("10+").ToString()
                                                finalrow("11+") = table.Rows(table.Rows.Count - 2)("11+").ToString()
                                                finalrow("12+") = table.Rows(table.Rows.Count - 2)("12+").ToString()
                                                finalrow("13+") = table.Rows(table.Rows.Count - 2)("13+").ToString()
                                                finalrow("14+") = table.Rows(table.Rows.Count - 2)("14+").ToString()
                                                finalrow("15+") = table.Rows(table.Rows.Count - 2)("15+").ToString()
                                                finalrow("16+") = table.Rows(table.Rows.Count - 2)("16+").ToString()
                                                finalrow("17+") = table.Rows(table.Rows.Count - 2)("17+").ToString()
                                                finalrow("18+") = table.Rows(table.Rows.Count - 2)("18+").ToString()
                                                finalrow("19+") = table.Rows(table.Rows.Count - 2)("19+").ToString()
                                                finalrow("20+") = table.Rows(table.Rows.Count - 2)("20+").ToString()
                                            Else

                                                For index = 1 To HRN
                                                    finalrow(index.ToString() + "+") = table.Rows(table.Rows.Count - 2)(index.ToString() + "+").ToString()
                                                Next

                                            End If


                                        Catch ex As Exception

                                        End Try
                                        table.Rows.Add(finalrow)
                                        Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "Rnf" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString())
                                        table.Columns.RemoveAt(0)
                                        table.Columns.RemoveAt(0)
                                        table.Columns.RemoveAt(0)
                                        table.Columns.RemoveAt(0)
                                        table.Columns.RemoveAt(0)

                                        If HRN > 20 Then

                                            For index = 21 To HRN
                                                table.Columns.Remove(index.ToString() + "+")
                                            Next

                                        End If

                                        listobject.DataSource = table
                                        rowcount += table.Rows.Count + 2
                                        listobject.AutoSetDataBoundColumnHeaders = True
                                        listobject.ShowAutoFilter = False
                                        listobject.ListColumns(2).Range.NumberFormat = "dd/MM/yyyy"
                                        'Dim rowstring As String = listobject.Range.Address.Split({":"c}, StringSplitOptions.None)(1)
                                        'Dim rowstring1 As String = rowstring.Substring(4, rowstring.Length - 4)
                                        'Dim row As Integer = Convert.ToInt32(rowstring1)
                                        'Dim totalspots As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(rowcount + ((index1 + 1) * 9), 1), vstoWorkbook.Cells(rowcount + ((index1 + 1) * 9), 1))
                                        'totalspots.Value2 = "Total Number of Spots: " + periodElement.Attribute("NumberOfSpots").Value
                                    End If

                                Else
                                    For index = 0 To dtweekss.Rows.Count - 1
                                        Dim periodElement As XElement = GetPeriodElementForMG(mgroup, dtweekss.Rows(index)("WeekNumber").ToString())
                                        Dim startdate As String = Convert.ToDateTime(dtweekss.Rows(index)("StartDate").ToString()).ToString("dd/MM/yyyy")
                                        Dim enddate As String = Convert.ToDateTime(dtweekss.Rows(index)("EndDate").ToString()).ToString("dd/MM/yyyy")
                                        Dim periodrange As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells((1 * rowcount + 3) + 2 + (2 * index), 1), vstoWorkbook.Cells((1 * rowcount + 3) + 2 + (2 * index), 1))
                                        Try
                                            periodrange.Value2 = "Period : " + startdate + "To " + enddate

                                        Catch ex As Exception
                                            periodrange.Value2 = "Period : "
                                        End Try
                                        Dim universeRange As Microsoft.Office.Interop.Excel.Range = periodrange.Next(2, 0)
                                        Try
                                            universeRange.Value2 = "Univserse: " + periodElement.Attribute("universe").Value
                                        Catch ex As Exception
                                            universeRange.Value2 = "Univserse: "
                                        End Try
                                        Dim scountrange As Microsoft.Office.Interop.Excel.Range = universeRange.Next(1, 2)
                                        Try
                                            scountrange.Value2 = "Sample Count: " + periodElement.Attribute("SampleCount").Value
                                        Catch ex As Exception
                                            scountrange.Value2 = "Sample Count: "
                                        End Try
                                        Dim middate As Microsoft.Office.Interop.Excel.Range = scountrange.Next(1, 2)
                                        Try
                                            middate.Value2 = "Mid Date: " + Convert.ToDateTime(periodElement.Attribute("MidDate").Value).ToString("dd/MM/yyyy")

                                        Catch ex As Exception
                                            middate.Value2 = "Mid Date: "
                                        End Try
                                        Dim filter As String = String.Format("TG='{0}' and MG='{1}' and WeekNum={2}", plantgname, mgroup, Convert.ToInt32(dtweekss.Rows(index)("WeekNumber").ToString()))
                                        Dim rnfrows As Data.DataRow() = RnFShowResultsTable.Select(filter)


                                        ' dttt.Rows.
                                        If rnfrows.Count > 0 Then
                                            'Dim table As Data.DataTable = New Data.DataTable()
                                            'table = RnfShowResultsTable.Clone()
                                            'For Each row1 As Data.DataRow In rnfrows
                                            '    table.ImportRow(row1)
                                            'Next
                                            Dim lrange As Microsoft.Office.Interop.Excel.Range = middate.Next(2, -4)
                                            Dim table As Data.DataTable = rnfrows.CopyToDataTable()

                                            ' table.Columns.Remove("End Time")
                                            Dim blankrow As Data.DataRow = table.NewRow()
                                            table.Rows.InsertAt(blankrow, table.Rows.Count)
                                            Dim finalrow As Data.DataRow = table.NewRow()
                                            periodcount += 1
                                            '                ' AvgFreq="1.03" CummCPRP="0" CummCost="0"
                                            'EndDate="20130810" GRP="64.97" GRP000s="13502.34"
                                            'MidDate="2013-08-07" NumberOfSpots="9" R1="62.85"
                                            'R2="25.07" R3="18.21" R4="14.98" R5="11.34" R6="8.18"
                                            'R7="5.58" R8="3.49" R9="0" Reach000s="9833.79"
                                            'SampleCount="2165.0" SpotCPRP="0"
                                            Try


                                                '  finalrow("GRP000s") = periodElement.Attribute("GRP000s").Value
                                                finalrow("GRP000s") = table.Rows(table.Rows.Count - 2)("GRP000s").ToString()
                                                ' finalrow("GRP") = periodElement.Attribute("GRP").Value
                                                '   finalrow("AvgFreq") = periodElement.Attribute("AvgFreq").Value
                                                finalrow("CummCost") = periodElement.Attribute("CummCost").Value
                                                finalrow("SpotCPRP") = periodElement.Attribute("SpotCPRP").Value
                                                finalrow("CummCPRP") = periodElement.Attribute("CummCPRP").Value
                                                ' finalrow("Reach000s") = periodElement.Attribute("Reach000s").Value
                                                '  finalrow("Reach000s") = table.AsEnumerable().Last()("Reach000s").ToString()
                                                finalrow("Reach000s") = table.Rows(table.Rows.Count - 2)("Reach000s").ToString()
                                                '   Dim reachValues As Decimal() = New Decimal() {periodElement.Attribute("R1").Value, periodElement.Attribute("R2").Value, periodElement.Attribute("R3").Value, periodElement.Attribute("R4").Value, periodElement.Attribute("R5").Value, periodElement.Attribute("R6").Value, periodElement.Attribute("R7").Value, periodElement.Attribute("R8").Value, periodElement.Attribute("R9").Value, periodElement.Attribute("R10").Value, periodElement.Attribute("R11").Value, periodElement.Attribute("R12").Value, periodElement.Attribute("R13").Value, periodElement.Attribute("R14").Value, periodElement.Attribute("R15").Value, periodElement.Attribute("R16").Value, periodElement.Attribute("R17").Value, periodElement.Attribute("R18").Value, periodElement.Attribute("R19").Value, periodElement.Attribute("R20").Value}
                                                '  Dim rValues As Decimal() = nbd(periodElement.Attribute("GRP").Value, reachValues)
                                                'finalrow("1+") = rValues(3)
                                                ''   finalrow("GRP") = Convert.ToDecimal(finalrow("AvgFreq").ToString) * Convert.ToDecimal(finalrow("1+").ToString())
                                                'finalrow("GRP") = rValues(0)
                                                'finalrow("AvgFreq") = rValues(1)
                                                'finalrow("2+") = rValues(4)
                                                'finalrow("3+") = rValues(5)
                                                'finalrow("4+") = rValues(6)
                                                'finalrow("5+") = rValues(7)
                                                'finalrow("6+") = rValues(8)
                                                'finalrow("7+") = rValues(9)
                                                'finalrow("8+") = rValues(10)
                                                'finalrow("9+") = rValues(11)
                                                'finalrow("10+") = rValues(12)
                                                'finalrow("11+") = rValues(13)
                                                'finalrow("12+") = rValues(14)
                                                'finalrow("13+") = rValues(15)
                                                'finalrow("14+") = rValues(16)
                                                'finalrow("15+") = rValues(17)
                                                'finalrow("16+") = rValues(18)
                                                'finalrow("17+") = rValues(19)
                                                'finalrow("18+") = rValues(20)
                                                'finalrow("19+") = rValues(21)
                                                'finalrow("20+") = rValues(22)
                                                finalrow("GRP") = table.Rows(table.Rows.Count - 2)("GRP").ToString()
                                                finalrow("AvgFreq") = table.Rows(table.Rows.Count - 2)("AvgFreq").ToString()
                                                If HRN > 20 Then


                                                    finalrow("1+") = table.Rows(table.Rows.Count - 2)("1+").ToString()
                                                    '  finalrow("GRP") = Convert.ToDecimal(finalrow("AvgFreq").ToString) * Convert.ToDecimal(finalrow("1+").ToString())

                                                    finalrow("2+") = table.Rows(table.Rows.Count - 2)("2+").ToString()
                                                    finalrow("3+") = table.Rows(table.Rows.Count - 2)("3+").ToString()
                                                    finalrow("4+") = table.Rows(table.Rows.Count - 2)("4+").ToString()
                                                    finalrow("5+") = table.Rows(table.Rows.Count - 2)("5+").ToString()
                                                    finalrow("6+") = table.Rows(table.Rows.Count - 2)("6+").ToString()
                                                    finalrow("7+") = table.Rows(table.Rows.Count - 2)("7+").ToString()
                                                    finalrow("8+") = table.Rows(table.Rows.Count - 2)("8+").ToString()
                                                    finalrow("9+") = table.Rows(table.Rows.Count - 2)("9+").ToString()
                                                    finalrow("10+") = table.Rows(table.Rows.Count - 2)("10+").ToString()
                                                    finalrow("11+") = table.Rows(table.Rows.Count - 2)("11+").ToString()
                                                    finalrow("12+") = table.Rows(table.Rows.Count - 2)("12+").ToString()
                                                    finalrow("13+") = table.Rows(table.Rows.Count - 2)("13+").ToString()
                                                    finalrow("14+") = table.Rows(table.Rows.Count - 2)("14+").ToString()
                                                    finalrow("15+") = table.Rows(table.Rows.Count - 2)("15+").ToString()
                                                    finalrow("16+") = table.Rows(table.Rows.Count - 2)("16+").ToString()
                                                    finalrow("17+") = table.Rows(table.Rows.Count - 2)("17+").ToString()
                                                    finalrow("18+") = table.Rows(table.Rows.Count - 2)("18+").ToString()
                                                    finalrow("19+") = table.Rows(table.Rows.Count - 2)("19+").ToString()
                                                    finalrow("20+") = table.Rows(table.Rows.Count - 2)("20+").ToString()
                                                Else
                                                    For index11 = 1 To HRN
                                                        finalrow(index11.ToString() + "+") = table.Rows(table.Rows.Count - 2)(index11.ToString() + "+").ToString()
                                                    Next

                                                End If
                                            Catch ex As Exception

                                            End Try
                                            table.Rows.Add(finalrow)
                                            Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "Rnf" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString())
                                            table.Columns.RemoveAt(0)
                                            table.Columns.RemoveAt(0)
                                            table.Columns.RemoveAt(0)
                                            table.Columns.RemoveAt(0)
                                            table.Columns.RemoveAt(0)

                                            If HRN > 20 Then

                                                For indeex = 21 To HRN
                                                    table.Columns.Remove(indeex.ToString() + "+")
                                                Next

                                            End If
                                            listobject.DataSource = table
                                            rowcount += table.Rows.Count + 2
                                            listobject.AutoSetDataBoundColumnHeaders = True
                                            listobject.ShowAutoFilter = False
                                            listobject.ListColumns(2).Range.NumberFormat = "dd/MM/yyyy"
                                            'Dim rowstring As String = listobject.Range.Address.Split({":"c}, StringSplitOptions.None)(1)
                                            'Dim rowstring1 As String = rowstring.Substring(4, rowstring.Length - 4)
                                            'Dim row As Integer = Convert.ToInt32(rowstring1)
                                            'Dim totalspots As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(rowcount + ((index1 + 1) * 9), 1), vstoWorkbook.Cells(rowcount + ((index1 + 1) * 9), 1))
                                            'totalspots.Value2 = "Total Number of Spots: " + periodElement.Attribute("NumberOfSpots").Value

                                        End If

                                    Next
                                End If
                            Next
                        Catch ex As Exception
                            LogMpsrintExException("Exception occured while displaying reach and frequency values." + ex.Message)
                        End Try

                        'Try

                        EnableSummaryButtons()
                        'If Not SpotSelectionPane Is Nothing Then
                        '    SpotSelectionPane.Visible = False
                        '    SpotSelectionPane.Dispose()
                        'End If

                        'mpTpSpotSelection = New ucSpotSelection()
                        'currentLineItem = xecellineItemsTable.AsEnumerable().First()("GUID").ToString()
                        'Dim filter As String = String.Format("GUID = '{0}'", xecellineItemsTable.AsEnumerable().First()("GUID").ToString())
                        'Dim rows As Data.DataRow() = RnFSelectedSpots.Select(filter)
                        ''Dim spotst As Data.DataTable = New Data.DataTable()
                        ''spotst = RnFSelectedSpots.Clone()
                        ''For Each row1 As Data.DataRow In rows
                        ''    spotst.ImportRow(row1)
                        ''Next
                        'mpTpSpotSelection.dgSelectedSpotsGrid.DataSource = rows.CopyToDataTable()
                        'HideSelectedSpotsGrid(mpTpSpotSelection.dgSelectedSpotsGrid)
                        ''  mpTpSpotSelection.Anchor = AnchorStyles.Bottom

                        ''mpTpSpotSelection.Anchor = (AnchorStyles.Bottom Or AnchorStyles.Right Or AnchorStyles.Top Or AnchorStyles.Left)
                        '' mpTpSpotSelection.dgSelectedSpotsGrid.Refresh()
                        'SpotSelectionPane = Globals.ThisAddIn.CustomTaskPanes.Add(mpTpSpotSelection, "Spot Selection")
                        'SpotSelectionPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
                        'SpotSelectionPane.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange
                        'mpTpSpotSelection.Dock = DockStyle.Fill
                        '' SpotSelectionPane.
                        'SpotSelectionPane.Width = 800
                        'SpotSelectionPane.Height = 400

                        ''               Dim paneSize As System.Drawing.Size = _
                        ''New System.Drawing.Size(SpotSelectionPane.Width, SpotSelectionPane.Height)
                        ''               mpTpSpotSelection.AutoSize = True
                        ''               mpTpSpotSelection.AutoSizeMode = AutoSizeMode.GrowAndShrink
                        '' mpTpSpotSelection.si.FlowPanel.Size = paneSize
                        ''SpotSelectionPane.Width = 300
                        ''SpotSelectionPane.Height = 200
                        'DisplayCurrentPlanItem()
                        'SpotSelectionPane.Visible = True
                        '                Catch ex As Exception
                        '    LogMpsrintExException("Exception occured while displaying spot selection pane." + ex.Message)

                        'End Try
                    End If
                Else
                    MessageBox.Show("Unable to retreive Reach and Frequency details from server.")
                End If
                SetNormalCursor()

                If Not (frm Is Nothing) Then
                    frm.Close()

                End If

            End If
        Catch ex As Exception
            SetNormalCursor()
            If Not (frm Is Nothing) Then
                frm.Close()

            End If
            LogMpsrintExException("Exception occured while retreiving Reach and Frequency details." + ex.Message)
            MessageBox.Show("Exception occured while retreiving Reach and Frequency details.Please view the Error log for more details")
        End Try
        ' ObjectDumper.Write(smallOrders)
    End Sub
    Public Function ReverseCalculateTVRReach(ByVal table As Data.DataTable) As Data.DataTable
        Dim finalRnFReachTable As Data.DataTable = New Data.DataTable

        Try
            Dim mgs As List(Of String) = New List(Of String)()
            For index1 = 0 To tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                mgs.Add(tpSelections.UcMarkets1.lbPlan.Items(index1).ToString())
            Next
            mgs.Add("TotalMarkets")
            For Each mg In mgs
                Dim rows As Data.DataRow() = table.Select(String.Format("MG='{0}'", mg))
                Dim tempTable As Data.DataTable = CalculateReverseReachForEachMG(rows.CopyToDataTable())
                finalRnFReachTable.Merge(tempTable)
            Next

        Catch ex As Exception
            LogMpsrintExException("Exception occured while reverse calculation of TVR and Reach.Message:" + ex.Message)
            Throw ex
        End Try
        Return finalRnFReachTable
    End Function
    Private Function CalculateReverseReachForEachMG(ByVal table As Data.DataTable)
        Try
            For index = 0 To table.Rows.Count - 1
                Dim row As Data.DataRow = table.Rows(index)
                Dim periodElement As XElement
                Dim universe As Decimal
                If table.Columns.Contains("MG") Then

                    If Convert.ToInt32(row("WeekNum").ToString()) = 0 Then
                        periodElement = GetPeriodElementForMG(row("MG").ToString(), String.Empty)
                    Else
                        periodElement = GetPeriodElementForMG(row("MG").ToString(), Convert.ToInt32(row("WeekNum").ToString()))
                    End If


                    universe = Convert.ToDecimal(periodElement.Attribute("universe").Value)
                Else
                    '  periodElement = GetPeriodElementForMG(row("Market").ToString(), String.Empty)
                    universe = Convert.ToDecimal(row("Universe").ToString())
                End If



                Dim r1 As Decimal = Convert.ToDecimal(row("1+").ToString())
                row("Reach000s") = Math.Round((universe * r1 / 100), 0)

                If index > 0 Then
                    row("TVR") = Convert.ToDecimal(row("GRP").ToString()) - Convert.ToDecimal(table.Rows(index - 1)("GRP").ToString())
                End If
                table.Rows(index)("TVR000s") = Math.Round((universe * Convert.ToDecimal(row("TVR").ToString()) / 100), 0)
                row("GRP000s") = Math.Round((universe * Convert.ToDecimal(row("GRP").ToString()) / 100), 0)
            Next
            table.AcceptChanges()
        Catch ex As Exception
            LogMpsrintExException("Exception occured while reverse calculating reach values for each MG.Message :" + ex.Message)
        End Try
        Return table
    End Function
    Private Function EnableSummaryButtons()
        Try
            btnChannelSummary.Enabled = True
            btnMarketSummary.Enabled = True
            btnCreativeSummary.Enabled = True
            btnDurationSummary.Enabled = True
            btnLogFile.Enabled = True
            btnAllSummary.Enabled = True
            btnSpotReplace.Enabled = True
            btnDeleteSpot.Enabled = True
            btnSpotSelect.Enabled = True
        Catch ex As Exception

        End Try
    End Function
    Public Function nbdmain(ByVal dt As Data.DataTable) As Data.DataTable
        '  Dim dtfinal As Data.DataTable = dt.Clone()

        'DataTable dsTak = new DataTable();

        'Commented to include grpfinal and avgfreq final values
        ' Dim flArray As Decimal() = New Decimal(19) {}
        '  Dim outputArray As Decimal() = New Decimal(19) {}
        Dim flArray1 As List(Of Decimal) = New List(Of Decimal)()
        ' Dim flArray2 As List(Of Decimal) = New List(Of Decimal)()
        Dim outputArray As Decimal() = New Decimal(19) {}
        For Each drnew As DataRow In dt.Rows
            Dim tp As Decimal = CDec(drnew("GRP"))
            'For i As Integer = 22 To dt.Columns.Count - 5
            '    flArray(i - 22) = CDec(drnew(i))
            'Next
            flArray1.Clear()
            For i As Integer = 22 To dt.Columns.Count - 1
                flArray1.Add(CDec(drnew(i)))
            Next
            '   Dim zero
            Dim length As Integer = 0
            For index = 1 To flArray1.Count

                If flArray1(index - 1) = 0 Then
                    ' flArray2.Add(0)
                    length = index
                    Exit For
                    'Else
                    ' flArray2.Add(flArray1(index - 1))
                End If

            Next

            'Dim flarray As Decimal() = New Decimal(flArray2.Count) {}
            'flArray.FindIndex(System.Predicate(Of Decimal))
            ' Dim items = From s In flArray _
            'Where s = 0 _
            'Select s Take 1
            outputArray = nbd(tp, flArray1.ToArray, length)
            'drnew("GRP") = outputArray
            'Dim drout As DataRow = dtfinal.NewRow()
            Dim d As Decimal = Convert.ToDecimal(drnew("GRP").ToString())
            ' Try
            '  Dim avgFreq As Decimal
            ' If Decimal.TryParse(drnew("AvgFreq"), avgFreq) Then
            '  drnew("GRP") = avgFreq * outputArray(1)

            If outputArray.Length > 1 Then
                drnew("GRP") = outputArray(0)
                drnew("AvgFreq") = outputArray(1)
                '  End If
                'Catch ex As Exception
                '    drnew("GRP") = d
                'End Try

                For x As Integer = 3 To outputArray.Length - 1
                    drnew(x + 19) = outputArray(x)
                Next
            End If



            'dtfinal.Rows.Add(drout)
            'dtfinal.AcceptChanges()
        Next
        dt.AcceptChanges()
        Return dt
    End Function
    Public Function nbdmainMarketSummary(ByVal dt As Data.DataTable) As Data.DataTable
        '  Dim dtfinal As Data.DataTable = dt.Clone()

        'DataTable dsTak = new DataTable();


        Dim flArray As List(Of Decimal) = New List(Of Decimal)()

        Dim outputArray As Decimal() = New Decimal(19) {}

        For Each drnew As DataRow In dt.Rows
            Dim tp As Decimal = CDec(drnew("GRP"))
            'For i As Integer = 5 To dt.Columns.Count - 3
            '    flArray(i - 5) = CDec(drnew(i))
            'Next

            For index = 1 To Globals.Ribbons.MSprintExRibbon.MSHRN
                Dim colname As String = index.ToString() + "+"
                flArray.Add(CDec(drnew(colname)))
            Next

            Dim length As Integer = 0
            For index = 1 To flArray.Count

                If flArray(index - 1) = 0 Then
                    ' flArray2.Add(0)
                    length = index
                    Exit For
                    'Else
                    ' flArray2.Add(flArray1(index - 1))
                End If

            Next
            outputArray = nbd(tp, flArray.ToArray(), length)
            ' drnew("GRP") = outputArray
            'Dim drout As DataRow = dtfinal.NewRow()
            drnew("GRP") = outputArray(0)
            If dt.Columns.Contains("AOTS") Then

                drnew("AOTS") = outputArray(1)
                'Dim aotsValue As Decimal
                'If Decimal.TryParse(drnew("AOTS").ToString, aotsValue) Then
                '    drnew("GRP") = aotsValue * outputArray(1)
                'End If


            End If


            For x As Integer = 3 To outputArray.Length - 1
                drnew(x + 2) = outputArray(x)
            Next

            'dtfinal.Rows.Add(drout)
            'dtfinal.AcceptChanges()
        Next
        dt.AcceptChanges()
        Return dt
    End Function
    Public Function nbdmainSummary(ByVal dt As Data.DataTable, ByVal hrn As Integer) As Data.DataTable
        '  Dim dtfinal As Data.DataTable = dt.Clone()

        'DataTable dsTak = new DataTable();


        Dim flArray As List(Of Decimal) = New List(Of Decimal)()
        Dim outputArray As Decimal() = New Decimal(19) {}


        For Each drnew As DataRow In dt.Rows
            Dim tp As Decimal = CDec(drnew("GRP"))
            'For i As Integer = 6 To dt.Columns.Count - 6
            '    flArray(i - 6) = CDec(drnew(i))
            'Next

            For index = 1 To hrn
                Dim colname As String = index.ToString() + "+"
                flArray.Add(CDec(drnew(colname)))
            Next

            Dim length As Integer = 0
            For index = 1 To flArray.Count

                If flArray(index - 1) = 0 Then
                    ' flArray2.Add(0)
                    length = index
                    Exit For
                    'Else
                    ' flArray2.Add(flArray1(index - 1))
                End If

            Next
            outputArray = nbd(tp, flArray.ToArray(), length)
            drnew("GRP") = outputArray(0)
            ' drnew("GRP") = outputArray
            'Dim drout As DataRow = dtfinal.NewRow()
            '   drnew("GRP") = drnew("AvgFreq") * outputArray(1)
            If dt.Columns.Contains("AOTS") Then
                'Dim aotsvalue As Decimal
                drnew("AOTS") = outputArray(1)
                'If Decimal.TryParse(drnew("AOTS").ToString(), aotsvalue) Then
                '    drnew("GRP") = drnew("AOTS") * outputArray(1)
                'End If

            End If
            For x As Integer = 3 To outputArray.Length - 1
                drnew(x + 3) = outputArray(x)
            Next

            'dtfinal.Rows.Add(drout)
            'dtfinal.AcceptChanges()
        Next
        dt.AcceptChanges()
        Return dt
    End Function

    Public Function nbd(ByVal tpval As Decimal, ByVal fcttemp As Decimal(), ByVal length As Integer) As Decimal()
        Dim [error] As Decimal() = New Decimal() {CDec(-1)}
        Try
            Dim tp As Decimal = tpval
            Dim len As Integer = length
            Dim fct As Decimal() = New Decimal(len) {}
            fct(0) = 0

            If fcttemp.Contains(CDec(100)) Then
                Return [error]
            Else
                For T As Integer = 1 To len - 1
                    fct(T) = fcttemp(T - 1) / 100
                Next
                fct(0) = 1 - fct(1)


                Dim fc As Decimal() = New Decimal(fct.Length - 1) {}
                fc(0) = fct(0)

                For T As Integer = 1 To fct.Length - 2
                    fc(T) = fct(T) - fct(T + 1)
                Next
                fc(fct.Length - 1) = fct(fct.Length - 1)

                Dim tc As Decimal = CDec(0)

                For T As Integer = 1 To fc.Length - 1
                    tc = tc + T * fc(T)
                Next

                tc = tc * 100

                Dim c As Decimal = tc / (100 * CDec(Math.Log(CDbl(fc(0)))))
                Dim pc As Decimal() = New Decimal(fc.Length - 1) {}
                Dim pp As Decimal() = New Decimal(fc.Length - 1) {}
                Dim fp As Decimal() = New Decimal(fc.Length - 1) {}
                Dim mysum As Decimal
                If c >= -1 Then
                    Dim lamb As Decimal = tc / 100
                    Dim lambS As Decimal = tp / 100
                    pc(0) = Math.Exp(-lamb)
                    pp(0) = Math.Exp(-lambS)
                    fp(0) = pp(0) + fc(0) - pc(0)
                    '  Dim mysum As Decimal
                    mysum = fp(0)

                    For index = 1 To fp.Length - 1
                        pc(index) = lamb * pc(index - 1) / (index)
                        pp(index) = lambS * pp(index - 1) / (index)
                        fp(index) = pp(index) + fc(index) - pc(index)
                        mysum = mysum + fp(index)
                    Next

                    '                    for loop in range(1,len(fp)-1):
                    'pc[loop] = lamb*pc[loop-1]/(loop)
                    'pp[loop] = lambS*pp[loop-1]/(loop)
                    'fp[loop] = pp[loop] + fc[loop] - pc[loop]
                    'mysum = mysum + fp[loop]
                    'Return [error]
                Else

                    Dim a As Decimal = -2 * (1 + c)
                    ' decimal a = (decimal)5.0529863205884977;
                    Dim nbd_a As Decimal = nbdparams(a, c)

                    Dim k As Decimal = tc / (100 * nbd_a)

                    Dim ap As Decimal = nbd_a * tp / tc

                    'Dim pc As Decimal() = New Decimal(fc.Length - 1) {}
                    'Dim pp As Decimal() = New Decimal(fc.Length - 1) {}
                    'Dim fp As Decimal() = New Decimal(fc.Length - 1) {}

                    pc(0) = fc(0)
                    pp(0) = CDec(Math.Pow(CDbl(1 / CDbl(1 + ap)), CDbl(k)))
                    fp(0) = pp(0)

                    Dim anew As Decimal = nbd_a / (1 + nbd_a)
                    Dim apnew As Decimal = ap / (1 + ap)
                    mysum = fp(0)
                    For T As Integer = 1 To fp.Length - 2
                        Dim x As Decimal = (k + T - 1) / T
                        pc(T) = x * anew * pc(T - 1)
                        pp(T) = x * apnew * pp(T - 1)
                        fp(T) = pp(T) + fc(T) - pc(T)
                        mysum = mysum + fp(T)
                    Next
                    'fp(fp.Length - 1) = 1 - mysum

                    'Dim fpf As Decimal() = New Decimal(fp.Length - 1) {}

                    'fpf(fp.Length - 1) = fp(fp.Length - 1)

                    'For S As Integer = fp.Length - 2 To 1 Step -1
                    '    fpf(S) = fp(S) + fpf(S + 1)
                    'Next
                    'fpf(0) = 1 - fpf(1)

                    'For i As Integer = 0 To fpf.Length - 1
                    '    fpf(i) = CDec(Math.Round(CDec(fpf(i) * 100)))
                    'Next

                    'Return fpf.ToArray()
                End If
                fp(fp.Length - 1) = 1 - mysum
                Dim fpf As Decimal() = New Decimal(fp.Length - 1) {}

                fpf(fp.Length - 1) = fp(fp.Length - 1)

                For S As Integer = fp.Length - 2 To 1 Step -1
                    fpf(S) = fp(S) + fpf(S + 1)
                Next
                fpf(0) = 1 - fpf(1)

                'Commented due to include grpfinal and avgfreq final

                'For i As Integer = 0 To fpf.Length - 1
                '    fpf(i) = CDec(Math.Round(CDec(fpf(i) * 100)))
                'Next

                'Return fpf.ToArray()

                ' ' GRP smoothening and frequency calculation 
                Dim fpf_grp As Decimal() = New Decimal(fpf.Length - 2) {}
                Dim grp_final As Decimal = CDec(0)
                For T As Integer = 1 To fpf.Length - 2

                    fpf_grp(T - 1) = fpf(T) - fpf(T + 1)
                Next

                fpf_grp(fpf_grp.Length - 1) = fpf(fpf.Length - 1)
                For N As Integer = 0 To fpf_grp.Length - 1
                    grp_final = grp_final + (fpf_grp(N) * (N + 1))
                Next

                Dim freq_final As Decimal = CDec(0)
                freq_final = grp_final / CDec(fpf(1))

                For i As Integer = 0 To fpf.Length - 1
                    fpf(i) = CDec(Math.Round(CDec(fpf(i) * 100)))
                Next

                Dim fpf_finalArray As Decimal() = New Decimal(fpf.Length + 1) {}
                fpf_finalArray(0) = grp_final * 100
                fpf_finalArray(1) = freq_final

                For p As Integer = 2 To fpf_finalArray.Length - 1
                    fpf_finalArray(p) = fpf(p - 2)
                Next

                Return fpf_finalArray.ToArray()
            End If

        Catch e As Exception
            Return [error]
        End Try

    End Function
    'Public Function nbd(ByVal tpval As Decimal, ByVal fcttemp As Decimal()) As Decimal()
    '    Dim [error] As Decimal() = New Decimal() {CDec(-1)}
    '    Try
    '        Dim tp As Decimal = tpval

    '        Dim fct As Decimal() = New Decimal(fcttemp.Length) {}
    '        fct(0) = 0

    '        If fcttemp.Contains(CDec(100)) Then
    '            Return [error]
    '        Else
    '            For T As Integer = 1 To fcttemp.Length - 1
    '                fct(T) = fcttemp(T - 1) / 100
    '            Next
    '            fct(0) = 1 - fct(1)


    '            Dim fc As Decimal() = New Decimal(fct.Length - 1) {}
    '            fc(0) = fct(0)

    '            For T As Integer = 1 To fct.Length - 2
    '                fc(T) = fct(T) - fct(T + 1)
    '            Next
    '            fc(fct.Length - 1) = fct(fct.Length - 1)

    '            Dim tc As Decimal = CDec(0)

    '            For T As Integer = 1 To fc.Length - 1
    '                tc = tc + T * fc(T)
    '            Next

    '            tc = tc * 100

    '            Dim c As Decimal = tc / (100 * CDec(Math.Log(CDbl(fc(0)))))

    '            If c >= -1 Then
    '                Return [error]
    '            Else

    '                Dim a As Decimal = -2 * (1 + c)
    '                ' decimal a = (decimal)5.0529863205884977;
    '                Dim nbd_a As Decimal = nbdparams(a, c)

    '                Dim k As Decimal = tc / (100 * nbd_a)

    '                Dim ap As Decimal = nbd_a * tp / tc

    '                Dim pc As Decimal() = New Decimal(fc.Length - 1) {}
    '                Dim pp As Decimal() = New Decimal(fc.Length - 1) {}
    '                Dim fp As Decimal() = New Decimal(fc.Length - 1) {}

    '                pc(0) = fc(0)
    '                pp(0) = CDec(Math.Pow(CDbl(1 / CDbl(1 + ap)), CDbl(k)))
    '                fp(0) = pp(0)

    '                Dim mysum As Decimal = fp(0)

    '                Dim anew As Decimal = nbd_a / (1 + nbd_a)
    '                Dim apnew As Decimal = ap / (1 + ap)

    '                For T As Integer = 1 To fp.Length - 2
    '                    Dim x As Decimal = (k + T - 1) / T
    '                    pc(T) = x * anew * pc(T - 1)
    '                    pp(T) = x * apnew * pp(T - 1)
    '                    fp(T) = pp(T) + fc(T) - pc(T)
    '                    mysum = mysum + fp(T)
    '                Next
    '                fp(fp.Length - 1) = 1 - mysum

    '                Dim fpf As Decimal() = New Decimal(fp.Length - 1) {}

    '                fpf(fp.Length - 1) = fp(fp.Length - 1)

    '                For S As Integer = fp.Length - 2 To 1 Step -1
    '                    fpf(S) = fp(S) + fpf(S + 1)
    '                Next
    '                fpf(0) = 1 - fpf(1)
    '                ' GRP smoothening and frequency calculation
    '                Dim fpf_grp As Decimal() = New Decimal(fpf.Length - 2) {}
    '                Dim grp_final As Decimal = CDec(0)
    '                For T As Integer = 1 To fpf.Length - 2

    '                    fpf_grp(T - 1) = fpf(T) - fpf(T + 1)
    '                Next

    '                fpf_grp(fpf_grp.Length - 1) = fpf(fpf.Length - 1)
    '                For N As Integer = 0 To fpf_grp.Length - 1
    '                    grp_final = grp_final + (fpf_grp(N) * (N + 1))
    '                Next

    '                Dim freq_final As Decimal = CDec(0)
    '                freq_final = grp_final / CDec(fpf(1))

    '                For i As Integer = 0 To fpf.Length - 1
    '                    fpf(i) = CDec(Math.Round(CDec(fpf(i) * 100)))
    '                Next

    '                Dim fpf_finalArray As Decimal() = New Decimal(fpf.Length + 1) {}
    '                fpf_finalArray(0) = grp_final * 100
    '                fpf_finalArray(1) = freq_final

    '                For p As Integer = 2 To fpf_finalArray.Length - 1
    '                    fpf_finalArray(p) = fpf(p - 2)
    '                Next

    '                Return fpf_finalArray.ToArray()
    '            End If
    '        End If
    '    Catch e As Exception
    '        Return [error]
    '    End Try

    'End Function
    'Old nbdparams method (Working fine)- commented due to inclusion of GRP_Final and AvgFreq
    'Public Function nbdparams(ByVal a As Decimal, ByVal c As Decimal) As Decimal
    '    Dim b As Decimal = a
    '    Dim a1 As Decimal = CDec(Math.Log(CDbl(1 + a)))
    '    Dim atmp As Decimal = c * (a - (1 + a) * a1) / (1 + a + c)
    '    If Math.Abs(CDbl(b - atmp)) < 0.001 Then
    '        Return CDec(atmp)
    '    Else
    '        Return nbdparams(CDec(atmp), c)
    '    End If
    'End Function
    'Public Function DisplayCurrentPlanItem()
    '    ' ExcelPlan.mpTpSpotSelection.lbCurrentItem.Text = "Current Plan Item:"
    '    '  Dim currentPlanItem As Data.DataTable = New Data.DataTable()
    '    '   Dim dr As Data.DataRow = currentPlanItem.NewRow()
    '    Try

    '        If Not (xecelTable Is Nothing) And Not (ExcelPlan.mpTpSpotSelection Is Nothing) Then
    '            ' For index = 1 To xecellineItemsTable.Columns.Count - 3
    '            ' Globals.Ribbons.MSprintExRibbon.mp()

    '            'If Not (currentPlanItem.Columns.Contains(xecellineItemsTable.Columns(index).ColumnName)) Then
    '            '    currentPlanItem.Columns.Add(xecellineItemsTable.Columns(index).ColumnName)
    '            'End If
    '            'dr(xecellineItemsTable.Columns(index).ColumnName) = xecellineItemsTable.AsEnumerable().First()(index).ToString()
    '            Dim row As Data.DataRow() = xecelTable.Select(String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem))
    '            'If xecellineItemsTable.Columns(index).ColumnName = "Start Time" Or xecellineItemsTable.Columns(index).ColumnName = "End Time" Then
    '            '    value = GetStartTime(value)
    '            'End If

    '            If row.Length > 0 Then
    '                ExcelPlan.mpTpSpotSelection.lbChannel.Text = row(0)("Channel").ToString()
    '                ExcelPlan.mpTpSpotSelection.lbProg.Text = row(0)("Programme").ToString()
    '                ExcelPlan.mpTpSpotSelection.lbStartTime.Text = GetStartTime(row(0)("Start Time").ToString())
    '                ExcelPlan.mpTpSpotSelection.lbEndtime.Text = GetStartTime(row(0)("End Time").ToString())
    '                ExcelPlan.mpTpSpotSelection.lbRate.Text = row(0)("RatePer10Sec").ToString()
    '                ExcelPlan.mpTpSpotSelection.lbDays.Text = row(0)("Day").ToString()
    '                ' ExcelPlan.mpTpSpotSelection.tbChannelValue.Text = row(0)("Channel").ToString()
    '            End If

    '            ' ExcelPlan.mpTpSpotSelection.lbCurrentItem.Text = ExcelPlan.mpTpSpotSelection.lbCurrentItem.Text + value

    '            'If index <> xecellineItemsTable.Columns.Count - 3 Then
    '            '    ExcelPlan.mpTpSpotSelection.lbCurrentItem.Text = ExcelPlan.mpTpSpotSelection.lbCurrentItem.Text + ","
    '            'End If

    '            '  Next
    '            'currentPlanItem.Rows.Add(dr)
    '            ' ExcelPlan.mpTpSpotSelection.dgvCurrentLineItem.DataSource = Nothing
    '            'ExcelPlan.mpTpSpotSelection.dgvCurrentLineItem.DataSource = currentPlanItem
    '            '   ExcelPlan.mpTpSpotSelection.lbCurrentItem.Refresh()
    '            ExcelPlan.mpTpSpotSelection.Refresh()
    '        End If


    '    Catch ex As Exception

    '    End Try
    '    'Return text
    'End Function
    Public Function nbdparams(ByVal a As Decimal, ByVal c As Decimal) As Decimal
        Dim b As Decimal = a
        Dim a1 As Decimal = CDec(Math.Log(CDbl(1 + a)))
        Dim atmp As Decimal = c * (a - (1 + a) * a1) / (1 + a + c)
        If Math.Abs(CDbl(b - atmp)) < 0.001 Then
            Return CDec(atmp)
        Else
            Return nbdparams(CDec(atmp), c)
        End If
    End Function

    Public Function GetPeriodElementForMG(ByVal mGroup As String, ByVal weeknum As String) As XElement
        Dim period As XElement = New XElement("period")
        For Each mg As XElement In rnfoutputXml.Element("tg").Elements

            If mg.Attribute("name").Value.Equals(mGroup) Then

                If weeknum.Length = 0 Then
                    Return mg.Element("period")
                Else
                    For Each p As XElement In mg.Elements
                        If p.Attributes().Any() Then
                            If p.Attribute("WeekNum").Value.Equals(weeknum) Then
                                Return p
                            End If

                        End If
                    Next

                End If




            End If

        Next
        Return period
    End Function
    Public Function HideSelectedSpotsGrid(ByVal grid As System.Windows.Forms.DataGridView)
        Try
            'RnFSelectedSpots.Columns.Add("ChannelName")
            'RnFSelectedSpots.Columns.Add("Date", System.Type.GetType("System.DateTime"))
            'RnFSelectedSpots.Columns.Add("StartTime")
            'RnFSelectedSpots.Columns.Add("EndTime")
            'RnFSelectedSpots.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
            '' RnFSelectedSpots.Columns.Add("Commercial")
            'RnFSelectedSpots.Columns.Add("Cost")
            'RnFSelectedSpots.Columns.Add("PA")
            'RnFSelectedSpots.Columns.Add("TA")
            'RnFSelectedSpots.Columns.Add("GUID", System.Type.GetType("System.Int32"))
            'RnFSelectedSpots.Columns.Add("Spot")
            'RnFSelectedSpots.Columns.Add("StartDate", System.Type.GetType("System.DateTime"))
            'RnFSelectedSpots.Columns.Add("EndDate", System.Type.GetType("System.DateTime"))
            'RnFSelectedSpots.Columns.Add("WeekNum", System.Type.GetType("System.Int32"))
            Try
                If grid.Columns.Contains("Date") Then
                    grid.Columns("Date").DefaultCellStyle.Format = "dd/MM/yyyy"
                End If

                '   grid.Columns(""
                'If grid.Columns.Contains("End Time") Then
                '    grid.Columns("End Time").Visible = False
                '    End If
            Catch ex As Exception

            End Try
            grid.Columns("GUID").Visible = False
            grid.Columns("Spot").Visible = False
            grid.Columns("Start Date").Visible = False
            grid.Columns("End Date").Visible = False
            grid.Columns("WeekNum").Visible = False

            '  grid.

        Catch ex As Exception

        End Try
    End Function
    Private Function GetConcatenatedDate(ByVal dateVal As Date)
        Dim shortDateString As String = String.Empty
        shortDateString = dateVal.Year.ToString() + dateVal.Month.ToString() + dateVal.Day.ToString()
        Return shortDateString
    End Function
    Public Function GetSpotRow(ByVal spotString As String) As Data.DataRow
        Dim spot As Data.DataTable = New Data.DataTable()
        Try
            '  This is the order : <!-- Col1=ChannelCode Integer width Col2=Date Text width Col3=TimeFrom Text width Col4=TimeTo Text width Col5=Cost Text width Col6=AP Text Width Col7=TA Text Width  Col8=Duration Text width -->
            ' spot.Columns.Add("ChannelCode")
            spot.Columns.Add("Date", System.Type.GetType("System.DateTime"))
            ' spot.Columns.Add("StartTime", System.Type.GetType("System.TimeSpan"))
            spot.Columns.Add("StartTime")
            '  spot.Columns.Add("EndTime", System.Type.GetType("System.TimeSpan"))
            ' spot.Columns.Add(
            spot.Columns.Add("Cost")
            spot.Columns.Add("PA")
            spot.Columns.Add("TA")
            spot.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
            Dim values As String() = spotString.Split({","c}, StringSplitOptions.None)
            Dim dr As Data.DataRow = spot.NewRow()
            'For index = 0 To values.Length - 1

            ' If index = 1 Then
            dr("Date") = New Date(Convert.ToInt32(values(1).Substring(0, 4)), Convert.ToInt32(values(1).Substring(4, 2)), Convert.ToInt32(values(1).Substring(6, 2))).ToShortDateString()
            ' ElseIf index = 2 Then
            ' dr("StartTime") = New TimeSpan(Convert.ToInt32(values(2).Substring(0, 2)), Convert.ToInt32(values(2).Substring(2, 2)), Convert.ToInt32(values(2).Substring(4, 2)))
            '  dr("StartTime") = values(2).Substring(0, 5)
            dr("StartTime") = String.Format("{0}:{1}", values(2).Substring(0, 2), values(2).Substring(2, 2))
            dr("Cost") = values(4)
            dr("PA") = values(5)
            dr("TA") = values(6)
            dr("Duration(Sec)") = values(7)
            'ElseIf index = 3 Then
            '    dr("EndTime") = New TimeSpan(Convert.ToInt32(values(3).Substring(0, 2)), Convert.ToInt32(values(3).Substring(2, 2)), Convert.ToInt32(values(3).Substring(4, 2)))

            ' Else
            ' dr(index) = values(index)
            ' End If

            ' Next
            spot.Rows.Add(dr)
        Catch ex As Exception

        End Try
        Return spot.Rows(0)
    End Function
    Public Function GetReachRow(ByVal reachVal As String)
        Dim reach As Data.DataTable = New Data.DataTable()
        reach.Columns.Add("TVR000s")
        reach.Columns.Add("TVR")
        reach.Columns.Add("GRP000s")
        reach.Columns.Add("GRP")
        reach.Columns.Add("AvgFreq")
        reach.Columns.Add("CummCost")
        reach.Columns.Add("SpotCPRP")
        reach.Columns.Add("CummCPRP")
        reach.Columns.Add("Reach000s")
        reach.Columns.Add("R1")
        reach.Columns.Add("R2")
        reach.Columns.Add("R3")
        reach.Columns.Add("R4")
        reach.Columns.Add("R5")
        reach.Columns.Add("R6")
        reach.Columns.Add("R7")
        reach.Columns.Add("R8")
        reach.Columns.Add("R9")
        reach.Columns.Add("R10")
        '  Dim reachval As String = "703.00,3.38,703.00,3.38,1.33,0,0,0,527.20,2.54,0,0,0,0,0,0,0,0,0"
        Dim vals As String() = reachVal.Split({","c}, StringSplitOptions.None)
        Dim drow As Data.DataRow = reach.NewRow()
        For index = 0 To vals.Length - 1
            drow(index) = vals(index)
        Next
        reach.Rows.Add(drow)
        Return reach.Rows(0)
    End Function

    Private Sub btnShwHideSpotSelection_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = CType(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        If Not (SpotSelectionPane Is Nothing) And sheet.Name.Equals("Plan Selection") Then
            If SpotSelectionPane.Visible Then
                SpotSelectionPane.Visible = False
            Else
                SpotSelectionPane.Visible = True
            End If

        End If
    End Sub
    Public Function PopulateLatestDates(ByVal latestDateXML As XElement)
        Try
            '<database date_from="2013-07-28" date_to="2014-08-16" week_to="33" />
            db_FromDate = Convert.ToDateTime(latestDateXML.Attribute("date_from").Value)
            '  Convert.ToDateTime()
            db_ToDate = Convert.ToDateTime(latestDateXML.Attribute("date_to").Value)
            db_WeekNo = Convert.ToInt32(latestDateXML.Attribute("week_to").Value)
            ' Convert.ToDateTime()
        Catch ex As Exception
            LogMpsrintExException("Exception occured while parsing latest Date xml" + ex.Message)
        End Try
    End Function
    Public Function UpdateUsageReport(ByVal methodName As String, ByVal clientValue As String, ByVal brandValue As String, ByVal xml As XElement, ByVal no_spots As Integer) As Boolean
        Dim inserted As Boolean = False
        Try
            Dim sqlConnection1 As New System.Data.SqlClient.SqlConnection("Server= MUMSQLP01107\GRMINDSQL01;Database=MsprintXTracker;User Id=MSXAdmin;Password=MSXAdmin@123;")
            Dim cmd As New System.Data.SqlClient.SqlCommand
            cmd.CommandType = System.Data.CommandType.Text
            Dim commandText As String = String.Format("INSERT UsageReport (NTUserName,MSprintX_Method_Invoked,Date,Client,Brand,InputXML,No_Of_Spots,SysDate) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}',{6},getdate())", loggedInUserName, methodName, Date.Now.ToString, clientValue, brandValue, xml.ToString(), no_spots)
            '  LogMpsrintExException(commandText)
            cmd.CommandText = commandText
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()
            '  Dim adapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(commandText,
            cmd.ExecuteNonQuery()
            sqlConnection1.Close()
        Catch ex As Exception
            LogMpsrintExException("Exception occured while updating usage report table.Message: " + ex.Message)
        End Try
        Return inserted
    End Function
    Public Function CheckLoggedInUserAccess() As Boolean
        Dim userIsValid As Boolean = False
        Try
            Dim dr As DirectoryEntry = New DirectoryEntry("LDAP://mumadcp01102.ad.insidemedia.net")
            Dim ds As DirectorySearcher = New DirectorySearcher(dr)

            ds.Filter = "(&(objectCategory=user)(memberOf=CN= MUMgrm-GRMITmSprintX-gs,OU=Groups,OU=MUM,OU=APAC,DC=ad,DC=insidemedia,DC=net))"
            Dim searchResultColl As SearchResultCollection = ds.FindAll()
            '  searchResultColl
            ' Dim userprincipalname = Environment.UserName + "@groupm.com"
            Dim userprincipalname = Environment.UserName
            For Each result As SearchResult In searchResultColl

                'If userprincipalname.ToLower().Equals(result.Properties("userprincipalname").Item(0).ToString().ToLower()) Then
                '    userIsValid = True
                '    '  MessageBox.Show("Valid")
                'End If

                If result.Properties("userprincipalname").Item(0).ToString().ToLower().Contains(userprincipalname.ToLower()) Then
                    userIsValid = True
                    loggedInUserName = result.Properties("userprincipalname").Item(0).ToString()
                End If

                ' MessageBox.Show(result.Properties("userprincipalname").Item(0).ToString())
                ' MessageBox.Show("www")
            Next
        Catch ex As Exception
            LogMpsrintExException("Exception occured while validating logged in user.Message: " + ex.Message)
        End Try
        Return userIsValid
    End Function

    Private Sub btnStartMsprint_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnStartMsprint.Click
        Try
            'backGroundWorker = New BackgroundWorker()
            'backGroundWorker.WorkerReportsProgress = True
            'backGroundWorker.WorkerSupportsCancellation = True
            'System.Threading.SynchronizationContext.SetSynchronizationContext(New WindowsFormsSynchronizationContext())
            'backGroundWorker.RunWorkerAsync()

            If CheckLoggedInUserAccess() Then

                Dim planFolderPath As FolderBrowserDialog = New FolderBrowserDialog()

                Globals.ThisAddIn.Application.StatusBar = "Loading MSprintEx application..."

                If Globals.ThisAddIn.Application.Workbooks.Count < 1 Then
                    '  Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(,,
                    Globals.ThisAddIn.Application.Workbooks.Add(System.Type.Missing)
                End If

                '  BackgroundWorker1.RunWorkerAsync()
                If Not (File.Exists(ExceptionLogFilePath)) Then
                    File.Create(ExceptionLogFilePath)
                End If

                If MachineConnectedToInternet() Then
                    btnStartMsprint.Enabled = False
                    'If VariantMasterIsUpdated(db_WeekNo) And System.Windows.Forms.MessageBox.Show("BSL Master seems to be updated.MsprintX will communicate with server to update.This may take some time.Click Yes to update now,No to update later", "BSL Update", MessageBoxButtons.YesNo, MessageBoxIcon.Information).Equals(DialogResult.Yes) Then

                    '    PopulateBSLTabs()
                    '    'Else
                    '    '    Dim bslmaster As XElement = XElement.Load(AppDomain.CurrentDomain.BaseDirectory & "\\Masters\\BSLMaster.xml")
                    '    '    InvokeWebService.PopulateVariantMaster(bslmaster)
                    '    'End If
                    'Else
                    '    Dim path As String = AppDomain.CurrentDomain.BaseDirectory & "Masters\BSLMaster.xml"
                    '    'Dim bslmaster As XElement = XElement.Load(path)
                    '    Dim encode As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")

                    '    ' Pipe the stream to a higher level stream reader with the required encoding format.
                    '    Dim readStream As New IO.StreamReader(path, encode)
                    '    '   Using reader As StreamReader = New StreamReader("C:\\ws\\Ptvr.xml.xml")
                    '    Dim bslmaster As XElement = XElement.Parse(readStream.ReadToEnd())
                    '    InvokeWebService.PopulateVariantMaster(bslmaster)
                    'End If
                    PopulateGenresChannelsMarkets()
                    '   Globals.ThisAddIn.MSprintExTaskPane.Visible = True

                    Globals.ThisAddIn.IsOpen = True
                    Dim latestStatus As XElement = GetLatestWeekDetails(Globals.Ribbons.MSprintExRibbon.GetURLForWS("LatestWeekWSURL_New"), "POST")
                    PopulateLatestDates(latestStatus)
                    frmPrepareServer = New frmPrepareServer()
                    frmPrepareServer.ShowDialog()
                    DisplayTVBuilder()
                    ' Dim rclick As Office.CommandBar = Globals.ThisAddIn.Application.CommandBars.Add("ContextMenu", Office.MsoBarPosition.msoBarPopup, Type.Missing, True)
                    '  EnableDisableButtons(True)
                    Dim pathvalue As XElement
                    Dim pathValid As Boolean = True
                    Try
                        pathvalue = XElement.Load(MasterFolderPath + "\LogDirecPath")

                        If Directory.Exists(pathvalue.Value) Then
                            LogDirectoryPath = pathvalue.Value
                            pathValid = True
                        Else
                            pathValid = False
                        End If

                    Catch ex As Exception
                        pathValid = False
                    End Try

                    If Not (pathValid) Then
                        MessageBox.Show("Please select working directory for plan")
                        If planFolderPath.ShowDialog = DialogResult.OK Then
                            Try
                                Dim path As String = planFolderPath.SelectedPath

                                If Not (Directory.Exists(path + "\Logs")) Then
                                    Directory.CreateDirectory(path + "\Logs")
                                End If
                                LogDirectoryPath = path + "\Logs\"
                                logDirectoryXML = <LogDirPath></LogDirPath>
                                logDirectoryXML.Value = LogDirectoryPath
                                logDirectoryXML.Save(MasterFolderPath + "\LogDirecPath")
                            Catch ex As Exception
                                If Not (Directory.Exists(MasterFolderPath + "\Logs")) Then
                                    Directory.CreateDirectory(MasterFolderPath + "\Logs")
                                End If
                                LogDirectoryPath = MasterFolderPath + "\Logs\"
                                logDirectoryXML = <LogDirPath></LogDirPath>
                                logDirectoryXML.Value = LogDirectoryPath
                                logDirectoryXML.Save(MasterFolderPath + "\LogDirecPath")
                            End Try

                            ' InputXMLFolderPath = planFolderPath.FileName
                        End If
                    End If



                Else
                    System.Windows.Forms.MessageBox.Show("MSprintEx communicates with Server over Internet.Please ensure Internet connectivity and Try again.")
                End If
            Else
                MessageBox.Show("Insufficient rights to use MsprintX.Please contact your system administrator.")
            End If
            Globals.ThisAddIn.Application.StatusBar = String.Empty
        Catch ex As Exception
            Globals.ThisAddIn.Application.StatusBar = String.Empty
            btnStartMsprint.Enabled = True
            LogMpsrintExException("Exception occured while starting MsprintApp." + ex.Message)
            System.Windows.Forms.MessageBox.Show(String.Format("Exception occured while loading MsprintRibbon.Please refer to {0} for details", ExceptionLogFilePath))
        End Try
    End Sub
    Public Function ReturnStoredXMLWeekNumber() As Integer
        Dim week_number As Integer
        Try

        Catch ex As Exception

        End Try
    End Function

    Public Function EnableDisableButtons(ByVal enable As Boolean)
        Try
            btnOpen.Enabled = enable
            ' btnGenreShare.Enabled = enable
            btnChannelShare.Enabled = enable
            btnBreakTVR.Enabled = enable
            btnProgramTVR.Enabled = enable
            menuTopPrograms.Enabled = enable
            menubtnOpenPlan.Enabled = enable
            menuBtnSavePlan.Enabled = enable
            btnChangeLogDir.Enabled = enable
            btnAddPlan.Enabled = enable
            '   btnReorderPlanChannels.Enabled = enable
            menGenreShare.Enabled = enable
            btnGenreShareAll.Enabled = enable
            btnGenreShareTopTen.Enabled = enable
            btnProgAvgTVR.Enabled = enable
            ' btnGenerateEndTime.Enabled = enable

        Catch ex As Exception

        End Try
    End Function
    Public Function EnableDisableEntirePane(ByVal enableDisable As Boolean)
        Try
            ' btnStartMsprint.Enabled = enableDisable
            btnChangeLogDir.Enabled = enableDisable

            btnOpen.Enabled = enableDisable
            ' btnGenreShare.Enabled = enable
            btnChannelShare.Enabled = enableDisable
            btnBreakTVR.Enabled = enableDisable
            btnProgramTVR.Enabled = enableDisable
            menuTopPrograms.Enabled = enableDisable
            menubtnOpenPlan.Enabled = enableDisable
            menuBtnSavePlan.Enabled = enableDisable
            btnChangeLogDir.Enabled = enableDisable
            btnAddPlan.Enabled = enableDisable
            '   btnReorderPlanChannels.Enabled = enable
            menGenreShare.Enabled = enableDisable
            btnGenreShareAll.Enabled = enableDisable
            btnGenreShareTopTen.Enabled = enableDisable
            btnProgAvgTVR.Enabled = enableDisable
            btnGenerateEndTime.Enabled = enableDisable



        Catch ex As Exception

        End Try
    End Function

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

    End Sub

    Private Sub btnTGSelectionShowHide_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnTGSelectionShowHide.Click

        'If Not (MSprintExChannelShare Is Nothing) Then

        '    If MSprintExChannelShare.Visible Then
        '        MSprintExChannelShare.Visible = False
        '    Else
        '        MSprintExChannelShare.Visible = True
        '    End If

        'End If

        'Try
        '    Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        '    Dim text As String = String.Empty
        '    If sheet.Name.Equals("Genre Share") Then

        '        If Not (CTPGenreShare Is Nothing) Then
        '            If CTPGenreShare.Visible Then
        '                CTPGenreShare.Visible = False
        '            Else
        '                CTPGenreShare.Visible = True
        '            End If
        '        End If

        '        ElseIf sheet.Name.Equals("Channel Share") Then

        '        If Not (CTPChannelShare Is Nothing) Then
        '            If CTPChannelShare.Visible Then
        '                CTPChannelShare.Visible = False
        '            Else
        '                CTPChannelShare.Visible = True
        '            End If
        '        End If

        '        ElseIf sheet.Name.Equals("Program TVR") Then

        '        If Not (CTPProgTVR Is Nothing) Then
        '            If CTPProgTVR.Visible Then
        '                CTPProgTVR.Visible = False
        '            Else
        '                CTPProgTVR.Visible = True
        '            End If
        '        End If

        '        ElseIf sheet.Name.Equals("Break TVR") Then

        '        If Not (CTPBreakTVR Is Nothing) Then

        '            If CTPBreakTVR.Visible Then
        '                CTPBreakTVR.Visible = False
        '            Else
        '                CTPBreakTVR.Visible = True
        '            End If

        '        End If

        '        End If
        'Catch ex As Exception

        'End Try
        DisplayRelevantTGPane()
    End Sub
    Public Function DisplayRelevantTGPane()
        Try
            Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            Dim text As String = String.Empty
            If sheet.Name.Equals("Genre Share") Then

                If Not (CTPGenreShare Is Nothing) Then
                    If CTPGenreShare.Visible Then
                        CTPGenreShare.Visible = False
                    Else
                        CTPGenreShare.Visible = True
                        CTPChannelShare.Visible = False
                        CTPProgTVR.Visible = False
                        CTPBreakTVR.Visible = False
                    End If
                End If

            ElseIf sheet.Name.Equals("Channel Share") Then

                If Not (CTPChannelShare Is Nothing) Then
                    If CTPChannelShare.Visible Then
                        CTPChannelShare.Visible = False

                    Else
                        CTPChannelShare.Visible = True
                        CTPGenreShare.Visible = False
                        ' CTPChannelShare.Visible = False
                        CTPProgTVR.Visible = False
                        CTPBreakTVR.Visible = False
                    End If
                End If

            ElseIf sheet.Name.Equals("Program TVR") Then

                If Not (CTPProgTVR Is Nothing) Then
                    If CTPProgTVR.Visible Then
                        CTPProgTVR.Visible = False
                    Else
                        CTPProgTVR.Visible = True
                        CTPGenreShare.Visible = False
                        CTPChannelShare.Visible = False
                        '  CTPProgTVR.Visible = False
                        CTPBreakTVR.Visible = False
                    End If
                End If

            ElseIf sheet.Name.Equals("Break TVR") Then

                If Not (CTPBreakTVR Is Nothing) Then

                    If CTPBreakTVR.Visible Then
                        CTPBreakTVR.Visible = False
                    Else
                        CTPBreakTVR.Visible = True
                        CTPGenreShare.Visible = False
                        CTPChannelShare.Visible = False
                        CTPProgTVR.Visible = False
                        ' CTPBreakTVR.Visible = False
                    End If

                End If

            End If
        Catch ex As Exception

        End Try
    End Function

    'Private Sub channelMapping_ShowMoreChannels() Handles channelMapping.ShowMoreChannels
    '    ShowMoreChannels()
    'End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged

    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

    End Sub

    Private Sub menubtnOpenPlan_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles menubtnOpenPlan.Click
        Dim msprintExPlan As XElement = New XElement("MsprintExPlan")
        Dim planopener As OpenFileDialog = New OpenFileDialog()
        Try
            If Globals.ThisAddIn.Application.Workbooks.Count < 1 Then
                '  Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(,,
                Globals.ThisAddIn.Application.Workbooks.Add(System.Type.Missing)
            End If

            btnAddPlan.Enabled = True
            '"Text files (*.txt)|*.txt|All files (*.*)|*.*"
            planopener.Filter = "MSprintEx Plan Files (*.xml)|*.xml"
            planopener.InitialDirectory = LogDirectoryPath
            btnCleanupplan.Enabled = True
            btnRnF.Enabled = True
            btnMapChannels.Enabled = True
            If planopener.ShowDialog() = DialogResult.OK Then
                Dim plan As XElement
                Try
                    plan = XElement.Load(planopener.FileName)
                Catch ex As Exception
                    MessageBox.Show("Only valid MsprintEx Plan files can be opened.Please retry again")
                End Try


                'If plan.Name.Equals("{MsprintExPlan}") Then

                msprintExPlan = XElement.Load(planopener.FileName)

                If ValidateChosenXML(msprintExPlan) Then
                    PopulateTGMGS(msprintExPlan)
                    PopulatePeriods(msprintExPlan)
                    'Next

                    If msprintExPlan.Element("SelectedSpots").Elements("NewDataSet").Any Then
                        Using reader As XmlReader = msprintExPlan.Element("SelectedSpots").Element("NewDataSet").CreateReader()

                            ' If RnFSelectedSpots Is Nothing Then
                            RnFSelectedSpots = New Data.DataTable
                            ' End If
                            RnFSelectedSpots.ReadXml(reader)

                        End Using

                    End If


                    If msprintExPlan.Element("MappedChannels").Elements("NewDataSet").Any Then
                        Using reader As XmlReader = msprintExPlan.Element("MappedChannels").Element("NewDataSet").CreateReader()

                            '  If mappedchannels Is Nothing Then
                            mappedchannels = New Data.DataTable

                            'End If

                            mappedchannels.ReadXml(reader)
                        End Using
                    End If

                    'If msprintExPlan.Element("AvailableSpots").Elements("NewDataSet").Any Then
                    '    Using reader As XmlReader = msprintExPlan.Element("AvailableSpots").Element("NewDataSet").CreateReader()
                    '        If RnFAvaiSpots Is Nothing Then
                    '            RnFAvaiSpots = New Data.DataTable
                    '        End If
                    '        RnFAvaiSpots.ReadXml(reader)
                    '    End Using
                    'End If
                    Using reader As XmlReader = msprintExPlan.Element("PlanItems").Element("NewDataSet").CreateReader()
                        '  If xecellineItemsTable Is Nothing Then
                        xecellineItemsTable = New Data.DataTable
                        'End If
                        xecellineItemsTable.ReadXml(reader)
                    End Using
                    Try


                        If Not CheckSheetExists("Plan Selection") Then
                            logCreator = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                            logCreator.Name = "Plan Selection"
                        Else
                            'logCreator = CheckAndReturnSheet("Plan Selection")
                            ''  newSheet.UsedRange.Clear()
                            'Globals.Ribbons.MSprintExRibbon.CleanSheet(logCreator)
                            ''logCreator.Delete()
                            ''logCreator = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                            ''logCreator.Name = "Plan Selection"
                            'logCreator.Activate()
                            Dim sheetcount As Integer = CheckAndReturnSheet("Plan Selection")
                            If sheetcount > 0 Then
                                logCreator = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                                Dim name As String = String.Format("Plan Selection({0})", sheetcount)
                                logCreator.Name = name
                            End If
                        End If
                        Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(logCreator)
                        ' vstoWorkbook.Name = "Plan Selection"

                        If vstoWorkbook.Controls.Contains("InputSpotSelection") Then
                            loSpotSelection = DirectCast(vstoWorkbook.Controls.Item("InputSpotSelection"), Microsoft.Office.Tools.Excel.ListObject)
                        Else
                            loSpotSelection = vstoWorkbook.Controls.AddListObject(vstoWorkbook.Range("$A$1"), "InputSpotSelection")
                        End If

                        loSpotSelection.AutoSetDataBoundColumnHeaders = True
                        loSpotSelection.DataSource = xecellineItemsTable
                        loSpotSelection.ListColumns(5).Range.NumberFormat = "hh:mm;@"
                        loSpotSelection.ListColumns(6).Range.NumberFormat = "hh:mm;@"
                        vstoWorkbook.get_Range("A:A", Type.Missing).EntireColumn.Hidden = True
                        PlanSheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    Catch ex As Exception

                    End Try
                    'Display Pane with spots
                    Try


                        'If Not SpotSelectionPane Is Nothing Then
                        '    SpotSelectionPane.Visible = False
                        '    SpotSelectionPane.Dispose()
                        'End If

                        'mpTpSpotSelection = New ucSpotSelection()
                        'Dim filter As String = String.Format("GUID = '{0}'", xecellineItemsTable.AsEnumerable().First()("GUID").ToString())
                        'Dim rows As Data.DataRow() = RnFSelectedSpots.Select(filter)
                        ''Dim spotst As Data.DataTable = New Data.DataTable()
                        ''spotst = RnFSelectedSpots.Clone()
                        ''For Each row1 As Data.DataRow In rows
                        ''    spotst.ImportRow(row1)
                        ''Next
                        'mpTpSpotSelection.dgSelectedSpotsGrid.DataSource = rows.CopyToDataTable()
                        'HideSelectedSpotsGrid(mpTpSpotSelection.dgSelectedSpotsGrid)
                        ''Dim avaiRows As Data.DataRow() = RnFAvaiSpots.Select(filter)
                        ''mpTpSpotSelection.dgvAvailableSpotsGrid.DataSource = avaiRows.CopyToDataTable()
                        ''HideSelectedSpotsGrid(mpTpSpotSelection.dgvAvailableSpotsGrid)
                        ''  mpTpSpotSelection.Anchor = AnchorStyles.Bottom

                        ''mpTpSpotSelection.Anchor = (AnchorStyles.Bottom Or AnchorStyles.Right Or AnchorStyles.Top Or AnchorStyles.Left)
                        '' mpTpSpotSelection.dgSelectedSpotsGrid.Refresh()
                        'SpotSelectionPane = Globals.ThisAddIn.CustomTaskPanes.Add(mpTpSpotSelection, "Spot Selection")
                        'SpotSelectionPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
                        'SpotSelectionPane.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange
                        'mpTpSpotSelection.Dock = DockStyle.Fill
                        '' SpotSelectionPane.
                        'SpotSelectionPane.Width = 800
                        'SpotSelectionPane.Height = 400

                        '               Dim paneSize As System.Drawing.Size = _
                        'New System.Drawing.Size(SpotSelectionPane.Width, SpotSelectionPane.Height)
                        '               mpTpSpotSelection.AutoSize = True
                        '               mpTpSpotSelection.AutoSizeMode = AutoSizeMode.GrowAndShrink
                        ' mpTpSpotSelection.si.FlowPanel.Size = paneSize
                        'SpotSelectionPane.Width = 300
                        'SpotSelectionPane.Height = 200
                        ' DisplayCurrentPlanItem()
                        '    SpotSelectionPane.Visible = True
                        isPlanClean = True
                        planOpenedSuccessfully = True
                        btnGetReqSpots.Enabled = True
                        btnRnF.Enabled = True
                        btnReorderPlanChannels.Enabled = True

                        If Not (RnFSelectedSpots Is Nothing) Then

                            If RnFSelectedSpots.Rows.Count > 0 Then
                                btnMarketSummary.Enabled = True
                                btnChannelSummary.Enabled = True
                                btnDurationSummary.Enabled = True
                                btnCreativeSummary.Enabled = True
                                btnAllSummary.Enabled = True
                                btnLogFile.Enabled = True
                            End If

                        End If

                    Catch ex As Exception

                    End Try
                Else
                    MessageBox.Show("Attempted to open corrupt MSprintEx plan. Please retry with valid MSprintEx plan")
                End If
                'Else
                '    MessageBox.Show("Attempted to open incorrect plan .Please retry opening MSprintEx plan")
                'End If
            End If

        Catch ex As Exception
            LogMpsrintExException("Exception occured while opening MsprintPlan." + ex.Message)
            MessageBox.Show("Exception occured while opening plan.Please refer to Error Log for more details")
        End Try
    End Sub
    Private Function PopulateTGMGS(ByVal msprintplan As XElement)
        Try
            Dim dt As System.Data.DataTable = DirectCast(tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            '  dt.Rows(0)(1) = msprintplan.Element("TG").Value
            Dim tgInPlan As XElement = XElement.Parse(msprintplan.Element("TG").Value)
            Dim found As Boolean = False
            Dim tgNameToBeSaved As String = String.Empty
            If Directory.Exists(tgDirectoryPath) Then
                For index = 0 To Directory.GetFiles(tgDirectoryPath, "*.xml").Count - 1
                    'fileList.Add(Path.GetFileNameWithoutExtension(Directory.GetFiles(System.IO.Path.GetTempPath() + "\\TGS", "*.xml")(index)))
                    Dim tgInLocalMachine As XElement = XElement.Load(Directory.GetFiles(tgDirectoryPath, "*.xml")(index))
                    If XElement.DeepEquals(tgInPlan, tgInLocalMachine) Then
                        found = True
                        tgNameToBeSaved = tgInPlan.Attribute("name").Value
                        Exit For
                    Else
                        Dim tgName As String = tgInPlan.Attribute("name").Value
                        tgInPlan.Attribute("name").Value = tgInLocalMachine.Attribute("name").Value

                        If XElement.DeepEquals(tgInPlan, tgInLocalMachine) Then
                            found = True
                            tgNameToBeSaved = tgInLocalMachine.Attribute("name").Value
                            '  tgInLocalMachine.Attribute("name").Value = tgName
                            ' tgInLocalMachine.Save(tgDirectoryPath + tgInLocalMachine.Attribute("name").Value + ".xml")
                            Exit For
                        Else
                            tgInPlan.Attribute("name").Value = tgName
                            found = False
                        End If
                    End If
                Next
            End If
            If found Then
                dt.Rows(0)(1) = tgNameToBeSaved
            Else
                tgInPlan.Save(tgDirectoryPath + tgInPlan.Attribute("name").Value + ".xml")
                dt.Rows(0)(1) = tgInPlan.Attribute("name").Value
            End If

            tpSelections.UcAudience1.DgPlanRefGrid.DataSource = dt
            tpSelections.UcMarkets1.lbPlan.Items.Clear()
            For Each market As XElement In msprintplan.Element("MGS").Elements
                Dim mgInPlan As XElement = XElement.Parse(msprintplan.Element("TG").Value)
                Dim foundMg As Boolean = False
                Dim mgNameToBeSaved As String = String.Empty
                If Directory.Exists(mgDirectoryPath) Then
                    For index = 0 To Directory.GetFiles(mgDirectoryPath, "*.xml").Count - 1
                        'fileList.Add(Path.GetFileNameWithoutExtension(Directory.GetFiles(System.IO.Path.GetTempPath() + "\\TGS", "*.xml")(index)))
                        Dim mgInLocalMachine As XElement = XElement.Load(Directory.GetFiles(mgDirectoryPath, "*.xml")(index))
                        If XElement.DeepEquals(market, mgInLocalMachine) Then
                            foundMg = True
                            mgNameToBeSaved = market.Attribute("name").Value
                            Exit For
                        Else
                            Dim mgName As String = market.Attribute("name").Value
                            market.Attribute("name").Value = mgInLocalMachine.Attribute("name").Value

                            If XElement.DeepEquals(market, mgInLocalMachine) Then
                                foundMg = True
                                mgNameToBeSaved = mgInLocalMachine.Attribute("name").Value
                                '  mgInLocalMachine.Attribute("name").Value = mgName
                                ' mgInLocalMachine.Save(mgDirectoryPath + mgInLocalMachine.Attribute("name").Value + ".xml")
                                Exit For
                            Else
                                market.Attribute("name").Value = mgName
                                foundMg = False
                            End If
                        End If
                    Next
                End If
                If foundMg Then
                    '  dt.Rows(0)(1) = tgNameToBeSaved
                    tpSelections.UcMarkets1.lbPlan.Items.Add(mgNameToBeSaved)
                Else
                    market.Save(mgDirectoryPath + market.Attribute("name").Value + ".xml")
                    tpSelections.UcMarkets1.lbPlan.Items.Add(market.Attribute("name").Value)
                End If
                ' tpSelections.UcMarkets1.lbPlan.Items.Add(market.Value)
            Next
        Catch ex As Exception
            LogMpsrintExException("Exception occured while populating TGMGS from opened plan.Message: " + ex.Message)
        End Try
    End Function
    Private Function PopulatePeriods(ByVal msprintPlan As XElement)
        Try
            tpSelections.TaskPaneLogFile1.dtFromDate.Value = msprintPlan.Element("STARTDATE").Value
            tpSelections.TaskPaneLogFile1.dtToDate.Value = msprintPlan.Element("ENDDATE").Value
            tpSelections.TaskPaneLogFile1.dtWeeks.Clear()
            '  tpSelections.TaskPaneLogFile1.dtWeeks.TableName = "Weeks"
            Using reader As XmlReader = msprintPlan.Element("PERIODS").Element("NewDataSet").CreateReader()
                tpSelections.TaskPaneLogFile1.dtWeeks.ReadXml(reader)
            End Using
            tpSelections.TaskPaneLogFile1.dgvWeeks.DataSource = tpSelections.TaskPaneLogFile1.dtWeeks

            tpSelections.TaskPaneLogFile1.LbDayParts.Items.Clear()
            For Each daypart As XElement In msprintPlan.Element("DayParts").Elements
                tpSelections.TaskPaneLogFile1.LbDayParts.Items.Add(daypart.Value)
            Next
        Catch ex As Exception
            LogMpsrintExException("Exception occured while populating periods from openedplan." + ex.Message)
        End Try
    End Function
    Private Function ValidateChosenXML(ByVal plan As XElement) As Boolean
        Dim validated As Boolean = False
        Try
            '<SelectedSpots></SelectedSpots>
            '   <AvailableSpots></AvailableSpots>
            '   <PlanItems></PlanItems>
            '   <TG></TG>
            '   <MGS></MGS>
            '   <STARTDATE></STARTDATE>
            '   <ENDDATE></ENDDATE>
            '   <PERIODS></PERIODS>
            If plan.HasElements() Then

                If plan.Elements("SelectedSpots").Any() And plan.Elements("PlanItems").Any And plan.Elements("TG").Any And plan.Elements("MGS").Any And plan.Elements("STARTDATE").Any And plan.Elements("ENDDATE").Any And plan.Elements("PERIODS").Any And plan.Elements("DayParts").Any And plan.Elements("MappedChannels").Any Then
                    validated = True
                Else
                    validated = False
                End If
            Else
                validated = False

            End If

        Catch ex As Exception
            LogMpsrintExException("Exception occured while Validating Msprint Plan." + ex.Message + plan.ToString())
            Throw ex
        End Try
        Return validated
    End Function


    Private Sub menuBtnSavePlan_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles menuBtnSavePlan.Click
        Dim planXmL As XElement =
            <MsprintExPlan>
                <SelectedSpots></SelectedSpots>
                <PlanItems></PlanItems>
                <MappedChannels></MappedChannels>
                <TG></TG>
                <MGS></MGS>
                <STARTDATE></STARTDATE>
                <ENDDATE></ENDDATE>
                <PERIODS></PERIODS>
                <DayParts></DayParts>
            </MsprintExPlan>
        Try
            xecellineItemsTable = CType(loSpotSelection.DataSource, Data.DataTable)
            If Not (xecellineItemsTable Is Nothing) Then

                ' RnFSelectedSpots Is Nothing And RnFAvaiSpots Is Nothing And

                If Not (RnFSelectedSpots Is Nothing) Then
                    Dim sspots As XElement = planXmL.Element("SelectedSpots")
                    RnFSelectedSpots.TableName = "SelecSpots"
                    Using sspotswriter As XmlWriter = sspots.CreateWriter()
                        RnFSelectedSpots.WriteXml(sspotswriter, XmlWriteMode.WriteSchema, True)
                    End Using
                End If

                'If Not (RnFAvaiSpots Is Nothing) Then
                '    Dim aspots As XElement = planXmL.Element("AvailableSpots")
                '    RnFAvaiSpots.TableName = "AvaiSpots"

                '    Using aspotswriter As XmlWriter = aspots.CreateWriter()
                '        RnFAvaiSpots.WriteXml(aspotswriter, XmlWriteMode.WriteSchema, True)
                '    End Using

                'End If
                mappedchannels = GetGridTable()
                Dim channels As XElement = planXmL.Element("MappedChannels")
                mappedchannels.TableName = "MappedChannels"
                Using channelsswriter As XmlWriter = channels.CreateWriter()
                    mappedchannels.WriteXml(channelsswriter, XmlWriteMode.WriteSchema, True)
                End Using
                'planXmL.Element("SelectedSpots").Add(mediaplan)
                'planXmL.Element("AvailableSpots").Add(rnfoutputXml)
                xecellineItemsTable.TableName = "PlanItemRows"
                ' xecellineItemsTable.WriteXmlSchema(MasterFolderPath + "PlanItems.xsd")
                Dim planItems As XElement = planXmL.Element("PlanItems")
                '           using (XmlWriter w = container.CreateWriter()) {
                Using planitemsWriter As XmlWriter = planItems.CreateWriter()
                    xecellineItemsTable.WriteXml(planitemsWriter, System.Data.XmlWriteMode.WriteSchema, True)
                    '  xecellineItemsTable.ReadXml(
                End Using

                Dim dtable As System.Data.DataTable = DirectCast(tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
                Dim plantg As String = dtable.Rows(0)(1).ToString().Trim()

                planXmL.Element("TG").Value = XElement.Load(tgDirectoryPath + plantg + ".xml").ToString()
                planXmL.Element("STARTDATE").Value = tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToString()
                planXmL.Element("ENDDATE").Value = tpSelections.TaskPaneLogFile1.dtToDate.Value.ToString()
                Dim dtweekscopy As Data.DataTable = tpSelections.TaskPaneLogFile1.dtWeeks.Copy()
                Dim weeks As XElement = planXmL.Element("PERIODS")
                ' dtweekscopy.TableName = "dtweeks"

                Using weekswriter As XmlWriter = weeks.CreateWriter()
                    dtweekscopy.WriteXml(weekswriter, XmlWriteMode.WriteSchema, True)
                End Using

                For index = 0 To tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                    ' Dim market As XElement = <Mgroup></Mgroup>
                    ' plantg
                    ' market.Value = tpSelections.UcMarkets1.lbPlan.Items(index).ToString()
                    Dim market As XElement = XElement.Load(mgDirectoryPath + tpSelections.UcMarkets1.lbPlan.Items(index).ToString() + ".xml")
                    planXmL.Element("MGS").Add(market)
                Next
                For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.LbDayParts.Items.Count - 1
                    '<day_part>0200-0200</day_part>
                    '  <day_part>0200-0200</day_part>
                    Dim dpart As XElement = New XElement("DayPart")
                    dpart.Value = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.LbDayParts.Items(index)
                    planXmL.Element("DayParts").Add(dpart)
                Next
                savePlanPath = New SaveFileDialog()
                savePlanPath.Filter = "MsprintEx Plan Files | *.xml"
                savePlanPath.DefaultExt = "xml"

                savePlanPath.InitialDirectory = LogDirectoryPath
                '  savePlanPath.RestoreDirectory = True
                Dim name As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                savePlanPath.FileName = "MsprintExPlan_" + name

                '' mediaplan.Save(LogDirectoryPath + "RNF_Inp_" + name)
                'Input.Save(LogDirectoryPath + "SpotsLogFile_Inp_" + name)
                If savePlanPath.ShowDialog() = DialogResult.OK Then
                    planXmL.Save(savePlanPath.FileName)
                    MessageBox.Show("Plan saved successfully")
                End If

            Else
                MessageBox.Show("Plan must have atleast one item to be saved.")
            End If
            'planXmL.Element("PlanItems").Add(xecellineItemsTable.WriteXml(
        Catch ex As Exception
            LogMpsrintExException("Exception occured while saving plan.Message" + ex.Message)
            MessageBox.Show("Exception occured while saving plan.Please refer Error log for more details")
        End Try
    End Sub

    Private Sub btnAddPlan_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnAddPlan.Click
        Dim msprintExPlan As XElement = New XElement("MsprintExPlan")
        Dim planopener As OpenFileDialog = New OpenFileDialog()
        Try
            If Globals.ThisAddIn.Application.Workbooks.Count < 1 Then
                '  Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(,,
                Globals.ThisAddIn.Application.Workbooks.Add(System.Type.Missing)
            End If

            ' btnAddPlan.Enabled = True
            '"Text files (*.txt)|*.txt|All files (*.*)|*.*"
            planopener.Filter = "MSprintEx Plan Files (*.xml)|*.xml"
            planopener.InitialDirectory = LogDirectoryPath
            btnCleanupplan.Enabled = True
            btnRnF.Enabled = True
            btnMapChannels.Enabled = True
            If planopener.ShowDialog() = DialogResult.OK Then
                Dim plan As XElement
                Try
                    plan = XElement.Load(planopener.FileName)
                Catch ex As Exception
                    MessageBox.Show("Only valid MsprintEx Plan files can be opened.Please retry again")
                End Try


                'If plan.Name.Equals("{MsprintExPlan}") Then

                msprintExPlan = XElement.Load(planopener.FileName)

                If ValidateChosenXML(msprintExPlan) Then
                    PopulateTGMGS(msprintExPlan)
                    PopulatePeriods(msprintExPlan)
                    'Next

                    If msprintExPlan.Element("SelectedSpots").Elements("NewDataSet").Any Then
                        Using reader As XmlReader = msprintExPlan.Element("SelectedSpots").Element("NewDataSet").CreateReader()

                            If RnFSelectedSpots Is Nothing Then
                                RnFSelectedSpots = New Data.DataTable

                            End If
                            Dim spottab As Data.DataTable = New Data.DataTable
                            spottab = RnFSelectedSpots.Clone()
                            spottab.ReadXml(reader)
                            RnFSelectedSpots.Merge(spottab)
                        End Using

                    End If


                    If msprintExPlan.Element("MappedChannels").Elements("NewDataSet").Any Then
                        Using reader As XmlReader = msprintExPlan.Element("MappedChannels").Element("NewDataSet").CreateReader()

                            If mappedchannels Is Nothing Then
                                mappedchannels = New Data.DataTable

                            End If
                            Dim chann As Data.DataTable = New Data.DataTable
                            chann = mappedchannels.Clone()
                            chann.ReadXml(reader)
                            mappedchannels.Merge(chann)
                        End Using
                    End If

                    'If msprintExPlan.Element("AvailableSpots").Elements("NewDataSet").Any Then
                    '    Using reader As XmlReader = msprintExPlan.Element("AvailableSpots").Element("NewDataSet").CreateReader()
                    '        If RnFAvaiSpots Is Nothing Then
                    '            RnFAvaiSpots = New Data.DataTable
                    '        End If
                    '        RnFAvaiSpots.ReadXml(reader)
                    '    End Using
                    'End If
                    Using reader As XmlReader = msprintExPlan.Element("PlanItems").Element("NewDataSet").CreateReader()
                        If xecellineItemsTable Is Nothing Then
                            xecellineItemsTable = New Data.DataTable
                        End If
                        Dim xecel As Data.DataTable = New Data.DataTable
                        xecel = xecellineItemsTable.Clone()
                        xecel.ReadXml(reader)
                        xecellineItemsTable.Merge(xecel)
                    End Using
                    isPlanClean = True
                    planOpenedSuccessfully = True
                    btnGetReqSpots.Enabled = True
                    btnRnF.Enabled = True
                    btnReorderPlanChannels.Enabled = True
                    'Try
                    If Not (RnFSelectedSpots Is Nothing) Then

                        If RnFSelectedSpots.Rows.Count > 0 Then
                            btnMarketSummary.Enabled = True
                            btnChannelSummary.Enabled = True
                            btnDurationSummary.Enabled = True
                            btnCreativeSummary.Enabled = True
                            btnAllSummary.Enabled = True
                            btnLogFile.Enabled = True
                        End If

                    End If

                    '    If Not CheckSheetExists("Plan Selection") Then
                    '        logCreator = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                    '        logCreator.Name = "Plan Selection"
                    '    Else
                    '        logCreator = CheckAndReturnSheet("Plan Selection")
                    '        '  newSheet.UsedRange.Clear()
                    '        Globals.Ribbons.MSprintExRibbon.CleanSheet(logCreator)
                    '        'logCreator.Delete()
                    '        'logCreator = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                    '        'logCreator.Name = "Plan Selection"
                    '        logCreator.Activate()
                    '    End If
                    '    Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(logCreator)
                    '    ' vstoWorkbook.Name = "Plan Selection"

                    '    If vstoWorkbook.Controls.Contains("InputSpotSelection") Then
                    '        loSpotSelection = DirectCast(vstoWorkbook.Controls.Item("InputSpotSelection"), Microsoft.Office.Tools.Excel.ListObject)
                    '    Else
                    '        loSpotSelection = vstoWorkbook.Controls.AddListObject(vstoWorkbook.Range("$A$1"), "InputSpotSelection")
                    '    End If

                    '    loSpotSelection.AutoSetDataBoundColumnHeaders = True
                    '    loSpotSelection.SetDataBinding(xecellineItemsTable)
                    '    loSpotSelection.ListColumns(5).Range.NumberFormat = "hh:mm;@"
                    '    loSpotSelection.ListColumns(6).Range.NumberFormat = "hh:mm;@"
                    '    vstoWorkbook.get_Range("A:A", Type.Missing).EntireColumn.Hidden = True
                    '    PlanSheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    'Catch ex As Exception

                    'End Try
                    ''Display Pane with spots
                    'Try


                    '    If Not SpotSelectionPane Is Nothing Then
                    '        SpotSelectionPane.Visible = False
                    '        SpotSelectionPane.Dispose()
                    '    End If

                    '    mpTpSpotSelection = New ucSpotSelection()
                    '    Dim filter As String = String.Format("GUID = '{0}'", xecellineItemsTable.AsEnumerable().First()("GUID").ToString())
                    '    Dim rows As Data.DataRow() = RnFSelectedSpots.Select(filter)
                    '    'Dim spotst As Data.DataTable = New Data.DataTable()
                    '    'spotst = RnFSelectedSpots.Clone()
                    '    'For Each row1 As Data.DataRow In rows
                    '    '    spotst.ImportRow(row1)
                    '    'Next
                    '    mpTpSpotSelection.dgSelectedSpotsGrid.DataSource = rows.CopyToDataTable()
                    '    HideSelectedSpotsGrid(mpTpSpotSelection.dgSelectedSpotsGrid)
                    '    'Dim avaiRows As Data.DataRow() = RnFAvaiSpots.Select(filter)
                    '    'mpTpSpotSelection.dgvAvailableSpotsGrid.DataSource = avaiRows.CopyToDataTable()
                    '    'HideSelectedSpotsGrid(mpTpSpotSelection.dgvAvailableSpotsGrid)
                    '    '  mpTpSpotSelection.Anchor = AnchorStyles.Bottom

                    '    'mpTpSpotSelection.Anchor = (AnchorStyles.Bottom Or AnchorStyles.Right Or AnchorStyles.Top Or AnchorStyles.Left)
                    '    ' mpTpSpotSelection.dgSelectedSpotsGrid.Refresh()
                    '    SpotSelectionPane = Globals.ThisAddIn.CustomTaskPanes.Add(mpTpSpotSelection, "Spot Selection")
                    '    SpotSelectionPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
                    '    SpotSelectionPane.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange
                    '    mpTpSpotSelection.Dock = DockStyle.Fill
                    '    ' SpotSelectionPane.
                    '    SpotSelectionPane.Width = 800
                    '    SpotSelectionPane.Height = 400

                    '    '               Dim paneSize As System.Drawing.Size = _
                    '    'New System.Drawing.Size(SpotSelectionPane.Width, SpotSelectionPane.Height)
                    '    '               mpTpSpotSelection.AutoSize = True
                    '    '               mpTpSpotSelection.AutoSizeMode = AutoSizeMode.GrowAndShrink
                    '    ' mpTpSpotSelection.si.FlowPanel.Size = paneSize
                    '    'SpotSelectionPane.Width = 300
                    '    'SpotSelectionPane.Height = 200
                    '    DisplayCurrentPlanItem()
                    '    SpotSelectionPane.Visible = True
                    'Catch ex As Exception

                    'End Try
                Else
                    MessageBox.Show("Attempted to open corrupt MSprintEx plan. Please retry with valid MSprintEx plan")
                End If
                'Else
                '    MessageBox.Show("Attempted to open incorrect plan .Please retry opening MSprintEx plan")
                'End If
            End If

        Catch ex As Exception
            LogMpsrintExException("Exception occured while adding plan to existing MsprintPlan." + ex.Message)
            MessageBox.Show("Exception occured while adding plan.Please refer to Error Log for more details")
        End Try
    End Sub

    Private Sub btnMarketSummary_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnMarketSummary.Click

        Dim frm As frmWait
        Try

            'Dim str As String = String.Empty
            'For Each row As Microsoft.Office.Interop.Excel.ListRow In loSpotSelection.ListRows
            '    '  loSpotSelection.ListRows(0).Range.Cells.
            '    'For Each cell As Microsoft.Office.Interop.Excel.Range In row.Range.Cells
            '    '    ' str = row.Range.Cells(row.Index, col.Index).ToString()
            '    '    str = cell.Text
            '    'Next

            'Next

            'If Not (RnFAvaiSpots Is Nothing) Then
            '    RnFAvaiSpots.Rows.Clear()
            'End If

            ' Dim dtt As BindingSource = CType(channelMapping.dgvChannels.DataSource, BindingSource)

            Dim plantgname As String = String.Empty
            Dim reftgname As String = String.Empty
            Dim dtable As System.Data.DataTable = DirectCast(tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            plantgname = dtable.Rows(0)(1).ToString().Trim()
            '     reftgname = dtable.Rows(1)(1).ToString().Trim()
            If plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please make Target Group and Market Group(s) Selections")
            ElseIf plantgname.Length > 1 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for chosen TG")
            ElseIf plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count > 0 Then
                System.Windows.Forms.MessageBox.Show("Please choose Planning Target group for Market groups")
            ElseIf loSpotSelection Is Nothing Then
                System.Windows.Forms.MessageBox.Show("Please enter plan to view Market Summary")
            ElseIf Not (isPlanClean) Then
                System.Windows.Forms.MessageBox.Show("Please clean plan and ensure no duplication of rows")
            ElseIf Not (AllChannelsMapped()) Then
                System.Windows.Forms.MessageBox.Show("Please map all channels with master channellist")
                'ElseIf reftgname.Length > 0 And tpSelections.UcMarkets1.lbRef.Items.Count < 1 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for reference Target Group chosen")
            Else
                SetWaitCursor("Getting Market Summary details..")

                '  Globals.ThisAddIn.Application.DoEvents()
                System.Windows.Forms.Application.DoEvents()
                'frm = New frmWait()
                'frm.Show()
                ' loSpotSelection.RefreshDataRows()
                xecellineItemsTable = CType(loSpotSelection.DataSource, Data.DataTable)
                mediaplan = ConstructMarketSummaryInputXML()
                ' mediaplan = ConstructInputRnFXML()



                'testmediaplan =
                '<mediaplan>
                '    <!-- TO BE FROZEN 26 Dec 13 -->
                '    <!-- common section start -->
                '    <PreEvalPeriod>
                '        <StartDate>20130804</StartDate>
                '        <EndDate>20130817</EndDate>
                '    </PreEvalPeriod>

                '    <DayParts>
                '        <DayPart>0800-1200</DayPart>
                '        <DayPart>2100-2200</DayPart>
                '    </DayParts>


                '    <!-- common section ends -->


                '    <tg name="CS 15-44" cs="1" sec="1,2,3,4" sex="1,2" age="3,4,5">

                '        <mg name="mg1" type="group">
                '            <market>1</market>
                '            <market>3</market>
                '        </mg>

                '        <mg name="mg2" type="single">
                '            <market>2</market>
                '        </mg>
                '    </tg>


                '    <!-- TVR000s,TVR,GRP000s,GRP,AvgFreq,CummCost,SpotCPRP,CummCPRP,Reach000s,R1,R2,R3,R4,R5,R6,R7,R8,R9,R10 -->

                '    <!-- all output -->
                '    <plan type="weekwise"><!-- clubbed-->
                '        <period StartDate="20130804" EndDate="20130810" year="2013" WeekNum="32">
                '            <programme guid="1" SeqNumber="1" ChannelCode="004" ChannelName="Star Plus" ProgName="Yeh Rishta Kya Kehlata Hai" days="Thu" StartTime="21:30" EndTime="22:00" CostPer10s="150" caption="Colgate Kids Jumping" AdDuration="30" NumberOfSpots="10">

                '                <selected_spots><!-- this is a demo output - spot count and dates may be incorrect. attributes are correctNO SPACES BETWEEN VALUES IN CSV -->
                '                    <spot log="004,10082013,032238,032308,0,001,005,00030"/>
                '                </selected_spots>
                '            </programme>
                '            <programme guid="2" SeqNumber="2" ChannelCode="004" ChannelName="Star Plus" ProgName="DIYA AUR BAATI HUM" days="Mon,Tue,Wed,Thu,Fri" StartTime="21:00" EndTime="21:30" CostPer10s="120" caption="Colgate Kids Jumping" AdDuration="20" NumberOfSpots="9">

                '                <selected_spots><!-- this is a demo output - spot count and dates may be incorrect. attributes are correctNO SPACES BETWEEN VALUES IN CSV -->
                '                    <spot log="004,10082013,032238,032308,0,001,005,00030"/>
                '                </selected_spots>
                '            </programme>
                '        </period>
                '    </plan>

                '</mediaplan>


                '  request = WebRequest.Create("")
                '

                'request.Method = "POST"
                'request.ContentType = "application/x-www-form-urlencoded"
                'request.Timeout = 300000
                'request.ServicePoint.MaxIdleTime = 300000
                ''  request.KeepAlive = True
                'inputstring = HttpUtility.UrlEncode(MyExtensionClasses.OuterXml(testmediaplan))

                'stream = request.GetRequestStream()
                ''  Dim bytes As Byte() = Encoding.Default.GetBytes(input.OuterXml)

                'Dim encoding As System.Text.UTF8Encoding = New UTF8Encoding()
                'postData = "inputXML=" + inputstring

                'data = encoding.GetBytes(postData)
                ''input.Save(stream)
                '' request.ContentLength = data.Length
                'stream.Write(data, 0, data.Length)

                '' request.Proxy = Nothing
                'ws = request.GetResponse()

                'oStream = ws.GetResponseStream()
                'Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")

                '' Pipe the stream to a higher level stream reader with the required encoding format.
                'Dim readStream As New StreamReader(oStream, encode)
                '     Dim separators() As String = {"Genre,Viewership"}
                ' Dim file As [String]() = readStream.ReadToEnd().Split(separators, StringSplitOptions.RemoveEmptyEntries)

                Dim name As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                mediaplan.Save(LogDirectoryPath + "MarketSummary_Inp_" + name)
                ' rnfoutputXml = GetOpXMLFromWS(mediaplan, "http://54.255.217.55:8080/GroupM/marketsummary/")
                rnfoutputXml = GetOpXMLFromWS(mediaplan, Globals.Ribbons.MSprintExRibbon.GetURLForWS("MarketSummaryWSURL_New"))
                'rnfoutputXml = XElement.Load("C:\ASR\op.xml")
                If Not (rnfoutputXml Is Nothing) Then
                    Dim name1 As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                    rnfoutputXml.Save(LogDirectoryPath + "MarketSummary_Op_" + name1)

                    If ConstructMarketSummaryTable(rnfoutputXml) Then
                        RnFMarketSummary = nbdmainMarketSummary(RnFMarketSummary)
                        '   RnFMarketSummary = ReverseCalculateTVRReach(RnFMarketSummary)
                        If Not CheckSheetExists("Market Summary") Then
                            nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                            nativeSheet.Name = "Market Summary"
                        Else
                            'nativeSheet = CheckAndReturnSheet("Market Summary")
                            ''  newSheet.UsedRange.Clear()
                            'Globals.Ribbons.MSprintExRibbon.CleanSheet(nativeSheet)
                            'nativeSheet.Activate()

                            Dim sheetcount As Integer = CheckAndReturnSheet("Market Summary")
                            If sheetcount > 0 Then
                                nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                                Dim sname As String = String.Format("Market Summary({0})", sheetcount)
                                nativeSheet.Name = sname
                            End If


                        End If
                        Dim tgcell As Microsoft.Office.Interop.Excel.Range = nativeSheet.Range("$A$2", Type.Missing)
                        tgcell.Value2 = "TG : " + plantgname
                        Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(nativeSheet)

                        If tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                            '  Dim periodcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.Range("$A$3", Type.Missing)
                            Dim periodcell As Microsoft.Office.Interop.Excel.Range = tgcell.Next(2, 0)
                            Dim period As String = String.Format("Period : {0} To {1}", tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToString("dd/MM/yyyy"), tpSelections.TaskPaneLogFile1.dtToDate.Value.ToString("dd/MM/yyyy"))
                            periodcell.Value = period
                            Dim lrange As Microsoft.Office.Interop.Excel.Range = periodcell.Next(2, 0)
                            Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "RnfMSummary" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString())
                            Dim table As Data.DataTable = RnFMarketSummary.Copy()
                            table.Columns.RemoveAt(0)
                            If MSHRN > 20 Then

                                For index = 21 To MSHRN
                                    table.Columns.Remove(index.ToString() + "+")
                                Next

                            End If
                            listobject.DataSource = table
                            'rowcount += table.Rows.Count + 2
                            listobject.AutoSetDataBoundColumnHeaders = True
                            listobject.ShowAutoFilter = False
                        ElseIf tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                            Dim count As Integer = 4
                            For index = 0 To tpSelections.TaskPaneLogFile1.dtWeeks.Rows.Count - 1
                                Dim periodcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count, 1), vstoWorkbook.Cells(count, 1))
                                Dim period As String = String.Format("Period : {0} To {1}", Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("StartDate").ToString()).ToString("dd/MM/yyyy"), Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("EndDate").ToString()).ToString("dd/MM/yyyy"))
                                periodcell.Value = period
                                Dim lrange As Microsoft.Office.Interop.Excel.Range = periodcell.Next(2, 0)
                                Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "RnfMSummary" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString())
                                Dim filter As String = String.Format("WeekNum='{0}'", tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("WeekNumber").ToString())
                                Dim rows As Data.DataRow() = RnFMarketSummary.Select(filter)
                                If rows.Length > 0 Then
                                    Dim table As Data.DataTable = rows.CopyToDataTable()
                                    table.Columns.RemoveAt(0)
                                    If MSHRN > 20 Then

                                        For index1 = 21 To MSHRN
                                            table.Columns.Remove(index1.ToString() + "+")
                                        Next

                                    End If
                                    listobject.DataSource = table
                                    count = count + table.Rows.Count + 3
                                    'rowcount += table.Rows.Count + 2
                                    listobject.AutoSetDataBoundColumnHeaders = True
                                    listobject.ShowAutoFilter = False
                                End If


                            Next

                        End If

                    Else

                    End If
                Else
                    MessageBox.Show("Unable to retreive requested Market Summary details from server.")
                    ' SetNormalCursor()
                End If
                SetNormalCursor()
            End If
        Catch ex As Exception
            SetNormalCursor()
            LogMpsrintExException("Exception occured while retreiving Market Summary details." + ex.Message)
            MessageBox.Show("Exception occured while retreiving Market summary details.Please refer error log for more details.")
        End Try
    End Sub

    Private Sub btnCreativeSummary_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnCreativeSummary.Click
        Dim frm As frmWait
        Try

            'Dim str As String = String.Empty
            'For Each row As Microsoft.Office.Interop.Excel.ListRow In loSpotSelection.ListRows
            '    '  loSpotSelection.ListRows(0).Range.Cells.
            '    'For Each cell As Microsoft.Office.Interop.Excel.Range In row.Range.Cells
            '    '    ' str = row.Range.Cells(row.Index, col.Index).ToString()
            '    '    str = cell.Text
            '    'Next

            'Next

            'If Not (RnFAvaiSpots Is Nothing) Then
            '    RnFAvaiSpots.Rows.Clear()
            'End If

            ' Dim dtt As BindingSource = CType(channelMapping.dgvChannels.DataSource, BindingSource)

            Dim plantgname As String = String.Empty
            Dim reftgname As String = String.Empty
            Dim dtable As System.Data.DataTable = DirectCast(tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            plantgname = dtable.Rows(0)(1).ToString().Trim()
            '  reftgname = dtable.Rows(1)(1).ToString().Trim()
            If plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please make Target Group and Market Group(s) Selections")
            ElseIf plantgname.Length > 1 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for chosen TG")
            ElseIf plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count > 0 Then
                System.Windows.Forms.MessageBox.Show("Please choose Planning Target group for Market groups")
            ElseIf loSpotSelection Is Nothing Then
                System.Windows.Forms.MessageBox.Show("Please enter plan to view Creative Summary")
            ElseIf Not (isPlanClean) Then
                System.Windows.Forms.MessageBox.Show("Please clean plan and ensure no duplication of rows")
            ElseIf Not (AllChannelsMapped()) Then
                System.Windows.Forms.MessageBox.Show("Please map all channels with master channellist")
                'ElseIf reftgname.Length > 0 And tpSelections.UcMarkets1.lbRef.Items.Count < 1 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for reference Target Group chosen")
            Else
                SetWaitCursor("Getting Creative Summary details..")

                '  Globals.ThisAddIn.Application.DoEvents()
                System.Windows.Forms.Application.DoEvents()
                'frm = New frmWait()
                'frm.Show()
                ' loSpotSelection.RefreshDataRows()
                xecellineItemsTable = CType(loSpotSelection.DataSource, Data.DataTable)
                Dim creatives As Data.DataTable = xecellineItemsTable.Copy()
                creatives = creatives.DefaultView.ToTable(True, "Creative")
                RnFCreativeSummary = New Data.DataTable()
                RnFCreativeSummary.Columns.Add("WeekNum")
                RnFCreativeSummary.Columns.Add("Creative")
                RnFCreativeSummary.Columns.Add("Market")
                RnFCreativeSummary.Columns.Add("Universe")
                RnFCreativeSummary.Columns.Add("GRP")
                RnFCreativeSummary.Columns.Add("AOTS")
                RnFCreativeSummary.Columns.Add("1+")
                RnFCreativeSummary.Columns.Add("2+")
                RnFCreativeSummary.Columns.Add("3+")
                RnFCreativeSummary.Columns.Add("4+")
                RnFCreativeSummary.Columns.Add("5+")
                RnFCreativeSummary.Columns.Add("6+")
                RnFCreativeSummary.Columns.Add("7+")
                RnFCreativeSummary.Columns.Add("8+")
                RnFCreativeSummary.Columns.Add("9+")
                RnFCreativeSummary.Columns.Add("10+")
                '   For Each creative As Data.DataRow In creatives.Rows
                'xecellineItemsTable = CType(loSpotSelection.DataSource, Data.DataTable)
                'Dim filter As String = String.Format("Creative ='{0}'", creative(0).ToString())
                'Dim rows As Data.DataRow() = xecellineItemsTable.Select(filter)
                'xecellineItemsTable = rows.CopyToDataTable()
                mediaplan = ConstructCreativeSummaryInputXML()
                ' mediaplan = ConstructInputRnFXML()



                'testmediaplan =
                '<mediaplan>
                '    <!-- TO BE FROZEN 26 Dec 13 -->
                '    <!-- common section start -->
                '    <PreEvalPeriod>
                '        <StartDate>20130804</StartDate>
                '        <EndDate>20130817</EndDate>
                '    </PreEvalPeriod>

                '    <DayParts>
                '        <DayPart>0800-1200</DayPart>
                '        <DayPart>2100-2200</DayPart>
                '    </DayParts>


                '    <!-- common section ends -->


                '    <tg name="CS 15-44" cs="1" sec="1,2,3,4" sex="1,2" age="3,4,5">

                '        <mg name="mg1" type="group">
                '            <market>1</market>
                '            <market>3</market>
                '        </mg>

                '        <mg name="mg2" type="single">
                '            <market>2</market>
                '        </mg>
                '    </tg>


                '    <!-- TVR000s,TVR,GRP000s,GRP,AvgFreq,CummCost,SpotCPRP,CummCPRP,Reach000s,R1,R2,R3,R4,R5,R6,R7,R8,R9,R10 -->

                '    <!-- all output -->
                '    <plan type="weekwise"><!-- clubbed-->
                '        <period StartDate="20130804" EndDate="20130810" year="2013" WeekNum="32">
                '            <programme guid="1" SeqNumber="1" ChannelCode="004" ChannelName="Star Plus" ProgName="Yeh Rishta Kya Kehlata Hai" days="Thu" StartTime="21:30" EndTime="22:00" CostPer10s="150" caption="Colgate Kids Jumping" AdDuration="30" NumberOfSpots="10">

                '                <selected_spots><!-- this is a demo output - spot count and dates may be incorrect. attributes are correctNO SPACES BETWEEN VALUES IN CSV -->
                '                    <spot log="004,10082013,032238,032308,0,001,005,00030"/>
                '                </selected_spots>
                '            </programme>
                '            <programme guid="2" SeqNumber="2" ChannelCode="004" ChannelName="Star Plus" ProgName="DIYA AUR BAATI HUM" days="Mon,Tue,Wed,Thu,Fri" StartTime="21:00" EndTime="21:30" CostPer10s="120" caption="Colgate Kids Jumping" AdDuration="20" NumberOfSpots="9">

                '                <selected_spots><!-- this is a demo output - spot count and dates may be incorrect. attributes are correctNO SPACES BETWEEN VALUES IN CSV -->
                '                    <spot log="004,10082013,032238,032308,0,001,005,00030"/>
                '                </selected_spots>
                '            </programme>
                '        </period>
                '    </plan>

                '</mediaplan>


                '  request = WebRequest.Create("")
                '

                'request.Method = "POST"
                'request.ContentType = "application/x-www-form-urlencoded"
                'request.Timeout = 300000
                'request.ServicePoint.MaxIdleTime = 300000
                ''  request.KeepAlive = True
                'inputstring = HttpUtility.UrlEncode(MyExtensionClasses.OuterXml(testmediaplan))

                'stream = request.GetRequestStream()
                ''  Dim bytes As Byte() = Encoding.Default.GetBytes(input.OuterXml)

                'Dim encoding As System.Text.UTF8Encoding = New UTF8Encoding()
                'postData = "inputXML=" + inputstring

                'data = encoding.GetBytes(postData)
                ''input.Save(stream)
                '' request.ContentLength = data.Length
                'stream.Write(data, 0, data.Length)

                '' request.Proxy = Nothing
                'ws = request.GetResponse()

                'oStream = ws.GetResponseStream()
                'Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")

                '' Pipe the stream to a higher level stream reader with the required encoding format.
                'Dim readStream As New StreamReader(oStream, encode)
                '     Dim separators() As String = {"Genre,Viewership"}
                ' Dim file As [String]() = readStream.ReadToEnd().Split(separators, StringSplitOptions.RemoveEmptyEntries)

                Dim name As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                mediaplan.Save(LogDirectoryPath + "CreativeSummary_Inp_" + name)
                'rnfoutputXml = GetOpXMLFromWS(mediaplan, "http://54.255.217.55:8080/GroupM/creativesummary/")
                rnfoutputXml = GetOpXMLFromWS(mediaplan, Globals.Ribbons.MSprintExRibbon.GetURLForWS("CreativeSummaryWSURL_New"))
                'rnfoutputXml = XElement.Load("C:\ASR\op.xml")
                If Not (rnfoutputXml Is Nothing) Then
                    Dim name1 As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                    rnfoutputXml.Save(LogDirectoryPath + "CreativeSummary_Op_" + name1)
                    ConstructCreativeSummaryTable(rnfoutputXml)
                    RnFCreativeSummary = nbdmainSummary(RnFCreativeSummary, CreativeSummaryHRN)
                    ' RnFCreativeSummary = ReverseCalculateTVRReach(RnFCreativeSummary)
                End If
                '   Next


                '  If ConstructMarketSummaryTable(rnfoutputXml) Then
                If Not CheckSheetExists("Creative Summary") Then
                    nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                    nativeSheet.Name = "Creative Summary"
                Else
                    'nativeSheet = CheckAndReturnSheet("Creative Summary")
                    ''  newSheet.UsedRange.Clear()
                    'Globals.Ribbons.MSprintExRibbon.CleanSheet(nativeSheet)
                    'nativeSheet.Activate()

                    Dim sheetcount As Integer = CheckAndReturnSheet("Creative Summary")
                    If sheetcount > 0 Then
                        nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                        Dim sname As String = String.Format("Creative Summary({0})", sheetcount)
                        nativeSheet.Name = sname
                    End If
                End If
                Dim tgcell As Microsoft.Office.Interop.Excel.Range = nativeSheet.Range("$A$2", Type.Missing)
                tgcell.Value2 = "TG: " + plantgname
                Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(nativeSheet)

                If tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                    '  Dim periodcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.Range("$A$3", Type.Missing)
                    Dim periodcell As Microsoft.Office.Interop.Excel.Range = tgcell.Next(2, 0)
                    Dim period As String = String.Format("Period : {0} To {1}", tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToString("dd/MM/yyyy"), tpSelections.TaskPaneLogFile1.dtToDate.Value.ToString("dd/MM/yyyy"))
                    periodcell.Value = period
                    Dim lrange As Microsoft.Office.Interop.Excel.Range = periodcell.Next(2, 0)
                    Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "RnfCreativeSummary" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString())
                    Dim table As Data.DataTable = RnFCreativeSummary.Copy()
                    table.Columns.RemoveAt(0)
                    If CreativeSummaryHRN > 20 Then

                        For index = 21 To CreativeSummaryHRN
                            table.Columns.Remove(index.ToString() + "+")
                        Next

                    End If
                    listobject.DataSource = table
                    'rowcount += table.Rows.Count + 2
                    listobject.AutoSetDataBoundColumnHeaders = True
                    listobject.ShowAutoFilter = False
                ElseIf tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                    Dim count As Integer = 4
                    For index = 0 To tpSelections.TaskPaneLogFile1.dtWeeks.Rows.Count - 1
                        Dim periodcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count, 1), vstoWorkbook.Cells(count, 1))
                        '  Dim period As String = String.Format("Period : '{0}' To '{1}'", tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToString("dd/MM/yyyy"), tpSelections.TaskPaneLogFile1.dtToDate.Value.ToString("dd/MM/yyyy"))
                        Dim period As String = String.Format("Period : {0} To {1}", Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("StartDate").ToString()).ToString("dd/MM/yyyy"), Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("EndDate").ToString()).ToString("dd/MM/yyyy"))
                        periodcell.Value = period
                        Dim lrange As Microsoft.Office.Interop.Excel.Range = periodcell.Next(2, 0)
                        Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "RnfCreativeSummary" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString())
                        Dim filter As String = String.Format("WeekNum='{0}'", tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("WeekNumber").ToString())
                        Dim rows As Data.DataRow() = RnFCreativeSummary.Select(filter)
                        If rows.Length > 0 Then
                            Dim table As Data.DataTable = rows.CopyToDataTable()
                            table.Columns.RemoveAt(0)
                            If CreativeSummaryHRN > 20 Then

                                For index1 = 21 To CreativeSummaryHRN
                                    table.Columns.Remove(index1.ToString() + "+")
                                Next

                            End If
                            listobject.DataSource = table
                            count = count + table.Rows.Count + 3
                            'rowcount += table.Rows.Count + 2
                            listobject.AutoSetDataBoundColumnHeaders = True
                            listobject.ShowAutoFilter = False
                        End If


                    Next

                End If
                ' End If
                SetNormalCursor()
            End If
            '  End If
        Catch ex As Exception
            SetNormalCursor()
            LogMpsrintExException("Exception occured while retreiving Creative summary details" + ex.Message)
            MessageBox.Show("Exception occured while retreiving Creative summary details.Please refer to Error log for more details")
        End Try
    End Sub

    Private Sub btnChannelSummary_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnChannelSummary.Click
        Dim frm As frmWait
        Try

            'Dim str As String = String.Empty
            'For Each row As Microsoft.Office.Interop.Excel.ListRow In loSpotSelection.ListRows
            '    '  loSpotSelection.ListRows(0).Range.Cells.
            '    'For Each cell As Microsoft.Office.Interop.Excel.Range In row.Range.Cells
            '    '    ' str = row.Range.Cells(row.Index, col.Index).ToString()
            '    '    str = cell.Text
            '    'Next

            'Next

            'If Not (RnFAvaiSpots Is Nothing) Then
            '    RnFAvaiSpots.Rows.Clear()
            'End If

            ' Dim dtt As BindingSource = CType(channelMapping.dgvChannels.DataSource, BindingSource)

            Dim plantgname As String = String.Empty
            Dim reftgname As String = String.Empty
            Dim dtable As System.Data.DataTable = DirectCast(tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            plantgname = dtable.Rows(0)(1).ToString().Trim()
            '   reftgname = dtable.Rows(1)(1).ToString().Trim()
            If plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please make Target Group and Market Group(s) Selections")
            ElseIf plantgname.Length > 1 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for chosen TG")
            ElseIf plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count > 0 Then
                System.Windows.Forms.MessageBox.Show("Please choose Planning Target group for Market groups")
            ElseIf loSpotSelection Is Nothing Then
                System.Windows.Forms.MessageBox.Show("Please enter plan to view Channel Summary")
            ElseIf Not (isPlanClean) Then
                System.Windows.Forms.MessageBox.Show("Please clean plan and ensure no duplication of rows")
            ElseIf Not (AllChannelsMapped()) Then
                System.Windows.Forms.MessageBox.Show("Please map all channels with master channellist")
                'ElseIf reftgname.Length > 0 And tpSelections.UcMarkets1.lbRef.Items.Count < 1 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for reference Target Group chosen")
            Else
                SetWaitCursor("Getting Channel Summary details..")

                '  Globals.ThisAddIn.Application.DoEvents()
                System.Windows.Forms.Application.DoEvents()
                'frm = New frmWait()
                'frm.Show()
                ' loSpotSelection.RefreshDataRows()
                xecellineItemsTable = CType(loSpotSelection.DataSource, Data.DataTable)
                mediaplan = ConstructChannelSummaryInputXML()
                '  mediaplan = ConstructInputRnFXML()



                'testmediaplan =
                '<mediaplan>
                '    <!-- TO BE FROZEN 26 Dec 13 -->
                '    <!-- common section start -->
                '    <PreEvalPeriod>
                '        <StartDate>20130804</StartDate>
                '        <EndDate>20130817</EndDate>
                '    </PreEvalPeriod>

                '    <DayParts>
                '        <DayPart>0800-1200</DayPart>
                '        <DayPart>2100-2200</DayPart>
                '    </DayParts>


                '    <!-- common section ends -->


                '    <tg name="CS 15-44" cs="1" sec="1,2,3,4" sex="1,2" age="3,4,5">

                '        <mg name="mg1" type="group">
                '            <market>1</market>
                '            <market>3</market>
                '        </mg>

                '        <mg name="mg2" type="single">
                '            <market>2</market>
                '        </mg>
                '    </tg>


                '    <!-- TVR000s,TVR,GRP000s,GRP,AvgFreq,CummCost,SpotCPRP,CummCPRP,Reach000s,R1,R2,R3,R4,R5,R6,R7,R8,R9,R10 -->

                '    <!-- all output -->
                '    <plan type="weekwise"><!-- clubbed-->
                '        <period StartDate="20130804" EndDate="20130810" year="2013" WeekNum="32">
                '            <programme guid="1" SeqNumber="1" ChannelCode="004" ChannelName="Star Plus" ProgName="Yeh Rishta Kya Kehlata Hai" days="Thu" StartTime="21:30" EndTime="22:00" CostPer10s="150" caption="Colgate Kids Jumping" AdDuration="30" NumberOfSpots="10">

                '                <selected_spots><!-- this is a demo output - spot count and dates may be incorrect. attributes are correctNO SPACES BETWEEN VALUES IN CSV -->
                '                    <spot log="004,10082013,032238,032308,0,001,005,00030"/>
                '                </selected_spots>
                '            </programme>
                '            <programme guid="2" SeqNumber="2" ChannelCode="004" ChannelName="Star Plus" ProgName="DIYA AUR BAATI HUM" days="Mon,Tue,Wed,Thu,Fri" StartTime="21:00" EndTime="21:30" CostPer10s="120" caption="Colgate Kids Jumping" AdDuration="20" NumberOfSpots="9">

                '                <selected_spots><!-- this is a demo output - spot count and dates may be incorrect. attributes are correctNO SPACES BETWEEN VALUES IN CSV -->
                '                    <spot log="004,10082013,032238,032308,0,001,005,00030"/>
                '                </selected_spots>
                '            </programme>
                '        </period>
                '    </plan>

                '</mediaplan>


                '  request = WebRequest.Create("")
                '

                'request.Method = "POST"
                'request.ContentType = "application/x-www-form-urlencoded"
                'request.Timeout = 300000
                'request.ServicePoint.MaxIdleTime = 300000
                ''  request.KeepAlive = True
                'inputstring = HttpUtility.UrlEncode(MyExtensionClasses.OuterXml(testmediaplan))

                'stream = request.GetRequestStream()
                ''  Dim bytes As Byte() = Encoding.Default.GetBytes(input.OuterXml)

                'Dim encoding As System.Text.UTF8Encoding = New UTF8Encoding()
                'postData = "inputXML=" + inputstring

                'data = encoding.GetBytes(postData)
                ''input.Save(stream)
                '' request.ContentLength = data.Length
                'stream.Write(data, 0, data.Length)

                '' request.Proxy = Nothing
                'ws = request.GetResponse()

                'oStream = ws.GetResponseStream()
                'Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")

                '' Pipe the stream to a higher level stream reader with the required encoding format.
                'Dim readStream As New StreamReader(oStream, encode)
                '     Dim separators() As String = {"Genre,Viewership"}
                ' Dim file As [String]() = readStream.ReadToEnd().Split(separators, StringSplitOptions.RemoveEmptyEntries)

                Dim name As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                mediaplan.Save(LogDirectoryPath + "ChannelSummary_Inp_" + name)
                ' rnfoutputXml = GetOpXMLFromWS(mediaplan, "http://54.255.217.55:8080/GroupM/channelsummary/")
                rnfoutputXml = GetOpXMLFromWS(mediaplan, Globals.Ribbons.MSprintExRibbon.GetURLForWS("ChannelSummaryWSURL_New"))
                'rnfoutputXml = XElement.Load("C:\ASR\op.xml")
                If Not (rnfoutputXml Is Nothing) Then
                    Dim name1 As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                    rnfoutputXml.Save(LogDirectoryPath + "ChannelSummary_Op_" + name1)

                    If ConstructChannelSummaryTable(rnfoutputXml) Then
                        RnFChannelSummary = nbdmainSummary(RnFChannelSummary, ChannelSummaryHRN)
                        '  RnFChannelSummary = ReverseCalculateTVRReach(RnFChannelSummary)
                        If Not CheckSheetExists("Channel Summary") Then
                            nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                            nativeSheet.Name = "Channel Summary"
                        Else
                            'nativeSheet = CheckAndReturnSheet("Channel Summary")
                            ''  newSheet.UsedRange.Clear()
                            'Globals.Ribbons.MSprintExRibbon.CleanSheet(nativeSheet)
                            'nativeSheet.Activate()

                            Dim sheetcount As Integer = CheckAndReturnSheet("Channel Summary")
                            If sheetcount > 0 Then
                                nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                                Dim sname As String = String.Format("Channel Summary({0})", sheetcount)
                                nativeSheet.Name = sname
                            End If

                        End If
                        Dim tgcell As Microsoft.Office.Interop.Excel.Range = nativeSheet.Range("$A$2", Type.Missing)
                        tgcell.Value2 = "TG : " + plantgname
                        Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(nativeSheet)

                        If tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                            '  Dim periodcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.Range("$A$3", Type.Missing)
                            Dim periodcell As Microsoft.Office.Interop.Excel.Range = tgcell.Next(2, 0)
                            Dim period As String = String.Format("Period : {0} To {1}", tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToString("dd/MM/yyyy"), tpSelections.TaskPaneLogFile1.dtToDate.Value.ToString("dd/MM/yyyy"))
                            periodcell.Value = period
                            Dim mgs As List(Of String) = New List(Of String)()
                            For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                                mgs.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString())
                            Next
                            mgs.Add("TotalMarkets")

                            ' For Each mg As String In mgs
                            Dim count As Integer = 0
                            For index = 0 To mgs.Count - 1

                                '   Next

                                Dim filter As String = String.Format("Market='{0}'", mgs(index))
                                Dim rows As Data.DataRow() = RnFChannelSummary.Select(filter)
                                Dim mgcell As Microsoft.Office.Interop.Excel.Range = periodcell.Next(2 + (count * index), 0)
                                mgcell.Value2 = mgs(index)
                                Dim lrange As Microsoft.Office.Interop.Excel.Range = mgcell.Next(2, 0)
                                Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "RnfCSummary" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString())
                                Dim table As Data.DataTable = rows.CopyToDataTable
                                table.Columns.RemoveAt(0)
                                table.Columns.RemoveAt(1)
                                If ChannelSummaryHRN > 20 Then

                                    For index1 = 21 To ChannelSummaryHRN
                                        table.Columns.Remove(index1.ToString() + "+")
                                    Next

                                End If
                                listobject.DataSource = table
                                count = table.Rows.Count + 3
                                'rowcount += table.Rows.Count + 2
                                listobject.AutoSetDataBoundColumnHeaders = True
                                listobject.ShowAutoFilter = False

                            Next
                        ElseIf tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                            Dim count As Integer = 4
                            Dim mgs As List(Of String) = New List(Of String)()
                            For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                                mgs.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString())
                            Next
                            mgs.Add("TotalMarkets")

                            For index = 0 To tpSelections.TaskPaneLogFile1.dtWeeks.Rows.Count - 1
                                For index1 = 0 To mgs.Count - 1

                                    Dim periodcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count, 1), vstoWorkbook.Cells(count, 1))
                                    ' Dim period As String = String.Format("Period : '{0}' To '{1}'", tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToString("dd/MM/yyyy"), tpSelections.TaskPaneLogFile1.dtToDate.Value.ToString("dd/MM/yyyy"))
                                    Dim period As String = String.Format("Period : {0} To {1}", Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("StartDate").ToString()).ToString("dd/MM/yyyy"), Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("EndDate").ToString()).ToString("dd/MM/yyyy"))
                                    periodcell.Value = period
                                    '  Dim mgcell As String = String.Format("Period : {0} To {1}", Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("StartDate").ToString()).ToString("dd/MM/yyyy"), Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("EndDate").ToString()).ToString("dd/MM/yyyy"))
                                    Dim mgcell As Microsoft.Office.Interop.Excel.Range = periodcell.Next(2, 0)

                                    mgcell.Value = mgs(index1)

                                    Dim lrange As Microsoft.Office.Interop.Excel.Range = mgcell.Next(2, 0)
                                    Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "RnfCSummary" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString())
                                    Dim filter As String = String.Format("WeekNum='{0}' and Market='{1}'", tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("WeekNumber").ToString(), mgs(index1))
                                    Dim rows As Data.DataRow() = RnFChannelSummary.Select(filter)
                                    If rows.Length > 0 Then
                                        Dim table As Data.DataTable = rows.CopyToDataTable()
                                        table.Columns.RemoveAt(0)
                                        table.Columns.RemoveAt(1)
                                        If ChannelSummaryHRN > 20 Then

                                            For index2 = 21 To ChannelSummaryHRN
                                                table.Columns.Remove(index2.ToString() + "+")
                                            Next

                                        End If
                                        listobject.DataSource = table
                                        count = count + table.Rows.Count + 6
                                        'rowcount += table.Rows.Count + 2
                                        listobject.AutoSetDataBoundColumnHeaders = True
                                        listobject.ShowAutoFilter = False
                                    End If

                                Next
                            Next

                        End If
                    End If
                Else
                    MessageBox.Show("Unable to retreive requested Channel Summary details from server.")

                End If
                SetNormalCursor()
            End If

        Catch ex As Exception
            SetNormalCursor()
            LogMpsrintExException("Exception occured while retreiving Channel Summary details." + ex.Message)
            MessageBox.Show("Exception occured while retreiving Channel Summary details.Please refer Error log for more details")
        End Try
    End Sub

    Private Sub btnDurationSummary_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnDurationSummary.Click
        Dim frm As frmWait
        Try

            'Dim str As String = String.Empty
            'For Each row As Microsoft.Office.Interop.Excel.ListRow In loSpotSelection.ListRows
            '    '  loSpotSelection.ListRows(0).Range.Cells.
            '    'For Each cell As Microsoft.Office.Interop.Excel.Range In row.Range.Cells
            '    '    ' str = row.Range.Cells(row.Index, col.Index).ToString()
            '    '    str = cell.Text
            '    'Next

            'Next

            'If Not (RnFAvaiSpots Is Nothing) Then
            '    RnFAvaiSpots.Rows.Clear()
            'End If

            ' Dim dtt As BindingSource = CType(channelMapping.dgvChannels.DataSource, BindingSource)

            Dim plantgname As String = String.Empty
            Dim reftgname As String = String.Empty
            Dim dtable As System.Data.DataTable = DirectCast(tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            plantgname = dtable.Rows(0)(1).ToString().Trim()
            '  reftgname = dtable.Rows(1)(1).ToString().Trim()
            If plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please make Target Group and Market Group(s) Selections")
            ElseIf plantgname.Length > 1 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for chosen TG")
            ElseIf plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count > 0 Then
                System.Windows.Forms.MessageBox.Show("Please choose Planning Target group for Market groups")
            ElseIf loSpotSelection Is Nothing Then
                System.Windows.Forms.MessageBox.Show("Please enter plan to view Duration Summary")
            ElseIf Not (isPlanClean) Then
                System.Windows.Forms.MessageBox.Show("Please clean plan and ensure no duplication of rows")
            ElseIf Not (AllChannelsMapped()) Then
                System.Windows.Forms.MessageBox.Show("Please map all channels with master channellist")
                'ElseIf reftgname.Length > 0 And tpSelections.UcMarkets1.lbRef.Items.Count < 1 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for reference Target Group chosen")
            Else
                SetWaitCursor("Getting Duration Summary details..")

                '  Globals.ThisAddIn.Application.DoEvents()
                System.Windows.Forms.Application.DoEvents()
                'frm = New frmWait()
                'frm.Show()
                ' loSpotSelection.RefreshDataRows()
                xecellineItemsTable = CType(loSpotSelection.DataSource, Data.DataTable)
                Dim durationtable As Data.DataTable = xecellineItemsTable.Copy()
                durationtable = durationtable.DefaultView.ToTable(True, "Duration")
                RnFDurationSummary = New Data.DataTable()
                RnFDurationSummary.Columns.Add("WeekNum")
                RnFDurationSummary.Columns.Add("Duration")
                RnFDurationSummary.Columns.Add("Market")
                RnFDurationSummary.Columns.Add("Universe")
                RnFDurationSummary.Columns.Add("GRP")
                RnFDurationSummary.Columns.Add("AOTS")
                RnFDurationSummary.Columns.Add("1+")
                RnFDurationSummary.Columns.Add("2+")
                RnFDurationSummary.Columns.Add("3+")
                RnFDurationSummary.Columns.Add("4+")
                RnFDurationSummary.Columns.Add("5+")
                RnFDurationSummary.Columns.Add("6+")
                RnFDurationSummary.Columns.Add("7+")
                RnFDurationSummary.Columns.Add("8+")
                RnFDurationSummary.Columns.Add("9+")
                RnFDurationSummary.Columns.Add("10+")
                '  For Each duration As Data.DataRow In durationtable.Rows
                ' xecellineItemsTable = CType(loSpotSelection.DataSource, Data.DataTable)
                'Dim filter As String = String.Format("Duration ='{0}'", duration(0).ToString())
                'Dim rows As Data.DataRow() = xecellineItemsTable.Select(filter)
                'xecellineItemsTable = rows.CopyToDataTable()
                mediaplan = ConstructDurationSummaryInputXML()
                ' mediaplan = ConstructInputRnFXML()



                'testmediaplan =
                '<mediaplan>
                '    <!-- TO BE FROZEN 26 Dec 13 -->
                '    <!-- common section start -->
                '    <PreEvalPeriod>
                '        <StartDate>20130804</StartDate>
                '        <EndDate>20130817</EndDate>
                '    </PreEvalPeriod>

                '    <DayParts>
                '        <DayPart>0800-1200</DayPart>
                '        <DayPart>2100-2200</DayPart>
                '    </DayParts>


                '    <!-- common section ends -->


                '    <tg name="CS 15-44" cs="1" sec="1,2,3,4" sex="1,2" age="3,4,5">

                '        <mg name="mg1" type="group">
                '            <market>1</market>
                '            <market>3</market>
                '        </mg>

                '        <mg name="mg2" type="single">
                '            <market>2</market>
                '        </mg>
                '    </tg>


                '    <!-- TVR000s,TVR,GRP000s,GRP,AvgFreq,CummCost,SpotCPRP,CummCPRP,Reach000s,R1,R2,R3,R4,R5,R6,R7,R8,R9,R10 -->

                '    <!-- all output -->
                '    <plan type="weekwise"><!-- clubbed-->
                '        <period StartDate="20130804" EndDate="20130810" year="2013" WeekNum="32">
                '            <programme guid="1" SeqNumber="1" ChannelCode="004" ChannelName="Star Plus" ProgName="Yeh Rishta Kya Kehlata Hai" days="Thu" StartTime="21:30" EndTime="22:00" CostPer10s="150" caption="Colgate Kids Jumping" AdDuration="30" NumberOfSpots="10">

                '                <selected_spots><!-- this is a demo output - spot count and dates may be incorrect. attributes are correctNO SPACES BETWEEN VALUES IN CSV -->
                '                    <spot log="004,10082013,032238,032308,0,001,005,00030"/>
                '                </selected_spots>
                '            </programme>
                '            <programme guid="2" SeqNumber="2" ChannelCode="004" ChannelName="Star Plus" ProgName="DIYA AUR BAATI HUM" days="Mon,Tue,Wed,Thu,Fri" StartTime="21:00" EndTime="21:30" CostPer10s="120" caption="Colgate Kids Jumping" AdDuration="20" NumberOfSpots="9">

                '                <selected_spots><!-- this is a demo output - spot count and dates may be incorrect. attributes are correctNO SPACES BETWEEN VALUES IN CSV -->
                '                    <spot log="004,10082013,032238,032308,0,001,005,00030"/>
                '                </selected_spots>
                '            </programme>
                '        </period>
                '    </plan>

                '</mediaplan>


                '  request = WebRequest.Create("")
                '

                'request.Method = "POST"
                'request.ContentType = "application/x-www-form-urlencoded"
                'request.Timeout = 300000
                'request.ServicePoint.MaxIdleTime = 300000
                ''  request.KeepAlive = True
                'inputstring = HttpUtility.UrlEncode(MyExtensionClasses.OuterXml(testmediaplan))

                'stream = request.GetRequestStream()
                ''  Dim bytes As Byte() = Encoding.Default.GetBytes(input.OuterXml)

                'Dim encoding As System.Text.UTF8Encoding = New UTF8Encoding()
                'postData = "inputXML=" + inputstring

                'data = encoding.GetBytes(postData)
                ''input.Save(stream)
                '' request.ContentLength = data.Length
                'stream.Write(data, 0, data.Length)

                '' request.Proxy = Nothing
                'ws = request.GetResponse()

                'oStream = ws.GetResponseStream()
                'Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")

                '' Pipe the stream to a higher level stream reader with the required encoding format.
                'Dim readStream As New StreamReader(oStream, encode)
                '     Dim separators() As String = {"Genre,Viewership"}
                ' Dim file As [String]() = readStream.ReadToEnd().Split(separators, StringSplitOptions.RemoveEmptyEntries)

                Dim name As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                mediaplan.Save(LogDirectoryPath + "DurationSummary_Inp_" + name)
                '  rnfoutputXml = GetOpXMLFromWS(mediaplan, "http://54.255.217.55:8080/GroupM/durationsummary/")
                rnfoutputXml = GetOpXMLFromWS(mediaplan, Globals.Ribbons.MSprintExRibbon.GetURLForWS("DurationSummaryWSURL_New"))
                'rnfoutputXml = XElement.Load("C:\ASR\op.xml")
                If Not (rnfoutputXml Is Nothing) Then
                    Dim name1 As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                    rnfoutputXml.Save(LogDirectoryPath + "DurationSummary_Op_" + name1)
                    ' ConstructCreativeSummaryTable(rnfoutputXml, creative(0).ToString())
                    ConstructDurationSummaryTable(rnfoutputXml)
                    RnFDurationSummary = nbdmainSummary(RnFDurationSummary, DSHRN)
                    '  RnFDurationSummary = ReverseCalculateTVRReach(RnFDurationSummary)
                End If
                '  Next


                '  If ConstructMarketSummaryTable(rnfoutputXml) Then
                If Not CheckSheetExists("Duration Summary") Then
                    nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                    nativeSheet.Name = "Duration Summary"
                Else
                    'nativeSheet = CheckAndReturnSheet("Duration Summary")
                    ''  newSheet.UsedRange.Clear()
                    'Globals.Ribbons.MSprintExRibbon.CleanSheet(nativeSheet)
                    'nativeSheet.Activate()

                    Dim sheetcount As Integer = CheckAndReturnSheet("Duration Summary")
                    If sheetcount > 0 Then
                        nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                        Dim sname As String = String.Format("Duration Summary({0})", sheetcount)
                        nativeSheet.Name = sname
                    End If
                End If
                Dim tgcell As Microsoft.Office.Interop.Excel.Range = nativeSheet.Range("$A$2", Type.Missing)
                tgcell.Value2 = "TG: " + plantgname
                Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(nativeSheet)

                If tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                    ' Dim periodcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.Range("$A$3", Type.Missing)
                    Dim periodcell As Microsoft.Office.Interop.Excel.Range = tgcell.Next(2, 0)
                    Dim period As String = String.Format("Period : {0} To {1}", tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToString("dd/MM/yyyy"), tpSelections.TaskPaneLogFile1.dtToDate.Value.ToString("dd/MM/yyyy"))
                    periodcell.Value = period
                    Dim lrange As Microsoft.Office.Interop.Excel.Range = periodcell.Next(2, 0)
                    Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "RnfDSummary" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString())
                    Dim table As Data.DataTable = RnFDurationSummary.Copy()
                    table.Columns.RemoveAt(0)
                    If DSHRN > 20 Then

                        For index = 21 To DSHRN
                            table.Columns.Remove(index.ToString() + "+")
                        Next

                    End If
                    listobject.DataSource = table
                    'rowcount += table.Rows.Count + 2
                    listobject.AutoSetDataBoundColumnHeaders = True
                    listobject.ShowAutoFilter = False
                ElseIf tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                    Dim count As Integer = 4
                    For index = 0 To tpSelections.TaskPaneLogFile1.dtWeeks.Rows.Count - 1
                        Dim periodcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count, 1), vstoWorkbook.Cells(count, 1))
                        '   Dim period As String = String.Format("Period : '{0}' To '{1}'", tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToString("dd/MM/yyyy"), tpSelections.TaskPaneLogFile1.dtToDate.Value.ToString("dd/MM/yyyy"))
                        Dim period As String = String.Format("Period : {0} To {1}", Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("StartDate").ToString()).ToString("dd/MM/yyyy"), Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("EndDate").ToString()).ToString("dd/MM/yyyy"))
                        periodcell.Value = period
                        Dim lrange As Microsoft.Office.Interop.Excel.Range = periodcell.Next(2, 0)
                        Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "RnfDSummary" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString())
                        Dim filter As String = String.Format("WeekNum='{0}'", tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("WeekNumber").ToString())
                        Dim rows As Data.DataRow() = RnFDurationSummary.Select(filter)
                        If rows.Length > 0 Then
                            Dim table As Data.DataTable = rows.CopyToDataTable()
                            table.Columns.RemoveAt(0)
                            If DSHRN > 20 Then

                                For index1 = 21 To DSHRN
                                    table.Columns.Remove(index1.ToString() + "+")
                                Next

                            End If
                            listobject.DataSource = table
                            count = count + table.Rows.Count + 3
                            'rowcount += table.Rows.Count + 2
                            listobject.AutoSetDataBoundColumnHeaders = True
                            listobject.ShowAutoFilter = False
                        End If


                    Next

                End If
                ' End If

            End If
            '  End If
            SetNormalCursor()
        Catch ex As Exception
            SetNormalCursor()
            LogMpsrintExException("Exception occured while retreiving duration summary details." + ex.Message)
            MessageBox.Show("Exception occured while retreiving duration summary details.Please refer to error log for more details")
        End Try
    End Sub

    Private Sub btnGetReqSpots_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnGetReqSpots.Click
        ' Try
        Dim frm As frmWait
        Try

            'Dim str As String = String.Empty
            'For Each row As Microsoft.Office.Interop.Excel.ListRow In loSpotSelection.ListRows
            '    '  loSpotSelection.ListRows(0).Range.Cells.
            '    'For Each cell As Microsoft.Office.Interop.Excel.Range In row.Range.Cells
            '    '    ' str = row.Range.Cells(row.Index, col.Index).ToString()
            '    '    str = cell.Text
            '    'Next

            'Next

            If Not (RnFAvaiSpots Is Nothing) Then
                RnFAvaiSpots.Rows.Clear()
            End If

            'If Not (planOpenedSuccessfully) Then

            'End If

            ' Dim dtt As BindingSource = CType(channelMapping.dgvChannels.DataSource, BindingSource)

            Dim plantgname As String = String.Empty
            Dim reftgname As String = String.Empty
            Dim dtable As System.Data.DataTable = DirectCast(tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            plantgname = dtable.Rows(0)(1).ToString().Trim()
            ' reftgname = dtable.Rows(1)(1).ToString().Trim()
            If Not (planOpenedSuccessfully) And plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please make Target Group and Market Group(s) Selections")
            ElseIf Not (planOpenedSuccessfully) And plantgname.Length > 1 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for chosen TG")
            ElseIf Not (planOpenedSuccessfully) And plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count > 0 Then
                System.Windows.Forms.MessageBox.Show("Please choose Planning Target group for Market groups")
            ElseIf Not (planOpenedSuccessfully) And loSpotSelection Is Nothing Then
                System.Windows.Forms.MessageBox.Show("Please enter plan to retreive spots")
            ElseIf Not (planOpenedSuccessfully) And Not (isPlanClean) Then
                System.Windows.Forms.MessageBox.Show("Please clean plan and ensure no duplication of rows")
            ElseIf Not (planOpenedSuccessfully) And Not (AllChannelsMapped()) Then
                System.Windows.Forms.MessageBox.Show("Please map all channels with master channellist")
                'ElseIf reftgname.Length > 0 And tpSelections.UcMarkets1.lbRef.Items.Count < 1 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for reference Target Group chosen")
            Else
                SetWaitCursor("Retreiving requested spots from server..")

                '  Globals.ThisAddIn.Application.DoEvents()
                System.Windows.Forms.Application.DoEvents()
                'frm = New frmWait()
                'frm.Show()
                ' loSpotSelection.RefreshDataRows()
                xecellineItemsTable = CType(loSpotSelection.DataSource, Data.DataTable)
                mediaplan = ConstructInputRnFXML("GetRequestedSpotsWS")
                Dim name As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                mediaplan.Save(LogDirectoryPath + "GetReqSpots_Inp_" + name)
                ' rnfoutputXml = GetOpXMLFromWS(mediaplan, "http://ec2-54-255-217-55.ap-southeast-1.compute.amazonaws.com:8080/GroupM-new/spotselectionnew/getselectedspot")
                rnfoutputXml = GetOpXMLFromWS(mediaplan, Globals.Ribbons.MSprintExRibbon.GetURLForWS("RnFTillZeroWSURL_New"))
                'rnfoutputXml = XElement.Load("C:\ASR\op.xml")
                If Not (rnfoutputXml Is Nothing) Then

                    Dim name1 As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                    rnfoutputXml.Save(LogDirectoryPath + "GetReqSpots_Op_" + name1)

                    If ConstructOpRnFTable(rnfoutputXml) Then
                        btnLogFile.Enabled = True
                        ' xecellTableCopy = xecellineItemsTable.Copy()
                        xecellTableCopy = xecelTable.Copy()
                        Try
                            If Not (xecelTable.Columns.Contains("Number of Spots Returned")) Then
                                xecelTable.Columns.Add("Number of Spots Returned")
                            End If
                            For Each row As Data.DataRow In xecelTable.Rows
                                Dim filter1 As String = String.Format("GUID='{0}'", row("GUID").ToString())
                                Dim rows1 As Data.DataRow() = RnFSelectedSpots.Select(filter1)
                                row("Number of Spots Returned") = rows1.Length
                            Next
                            loSpotSelection.SetDataBinding(xecelTable)
                        Catch ex As Exception
                            LogMpsrintExException("Exception occured while calculating number of returned spots")
                        End Try
                        If Not CheckSheetExists("Selected_Spots") Then
                            nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                            nativeSheet.Name = "Selected_Spots"
                        Else
                            nativeSheet = ReturnActualSheet("Selected_Spots")
                            '  newSheet.UsedRange.Clear()
                            Globals.Ribbons.MSprintExRibbon.CleanSheet(nativeSheet)
                            nativeSheet.Activate()
                            'Dim sheetcount As Integer = CheckAndReturnSheet("Selected_Spots")
                            'If sheetcount > 0 Then
                            '    nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                            '    Dim sname As String = String.Format("Selected_Spots({0})", sheetcount)
                            '    nativeSheet.Name = sname
                            'End If
                        End If
                        Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(nativeSheet)
                        '  Dim lrange As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(4, 1 + ((ii + 1) * index)), vstoWorkbook.Cells(4, 1 + (ii * index) + ii)), Microsoft.Office.Interop.Excel.Range)
                        Dim lrange As Microsoft.Office.Interop.Excel.Range = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(1, 1), vstoWorkbook.Cells(1, 1)), Microsoft.Office.Interop.Excel.Range)
                        'lrange.Validation.InputMessage = "Hello..Hi"
                        'lrange.Validation.ShowInput = True
                        selectedSpotsListObject = vstoWorkbook.Controls.AddListObject(lrange, "Selected_Spots" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString())
                        selectedSpotsListObject.AutoSetDataBoundColumnHeaders = True

                        '  listobject.ListColumns(0).

                        'If reftg.Length > 0 Then
                        'grid.Columns("GUID").Visible = False
                        'grid.Columns("Spot").Visible = False
                        'grid.Columns("Start Date").Visible = False
                        'grid.Columns("End Date").Visible = False
                        'grid.Columns("WeekNum").Visible = False
                        Dim rnfselectedCopy As Data.DataTable = RnFSelectedSpots.Copy()
                        rnfselectedCopy.Columns.Remove("GUID")
                        rnfselectedCopy.Columns.Remove("Spot")
                        rnfselectedCopy.Columns.Remove("Start Date")
                        rnfselectedCopy.Columns.Remove("End Date")

                        If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                            rnfselectedCopy.Columns.Remove("WeekNum")
                        End If


                        'Dim range As Microsoft.Office.Interop.Excel.Range = selectedSpotsListObject.ListRows(1).Range
                        'range.Validation.ShowInput = True
                        'range.Validation.InputMessage = "Hello..Hi"
                        selectedSpotsListObject.DataSource = rnfselectedCopy
                        selectedSpotsListObject.ListColumns(5).Range.NumberFormat = "dd/MM/yyyy"
                        EnableSummaryButtons()
                        'btnChannelSummary.Enabled = True
                        'btnMarketSummary.Enabled = True
                        'btnCreativeSummary.Enabled = True
                        'btnDurationSummary.Enabled = True
                        'btnLogFile.Enabled = True
                        ' vstoWorkbook.Protect(Type.Missing, Type.Missing)
                        'If Not SpotSelectionPane Is Nothing Then
                        '    SpotSelectionPane.Visible = False
                        '    SpotSelectionPane.Dispose()
                        'End If

                        ''mpTpSpotSelection = New ucSpotSelection()
                        ''currentLineItem = xecelTable.AsEnumerable().First()("GUID").ToString()
                        ''Dim filter As String = String.Format("GUID = '{0}'", xecelTable.AsEnumerable().First()("GUID").ToString())
                        ''Dim rows As Data.DataRow() = RnFSelectedSpots.Select(filter)
                        ' ''Dim spotst As Data.DataTable = New Data.DataTable()
                        ' ''spotst = RnFSelectedSpots.Clone()
                        ' ''For Each row1 As Data.DataRow In rows
                        ' ''    spotst.ImportRow(row1)
                        ' ''Next
                        ''mpTpSpotSelection.dgSelectedSpotsGrid.DataSource = rows.CopyToDataTable()
                        ''HideSelectedSpotsGrid(mpTpSpotSelection.dgSelectedSpotsGrid)
                        ' ''  mpTpSpotSelection.Anchor = AnchorStyles.Bottom

                        ' ''mpTpSpotSelection.Anchor = (AnchorStyles.Bottom Or AnchorStyles.Right Or AnchorStyles.Top Or AnchorStyles.Left)
                        ' '' mpTpSpotSelection.dgSelectedSpotsGrid.Refresh()
                        ''SpotSelectionPane = Globals.ThisAddIn.CustomTaskPanes.Add(mpTpSpotSelection, "Spot Selection")
                        ''SpotSelectionPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
                        ''SpotSelectionPane.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange
                        ''mpTpSpotSelection.Dock = DockStyle.Fill
                        ' SpotSelectionPane.
                        'SpotSelectionPane.Width = 800
                        'SpotSelectionPane.Height = 400

                        '               Dim paneSize As System.Drawing.Size = _
                        'New System.Drawing.Size(SpotSelectionPane.Width, SpotSelectionPane.Height)
                        '               mpTpSpotSelection.AutoSize = True
                        '               mpTpSpotSelection.AutoSizeMode = AutoSizeMode.GrowAndShrink
                        ' mpTpSpotSelection.si.FlowPanel.Size = paneSize
                        'SpotSelectionPane.Width = 300
                        ''SpotSelectionPane.Height = 200
                        'DisplayCurrentPlanItem()
                        'SpotSelectionPane.Visible = True
                        '                Catch ex As Exception
                        '    LogMpsrintExException("Exception occured while displaying spot selection pane." + ex.Message)

                        'End Try
                    End If
                Else
                    MessageBox.Show("Unable to retreive requested logs from server.")
                End If
                '   SetNormalCursor()

                'If Not (frm Is Nothing) Then
                '    frm.Close()

                'End If

            End If
        Catch ex As Exception
            ' SetNormalCursor()
            LogMpsrintExException("Exception occured while getting requested spots" + ex.Message)
            MessageBox.Show("Exception occured while retreiving requested spots.Please refer to error log for more details.")
        Finally
            SetNormalCursor()
        End Try

    End Sub

    Private Sub btnChangeLogDir_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnChangeLogDir.Click
        Try
            Dim frmchangeLog As ChangeLogDir = New ChangeLogDir()
            frmchangeLog.ShowDialog()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnReorderPlanChannels_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnReorderPlanChannels.Click
        Try
            Dim frmReorderPlan As frmRearrangePlanChannels = New frmRearrangePlanChannels()
            frmReorderPlan.ShowDialog()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnGenreShareAll_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnGenreShareAll.Click
        Dim waitFrm As frmWait
        Try
            Dim plantg As String = String.Empty
            Dim reftg As String = String.Empty


            Dim dtable As System.Data.DataTable = DirectCast(tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            plantg = dtable.Rows(0)(1).ToString().Trim()
            'reftg = dtable.Rows(1)(1).ToString()

            If Not (MachineConnectedToInternet()) Then
                MessageBox.Show("MSprintEx communicates with Server over Internet.Please ensure Internet connectivity and Try again.")

            ElseIf plantg.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please make Planning Target Group and Market(s) and/or Market Group(s) Selections")
            ElseIf plantg.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count > 0 Then
                System.Windows.Forms.MessageBox.Show("Please choose Planning Target group for Market groups")
            ElseIf plantg.Length > 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please choose Market(s) and/or Market Group(s) for chosen Planning Target Group")
                'ElseIf reftg.Length = 0 And tpSelections.UcMarkets1.lbRef.Items.Count > 0 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose Reference Target group for Market groups")
                'ElseIf reftg.Length > 0 And tpSelections.UcMarkets1.lbRef.Items.Count < 1 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose  Market(s) and/or Market Group(s) for reference Target Group chosen")
            Else

                Globals.ThisAddIn.Application.StatusBar = "Getting requested Genre Share details..."
                System.Windows.Forms.Application.DoEvents()
                ' Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlWait
                '  waitFrm = New frmWait()
                'waitFrm.Show()
                'waitFrm.Refresh()
                gshareds = New DataSet()

                If reftg.Length > 0 Then
                    gsharerefds = New DataSet()
                End If

                GetGenreShare(plantg, reftg, gshareds, gsharerefds, GenreShareView.All)


                If reftg.Length > 0 And gshareds.Tables.Count > 0 Then
                    'Dim plan, ref As List(Of String)
                    'plan = New List(Of String)()
                    'ref = New List(Of String)()
                    'For index = 0 To tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                    '    plan.Add(tpSelections.UcMarkets1.lbPlan.Items(index))
                    'Next
                    'For index = 0 To tpSelections.UcMarkets1.lbRef.Items.Count - 1
                    '    ref.Add(tpSelections.UcMarkets1.lbRef.Items(index))
                    'Next
                    If CTPGenreShare Is Nothing Then
                        mpChannelShare = New ucChannelShare()
                        CTPGenreShare = Globals.ThisAddIn.CustomTaskPanes.Add(mpChannelShare, "Genre Share Selections")
                        CTPGenreShare.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
                        CTPGenreShare.Height = 226
                        CTPGenreShare.Width = 300
                        CTPGenreShare.Visible = True
                    Else
                        CTPGenreShare.Visible = True
                        '  MSprintExChannelShare.Title = "Genre Share Selections"
                    End If

                    'tpSelections.TaskPaneLogFile1.scMain.Panel2Collapsed = False
                    'If Not tpSelections.TaskPaneLogFile1.showingChannels Then
                    '    tpSelections.TaskPaneLogFile1.ucChannelsMapping.btnShowHide.Text = "Genre Share Selections"
                    '    tpSelections.TaskPaneLogFile1.ucChannelsMapping.btnShowHide.Refresh()
                    '    tpSelections.TaskPaneLogFile1.showChannelMapping(True)
                    'End If
                    'tpSelections.TaskPaneLogFile1.Show()
                End If

                'waitFrm.Close()
                'waitFrm.Dispose()
                Globals.ThisAddIn.Application.StatusBar = String.Empty
                Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault

            End If
        Catch ex As Exception

            'If Not (waitFrm) Is Nothing Then
            '    waitFrm.Close()
            '    waitFrm.Dispose()
            'End If

            Globals.ThisAddIn.Application.StatusBar = String.Empty
            Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault
            LogMpsrintExException("Exception occured while retreiving requested Genre Share details" + ex.Message)
            System.Windows.Forms.MessageBox.Show("Exception occured while getting requested Genre share details.Please refer to Error log for more details")


        End Try
    End Sub

    Private Sub btnGenreShareTopTen_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnGenreShareTopTen.Click
        Dim waitFrm As frmWait
        Try
            Dim plantg As String = String.Empty
            Dim reftg As String = String.Empty


            Dim dtable As System.Data.DataTable = DirectCast(tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            plantg = dtable.Rows(0)(1).ToString().Trim()
            ' reftg = dtable.Rows(1)(1).ToString()

            If Not (MachineConnectedToInternet()) Then
                MessageBox.Show("MSprintEx communicates with Server over Internet.Please ensure Internet connectivity and Try again.")

            ElseIf plantg.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please make Planning Target Group and Market(s) and/or Market Group(s) Selections")
            ElseIf plantg.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count > 0 Then
                System.Windows.Forms.MessageBox.Show("Please choose Planning Target group for Market groups")
            ElseIf plantg.Length > 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please choose Market(s) and/or Market Group(s) for chosen Planning Target Group")
                'ElseIf reftg.Length = 0 And tpSelections.UcMarkets1.lbRef.Items.Count > 0 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose Reference Target group for Market groups")
                'ElseIf reftg.Length > 0 And tpSelections.UcMarkets1.lbRef.Items.Count < 1 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose  Market(s) and/or Market Group(s) for reference Target Group chosen")
            Else

                Globals.ThisAddIn.Application.StatusBar = "Getting requested Genre Share details..."
                System.Windows.Forms.Application.DoEvents()
                ' Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlWait
                '  waitFrm = New frmWait()
                'waitFrm.Show()
                'waitFrm.Refresh()
                gshareds = New DataSet()

                If reftg.Length > 0 Then
                    gsharerefds = New DataSet()
                End If

                GetGenreShare(plantg, reftg, gshareds, gsharerefds, GenreShareView.TopTen)


                If reftg.Length > 0 And gshareds.Tables.Count > 0 Then
                    'Dim plan, ref As List(Of String)
                    'plan = New List(Of String)()
                    'ref = New List(Of String)()
                    'For index = 0 To tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                    '    plan.Add(tpSelections.UcMarkets1.lbPlan.Items(index))
                    'Next
                    'For index = 0 To tpSelections.UcMarkets1.lbRef.Items.Count - 1
                    '    ref.Add(tpSelections.UcMarkets1.lbRef.Items(index))
                    'Next
                    If CTPGenreShare Is Nothing Then
                        mpChannelShare = New ucChannelShare()
                        CTPGenreShare = Globals.ThisAddIn.CustomTaskPanes.Add(mpChannelShare, "Genre Share Selections")
                        CTPGenreShare.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
                        CTPGenreShare.Height = 226
                        CTPGenreShare.Width = 300
                        CTPGenreShare.Visible = True
                    Else
                        CTPGenreShare.Visible = True
                        '  MSprintExChannelShare.Title = "Genre Share Selections"
                    End If

                    'tpSelections.TaskPaneLogFile1.scMain.Panel2Collapsed = False
                    'If Not tpSelections.TaskPaneLogFile1.showingChannels Then
                    '    tpSelections.TaskPaneLogFile1.ucChannelsMapping.btnShowHide.Text = "Genre Share Selections"
                    '    tpSelections.TaskPaneLogFile1.ucChannelsMapping.btnShowHide.Refresh()
                    '    tpSelections.TaskPaneLogFile1.showChannelMapping(True)
                    'End If
                    'tpSelections.TaskPaneLogFile1.Show()
                End If

                'waitFrm.Close()
                'waitFrm.Dispose()
                Globals.ThisAddIn.Application.StatusBar = String.Empty
                Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault

            End If
        Catch ex As Exception

            'If Not (waitFrm) Is Nothing Then
            '    waitFrm.Close()
            '    waitFrm.Dispose()
            'End If

            Globals.ThisAddIn.Application.StatusBar = String.Empty
            Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault
            LogMpsrintExException("Exception occured while retreiving requested Genre Share details" + ex.Message)
            System.Windows.Forms.MessageBox.Show("Exception occured while getting requested Genre share details.Please refer to Error log for more details")


        End Try
    End Sub
    Public Function GetSpotsTable()
        Try
            If Not CheckSheetExists("SpotsPercentage") Then
                logCreator = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                logCreator.Name = "SpotsPercentage"
            Else
                'logCreator = CheckAndReturnSheet("Plan Selection")
                ''  newSheet.UsedRange.Clear()
                'Globals.Ribbons.MSprintExRibbon.CleanSheet(logCreator)
                'logCreator.Activate()
                Dim sheetcount As Integer = CheckAndReturnSheet("SpotsPercentage")
                If sheetcount > 0 Then
                    logCreator = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                    Dim name As String = String.Format("Plan Selection({0})", sheetcount)
                    logCreator.Name = name
                End If

            End If

            Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(logCreator)
            ' vstoWorkbook.Name = "Plan Selection"
            loSpotSelection = vstoWorkbook.Controls.AddListObject(vstoWorkbook.Range("$A$1"), "SpotsPercentage")
            loSpotSelection.Application.CutCopyMode = False
            loSpotSelection.Application.ScreenUpdating = False
            loSpotSelection.AutoSetDataBoundColumnHeaders = True
            loSpotSelection.DataSource = GetInpSpotDataTable()
            vstoWorkbook.get_Range("A:A", Type.Missing).EntireColumn.Hidden = True
            vstoWorkbook.get_Range("G:G", Type.Missing).EntireColumn.NumberFormat = "#"
            'loSpotSelection.ListColumns(7).DataBodyRange.NumberFormat = "0"
            'loSpotSelection.ListColumns(6).Range.NumberFormat = "hh:mm;@"
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Function
        Finally
            loSpotSelection.Application.CutCopyMode = True
            loSpotSelection.Application.ScreenUpdating = True
        End Try
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        Try

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        Try

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button1.Click
        Dim plantgname As String = String.Empty
        Dim reftgname As String = String.Empty
        Dim dtable As System.Data.DataTable = DirectCast(tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
        plantgname = dtable.Rows(0)(1).ToString().Trim()
        ' reftgname = dtable.Rows(1)(1).ToString().Trim()
        If plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
            System.Windows.Forms.MessageBox.Show("Please make Target Group and Market Group(s) Selections")
        ElseIf plantgname.Length > 1 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
            System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for chosen TG")
        ElseIf plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count > 0 Then
            System.Windows.Forms.MessageBox.Show("Please choose Planning Target group for Market groups")
        ElseIf loSpotSelection Is Nothing Then
            System.Windows.Forms.MessageBox.Show("Please enter plan to view Reach and Frequency")
        ElseIf Not (isPlanClean) Then
            System.Windows.Forms.MessageBox.Show("Please clean plan to view Reach and Frequency and ensure no duplication of rows")
        ElseIf Not (AllChannelsMapped()) Then
            System.Windows.Forms.MessageBox.Show("Please map all channels with master channellist")
            'ElseIf reftgname.Length > 0 And tpSelections.UcMarkets1.lbRef.Items.Count < 1 Then
            '    System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for reference Target Group chosen")
        Else
            SetWaitCursor("Getting Average TVR for each program..")

            '  Globals.ThisAddIn.Application.DoEvents()
            System.Windows.Forms.Application.DoEvents()
            'frm = New frmWait()
            'frm.Show()
            ' loSpotSelection.RefreshDataRows()
            xecellineItemsTable = CType(loSpotSelection.DataSource, Data.DataTable)
            ' mediaplan = ConstructInputRnFXML()
            mediaplan = ConstructProgAvgTVR()
            ' mediaplan = XElement.Load("Z:\\asr\\Errors\\RNF_Inp_20032014_150040.xml")


            'testmediaplan =
            '<mediaplan>
            '    <!-- TO BE FROZEN 26 Dec 13 -->
            '    <!-- common section start -->
            '    <PreEvalPeriod>
            '        <StartDate>20130804</StartDate>
            '        <EndDate>20130817</EndDate>
            '    </PreEvalPeriod>

            '    <DayParts>
            '        <DayPart>0800-1200</DayPart>
            '        <DayPart>2100-2200</DayPart>
            '    </DayParts>


            '    <!-- common section ends -->


            '    <tg name="CS 15-44" cs="1" sec="1,2,3,4" sex="1,2" age="3,4,5">

            '        <mg name="mg1" type="group">
            '            <market>1</market>
            '            <market>3</market>
            '        </mg>

            '        <mg name="mg2" type="single">
            '            <market>2</market>
            '        </mg>
            '    </tg>


            '    <!-- TVR000s,TVR,GRP000s,GRP,AvgFreq,CummCost,SpotCPRP,CummCPRP,Reach000s,R1,R2,R3,R4,R5,R6,R7,R8,R9,R10 -->

            '    <!-- all output -->
            '    <plan type="weekwise"><!-- clubbed-->
            '        <period StartDate="20130804" EndDate="20130810" year="2013" WeekNum="32">
            '            <programme guid="1" SeqNumber="1" ChannelCode="004" ChannelName="Star Plus" ProgName="Yeh Rishta Kya Kehlata Hai" days="Thu" StartTime="21:30" EndTime="22:00" CostPer10s="150" caption="Colgate Kids Jumping" AdDuration="30" NumberOfSpots="10">

            '                <selected_spots><!-- this is a demo output - spot count and dates may be incorrect. attributes are correctNO SPACES BETWEEN VALUES IN CSV -->
            '                    <spot log="004,10082013,032238,032308,0,001,005,00030"/>
            '                </selected_spots>
            '            </programme>
            '            <programme guid="2" SeqNumber="2" ChannelCode="004" ChannelName="Star Plus" ProgName="DIYA AUR BAATI HUM" days="Mon,Tue,Wed,Thu,Fri" StartTime="21:00" EndTime="21:30" CostPer10s="120" caption="Colgate Kids Jumping" AdDuration="20" NumberOfSpots="9">

            '                <selected_spots><!-- this is a demo output - spot count and dates may be incorrect. attributes are correctNO SPACES BETWEEN VALUES IN CSV -->
            '                    <spot log="004,10082013,032238,032308,0,001,005,00030"/>
            '                </selected_spots>
            '            </programme>
            '        </period>
            '    </plan>

            '</mediaplan>


            '  request = WebRequest.Create("")
            '

            'request.Method = "POST"
            'request.ContentType = "application/x-www-form-urlencoded"
            'request.Timeout = 300000
            'request.ServicePoint.MaxIdleTime = 300000
            ''  request.KeepAlive = True
            'inputstring = HttpUtility.UrlEncode(MyExtensionClasses.OuterXml(testmediaplan))

            'stream = request.GetRequestStream()
            ''  Dim bytes As Byte() = Encoding.Default.GetBytes(input.OuterXml)

            'Dim encoding As System.Text.UTF8Encoding = New UTF8Encoding()
            'postData = "inputXML=" + inputstring

            'data = encoding.GetBytes(postData)
            ''input.Save(stream)
            '' request.ContentLength = data.Length
            'stream.Write(data, 0, data.Length)

            '' request.Proxy = Nothing
            'ws = request.GetResponse()

            'oStream = ws.GetResponseStream()
            'Dim encode As Encoding = System.Text.Encoding.GetEncoding("utf-8")

            '' Pipe the stream to a higher level stream reader with the required encoding format.
            'Dim readStream As New StreamReader(oStream, encode)
            '     Dim separators() As String = {"Genre,Viewership"}
            ' Dim file As [String]() = readStream.ReadToEnd().Split(separators, StringSplitOptions.RemoveEmptyEntries)

            Dim name As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
            mediaplan.Save(LogDirectoryPath + "ProgAvgTVR_Inp_" + name)
            ' rnfoutputXml = GetOpXMLFromWS(mediaplan, "http://54.255.217.55:8080/GroupM/breaktvr/")
            rnfoutputXml = GetOpXMLFromWS(mediaplan, Globals.Ribbons.MSprintExRibbon.GetURLForWS("AvgTVRWSURL_New"))

            'rnfoutputXml = XElement.Load("C:\ASR\op.xml")
            If Not (rnfoutputXml Is Nothing) Then
                Dim name1 As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"

                rnfoutputXml.Save(LogDirectoryPath + "ProgAvgTVR_Op_" + name1)

                If ConstructOpProgAvgTVRTable(rnfoutputXml) Then

                    btnLogFile.Enabled = True
                    ' xecellTableCopy = xecellineItemsTable.Copy()
                    xecellTableCopy = xecelTable.Copy()
                    Try
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("MG")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Avg TVR")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Std Deviation")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("Total available breaks")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("0 to m - 2s")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("(m - 2s) to (m - s)")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("(m - s) to m")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("m to (m + s)")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("(m + s) to (m + 2s)")
                        'Globals.Ribbons.MSprintExRibbon.RnFProgAvgTVRTable.Columns.Add("> (m + 2s)")
                        If Not (xecelTable.Columns.Contains("MG")) Then
                            xecelTable.Columns.Add("MG")
                        End If
                        If Not (xecelTable.Columns.Contains("0TVR Spots")) Then
                            xecelTable.Columns.Add("0TVR Spots")
                        End If
                        If Not (xecelTable.Columns.Contains("Avg TVR")) Then
                            xecelTable.Columns.Add("Avg TVR")
                        End If
                        If Not (xecelTable.Columns.Contains("Std Deviation")) Then
                            xecelTable.Columns.Add("Std Deviation")
                        End If
                        If Not (xecelTable.Columns.Contains("Total available breaks")) Then
                            xecelTable.Columns.Add("Total available breaks")
                        End If



                        'If Not (xecelTable.Columns.Contains("0 to m - 2s")) Then
                        '    xecelTable.Columns.Add("0 to m - 2s")
                        'End If
                        'If Not (xecelTable.Columns.Contains("(m - 2s) to (m - s)")) Then
                        '    xecelTable.Columns.Add("(m - 2s) to (m - s)")
                        'End If
                        If Not (xecelTable.Columns.Contains("(Avg TVR- SD) to  Avg TVR")) Then
                            xecelTable.Columns.Add("(Avg TVR- SD) to  Avg TVR")
                        End If
                        If Not (xecelTable.Columns.Contains("Avg TVR to (Avg TVR + SD)")) Then
                            xecelTable.Columns.Add("Avg TVR to (Avg TVR + SD)")
                        End If
                        'If Not (xecelTable.Columns.Contains("(m + s) to (m + 2s)")) Then
                        '    xecelTable.Columns.Add("(m + s) to (m + 2s)")
                        'End If
                        'If Not (xecelTable.Columns.Contains("> (m + 2s)")) Then
                        '    xecelTable.Columns.Add("> (m + 2s)")
                        'End If

                        For Each row As Data.DataRow In xecelTable.Rows
                            Dim filter As String = String.Format("GUID='{0}'", row("GUID").ToString())
                            Dim rows As Data.DataRow() = RnFProgAvgTVRTable.Select(filter)

                            If rows.Count > 0 Then
                                row("MG") = rows(0)("MG").ToString()
                                row("0TVR Spots") = rows(0)("0TVR Spots").ToString()
                                row("Avg TVR") = rows(0)("Avg TVR").ToString()
                                row("Std Deviation") = rows(0)("Std Deviation").ToString()
                                row("Total available breaks") = rows(0)("Total available breaks").ToString()
                                '  row("0 to m - 2s") = rows(0)("Break 0 to m - 2s").ToString() 'Break 0 to m - 2s
                                '   row("(m - 2s) to (m - s)") = rows(0)("Break (m - 2s) to (m - s)").ToString() 'Break (m - 2s) to (m - s)
                                row("(Avg TVR- SD) to  Avg TVR") = rows(0)("Break (m - s) to m").ToString() 'Break (m - s) to m 
                                row("Avg TVR to (Avg TVR + SD)") = rows(0)("Break m to (m + s)").ToString() 'Break m to (m + s) 
                                '  row("(m + s) to (m + 2s)") = rows(0)("Break (m + s) to (m + 2s)").ToString() 'Break (m + s) to (m + 2s)
                                '  row("> (m + 2s)") = rows(0)("Break > (m + 2s)").ToString() 'Break > (m + 2s) 
                            End If
                        Next
                        loSpotSelection.SetDataBinding(xecelTable)

                        If Not (ChannelPane Is Nothing) Then ChannelPane.Dispose()
                        Dim ucAvgTVRHandle As ucAvgTVRMGSelection = New ucAvgTVRMGSelection()
                        ucAvgTVRHandle.rbBreakCount.Checked = True
                        AvgTVRMGPane = Globals.ThisAddIn.CustomTaskPanes.Add(ucAvgTVRHandle, "MG selection for AvgTVR")
                        AvgTVRMGPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
                        AvgTVRMGPane.Height = 226
                        AvgTVRMGPane.Width = 320
                        AvgTVRMGPane.Visible = True
                        'Else
                        'ChannelPane.Visible = True
                    Catch ex As Exception
                        LogMpsrintExException("Exception occured while displaying Avg TVR for plan.Message : " + ex.Message)
                    Finally
                        SetNormalCursor()
                    End Try
                End If
            End If
        End If
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button2.Click

        Try
            Dim plantgname As String = String.Empty
            Dim reftgname As String = String.Empty
            Dim dtable As System.Data.DataTable = DirectCast(tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            plantgname = dtable.Rows(0)(1).ToString().Trim()
            reftgname = dtable.Rows(1)(1).ToString()
            'Dim chds As DataSet = New DataSet()

            If plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please make Target Group and Market Group(s) Selections")
            ElseIf plantgname.Length > 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please choose Market(s) and/or Market Group(s) for chosen Primary target group")
            ElseIf plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count > 0 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose Planning Target group for Market groups")
                'ElseIf reftgname.Length > 0 And tpSelections.UcMarkets1.lbRef.Items.Count < 1 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for reference Target Group chosen")
                'ElseIf reftgname.Length = 0 And tpSelections.UcMarkets1.lbRef.Items.Count > 0 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose Reference Target group for Market groups")
                'ElseIf tpSelections.UcChannels.lbSelectedChannels.Items.Count < 1 Or tpSelections.UcAdvertiser1.lbSelectedAdvertisers.Items.Count < 1 Or tpSelections.UcBrand1.lbSelectedBrands.Items.Count < 1 Or tpSelections.UcCategory1.lbSelectedCategories.Items.Count < 1 Or tpSelections.UcVariant1.lbSelectedVariants.Items.Count < 1 Then
                '    System.Windows.Forms.MessageBox.Show("Neither Advertiser/Variant/Brand/Category/Channel can not be chosen.Please do selections.")
            Else
                System.Windows.Forms.Application.DoEvents()
                ' Dim input As XElement = ConstructBSLSummaryInputXML()
                Dim input As XElement
                'Globals.ThisAddIn.Application.DoEvents()
                SetWaitCursor("Getting requested BSL Summary details..")
                Dim name As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                input.Save(LogDirectoryPath + "BSLSummary_Inp_" + name)
                ptvrRootNode = GetOpXMLFromWS(input, "http://ec2-54-255-217-55.ap-southeast-1.compute.amazonaws.com:8080/GroupM/programtvr/")
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles Button3.Click
        CreateLog()
    End Sub

    Private Sub btnSpotSplit_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnSpotSplit.Click
        Try
            If Not CheckSheetExists("Spots Split%") Then
                logCreator = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                logCreator.Name = "Spots Split%"
            Else
                'logCreator = CheckAndReturnSheet("Plan Selection")
                ''  newSheet.UsedRange.Clear()
                'Globals.Ribbons.MSprintExRibbon.CleanSheet(logCreator)
                'logCreator.Activate()
                Dim sheetcount As Integer = CheckAndReturnSheet("Spots Split%")
                If sheetcount > 0 Then
                    logCreator = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                    Dim name As String = String.Format("Spots Split%({0})", sheetcount)
                    logCreator.Name = name
                End If

            End If

            Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(logCreator)
            ' vstoWorkbook.Name = "Plan Selection"
            loCrDura = vstoWorkbook.Controls.AddListObject(vstoWorkbook.Range("$A$1"), "CreativeDuration")
            loCrDura.Application.CutCopyMode = False
            loCrDura.Application.ScreenUpdating = False
            loCrDura.AutoSetDataBoundColumnHeaders = True
            loCrDura.DataSource = GetCreativeDuration()

            loWeekWise = vstoWorkbook.Controls.AddListObject(vstoWorkbook.Range("$E$1"), "WeekWise")
            loWeekWise.Application.CutCopyMode = False
            loWeekWise.Application.ScreenUpdating = False
            loWeekWise.AutoSetDataBoundColumnHeaders = True
            loWeekWise.DataSource = GetWeekTable()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnProgAvgTVR_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnProgAvgTVR.Click
        Try
            btnCleanupplan.Enabled = True
            btnMapChannels.Enabled = True
            btnRnF.Enabled = True
            Button1.Enabled = True
            btnGetReqSpots.Enabled = True

            btnReorderPlanChannels.Enabled = True
            dtChannelMaster = New Plandata.ChannelMasterDataTable
            '  WeekColumns = New Dictionary(Of WeekYear, Data.DataColumn)(myc)
            If Not CheckSheetExists("Average TVR") Then
                logCreator = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                logCreator.Name = "Average TVR"
            Else
                'logCreator = CheckAndReturnSheet("Plan Selection")
                ''  newSheet.UsedRange.Clear()
                'Globals.Ribbons.MSprintExRibbon.CleanSheet(logCreator)
                'logCreator.Activate()
                Dim sheetcount As Integer = CheckAndReturnSheet("Average TVR")
                If sheetcount > 0 Then
                    logCreator = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                    Dim name As String = String.Format("Average TVR({0})", sheetcount)
                    logCreator.Name = name
                End If

            End If

            Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(logCreator)
            ' vstoWorkbook.Name = "Plan Selection"
            loSpotSelection = vstoWorkbook.Controls.AddListObject(vstoWorkbook.Range("$A$1"), "ProgAvgTVR")
            loSpotSelection.Application.CutCopyMode = False
            loSpotSelection.Application.ScreenUpdating = False
            loSpotSelection.AutoSetDataBoundColumnHeaders = True
            loSpotSelection.DataSource = GetInpProgAvgTVRTable()
            vstoWorkbook.get_Range("A:A", Type.Missing).EntireColumn.Hidden = True
            vstoWorkbook.get_Range("G:G", Type.Missing).EntireColumn.NumberFormat = "#"
            'loSpotSelection.ListColumns(7).DataBodyRange.NumberFormat = "0"
            'loSpotSelection.ListColumns(6).Range.NumberFormat = "hh:mm;@"
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        Finally
            loSpotSelection.Application.CutCopyMode = True
            loSpotSelection.Application.ScreenUpdating = True
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs)
        Try

        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnShowHideAvgTVrPlan_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnShowHideAvgTVrPlan.Click
        Try
            If Not (Globals.Ribbons.MSprintExRibbon.AvgTVRMGPane Is Nothing) Then

                If Globals.Ribbons.MSprintExRibbon.AvgTVRMGPane.Visible Then
                    Globals.Ribbons.MSprintExRibbon.AvgTVRMGPane.Visible = False
                Else
                    Globals.Ribbons.MSprintExRibbon.AvgTVRMGPane.Visible = True
                End If
                'AvgTVRMGPane
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnAllSummary_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnAllSummary.Click
        Dim frm As frmWait
        Try

            'Dim str As String = String.Empty
            'For Each row As Microsoft.Office.Interop.Excel.ListRow In loSpotSelection.ListRows
            '    '  loSpotSelection.ListRows(0).Range.Cells.
            '    'For Each cell As Microsoft.Office.Interop.Excel.Range In row.Range.Cells
            '    '    ' str = row.Range.Cells(row.Index, col.Index).ToString()
            '    '    str = cell.Text
            '    'Next

            'Next

            'If Not (RnFAvaiSpots Is Nothing) Then
            '    RnFAvaiSpots.Rows.Clear()
            'End If

            ' Dim dtt As BindingSource = CType(channelMapping.dgvChannels.DataSource, BindingSource)

            Dim plantgname As String = String.Empty
            Dim reftgname As String = String.Empty
            Dim dtable As System.Data.DataTable = DirectCast(tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            plantgname = dtable.Rows(0)(1).ToString().Trim()
            '     reftgname = dtable.Rows(1)(1).ToString().Trim()
            If plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please make Target Group and Market Group(s) Selections")
            ElseIf plantgname.Length > 1 And tpSelections.UcMarkets1.lbPlan.Items.Count < 1 Then
                System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for chosen TG")
            ElseIf plantgname.Length = 0 And tpSelections.UcMarkets1.lbPlan.Items.Count > 0 Then
                System.Windows.Forms.MessageBox.Show("Please choose Planning Target group for Market groups")
            ElseIf loSpotSelection Is Nothing Then
                System.Windows.Forms.MessageBox.Show("Please enter plan to view Market Summary")
            ElseIf Not (isPlanClean) Then
                System.Windows.Forms.MessageBox.Show("Please clean plan and ensure no duplication of rows")
            ElseIf Not (AllChannelsMapped()) Then
                System.Windows.Forms.MessageBox.Show("Please map all channels with master channellist")
                'ElseIf reftgname.Length > 0 And tpSelections.UcMarkets1.lbRef.Items.Count < 1 Then
                '    System.Windows.Forms.MessageBox.Show("Please choose Market Group(s) for reference Target Group chosen")
            Else
                SetWaitCursor("Getting All Summary details..")

                '  Globals.ThisAddIn.Application.DoEvents()
                System.Windows.Forms.Application.DoEvents()
                'frm = New frmWait()
                'frm.Show()
                ' loSpotSelection.RefreshDataRows()
                xecellineItemsTable = CType(loSpotSelection.DataSource, Data.DataTable)
                mediaplan = ConstructInputRnFXML("All Summary")

                Dim name As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                mediaplan.Save(LogDirectoryPath + "AllSummary_Inp_" + name)
                ' rnfoutputXml = GetOpXMLFromWS(mediaplan, "http://54.255.217.55:8080/GroupM/mergedsummary/")
                rnfoutputXml = GetOpXMLFromWS(mediaplan, Globals.Ribbons.MSprintExRibbon.GetURLForWS("AllSummaryWSURL_New"))
                'rnfoutputXml = XElement.Load("C:\ASR\op.xml")
                If Not (rnfoutputXml Is Nothing) Then
                    Dim name1 As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                    rnfoutputXml.Save(LogDirectoryPath + "AllSummary_Op_" + name1)
                    ConstructAllSummaryTable(rnfoutputXml)

                    '  If ConstructAllSummaryTable(rnfoutputXml) Then
                    Try


                        RnFMarketSummary = nbdmainMarketSummary(RnFMarketSummary)
                        '  RnFMarketSummary = ReverseCalculateTVRReach(RnFMarketSummary)
                        If Not CheckSheetExists("Market Summary") Then
                            nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                            nativeSheet.Name = "Market Summary"
                        Else
                            'nativeSheet = CheckAndReturnSheet("Market Summary")
                            ''  newSheet.UsedRange.Clear()
                            'Globals.Ribbons.MSprintExRibbon.CleanSheet(nativeSheet)
                            'nativeSheet.Activate()

                            Dim sheetcount As Integer = CheckAndReturnSheet("Market Summary")
                            If sheetcount > 0 Then
                                nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                                Dim sname As String = String.Format("Market Summary({0})", sheetcount)
                                nativeSheet.Name = sname
                            End If


                        End If
                        Dim tgcell As Microsoft.Office.Interop.Excel.Range = nativeSheet.Range("$A$2", Type.Missing)
                        tgcell.Value2 = "TG : " + plantgname
                        Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(nativeSheet)

                        If tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                            '  Dim periodcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.Range("$A$3", Type.Missing)
                            Dim periodcell As Microsoft.Office.Interop.Excel.Range = tgcell.Next(2, 0)
                            Dim period As String = String.Format("Period : {0} To {1}", tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToString("dd/MM/yyyy"), tpSelections.TaskPaneLogFile1.dtToDate.Value.ToString("dd/MM/yyyy"))
                            periodcell.Value = period
                            Dim lrange As Microsoft.Office.Interop.Excel.Range = periodcell.Next(2, 0)
                            Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "RnfMSummary" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString())
                            Dim table As Data.DataTable = RnFMarketSummary.Copy()
                            table.Columns.RemoveAt(0)
                            If MSHRN > 20 Then

                                For index = 21 To MSHRN
                                    table.Columns.Remove(index.ToString() + "+")
                                Next

                            End If
                            listobject.DataSource = table
                            'rowcount += table.Rows.Count + 2
                            listobject.AutoSetDataBoundColumnHeaders = True
                            listobject.ShowAutoFilter = False
                        ElseIf tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                            Dim count As Integer = 4
                            For index = 0 To tpSelections.TaskPaneLogFile1.dtWeeks.Rows.Count - 1
                                Dim periodcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count, 1), vstoWorkbook.Cells(count, 1))
                                Dim period As String = String.Format("Period : {0} To {1}", Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("StartDate").ToString()).ToString("dd/MM/yyyy"), Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("EndDate").ToString()).ToString("dd/MM/yyyy"))
                                periodcell.Value = period
                                Dim lrange As Microsoft.Office.Interop.Excel.Range = periodcell.Next(2, 0)
                                Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "RnfMSummary" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString())
                                Dim filter As String = String.Format("WeekNum='{0}'", tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("WeekNumber").ToString())
                                Dim rows As Data.DataRow() = RnFMarketSummary.Select(filter)
                                If rows.Length > 0 Then
                                    Dim table As Data.DataTable = rows.CopyToDataTable()
                                    table.Columns.RemoveAt(0)
                                    If MSHRN > 20 Then

                                        For index1 = 21 To MSHRN
                                            table.Columns.Remove(index1.ToString() + "+")
                                        Next

                                    End If
                                    listobject.DataSource = table
                                    count = count + table.Rows.Count + 3
                                    'rowcount += table.Rows.Count + 2
                                    listobject.AutoSetDataBoundColumnHeaders = True
                                    listobject.ShowAutoFilter = False
                                End If
                            Next
                        End If
                        'Duration summary
                    Catch ex As Exception
                        LogMpsrintExException("Exception occured while displaying market summary.Message:" + ex.Message)
                    End Try
                    Try


                        RnFDurationSummary = nbdmainSummary(RnFDurationSummary, DSHRN)
                        ' RnFDurationSummary = ReverseCalculateTVRReach(RnFDurationSummary)
                        '  Next


                        '  If ConstructMarketSummaryTable(rnfoutputXml) Then
                        If Not CheckSheetExists("Duration Summary") Then
                            nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                            nativeSheet.Name = "Duration Summary"
                        Else
                            'nativeSheet = CheckAndReturnSheet("Duration Summary")
                            ''  newSheet.UsedRange.Clear()
                            'Globals.Ribbons.MSprintExRibbon.CleanSheet(nativeSheet)
                            'nativeSheet.Activate()

                            Dim sheetcount As Integer = CheckAndReturnSheet("Duration Summary")
                            If sheetcount > 0 Then
                                nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                                Dim sname As String = String.Format("Duration Summary({0})", sheetcount)
                                nativeSheet.Name = sname
                            End If
                        End If
                        Dim tgcell1 As Microsoft.Office.Interop.Excel.Range = nativeSheet.Range("$A$2", Type.Missing)
                        tgcell1.Value2 = "TG: " + plantgname
                        Dim vstoWorkbook1 As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(nativeSheet)

                        If tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                            ' Dim periodcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.Range("$A$3", Type.Missing)
                            Dim periodcell As Microsoft.Office.Interop.Excel.Range = tgcell1.Next(2, 0)
                            Dim period As String = String.Format("Period : {0} To {1}", tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToString("dd/MM/yyyy"), tpSelections.TaskPaneLogFile1.dtToDate.Value.ToString("dd/MM/yyyy"))
                            periodcell.Value = period
                            Dim lrange As Microsoft.Office.Interop.Excel.Range = periodcell.Next(2, 0)
                            Dim listobject = vstoWorkbook1.Controls.AddListObject(lrange, "RnfDSummary" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString())
                            Dim table As Data.DataTable = RnFDurationSummary.Copy()
                            table.Columns.RemoveAt(0)
                            If DSHRN > 20 Then

                                For index = 21 To DSHRN
                                    table.Columns.Remove(index.ToString() + "+")
                                Next

                            End If
                            listobject.DataSource = table
                            'rowcount += table.Rows.Count + 2
                            listobject.AutoSetDataBoundColumnHeaders = True
                            listobject.ShowAutoFilter = False
                        ElseIf tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                            Dim count As Integer = 4
                            For index = 0 To tpSelections.TaskPaneLogFile1.dtWeeks.Rows.Count - 1
                                Dim periodcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count, 1), vstoWorkbook.Cells(count, 1))
                                '   Dim period As String = String.Format("Period : '{0}' To '{1}'", tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToString("dd/MM/yyyy"), tpSelections.TaskPaneLogFile1.dtToDate.Value.ToString("dd/MM/yyyy"))
                                Dim period As String = String.Format("Period : {0} To {1}", Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("StartDate").ToString()).ToString("dd/MM/yyyy"), Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("EndDate").ToString()).ToString("dd/MM/yyyy"))
                                periodcell.Value = period
                                Dim lrange As Microsoft.Office.Interop.Excel.Range = periodcell.Next(2, 0)
                                Dim listobject = vstoWorkbook1.Controls.AddListObject(lrange, "RnfDSummary" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString())
                                Dim filter As String = String.Format("WeekNum='{0}'", tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("WeekNumber").ToString())
                                Dim rows As Data.DataRow() = RnFDurationSummary.Select(filter)
                                If rows.Length > 0 Then
                                    Dim table As Data.DataTable = rows.CopyToDataTable()
                                    table.Columns.RemoveAt(0)
                                    If DSHRN > 20 Then

                                        For index1 = 21 To DSHRN
                                            table.Columns.Remove(index1.ToString() + "+")
                                        Next

                                    End If
                                    listobject.DataSource = table
                                    count = count + table.Rows.Count + 3
                                    'rowcount += table.Rows.Count + 2
                                    listobject.AutoSetDataBoundColumnHeaders = True
                                    listobject.ShowAutoFilter = False
                                End If

                            Next

                        End If
                        'Channel Summary
                    Catch ex As Exception
                        LogMpsrintExException("Exception occured while displaying duration summary.Message:" + ex.Message)

                    End Try
                    Try


                        RnFChannelSummary = nbdmainSummary(RnFChannelSummary, ChannelSummaryHRN)
                        '  RnFChannelSummary = ReverseCalculateTVRReach(RnFChannelSummary)
                        If Not CheckSheetExists("Channel Summary") Then
                            nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                            nativeSheet.Name = "Channel Summary"
                        Else
                            'nativeSheet = CheckAndReturnSheet("Channel Summary")
                            ''  newSheet.UsedRange.Clear()
                            'Globals.Ribbons.MSprintExRibbon.CleanSheet(nativeSheet)
                            'nativeSheet.Activate()

                            Dim sheetcount As Integer = CheckAndReturnSheet("Channel Summary")
                            If sheetcount > 0 Then
                                nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                                Dim sname As String = String.Format("Channel Summary({0})", sheetcount)
                                nativeSheet.Name = sname
                            End If

                        End If
                        Dim tgcell2 As Microsoft.Office.Interop.Excel.Range = nativeSheet.Range("$A$2", Type.Missing)
                        tgcell2.Value2 = "TG : " + plantgname
                        Dim vstoWorkbook2 As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(nativeSheet)

                        If tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                            '  Dim periodcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.Range("$A$3", Type.Missing)
                            Dim periodcell As Microsoft.Office.Interop.Excel.Range = tgcell2.Next(2, 0)
                            Dim period As String = String.Format("Period : {0} To {1}", tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToString("dd/MM/yyyy"), tpSelections.TaskPaneLogFile1.dtToDate.Value.ToString("dd/MM/yyyy"))
                            periodcell.Value = period
                            Dim mgs As List(Of String) = New List(Of String)()
                            For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                                mgs.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString())
                            Next
                            mgs.Add("TotalMarkets")

                            ' For Each mg As String In mgs
                            Dim count As Integer = 0
                            For index = 0 To mgs.Count - 1

                                '   Next

                                Dim filter As String = String.Format("Market='{0}'", mgs(index))
                                Dim rows As Data.DataRow() = RnFChannelSummary.Select(filter)
                                Dim mgcell As Microsoft.Office.Interop.Excel.Range = periodcell.Next(2 + (count * index), 0)
                                mgcell.Value2 = mgs(index)
                                Dim lrange As Microsoft.Office.Interop.Excel.Range = mgcell.Next(2, 0)
                                Dim listobject = vstoWorkbook2.Controls.AddListObject(lrange, "RnfCSummary" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString())
                                Dim table As Data.DataTable = rows.CopyToDataTable
                                table.Columns.RemoveAt(0)
                                table.Columns.RemoveAt(1)
                                If ChannelSummaryHRN > 20 Then

                                    For index1 = 21 To ChannelSummaryHRN
                                        table.Columns.Remove(index1.ToString() + "+")
                                    Next

                                End If
                                listobject.DataSource = table
                                count = table.Rows.Count + 3
                                'rowcount += table.Rows.Count + 2
                                listobject.AutoSetDataBoundColumnHeaders = True
                                listobject.ShowAutoFilter = False

                            Next
                        ElseIf tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                            Dim count As Integer = 4
                            Dim mgs As List(Of String) = New List(Of String)()
                            For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                                mgs.Add(Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index1).ToString())
                            Next
                            mgs.Add("TotalMarkets")

                            For index = 0 To tpSelections.TaskPaneLogFile1.dtWeeks.Rows.Count - 1
                                For index1 = 0 To mgs.Count - 1

                                    Dim periodcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count, 1), vstoWorkbook.Cells(count, 1))
                                    ' Dim period As String = String.Format("Period : '{0}' To '{1}'", tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToString("dd/MM/yyyy"), tpSelections.TaskPaneLogFile1.dtToDate.Value.ToString("dd/MM/yyyy"))
                                    Dim period As String = String.Format("Period : {0} To {1}", Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("StartDate").ToString()).ToString("dd/MM/yyyy"), Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("EndDate").ToString()).ToString("dd/MM/yyyy"))
                                    periodcell.Value = period
                                    '  Dim mgcell As String = String.Format("Period : {0} To {1}", Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("StartDate").ToString()).ToString("dd/MM/yyyy"), Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("EndDate").ToString()).ToString("dd/MM/yyyy"))
                                    Dim mgcell As Microsoft.Office.Interop.Excel.Range = periodcell.Next(2, 0)

                                    mgcell.Value = mgs(index1)

                                    Dim lrange As Microsoft.Office.Interop.Excel.Range = mgcell.Next(2, 0)
                                    Dim listobject = vstoWorkbook2.Controls.AddListObject(lrange, "RnfCSummary" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString())
                                    Dim filter As String = String.Format("WeekNum='{0}' and Market='{1}'", tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("WeekNumber").ToString(), mgs(index1))
                                    Dim rows As Data.DataRow() = RnFChannelSummary.Select(filter)
                                    If rows.Length > 0 Then
                                        Dim table As Data.DataTable = rows.CopyToDataTable()
                                        table.Columns.RemoveAt(0)
                                        table.Columns.RemoveAt(1)
                                        If ChannelSummaryHRN > 20 Then

                                            For index2 = 21 To ChannelSummaryHRN
                                                table.Columns.Remove(index2.ToString() + "+")
                                            Next

                                        End If
                                        listobject.DataSource = table
                                        count = count + table.Rows.Count + 6
                                        'rowcount += table.Rows.Count + 2
                                        listobject.AutoSetDataBoundColumnHeaders = True
                                        listobject.ShowAutoFilter = False
                                    End If

                                Next
                            Next

                        End If
                    Catch ex As Exception
                        LogMpsrintExException("Exception occured while displaying channel summary.Message:" + ex.Message)

                    End Try
                    'Creative Summary
                    Try


                        RnFCreativeSummary = nbdmainSummary(RnFCreativeSummary, CreativeSummaryHRN)
                        ' RnFCreativeSummary = ReverseCalculateTVRReach(RnFCreativeSummary)
                        '  End If
                        '   Next


                        '  If ConstructMarketSummaryTable(rnfoutputXml) Then
                        If Not CheckSheetExists("Creative Summary") Then
                            nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                            nativeSheet.Name = "Creative Summary"
                        Else
                            'nativeSheet = CheckAndReturnSheet("Creative Summary")
                            ''  newSheet.UsedRange.Clear()
                            'Globals.Ribbons.MSprintExRibbon.CleanSheet(nativeSheet)
                            'nativeSheet.Activate()

                            Dim sheetcount As Integer = CheckAndReturnSheet("Creative Summary")
                            If sheetcount > 0 Then
                                nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                                Dim sname As String = String.Format("Creative Summary({0})", sheetcount)
                                nativeSheet.Name = sname
                            End If
                        End If
                        Dim tgcell3 As Microsoft.Office.Interop.Excel.Range = nativeSheet.Range("$A$2", Type.Missing)
                        tgcell3.Value2 = "TG: " + plantgname
                        Dim vstoWorkbook3 As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(nativeSheet)

                        If tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                            '  Dim periodcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.Range("$A$3", Type.Missing)
                            Dim periodcell As Microsoft.Office.Interop.Excel.Range = tgcell3.Next(2, 0)
                            Dim period As String = String.Format("Period : {0} To {1}", tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToString("dd/MM/yyyy"), tpSelections.TaskPaneLogFile1.dtToDate.Value.ToString("dd/MM/yyyy"))
                            periodcell.Value = period
                            Dim lrange As Microsoft.Office.Interop.Excel.Range = periodcell.Next(2, 0)
                            Dim listobject = vstoWorkbook3.Controls.AddListObject(lrange, "RnfCreativeSummary" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString())
                            Dim table As Data.DataTable = RnFCreativeSummary.Copy()
                            table.Columns.RemoveAt(0)
                            If CreativeSummaryHRN > 20 Then

                                For index1 = 21 To CreativeSummaryHRN
                                    table.Columns.Remove(index1.ToString() + "+")
                                Next

                            End If
                            listobject.DataSource = table
                            'rowcount += table.Rows.Count + 2
                            listobject.AutoSetDataBoundColumnHeaders = True
                            listobject.ShowAutoFilter = False
                        ElseIf tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                            Dim count As Integer = 4
                            For index = 0 To tpSelections.TaskPaneLogFile1.dtWeeks.Rows.Count - 1
                                Dim periodcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count, 1), vstoWorkbook.Cells(count, 1))
                                '  Dim period As String = String.Format("Period : '{0}' To '{1}'", tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToString("dd/MM/yyyy"), tpSelections.TaskPaneLogFile1.dtToDate.Value.ToString("dd/MM/yyyy"))
                                Dim period As String = String.Format("Period : {0} To {1}", Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("StartDate").ToString()).ToString("dd/MM/yyyy"), Convert.ToDateTime(tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("EndDate").ToString()).ToString("dd/MM/yyyy"))
                                periodcell.Value = period
                                Dim lrange As Microsoft.Office.Interop.Excel.Range = periodcell.Next(2, 0)
                                Dim listobject = vstoWorkbook3.Controls.AddListObject(lrange, "RnfCreativeSummary" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString())
                                Dim filter As String = String.Format("WeekNum='{0}'", tpSelections.TaskPaneLogFile1.dtWeeks.Rows(index)("WeekNumber").ToString())
                                Dim rows As Data.DataRow() = RnFCreativeSummary.Select(filter)
                                If rows.Length > 0 Then
                                    Dim table As Data.DataTable = rows.CopyToDataTable()
                                    table.Columns.RemoveAt(0)
                                    If CreativeSummaryHRN > 20 Then

                                        For index1 = 21 To CreativeSummaryHRN
                                            table.Columns.Remove(index1.ToString() + "+")
                                        Next

                                    End If
                                    listobject.DataSource = table
                                    count = count + table.Rows.Count + 3
                                    'rowcount += table.Rows.Count + 2
                                    listobject.AutoSetDataBoundColumnHeaders = True
                                    listobject.ShowAutoFilter = False
                                End If


                            Next

                        End If
                    Catch ex As Exception
                        LogMpsrintExException("Exception occured while displaying creative summary.Message:" + ex.Message)

                    End Try
                    'End If
                End If

            End If

            'End If

            SetNormalCursor()
        Catch ex As Exception
            SetNormalCursor()
            LogMpsrintExException("Exception occured while displaying All summary details.Message :" + ex.Message)
            MessageBox.Show("Exception occured while displaying ALL summary details.Please refer to error log for more details")
        End Try

    End Sub

    Private Sub btnSpotSelect_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnSpotSelect.Click
        Try
            Dim range As Microsoft.Office.Interop.Excel.Range = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection
            Dim avaiRow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Rows(range.Row - 2)
            Globals.Ribbons.MSprintExRibbon.currentLineItem = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Rows(range.Row - 2)("GUID").ToString()

            Dim cbWeeks As String = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Rows(range.Row - 2)("WeekNum").ToString()

            Dim selecpots As Data.DataTable = New Data.DataTable()
            Dim fil As String = fil = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
            If cbWeeks.Equals(String.Empty) Then
                fil = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
            Else
                fil = String.Format("GUID='{0}' and WeekNum={1}", Globals.Ribbons.MSprintExRibbon.currentLineItem, Int32.Parse(cbWeeks).ToString())
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

            Dim dt As Data.DataTable = CType(loSpotSelection.DataSource, Data.DataTable)
            Dim count As Integer
            Try
                If dt.Columns.Contains("Total Spots") Then
                    Dim filter As String = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
                    count = Convert.ToInt32(dt.Select(filter)(0)("Total Spots").ToString())
                Else
                    Dim filter As String = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
                    count = Convert.ToInt32(dt.Select(filter)(0)("Week " & Int32.Parse(cbWeeks)).ToString())

                End If
            Catch ex As Exception
                Dim filter As String = String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem)
                If Globals.Ribbons.MSprintExRibbon.xecelTable.Columns.Contains("Total Spots") Then
                    count = Convert.ToInt32(Globals.Ribbons.MSprintExRibbon.xecelTable.Select(filter)(0)("Total Spots").ToString())
                Else
                    count = Convert.ToInt32(Globals.Ribbons.MSprintExRibbon.xecelTable.Select(filter)(0)("Week " & Int32.Parse(cbWeeks)).ToString())

                End If
            End Try

            If count > selecpots.Rows.Count And (count - selecpots.Rows.Count) >= 1 Then
                '  For Each row As DataGridViewRow In dgvAvailableSpotsGrid.SelectedRows
                Dim row As Data.DataRow = avaiRow
                Dim dr As DataRow = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.NewRow()
                '  Dim dr1 As DataRow = selecpots.NewRow()
                dr("GUID") = row("GUID").ToString()
                ' dr1("GUID") = row("GUID").ToString
                dr("Spot") = row("Spot").ToString
                ' dr1("Spot") = row.Cells.Item("Spot").Value
                dr("Start Date") = row("Start Date").ToString
                ' dr1("Start Date") = row.Cells.Item("Start Date").Value
                dr("End Date") = row("End Date").ToString
                ' dr1("End Date") = row.Cells.Item("End Date").Value
                dr("WeekNum") = row("WeekNum").ToString
                ' dr1("WeekNum") = row.Cells.Item("WeekNum").Value
                dr("Channel") = row("Channel").ToString
                '  dr1("Channel") = row.Cells.Item("Channel").Value
                '  Dim spotrow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.GetSpotRow(spots(index).Spot)
                dr("Date") = row("Date").ToString
                ' dr1("Date") = row.Cells.Item("Date").Value
                dr("Start Time") = row("Start Time").ToString
                '  dr1("Start Time") = row.Cells.Item("Start Time").Value
                'dRow("StartTime") = spotrow("StartTime")
                'dRow("EndTime") = spotrow("EndTime")
                'dr("End Time") = row.Cells.Item("End Time").Value
                'dr1("End Time") = row.Cells.Item("End Time").Value
                dr("Duration(Sec)") = row("Duration(Sec)").ToString
                ' dr1("Duration(Sec)") = row.Cells.Item("Duration(Sec)").Value
                'dRow("Duration(Sec)") = TimeSpan.Parse(dRow("EndTime").ToString()).Subtract(TimeSpan.Parse(dRow("StartTime").ToString())).TotalSeconds
                'dRow("PA") = spotrow("PA")
                dr("PA") = row("PA").ToString()
                'dr1("PA") = row.Cells.Item("PA").Value
                '  dr(spots(index).MG + "TVR") = spots(index).TVRVal.Split({","c}, StringSplitOptions.None)(0)
                ' dRow("TA") = spotrow("TA")
                dr("TA") = row("TA").ToString()
                ' dr1("TA") = row.Cells.Item("TA").Value
                '  dr("Commercial") = row.Cells("Commercial").Value
                dr("Cost") = row("Cost").ToString
                '   dr1("Cost") = row.Cells.Item("Cost").Value
                For Each market As String In Globals.Ribbons.MSprintExRibbon.markets
                    dr(market) = row(market).ToString
                    '   dr1(market) = row.Cells.Item(market).Value
                Next
                Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Add(dr)
                Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.AcceptChanges()
                Dim rnfselectedCopy As Data.DataTable = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Copy()
                rnfselectedCopy.Columns.Remove("GUID")
                rnfselectedCopy.Columns.Remove("Spot")
                rnfselectedCopy.Columns.Remove("Start Date")
                rnfselectedCopy.Columns.Remove("End Date")
                rnfselectedCopy.Columns.Remove("WeekNum")
                Globals.Ribbons.MSprintExRibbon.selectedSpotsListObject.SetDataBinding(rnfselectedCopy)
                System.Windows.Forms.MessageBox.Show("Selected spot replaced successfully.")
                Globals.Ribbons.MSprintExRibbon.selectedSpotsListObject.ListRows(Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count).Range.Cells.Interior.Color = RGB(112, 238, 233)
                Globals.Ribbons.MSprintExRibbon.selectedSpotsListObject.ListColumns(5).Range.NumberFormat = "dd/MM/yyyy"
                Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = ReturnActualSheet("Selected_Spots")
                sheet.Activate()
                'selecpots.Rows.Add(dr1)
                '   dgvAvailableSpotsGrid.Rows.Remove(row)
                '  Next
            ElseIf ((count - selecpots.Rows.Count) > 0 And ((count - selecpots.Rows.Count) < 1)) Then
                System.Windows.Forms.MessageBox.Show(String.Format("Maximum of {0} spot(s) can be selected.Please reselect {0} spot(s)", count - selecpots.Rows.Count))
                ' ShowErrorLabel(String.Format("Maximum of {0} spot(s) can be selected.Please reselect {0} spot(s)", count - selecpots.Rows.Count))
            Else
                System.Windows.Forms.MessageBox.Show("Maximum number of spot(s) has been selected.Please remove unwanted and/or increase number of required spot(s) count and Try again")
                ' ShowErrorLabel("Maximum number of spot(s) has been selected.Please remove unwanted and/or increase number of required spot(s) count and Try again")
            End If
        Catch ex As Exception
            LogMpsrintExException("Exception occured while selecting spot.Message :" + ex.Message)
        End Try
    End Sub

    Private Sub btnSpotReplace_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnSpotReplace.Click
        Try
            '  Dim active_sheet As Microsoft.Office.Interop.Excel.Worksheet = CType(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            Dim range As Microsoft.Office.Interop.Excel.Range = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection
            Globals.Ribbons.MSprintExRibbon.currentLineItem = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows(range.Row - 2)("GUID").ToString()
            '  Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.RemoveAt(range.Row - 2)
            Globals.Ribbons.MSprintExRibbon.Selectedrow = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows(range.Row - 2)

            ' Globals.Ribbons.MSprintExRibbon.selectedRowIndex = range.Row - 2

            'Globals.Ribbons.MSprintExRibbon.selectedSpotsListObject.SetDataBinding(Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots)
            '  Dim drows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.xecelTable.Select(String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem))
            Globals.ThisAddIn.GetAvailableSpots()
        Catch ex As Exception
            LogMpsrintExException("Exception occured while replacing the spot.Message :" + ex.Message)
        End Try
    End Sub
    Private Sub btnDeleteSpot_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnDeleteSpot.Click
        Try
            Dim range As Microsoft.Office.Interop.Excel.Range = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection
            Globals.Ribbons.MSprintExRibbon.currentLineItem = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows(range.Row - 2)("GUID").ToString()
            Dim weekNum As Integer = 0
            Dim colname As String = String.Empty
            If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                weekNum = Convert.ToInt32(Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows(range.Row - 2)("WeekNum").ToString())
                colname = "Week " & weekNum
            End If

            '  Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.RemoveAt(range.Row - 2)
            Globals.Ribbons.MSprintExRibbon.Selectedrow = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows(range.Row - 2)

            Globals.Ribbons.MSprintExRibbon.xecellineItemsTable = CType(loSpotSelection.DataSource, Data.DataTable)
            Dim row As Data.DataRow = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.Select(String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem))(0)
            ' row(colname) = Convert.ToInt32(row(colname).ToString()) - 1
            If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                'weekNum = Convert.ToInt32(Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows(range.Row - 2)("WeekNum").ToString())
                'colname = "Week " & weekNum
                row(colname) = Convert.ToInt32(row(colname).ToString()) - 1
            Else
                row("Total Spots") = Convert.ToInt32(row("Total Spots").ToString()) - 1
            End If
            Globals.Ribbons.MSprintExRibbon.xecellineItemsTable.AcceptChanges()
            loSpotSelection.DataSource = Globals.Ribbons.MSprintExRibbon.xecellineItemsTable
            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Remove(Globals.Ribbons.MSprintExRibbon.Selectedrow)
            Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.AcceptChanges()
            Globals.Ribbons.MSprintExRibbon.selectedSpotsListObject.SetDataBinding(Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots)
            MessageBox.Show("Selected spot deleted successfully.")
        Catch ex As Exception
            LogMpsrintExException("Exception occured while deleting spot.Message :" + ex.Message)
        End Try
    End Sub

    Private Sub btnGenerateEndTime_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnGenerateEndTime.Click
        Try
            mpGenEndTime = New ucGenEndTime()
            CTPChannelShare = Globals.ThisAddIn.CustomTaskPanes.Add(mpGenEndTime, "Generate End Time")
            CTPChannelShare.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
            CTPChannelShare.Height = 150
            CTPChannelShare.Width = 315
            CTPChannelShare.Visible = True
        Catch ex As Exception
            LogMpsrintExException("Exception occured while generating End time.Message:" + ex.Message)
        End Try
    End Sub
End Class
Public NotInheritable Class MyExtensionClasses
    Private Sub New()
    End Sub
    '  <System.Runtime.CompilerServices.Extension()> _
    Public Shared Function OuterXml(ByVal thiz As XElement) As String
        Dim xReader = thiz.CreateReader()
        xReader.MoveToContent()
        Return xReader.ReadOuterXml()
    End Function
    ' <System.Runtime.CompilerServices.Extension()> _
    Public Shared Function InnerXml(ByVal thiz As XElement) As String
        Dim xReader = thiz.CreateReader()
        xReader.MoveToContent()
        Return xReader.ReadInnerXml()
    End Function
End Class
