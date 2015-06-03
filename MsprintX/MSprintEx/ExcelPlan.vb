Imports System.Globalization
Imports System.Windows.Forms
Imports System.IO
Imports System
Imports System.Data
Imports System.Linq
Imports System.Management
Imports System.Security.Cryptography
Imports System.Text
Module ExcelPlan
    Dim WithEvents wbLogCreator As Excel.Workbook
    Friend WithEvents logCreator As Excel.Worksheet
    Friend logTaskPane As TaskPaneLogFile
    Friend mpTaskPane As ucPlanSelections
    '  Friend mpTpSpotSelection As ucSpotSelection
    Friend mpTopPrograms As ucTopPrograms

    Friend WithEvents MSprintExTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Friend WithEvents MSprintExTAMTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Friend lucMarkets As ucMarkets
    Const wbName As String = "LogCreator"
    Friend WithEvents PlanSheet As Excel.Worksheet
    Friend loPlanData As Excel.ListObject
    Friend WithEvents loSpotSelection, loProgAvgTVr, loCrDura, loWeekWise As Microsoft.Office.Tools.Excel.ListObject
    Dim ChannelColumn As Excel.ListColumn
    Dim DayCoulmn As Excel.ListColumn
    Dim StartTimeColumn As Excel.ListColumn
    Dim EndTimeColumn As Excel.ListColumn
    Dim ChannelCells As Excel.Range
    Friend SubtotalRows As Excel.Range
    Dim BlankFields, WrongTimeCells, WrongDayCells As Excel.Range
    Friend isPlanClean As Boolean = False
    Dim PlanChecking As Boolean = False
    Dim DayChecking As Boolean = False
    Dim TimeChecking As Boolean = False
    Dim dtErrors As New Plandata.ErrorRangeDataTable
    Friend dtChannelMaster As Plandata.ChannelMasterDataTable
    Friend dtPlanChannels As New Plandata.PlanChannelsDataTable

    Dim myc As New WeekYearEqualityComparer
    Friend WeekColumns As New Dictionary(Of WeekYear, Data.DataColumn)(myc)
    Friend isLoaded As Boolean = False
    Friend TotalSpotsWritten As Integer = 0
    Friend TotalPlanSpots As Integer = 0
    Friend daChannelMap As New BrkPerfDataSetTableAdapters.ChannelMapTableAdapter
    Friend dtChannelMap As BrkPerfDataSet.ChannelMapDataTable



    Class WeekYear
        Implements IEquatable(Of WeekYear)

        Private _WeekNumber, _WeekYear As Integer
        Public Sub New(ByVal wn As Integer, ByVal wy As Integer)
            Me.WeekNumber = wn
            Me.WeekYear = wy
        End Sub
        Property WeekNumber() As Integer
            Get
                Return _WeekNumber
            End Get
            Set(ByVal value As Integer)
                _WeekNumber = value
            End Set
        End Property
        Property WeekYear() As Integer
            Get
                Return _WeekYear
            End Get
            Set(ByVal value As Integer)
                _WeekYear = value
            End Set
        End Property

        Public Overloads Function Equals(ByVal other As WeekYear) As Boolean Implements System.IEquatable(Of WeekYear).Equals
            If Me.WeekNumber = other.WeekNumber And Me.WeekYear = other.WeekYear Then
                Return True
            Else
                Return False
            End If
        End Function
    End Class
    Class WeekYearEqualityComparer

        Inherits EqualityComparer(Of WeekYear)

        'Public Overloads Overrides Function Equals(ByVal w1 As WeekYear, ByVal w2 As WeekYear) _
        '           As Boolean Implements IEqualityComparer(Of WeekYear).Equals
        'End Function
        'Public Overloads Overrides Function GetHashCode(ByVal wy As WeekYear) _
        '        As Integer Implements IEqualityComparer(Of WeekYear).GetHashCode
        'End Function

        Public Overloads Overrides Function Equals(ByVal x As WeekYear, ByVal y As WeekYear) As Boolean
            If x.WeekNumber = y.WeekNumber And x.WeekYear = _
                    y.WeekYear Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Overloads Overrides Function GetHashCode(ByVal obj As WeekYear) As Integer
            Dim hCode As Integer = obj.WeekNumber Xor obj.WeekYear
            Return hCode.GetHashCode()
        End Function
    End Class
    'Friend Sub listObjectBefAddrow(ByVal sender As Object, ByVal e As Microsoft.Office.Tools.Excel.BeforeAddDataBoundRowEventArgs) Handles loSpotSelection.BeforeAddDataBoundRow
    '    ' MessageBox.Show(e.Item.ToString())

    'End Sub
    Public Function GetWeekTable() As Data.DataTable
        Dim weekTable As Data.DataTable = New Data.DataTable()
        Try
            For Each row As Data.DataRow In Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks.Rows
                weekTable.Columns.Add("Week " & row("WeekNumber"), Type.GetType("System.Int32"))
            Next
            weekTable.AcceptChanges()
        Catch ex As Exception

        End Try
        Return weekTable
    End Function
    Public Function GetCreativeDuration() As Data.DataTable
        Dim inpProgTVRTable As Data.DataTable = New Data.DataTable()
        inpProgTVRTable.Columns.Add("Creative")
        'inpSpotTable.Columns(0).AutoIncrement = True
        'inpSpotTable.Columns(0).AutoIncrementSeed = 1
        inpProgTVRTable.Columns.Add("Duration")
        inpProgTVRTable.Columns.Add("Spots%")
        'inpProgTVRTable.Columns.Add("Day")
        'inpProgTVRTable.Columns.Add("Start Time")
        'inpProgTVRTable.Columns.Add("End Time")
        'inpProgTVRTable.Columns.Add("RatePer10Sec", Type.GetType("System.Int32"))
        'inpProgTVRTable.Columns.Add("Total Spots", Type.GetType("System.Int32"))
        Return inpProgTVRTable
    End Function
    Public Function GetInpProgAvgTVRTable() As Data.DataTable
        Dim inpProgTVRTable As Data.DataTable = New Data.DataTable()
        inpProgTVRTable.Columns.Add("GUID")
        'inpSpotTable.Columns(0).AutoIncrement = True
        'inpSpotTable.Columns(0).AutoIncrementSeed = 1
        inpProgTVRTable.Columns.Add("Channel")
        inpProgTVRTable.Columns.Add("Programme")
        inpProgTVRTable.Columns.Add("Day")
        inpProgTVRTable.Columns.Add("Start Time")
        inpProgTVRTable.Columns.Add("End Time")
        inpProgTVRTable.Columns.Add("RatePer10Sec", Type.GetType("System.Int32"))
        ' inpProgTVRTable.Columns.Add("Total Spots", Type.GetType("System.Int32"))
        Return inpProgTVRTable
    End Function
    Public Function GetInpSpotDataTable() As Data.DataTable
        Dim inpSpotTable As Data.DataTable = New Data.DataTable()
        inpSpotTable.Columns.Add("GUID")
        'inpSpotTable.Columns(0).AutoIncrement = True
        'inpSpotTable.Columns(0).AutoIncrementSeed = 1
        inpSpotTable.Columns.Add("Channel")
        inpSpotTable.Columns.Add("Programme")
        inpSpotTable.Columns.Add("Day")
        inpSpotTable.Columns.Add("Start Time")
        inpSpotTable.Columns.Add("End Time")
        inpSpotTable.Columns.Add("RatePer10Sec", Type.GetType("System.Int32"))
        inpSpotTable.Columns.Add("Creative")
        inpSpotTable.Columns.Add("Duration", Type.GetType("System.Int32"))
        inpSpotTable.Columns.Add("Total Cost", Type.GetType("System.Int32"))
        inpSpotTable.Columns.Add("Min TVR", Type.GetType("System.Decimal"))
        inpSpotTable.Columns.Add("Max TVR", Type.GetType("System.Decimal"))
        If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
            inpSpotTable.Columns.Add("Total Spots", Type.GetType("System.Int32"))
        Else
            For Each row As Data.DataRow In Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtWeeks.Rows
                inpSpotTable.Columns.Add("Week " & row("WeekNumber"), Type.GetType("System.Int32"))
            Next
            'Dim col As Data.DataColumn
            ''  If Not (loSpotSelection Is Nothing) Then
            ''AddWeekColumn = loSpotSelection.ListColumns.Add()
            ''AddWeekColumn.Name = "Week " & WeekNumber
            'col = listObjectdt.Columns.Add()
            'col.ColumnName = "Week " & WeekNumber
        End If
        Return inpSpotTable
        'Channel	Programme	Day	Start Time	End Time	RatePer10Sec	Creative	Duration	Total Spots

    End Function

    Private Sub wbLogCreator_BeforeClose(ByRef Cancel As Boolean) Handles wbLogCreator.BeforeClose
        wbLogCreator.Saved = True
        ManageButtons(True)
        If Not MSprintExTaskPane Is Nothing Then
            MSprintExTaskPane.Dispose()
            logTaskPane.Dispose()
        End If
        ResetVariables()
        'BreakPerformance.daProgRFMT.Delete()
        wbLogCreator = Nothing
    End Sub
    'Friend Sub logCreatorSelectionChange(ByVal Target As Excel.Range) Handles logCreator.SelectionChange
    '    Try
    '        If Not (mpTpSpotSelection Is Nothing) And Target.Row > 1 Then
    '            Globals.Ribbons.MSprintExRibbon.currentLineItem = Globals.Ribbons.MSprintExRibbon.xecelTable.Rows(Target.Row - 2)("GUID").ToString()
    '            Dim drows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Select(String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem))

    '            If drows.Count > 0 Then
    '                mpTpSpotSelection.dgSelectedSpotsGrid.DataSource = drows.CopyToDataTable()
    '            Else
    '                mpTpSpotSelection.dgSelectedSpotsGrid.DataSource = Nothing
    '            End If
    '            mpTpSpotSelection.dgvAvailableSpotsGrid.DataSource = Nothing
    '            'If Not (Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots Is Nothing) Then
    '            '    Dim rows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Select(String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem))

    '            '    If rows.Count > 0 Then
    '            '        mpTpSpotSelection.dgvAvailableSpotsGrid.DataSource = rows.CopyToDataTable()
    '            '    End If
    '            'End If


    '            Globals.Ribbons.MSprintExRibbon.DisplayCurrentPlanItem()
    '        End If
    '    Catch ex As Exception
    '        LogMpsrintExException("Exception occured while displaying spots of selected line item")
    '        ' Throw ex
    '    End Try
    'End Sub

    Friend Sub OpenMSprintEx()
        Try
            '  wbLogCreator = Globals.ThisAddIn.Application.Workbooks.Add(AppDomain.CurrentDomain.BaseDirectory & "\LogCreator.xltx")
            ' logCreator = Globals.ThisAddIn.Application.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet, Type.Missing)

            If Not CheckSheetExists("Plan Selection") Then
                logCreator = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                logCreator.Name = "Plan Selection"
            Else
                'logCreator = CheckAndReturnSheet("Plan Selection")
                ''  newSheet.UsedRange.Clear()
                'Globals.Ribbons.MSprintExRibbon.CleanSheet(logCreator)
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
            loSpotSelection = vstoWorkbook.Controls.AddListObject(vstoWorkbook.Range("$A$1"), "InputSpotSelection")
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
            Exit Sub
        Finally
            loSpotSelection.Application.CutCopyMode = True
            loSpotSelection.Application.ScreenUpdating = True
        End Try
        ResetVariables()
        ManageButtons(False)

        ' MSprintExTaskPane.Width = 550
        'MSprintExTaskPane.Height = 348
        'myCustomTaskPane.Width = 300
        CreateLogFilePane()
        '   ActivateTaskPane(wbLogCreator)
        PlanSheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        '    loSpotSelection = PlanSheet.ListObjects("InputSpotSelection") '.GetVstoObject
        dtChannelMaster = New Plandata.ChannelMasterDataTable
        WeekColumns = New Dictionary(Of WeekYear, Data.DataColumn)(myc)
        'logTaskPane = mpTaskPane.lucPeriod
        'logTaskPane.dtFromDate.Value = Now.Date()
    End Sub
    Friend Sub CreateLogFilePane()
        ' logTaskPane = New TaskPaneLogFile
        '  mpTaskPane.Panel6.Controls.Add(logTaskPane)
        'logTaskPane.Dock = Windows.Forms.DockStyle.Fill

        'lucMarkets = New ucMarkets
        'mpTaskPane.Panel2.Controls.Add(lucMarkets)
        'lucMarkets.Dock = DockStyle.Fill

        'mpTaskPane.Panel3.AutoScroll = True
        'myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(logTaskPane, "Create TAM logfile")
        'myCustomTaskPane.Width = 300
        'logTaskPane.dtFromDate.Value = Now.Date
    End Sub
    Friend Sub ActivateTaskPane(ByVal Wb As Excel.Workbook)
        If InStr(Wb.Name, wbName) > 0 Then
            MSprintExTaskPane.Visible = True
        Else
            MSprintExTaskPane.Visible = False
        End If
    End Sub
    Friend Sub ManageButtons(ByVal wbClosed As Boolean)
        Globals.Ribbons.MSprintExRibbon.btnLogFile.Enabled = wbClosed
        Globals.Ribbons.MSprintExRibbon.btnMapChannels.Enabled = Not wbClosed
        'Globals.Ribbons.MSprintExRibbon.btnCleanUp.Enabled = Not wbClosed
        'Globals.Ribbons.MSprintExRibbon.btnExistingLog.Enabled = Not wbClosed
    End Sub
    Private Sub ListObjectChangeMethod(ByVal target As Microsoft.Office.Interop.Excel.Range, ByVal changedRanges As Microsoft.Office.Tools.Excel.ListRanges) Handles loSpotSelection.Change
        Try
            'For Each range As Microsoft.Office.Interop.Excel.Range In target.Cells
            '    MessageBox.Show(range.Value2.ToString())
            'Next
            Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(sheet)
            ' vstoWorkbook.Name = "Plan Selection"
            loSpotSelection = DirectCast(vstoWorkbook.Controls.Item("InputSpotSelection"), Microsoft.Office.Tools.Excel.ListObject)
            'ChannelColumn = loSpotSelection.ListColumns("Channel")
            'DayCoulmn = loSpotSelection.ListColumns("Day")
            'StartTimeColumn = loSpotSelection.ListColumns("Start Time")
            'EndTimeColumn = loSpotSelection.ListColumns("End Time")
            'ChannelCells = ChannelColumn.DataBodyRange
            '  logTaskPane = New TaskPaneLogFile
            If Not PlanChecking Then
                If Not loSpotSelection.Application.Intersect(target, DayCoulmn.DataBodyRange) Is Nothing Then
                    If Not DayChecking Then
                        For Each DayCell As Excel.Range In loSpotSelection.Application.Intersect(target, DayCoulmn.DataBodyRange).Cells
                            If SubtotalRows Is Nothing Then
                                If CheckDayValue(DayCell) Then
                                    If Not WrongDayCells Is Nothing Then
                                        If Not loSpotSelection.Application.Intersect(WrongDayCells, DayCell) Is Nothing Then
                                            WrongDayCells = RemoveCellFromRange(WrongDayCells, DayCell)
                                        End If
                                    End If
                                Else
                                    If Not WrongDayCells Is Nothing Then
                                        If loSpotSelection.Application.Intersect(WrongDayCells, DayCell) Is Nothing Then
                                            WrongDayCells = loSpotSelection.Application.Union(WrongDayCells, DayCell)
                                        End If
                                    Else
                                        WrongDayCells = DayCell
                                    End If
                                End If
                            Else
                                If loSpotSelection.Application.Intersect(DayCell, SubtotalRows) Is Nothing Then
                                    If CheckDayValue(DayCell) Then
                                        If Not WrongDayCells Is Nothing Then
                                            If Not loSpotSelection.Application.Intersect(WrongDayCells, DayCell) Is Nothing Then
                                                WrongDayCells = RemoveCellFromRange(WrongDayCells, DayCell)
                                            End If
                                        End If
                                    Else
                                        If Not WrongDayCells Is Nothing Then
                                            If loSpotSelection.Application.Intersect(WrongDayCells, DayCell) Is Nothing Then
                                                WrongDayCells = loSpotSelection.Application.Union(WrongDayCells, DayCell)
                                            End If
                                        Else
                                            WrongDayCells = DayCell
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        DayChecking = True
                        CheckEmptyDayTime()
                        showErrorRange()
                        DayChecking = False
                    End If
                End If
                If Not loSpotSelection.Application.Intersect(target, loSpotSelection.Application.Union(StartTimeColumn.DataBodyRange, EndTimeColumn.DataBodyRange)) Is Nothing _
                    Then
                    If Not TimeChecking Then
                        For Each TimeCell As Excel.Range In loSpotSelection.Application.Intersect(target, loSpotSelection.Application.Union(StartTimeColumn.DataBodyRange, EndTimeColumn.DataBodyRange)).Cells
                            If SubtotalRows Is Nothing Then
                                Dim CurrTime As Date
                                If DateTime.TryParse(TimeCell.Text, CurrTime) Then

                                    TimeChecking = True
                                    TimeCell.Value = CurrTime.ToOADate - Math.Floor(CurrTime.ToOADate) + 2
                                    If Not loSpotSelection.Application.Intersect(StartTimeColumn.DataBodyRange, TimeCell) Is Nothing Then
                                        If TypeOf TimeCell.Offset(0, 1).Value Is Double Then
                                            If TimeCell.Value > TimeCell.Offset(0, 1).Value Then
                                                TimeCell.Offset(0, 1).Value = TimeCell.Offset(0, 1).Value + 1
                                            End If
                                        End If
                                    End If
                                    If Not loSpotSelection.Application.Intersect(EndTimeColumn.DataBodyRange, TimeCell) Is Nothing Then
                                        If TypeOf TimeCell.Offset(0, -1).Value Is Double Then
                                            If TimeCell.Value < TimeCell.Offset(0, -1).Value Then
                                                TimeCell.Value = TimeCell.Value + 1
                                            End If
                                        End If
                                    End If
                                    TimeChecking = False

                                    If Not WrongTimeCells Is Nothing Then
                                        If Not loSpotSelection.Application.Intersect(WrongTimeCells, TimeCell) Is Nothing Then
                                            WrongTimeCells = RemoveCellFromRange(WrongTimeCells, TimeCell)
                                        End If
                                    End If
                                    TimeChecking = True
                                    With TimeCell
                                        .ClearFormats()
                                        .NumberFormat = "hh:mm;@"
                                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                                    End With
                                    TimeChecking = False
                                Else
                                    TimeChecking = True
                                    Dim aTime As String
                                    aTime = TimeCell.Text.ToString.Trim()
                                    If aTime.Length = 5 Then
                                        Try
                                            Dim oldDate As New Date(1900, 1, 1)
                                            Dim dblHour, dblMins As Double
                                            Dim tmStartTime As Date
                                            dblHour = aTime.Substring(0, 2)
                                            dblMins = aTime.Substring(3, 2)

                                            tmStartTime = DateAdd(DateInterval.Hour, dblHour, oldDate)
                                            tmStartTime = DateAdd(DateInterval.Minute, dblMins, tmStartTime)
                                            TimeCell.Value = tmStartTime
                                            Continue For
                                        Catch ex As Exception

                                        End Try
                                    End If
                                    TimeChecking = False
                                    If Not WrongTimeCells Is Nothing Then
                                        If loSpotSelection.Application.Intersect(WrongTimeCells, TimeCell) Is Nothing Then
                                            WrongTimeCells = loSpotSelection.Application.Union(WrongTimeCells, TimeCell)
                                        End If
                                    Else
                                        WrongTimeCells = TimeCell
                                    End If

                                End If
                            Else
                                If loSpotSelection.Application.Intersect(TimeCell, SubtotalRows) Is Nothing Then
                                    Dim CurrTime As Date
                                    If DateTime.TryParse(TimeCell.Text, CurrTime) Then

                                        TimeChecking = True
                                        TimeCell.Value = CurrTime.ToOADate - Math.Floor(CurrTime.ToOADate) + 2
                                        If Not loSpotSelection.Application.Intersect(StartTimeColumn.DataBodyRange, TimeCell) Is Nothing Then
                                            If TypeOf TimeCell.Offset(0, 1).Value Is Double Then
                                                If TimeCell.Value > TimeCell.Offset(0, 1).Value Then
                                                    TimeCell.Offset(0, 1).Value = TimeCell.Offset(0, 1).Value + 1
                                                End If
                                            End If
                                        End If
                                        If Not loSpotSelection.Application.Intersect(EndTimeColumn.DataBodyRange, TimeCell) Is Nothing Then
                                            If TypeOf TimeCell.Offset(0, -1).Value Is Double Then
                                                If TimeCell.Value < TimeCell.Offset(0, -1).Value Then
                                                    TimeCell.Value = TimeCell.Value + 1
                                                End If
                                            End If
                                        End If
                                        TimeChecking = False

                                        If Not WrongTimeCells Is Nothing Then
                                            If Not loSpotSelection.Application.Intersect(WrongTimeCells, TimeCell) Is Nothing Then
                                                WrongTimeCells = RemoveCellFromRange(WrongTimeCells, TimeCell)
                                            End If
                                        End If
                                        TimeChecking = True
                                        With TimeCell
                                            .ClearFormats()
                                            .NumberFormat = "hh:mm;@"
                                            .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                                        End With
                                        TimeChecking = False
                                    Else
                                        TimeChecking = True
                                        Dim aTime As String
                                        aTime = TimeCell.Text.ToString.Trim()
                                        If aTime.Length = 5 Then
                                            Try
                                                Dim oldDate As New Date(1900, 1, 1)
                                                Dim dblHour, dblMins As Double
                                                Dim tmStartTime As Date
                                                dblHour = aTime.Substring(0, 2)
                                                dblMins = aTime.Substring(3, 2)

                                                tmStartTime = DateAdd(DateInterval.Hour, dblHour, oldDate)
                                                tmStartTime = DateAdd(DateInterval.Minute, dblMins, tmStartTime)
                                                TimeCell.Value = tmStartTime
                                                Continue For
                                            Catch ex As Exception

                                            End Try
                                        End If
                                        TimeChecking = False
                                        If Not WrongTimeCells Is Nothing Then
                                            If loSpotSelection.Application.Intersect(WrongTimeCells, TimeCell) Is Nothing Then
                                                WrongTimeCells = loSpotSelection.Application.Union(WrongTimeCells, TimeCell)
                                            End If
                                        Else
                                            WrongTimeCells = TimeCell
                                        End If

                                    End If
                                End If
                            End If
                        Next
                        DayChecking = True
                        CheckEmptyDayTime()
                        showErrorRange()
                        DayChecking = False
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub PlanSheet_Change(ByVal Target As Microsoft.Office.Interop.Excel.Range) Handles PlanSheet.Change
        Try
            'Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            'Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(sheet)
            '' vstoWorkbook.Name = "Plan Selection"
            'loSpotSelection = DirectCast(vstoWorkbook.Controls.Item("InputSpotSelection"), Microsoft.Office.Tools.Excel.ListObject)
            ''ChannelColumn = loSpotSelection.ListColumns("Channel")
            ''DayCoulmn = loSpotSelection.ListColumns("Day")
            ''StartTimeColumn = loSpotSelection.ListColumns("Start Time")
            ''EndTimeColumn = loSpotSelection.ListColumns("End Time")
            ''ChannelCells = ChannelColumn.DataBodyRange
            ''  logTaskPane = New TaskPaneLogFile
            'If Not PlanChecking Then
            '    If Not loSpotSelection.Application.Intersect(Target, DayCoulmn.DataBodyRange) Is Nothing Then
            '        If Not DayChecking Then
            '            For Each DayCell As Excel.Range In loSpotSelection.Application.Intersect(Target, DayCoulmn.DataBodyRange).Cells
            '                If SubtotalRows Is Nothing Then
            '                    If CheckDayValue(DayCell) Then
            '                        If Not WrongDayCells Is Nothing Then
            '                            If Not loSpotSelection.Application.Intersect(WrongDayCells, DayCell) Is Nothing Then
            '                                WrongDayCells = RemoveCellFromRange(WrongDayCells, DayCell)
            '                            End If
            '                        End If
            '                    Else
            '                        If Not WrongDayCells Is Nothing Then
            '                            If loSpotSelection.Application.Intersect(WrongDayCells, DayCell) Is Nothing Then
            '                                WrongDayCells = loSpotSelection.Application.Union(WrongDayCells, DayCell)
            '                            End If
            '                        Else
            '                            WrongDayCells = DayCell
            '                        End If
            '                    End If
            '                Else
            '                    If loSpotSelection.Application.Intersect(DayCell, SubtotalRows) Is Nothing Then
            '                        If CheckDayValue(DayCell) Then
            '                            If Not WrongDayCells Is Nothing Then
            '                                If Not loSpotSelection.Application.Intersect(WrongDayCells, DayCell) Is Nothing Then
            '                                    WrongDayCells = RemoveCellFromRange(WrongDayCells, DayCell)
            '                                End If
            '                            End If
            '                        Else
            '                            If Not WrongDayCells Is Nothing Then
            '                                If loSpotSelection.Application.Intersect(WrongDayCells, DayCell) Is Nothing Then
            '                                    WrongDayCells = loSpotSelection.Application.Union(WrongDayCells, DayCell)
            '                                End If
            '                            Else
            '                                WrongDayCells = DayCell
            '                            End If
            '                        End If
            '                    End If
            '                End If
            '            Next
            '            DayChecking = True
            '            CheckEmptyDayTime()
            '            ' showErrorRange()
            '            DayChecking = False
            '        End If
            '    End If
            '    If Not loSpotSelection.Application.Intersect(Target, loSpotSelection.Application.Union(StartTimeColumn.DataBodyRange, EndTimeColumn.DataBodyRange)) Is Nothing _
            '        Then
            '        If Not TimeChecking Then
            '            For Each TimeCell As Excel.Range In loSpotSelection.Application.Intersect(Target, loSpotSelection.Application.Union(StartTimeColumn.DataBodyRange, EndTimeColumn.DataBodyRange)).Cells
            '                If SubtotalRows Is Nothing Then
            '                    Dim CurrTime As Date
            '                    If DateTime.TryParse(TimeCell.Text, CurrTime) Then

            '                        TimeChecking = True
            '                        TimeCell.Value = CurrTime.ToOADate - Math.Floor(CurrTime.ToOADate) + 2
            '                        If Not loSpotSelection.Application.Intersect(StartTimeColumn.DataBodyRange, TimeCell) Is Nothing Then
            '                            If TypeOf TimeCell.Offset(0, 1).Value Is Double Then
            '                                If TimeCell.Value > TimeCell.Offset(0, 1).Value Then
            '                                    TimeCell.Offset(0, 1).Value = TimeCell.Offset(0, 1).Value + 1
            '                                End If
            '                            End If
            '                        End If
            '                        If Not loSpotSelection.Application.Intersect(EndTimeColumn.DataBodyRange, TimeCell) Is Nothing Then
            '                            If TypeOf TimeCell.Offset(0, -1).Value Is Double Then
            '                                If TimeCell.Value < TimeCell.Offset(0, -1).Value Then
            '                                    TimeCell.Value = TimeCell.Value + 1
            '                                End If
            '                            End If
            '                        End If
            '                        TimeChecking = False

            '                        If Not WrongTimeCells Is Nothing Then
            '                            If Not loSpotSelection.Application.Intersect(WrongTimeCells, TimeCell) Is Nothing Then
            '                                WrongTimeCells = RemoveCellFromRange(WrongTimeCells, TimeCell)
            '                            End If
            '                        End If
            '                        TimeChecking = True
            '                        With TimeCell
            '                            .ClearFormats()
            '                            .NumberFormat = "hh:mm;@"
            '                            .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            '                        End With
            '                        TimeChecking = False
            '                    Else
            '                        TimeChecking = True
            '                        Dim aTime As String
            '                        aTime = TimeCell.Text.ToString.Trim()
            '                        If aTime.Length = 5 Then
            '                            Try
            '                                Dim oldDate As New Date(1900, 1, 1)
            '                                Dim dblHour, dblMins As Double
            '                                Dim tmStartTime As Date
            '                                dblHour = aTime.Substring(0, 2)
            '                                dblMins = aTime.Substring(3, 2)

            '                                tmStartTime = DateAdd(DateInterval.Hour, dblHour, oldDate)
            '                                tmStartTime = DateAdd(DateInterval.Minute, dblMins, tmStartTime)
            '                                TimeCell.Value = tmStartTime
            '                                Continue For
            '                            Catch ex As Exception

            '                            End Try
            '                        End If
            '                        TimeChecking = False
            '                        If Not WrongTimeCells Is Nothing Then
            '                            If loSpotSelection.Application.Intersect(WrongTimeCells, TimeCell) Is Nothing Then
            '                                WrongTimeCells = loSpotSelection.Application.Union(WrongTimeCells, TimeCell)
            '                            End If
            '                        Else
            '                            WrongTimeCells = TimeCell
            '                        End If

            '                    End If
            '                Else
            '                    If loSpotSelection.Application.Intersect(TimeCell, SubtotalRows) Is Nothing Then
            '                        Dim CurrTime As Date
            '                        If DateTime.TryParse(TimeCell.Text, CurrTime) Then

            '                            TimeChecking = True
            '                            TimeCell.Value = CurrTime.ToOADate - Math.Floor(CurrTime.ToOADate) + 2
            '                            If Not loSpotSelection.Application.Intersect(StartTimeColumn.DataBodyRange, TimeCell) Is Nothing Then
            '                                If TypeOf TimeCell.Offset(0, 1).Value Is Double Then
            '                                    If TimeCell.Value > TimeCell.Offset(0, 1).Value Then
            '                                        TimeCell.Offset(0, 1).Value = TimeCell.Offset(0, 1).Value + 1
            '                                    End If
            '                                End If
            '                            End If
            '                            If Not loSpotSelection.Application.Intersect(EndTimeColumn.DataBodyRange, TimeCell) Is Nothing Then
            '                                If TypeOf TimeCell.Offset(0, -1).Value Is Double Then
            '                                    If TimeCell.Value < TimeCell.Offset(0, -1).Value Then
            '                                        TimeCell.Value = TimeCell.Value + 1
            '                                    End If
            '                                End If
            '                            End If
            '                            TimeChecking = False

            '                            If Not WrongTimeCells Is Nothing Then
            '                                If Not loSpotSelection.Application.Intersect(WrongTimeCells, TimeCell) Is Nothing Then
            '                                    WrongTimeCells = RemoveCellFromRange(WrongTimeCells, TimeCell)
            '                                End If
            '                            End If
            '                            TimeChecking = True
            '                            With TimeCell
            '                                .ClearFormats()
            '                                .NumberFormat = "hh:mm;@"
            '                                .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            '                            End With
            '                            TimeChecking = False
            '                        Else
            '                            TimeChecking = True
            '                            Dim aTime As String
            '                            aTime = TimeCell.Text.ToString.Trim()
            '                            If aTime.Length = 5 Then
            '                                Try
            '                                    Dim oldDate As New Date(1900, 1, 1)
            '                                    Dim dblHour, dblMins As Double
            '                                    Dim tmStartTime As Date
            '                                    dblHour = aTime.Substring(0, 2)
            '                                    dblMins = aTime.Substring(3, 2)

            '                                    tmStartTime = DateAdd(DateInterval.Hour, dblHour, oldDate)
            '                                    tmStartTime = DateAdd(DateInterval.Minute, dblMins, tmStartTime)
            '                                    TimeCell.Value = tmStartTime
            '                                    Continue For
            '                                Catch ex As Exception

            '                                End Try
            '                            End If
            '                            TimeChecking = False
            '                            If Not WrongTimeCells Is Nothing Then
            '                                If loSpotSelection.Application.Intersect(WrongTimeCells, TimeCell) Is Nothing Then
            '                                    WrongTimeCells = loSpotSelection.Application.Union(WrongTimeCells, TimeCell)
            '                                End If
            '                            Else
            '                                WrongTimeCells = TimeCell
            '                            End If

            '                        End If
            '                    End If
            '                End If
            '            Next
            '            DayChecking = True
            '            CheckEmptyDayTime()
            '            ' showErrorRange()
            '            DayChecking = False
            '        End If
            '    End If
            'End If
            'CleanUpPlan()
        Catch ex As Exception

        End Try
    End Sub
    Public Function GetSpotSelecListObject() As Microsoft.Office.Tools.Excel.ListObject
        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(sheet)
        ' vstoWorkbook.Name = "Plan Selection"
        loSpotSelection = DirectCast(vstoWorkbook.Controls.Item("InputSpotSelection"), Microsoft.Office.Tools.Excel.ListObject)
        Return loSpotSelection
    End Function
    Private Function RemoveCellFromRange(ByVal MainRange As Excel.Range, ByVal Cell As Excel.Range) As Excel.Range
        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(sheet)
        ' vstoWorkbook.Name = "Plan Selection"
        loSpotSelection = DirectCast(vstoWorkbook.Controls.Item("InputSpotSelection"), Microsoft.Office.Tools.Excel.ListObject)
        If MainRange Is Nothing Or Cell Is Nothing Then Return MainRange
        For Each MainCell As Excel.Range In MainRange
            If MainCell.AddressLocal = Cell.AddressLocal Then Continue For
            If RemoveCellFromRange Is Nothing Then
                RemoveCellFromRange = MainCell
                Continue For
            End If
            RemoveCellFromRange = loSpotSelection.Application.Union(RemoveCellFromRange, MainCell)
        Next
    End Function
    Friend Sub ResetVariables()
        PlanSheet = Nothing
        loSpotSelection = Nothing
        ChannelColumn = Nothing
        DayCoulmn = Nothing
        StartTimeColumn = Nothing
        EndTimeColumn = Nothing
        ChannelCells = Nothing
        SubtotalRows = Nothing
        BlankFields = Nothing
        WrongDayCells = Nothing
        WrongTimeCells = Nothing
        isPlanClean = False
        PlanChecking = False
        DayChecking = False
        dtChannelMaster = Nothing
        WeekColumns = Nothing
        isLoaded = False
        TotalSpotsWritten = 0
        TotalPlanSpots = 0
        dtPlanChannels = Nothing
        'dtExistingLog = New BrkPerfDataSet.ExistingLogDataTable
    End Sub
    Friend Sub CleanUpPlan()
        Try
            Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(sheet)
            Dim loname As String = String.Empty

            If vstoWorkbook.Name = "Average TVR" Then
                loname = "ProgAvgTVR"
            ElseIf vstoWorkbook.Name = "Plan Selection" Then
                loname = "InputSpotSelection"

            End If

            ' vstoWorkbook.Name = "Plan Selection"
            loSpotSelection = DirectCast(vstoWorkbook.Controls.Item(loname), Microsoft.Office.Tools.Excel.ListObject)
            loSpotSelection.Application.CutCopyMode = False
            loSpotSelection.Application.ScreenUpdating = False
            '   loSpotSelection.Refresh()
            PlanChecking = True
            '   vstoWorkbook.
            Try
                ChannelColumn = loSpotSelection.ListColumns("Channel")
                DayCoulmn = loSpotSelection.ListColumns("Day")
                StartTimeColumn = loSpotSelection.ListColumns("Start Time")
                EndTimeColumn = loSpotSelection.ListColumns("End Time")
                ChannelCells = ChannelColumn.DataBodyRange

                ClearFormats()
                FillEmptyChannels()
                HideTotals()
                FormatTimeBand()
                isPlanClean = Not CheckEmptyDayTime() And CheckDayFormats() And CheckWrongTimes()
                '  logTaskPane.showingErrors = False
                If Not isPlanClean Then
                    showErrorRange()
                Else
                    '  logTaskPane.scMain.Panel2Collapsed = True

                    If Not (Globals.Ribbons.MSprintExRibbon.ErrorPane Is Nothing) Then
                        Globals.Ribbons.MSprintExRibbon.ErrorPane.Visible = False
                    End If

                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
            End Try
            PlanChecking = False
            loSpotSelection.Application.CutCopyMode = True
            loSpotSelection.Application.ScreenUpdating = True
        Catch ex As Exception
            LogMpsrintExException("Exception occured while cleaning plan." + ex.Message)
        End Try
    End Sub
    Private Sub CollectErrors(ByRef ErrRange As Excel.Range, ByRef WrongCells As Excel.Range)
        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(sheet)
        ' vstoWorkbook.Name = "Plan Selection"
        Dim loname As String = String.Empty

        If vstoWorkbook.Name = "Average TVR" Then
            loname = "ProgAvgTVR"
        ElseIf vstoWorkbook.Name = "Plan Selection" Then
            loname = "InputSpotSelection"

        End If
        loSpotSelection = DirectCast(vstoWorkbook.Controls.Item(loname), Microsoft.Office.Tools.Excel.ListObject)
        If Not ErrRange Is Nothing And Not WrongCells Is Nothing Then

            ErrRange = loSpotSelection.Application.Union(ErrRange, WrongCells)
        ElseIf ErrRange Is Nothing And Not WrongCells Is Nothing Then
            ErrRange = WrongCells
        End If
    End Sub
    Private Sub ClearFormats()
        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(sheet)
        ' vstoWorkbook.Name = "Plan Selection"
        Dim loname As String = String.Empty

        If vstoWorkbook.Name = "Average TVR" Then
            loname = "ProgAvgTVR"
        ElseIf vstoWorkbook.Name = "Plan Selection" Then
            loname = "InputSpotSelection"

        End If
        loSpotSelection = DirectCast(vstoWorkbook.Controls.Item(loname), Microsoft.Office.Tools.Excel.ListObject)
        loSpotSelection.DataBodyRange.ClearFormats()
        Try
            loSpotSelection.TableStyle = ""
        Catch ex As Exception

        End Try
        loSpotSelection.TableStyle = "TableStyleMedium2"
        loSpotSelection.ShowAutoFilter = False
    End Sub
    Private Sub FillEmptyChannels()
        Dim rowcount As Integer = ChannelCells.Count
        Dim remainder As Integer
        Dim BlankCellCount As Integer = Math.DivRem(rowcount, 8192, remainder)
        Dim LoopEnd As Integer
        If remainder > 0 Then LoopEnd = BlankCellCount + 1 Else LoopEnd = BlankCellCount
        Dim BlankChannelCells As Excel.Range
        Dim count As Integer = 1
        If (BlankCellCount > 1) Then
            For count = 1 To BlankCellCount
                Try
                    BlankChannelCells = PlanSheet.Range(ChannelCells(((count - 1) * 8192 + 1)), ChannelCells(count * 8192)).SpecialCells(Excel.XlCellType.xlCellTypeBlanks)
                    BlankChannelCells.FormulaR1C1 = "=R[-1]C"
                Catch ex As Exception
                    Exit Sub
                End Try
            Next
        Else

            Try
                BlankChannelCells = PlanSheet.Range(ChannelCells(((count - 1) * 8192 + 1)), ChannelCells((count - 1) * 8192 + 1)).SpecialCells(Excel.XlCellType.xlCellTypeBlanks)
                BlankChannelCells.FormulaR1C1 = ""
            Catch ex As Exception
                Exit Sub
            End Try
        End If

        ChannelCells.Copy()
        Dim ChannelRngStart As Excel.Range = ChannelCells(1)
        ChannelRngStart.PasteSpecial(Excel.XlPasteType.xlPasteValues)
        ChannelRngStart.Select()
    End Sub
    'Private Sub FillEmptyChannels()
    '    Dim rowcount As Integer = ChannelCells.Count
    '    Dim remainder As Integer
    '    Dim BlankCellCount As Integer = Math.DivRem(rowcount, 8192, remainder)
    '    Dim LoopEnd As Integer
    '    If remainder > 0 Then LoopEnd = BlankCellCount + 1 Else LoopEnd = BlankCellCount
    '    Dim BlankChannelCells As Excel.Range
    '    Dim count As Integer
    '    For count = 1 To BlankCellCount
    '        Try
    '            BlankChannelCells = PlanSheet.Range(ChannelCells(((count - 1) * 8192 + 1)), ChannelCells(count * 8192)).SpecialCells(Excel.XlCellType.xlCellTypeBlanks)
    '            BlankChannelCells.FormulaR1C1 = "=R[-1]C"
    '        Catch ex As Exception
    '            Exit Sub
    '        End Try
    '    Next
    '    Try
    '        BlankChannelCells = PlanSheet.Range(ChannelCells(((count - 1) * 8192 + 1)), ChannelCells(count * 8192 + remainder)).SpecialCells(Excel.XlCellType.xlCellTypeBlanks)
    '        BlankChannelCells.FormulaR1C1 = "=R[-1]C"
    '    Catch ex As Exception
    '        Exit Sub
    '    End Try

    '    ChannelCells.Copy()
    '    Dim ChannelRngStart As Excel.Range = ChannelCells(1)
    '    ChannelRngStart.PasteSpecial(Excel.XlPasteType.xlPasteValues)
    '    ChannelRngStart.Select()
    'End Sub

    Private Sub HideTotals()
        Dim CurrTotalCell As Excel.Range
        Dim TotalCells As Excel.Range
        If ChannelCells.Rows.Count <= 1 Then
            If ChannelCells.Value <> Nothing AndAlso ChannelCells.Value.ToString.Contains("total") Then
                SubtotalRows = ChannelCells.EntireRow
            End If
        Else
            Dim TotalRows As Integer = ChannelCells.Rows(ChannelCells.Rows.Count).Row
            TotalCells = ChannelCells.Find(What:="total", MatchCase:=False, SearchDirection:=Excel.XlSearchDirection.xlNext)
            If CurrTotalCell Is Nothing And Not TotalCells Is Nothing Then CurrTotalCell = TotalCells
            Do Until TotalCells Is Nothing
                If SubtotalRows Is Nothing Then SubtotalRows = TotalCells.EntireRow Else SubtotalRows = loSpotSelection.Application.Union(SubtotalRows, TotalCells.EntireRow)
                If TotalCells.Row = TotalRows Then Exit Do
                TotalCells = ChannelCells.FindNext(TotalCells)
                If Not TotalCells Is Nothing Then
                    If TotalCells.AddressLocal = CurrTotalCell.AddressLocal Then
                        Exit Do
                    End If
                End If
            Loop
        End If
        If Not SubtotalRows Is Nothing Then
            SubtotalRows.EntireRow.Hidden = True
        End If
    End Sub
    'Private Sub HideTotals()
    '    Dim CurrTotalCell As Excel.Range
    '    Dim TotalCells As Excel.Range
    '    If ChannelCells.Rows.Count <= 1 Then
    '        If ChannelCells.Value.ToString.Contains("total") Then
    '            SubtotalRows = ChannelCells.EntireRow
    '        End If
    '    Else
    '        Dim TotalRows As Integer = ChannelCells.Rows(ChannelCells.Rows.Count).Row
    '        TotalCells = ChannelCells.Find(What:="total", MatchCase:=False, SearchDirection:=Excel.XlSearchDirection.xlNext)
    '        If CurrTotalCell Is Nothing And Not TotalCells Is Nothing Then CurrTotalCell = TotalCells
    '        Do Until TotalCells Is Nothing
    '            If SubtotalRows Is Nothing Then SubtotalRows = TotalCells.EntireRow Else SubtotalRows = loSpotSelection.Application.Union(SubtotalRows, TotalCells.EntireRow)
    '            If TotalCells.Row = TotalRows Then Exit Do
    '            TotalCells = ChannelCells.FindNext(TotalCells)
    '            If Not TotalCells Is Nothing Then
    '                If TotalCells.AddressLocal = CurrTotalCell.AddressLocal Then
    '                    Exit Do
    '                End If
    '            End If
    '        Loop
    '    End If
    '    If Not SubtotalRows Is Nothing Then
    '        SubtotalRows.EntireRow.Hidden = True
    '    End If
    'End Sub

    Private Sub FormatTimeBand()
        With loSpotSelection.Application.Union(StartTimeColumn.DataBodyRange, EndTimeColumn.DataBodyRange)
            .NumberFormat = "hh:mm;@"
            .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        End With
    End Sub

    Private Function CheckDayFormats() As Boolean
        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(sheet)
        ' vstoWorkbook.Name = "Plan Selection"
        Dim loname As String = String.Empty

        If vstoWorkbook.Name = "Average TVR" Then
            loname = "ProgAvgTVR"
        ElseIf vstoWorkbook.Name = "Plan Selection" Then
            loname = "InputSpotSelection"

        End If
        loSpotSelection = DirectCast(vstoWorkbook.Controls.Item(loname), Microsoft.Office.Tools.Excel.ListObject)
        CheckDayFormats = True
        Dim DayCells As Excel.Range
        DayCells = loSpotSelection.ListColumns("Day").DataBodyRange
        WrongDayCells = Nothing
        For Each Day As Excel.Range In DayCells.Cells
            If Not SubtotalRows Is Nothing Then
                If Not loSpotSelection.Application.Intersect(Day, SubtotalRows) Is Nothing Then
                    Continue For
                End If
            End If
            If Not CheckDayValue(Day) Then
                If WrongDayCells Is Nothing Then WrongDayCells = Day Else WrongDayCells = loSpotSelection.Application.Union(WrongDayCells, Day)
                CheckDayFormats = False
            End If
        Next
    End Function
    Private Function CheckDayValue(ByRef Day As Excel.Range) As Boolean
        With Day
            Try
                Dim ErrorFound As Boolean = False
                DayChecking = True
                .Value = formatDays(.Value, ErrorFound)
                DayChecking = False
                If .Value Is Nothing Or ErrorFound Then
                    CheckDayValue = False
                Else
                    DayChecking = True
                    .ClearFormats()
                    DayChecking = False
                    CheckDayValue = True
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
            End Try
        End With
    End Function
    Private Function formatDays(ByVal Days As String, ByRef ErrorFound As Boolean) As String
        Dim WeekDays() As String = {"MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"}
        '  Dim WeekDays() As String = {"Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"}
        Dim arrGrpSeparator() As Char = {"-", "~", ":"}
        Dim arrCommaSeparator() As Char = {".", ",", "/"}
        Dim returnDays As String = ""
        If Days Is Nothing Then
            ErrorFound = True
            Return ""
            Exit Function
        End If

        Dim arrDays() As String
        formatDays = ""
        If Days.ToUpper().Contains("-") Or Days.ToUpper().Contains("~") Or Days.ToUpper().Contains(":") Then
            arrDays = Days.ToUpper().Split(arrGrpSeparator)
            Try
                If arrDays.Length = 1 Then GoTo Singleton

                If Trim(arrDays(0)).ToUpper().Equals("THUR") Then
                    arrDays(0) = "THU"
                End If

                If Trim(arrDays(1)).ToUpper().Equals("THUR") Then
                    arrDays(1) = "THU"
                End If

                Dim intStart As Integer = Array.IndexOf(WeekDays, Trim(arrDays(0)))
                Dim intEnd As Integer = Array.IndexOf(WeekDays, Trim(arrDays(1)))
                If intStart > intEnd Then
                    For i As Integer = intStart To 6
                        If i <> 6 Then
                            formatDays += WeekDays(i).ToString() & ","
                        Else
                            formatDays += WeekDays(i).ToString()
                        End If
                    Next
                    For i As Integer = 0 To intEnd
                        If i <> intEnd Then
                            formatDays += "," & WeekDays(i).ToString() & ","
                        Else
                            formatDays += "," & WeekDays(i).ToString()
                        End If
                    Next
                Else
                    For i As Integer = intStart To intEnd
                        If i <> intEnd Then
                            formatDays += WeekDays(i).ToString() & ","
                        Else
                            formatDays += WeekDays(i).ToString()
                        End If
                    Next
                End If
            Catch ex As Exception
                ErrorFound = True
                Return Days
                Exit Function
            End Try
            Exit Function
        ElseIf Days.ToUpper().Contains(".") Or Days.ToUpper().Contains(",") Or Days.ToUpper().Contains("/") Then
            arrDays = Days.ToUpper().Split(arrCommaSeparator)
            Try
                If arrDays.Length = 1 Then GoTo Singleton
                For Each sDay In arrDays
                    If Array.IndexOf(WeekDays, Trim(sDay)) < 0 Then
                        ErrorFound = True
                        Return Days
                        Exit Function
                    End If
                Next
                Return Days.Replace("/", ",").Replace(".", ",").ToUpper()
            Catch ex As Exception
                ErrorFound = True
                Return Days
                Exit Function
            End Try
        End If

Singleton:
        If Days.ToUpper().Contains("ALL") Or Days.ToUpper().Contains("DLY") Or Days.ToUpper().Contains("DAILY") Or Days.ToUpper().Contains("ALL DAYS") Then
            formatDays = "MON,TUE,WED,THU,FRI,SAT,SUN"
            '  formatDays = "Mon,Tue,Wed,Thu,Fri,Sat,Sun"
        ElseIf Array.IndexOf(WeekDays, Days.ToUpper()) >= 0 Then
            formatDays = Days.ToUpper()
        ElseIf Days.ToUpper().Equals("SUNDAY") Then
            formatDays = "SUN"
        ElseIf Days.ToUpper().Equals("SATURDAY") Then
            formatDays = "SAT"
        Else
            ErrorFound = True
            formatDays = Days
        End If
    End Function
    Private Function CheckEmptyDayTime() As Boolean
        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(sheet)
        ' vstoWorkbook.Name = "Plan Selection"
        Dim loname As String = String.Empty

        If vstoWorkbook.Name = "Average TVR" Then
            loname = "ProgAvgTVR"
        ElseIf vstoWorkbook.Name = "Plan Selection" Then
            loname = "InputSpotSelection"

        End If
        loSpotSelection = DirectCast(vstoWorkbook.Controls.Item(loname), Microsoft.Office.Tools.Excel.ListObject)
        Dim AllBlankFields As Excel.Range
        BlankFields = Nothing
        Try
            AllBlankFields = loSpotSelection.Application.Union(loSpotSelection.ListColumns("Day").DataBodyRange, loSpotSelection.ListColumns("Start Time").DataBodyRange, loSpotSelection.ListColumns("End Time").DataBodyRange).SpecialCells(Excel.XlCellType.xlCellTypeBlanks)
        Catch ex As Exception
            CheckEmptyDayTime = False
            Exit Function
        End Try
        Dim TotalCellsInBlankFields As Excel.Range
        If Not SubtotalRows Is Nothing Then
            TotalCellsInBlankFields = loSpotSelection.Application.Intersect(AllBlankFields, SubtotalRows)
        End If
        If Not TotalCellsInBlankFields Is Nothing Then
            For Each cell As Excel.Range In AllBlankFields
                If loSpotSelection.Application.Intersect(cell, TotalCellsInBlankFields) Is Nothing Then
                    If BlankFields Is Nothing Then
                        BlankFields = cell
                    Else
                        BlankFields = loSpotSelection.Application.Union(BlankFields, cell)
                    End If
                End If
            Next cell
        Else
            For Each cell As Excel.Range In AllBlankFields
                If BlankFields Is Nothing Then
                    BlankFields = cell
                Else
                    BlankFields = loSpotSelection.Application.Union(BlankFields, cell)
                End If
            Next cell
        End If
        If Not BlankFields Is Nothing Then
            CheckEmptyDayTime = True
        Else
            CheckEmptyDayTime = False
        End If
    End Function
    Private Function CheckWrongTimes() As Boolean
        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(sheet)
        ' vstoWorkbook.Name = "Plan Selection"
        Dim loname As String = String.Empty

        If vstoWorkbook.Name = "Average TVR" Then
            loname = "ProgAvgTVR"
        ElseIf vstoWorkbook.Name = "Plan Selection" Then
            loname = "InputSpotSelection"

        End If
        loSpotSelection = DirectCast(vstoWorkbook.Controls.Item(loname), Microsoft.Office.Tools.Excel.ListObject)
        Dim TimeColumn As Excel.Range
        Try
            TimeColumn = loSpotSelection.Application.Union(loSpotSelection.ListColumns("Start Time").DataBodyRange, loSpotSelection.ListColumns("End Time").DataBodyRange)
        Catch ex As Exception
            MsgBox(ex.Message)
            CheckWrongTimes = True
            Exit Function
        End Try
        WrongTimeCells = Nothing
        For Each Cell As Excel.Range In TimeColumn.Cells
            If Not SubtotalRows Is Nothing Then
                If Not loSpotSelection.Application.Intersect(Cell, SubtotalRows) Is Nothing Then
                    Continue For
                End If
            End If
            Dim StartTime As Date
            If Not DateTime.TryParse(Cell.Text, StartTime) Then
                Dim currTime As String
                currTime = Cell.Text.ToString.Trim()
                If currTime.Length = 5 Then
                    Try
                        Dim oldDate As New Date(1900, 1, 1)
                        Dim dblHour, dblMins As Double
                        Dim tmStartTime As Date
                        dblHour = currTime.Substring(0, 2)
                        dblMins = currTime.Substring(3, 2)

                        tmStartTime = DateAdd(DateInterval.Hour, dblHour, oldDate)
                        tmStartTime = DateAdd(DateInterval.Minute, dblMins, tmStartTime)
                        Cell.Value = tmStartTime
                        Continue For
                    Catch ex As Exception

                    End Try
                End If
                If WrongTimeCells Is Nothing Then
                    WrongTimeCells = Cell
                Else
                    WrongTimeCells = loSpotSelection.Application.Union(WrongTimeCells, Cell)
                End If
            Else
                Cell.Value = StartTime.ToOADate - Math.Floor(StartTime.ToOADate) + 2
                If Not loSpotSelection.Application.Intersect(EndTimeColumn.DataBodyRange, Cell) Is Nothing Then
                    If TypeOf Cell.Offset(0, -1).Value Is Double Then
                        If Cell.Value < Cell.Offset(0, -1).Value Then
                            Cell.Value = Cell.Value + 1
                        End If
                    End If
                End If
            End If
        Next
        If Not WrongTimeCells Is Nothing Then
            CheckWrongTimes = False
        Else
            CheckWrongTimes = True
        End If
    End Function
    Private Function showErrorRange() As Boolean
        Try


            '   Dim noerrors As Boolean = True
            Dim ErrRange As Excel.Range

            CollectErrors(ErrRange, BlankFields)
            CollectErrors(ErrRange, WrongDayCells)
            CollectErrors(ErrRange, WrongTimeCells)

            DayCoulmn.DataBodyRange.ClearFormats()

            If Not ErrRange Is Nothing Then
                With ErrRange
                    With .Interior
                        .Pattern = Excel.XlPattern.xlPatternSolid
                        .PatternColorIndex = Excel.XlPattern.xlPatternAutomatic 'xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                End With
                Dim drAddress As Plandata.ErrorRangeRow
                dtErrors.Clear()
                For Each cell As Excel.Range In ErrRange
                    drAddress = dtErrors.NewErrorRangeRow
                    drAddress.Address = cell.AddressLocal
                    drAddress.Value = cell.Text
                    '   MessageBox.Show(cell.FormatConditions.Count.ToString())
                    dtErrors.AddErrorRangeRow(drAddress)

                    'If dtErrors.Rows.Count = 1 And Not (cell.Interior.Color.ToString().Equals("255")) Then
                    '    noerrors = True
                    'Else
                    '    noerrors = False
                    'End If

                Next
                ' logTaskPane.ucDataErrors.DataSource = dtErrors
                If Not (Globals.Ribbons.MSprintExRibbon.ErrorPane Is Nothing) Then
                    Globals.Ribbons.MSprintExRibbon.ErrorPane.Visible = False
                    Globals.Ribbons.MSprintExRibbon.ErrorPane.Dispose()
                End If

                '  If Not (noerrors) Then
                Globals.Ribbons.MSprintExRibbon.errors = New DataErrors()
                Globals.Ribbons.MSprintExRibbon.errors.DataSource = dtErrors
                Globals.Ribbons.MSprintExRibbon.ErrorPane = Globals.ThisAddIn.CustomTaskPanes.Add(Globals.Ribbons.MSprintExRibbon.errors, "Errors")
                Globals.Ribbons.MSprintExRibbon.ErrorPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
                Globals.Ribbons.MSprintExRibbon.ErrorPane.Height = 226
                Globals.Ribbons.MSprintExRibbon.ErrorPane.Width = 271
                Globals.Ribbons.MSprintExRibbon.ErrorPane.Visible = True
                'End If
                ' logTaskPane.scMain.Panel2Collapsed = False
                'If Not logTaskPane.showingErrors Then logTaskPane.showErrorPanel(True)
            Else
                isPlanClean = True
                Globals.Ribbons.MSprintExRibbon.ErrorPane.Visible = False
                ' logTaskPane.showErrorPanel(False)
                ' logTaskPane.scMain.Panel2Collapsed = True
            End If
        Catch ex As Exception

        End Try
        'ErrRange
    End Function
   
    Private Function GetChannelCodeFromMapping(ByVal ChannelName As String) As String
        dtChannelMap = daChannelMap.GetChannels(ChannelName)
        If dtChannelMap.Count > 0 Then
            GetChannelCodeFromMapping = dtChannelMap(0).TAMChannelCode
        Else
            GetChannelCodeFromMapping = "000"
        End If
    End Function

    Private Function GetChannelCodeFromMaster(ByVal ChannelName As String) As String
        Dim drChannelMaster() As Plandata.ChannelMasterRow
        drChannelMaster = dtChannelMaster.Select("ChannelName = '" & ChannelName & "'")
        If drChannelMaster.Length > 0 Then
            GetChannelCodeFromMaster = drChannelMaster(0).ChannelCode
        Else
            GetChannelCodeFromMaster = "000"
        End If
    End Function

    'Friend Sub ShowMoreChannels()
    '    Try

    '        If Not (Globals.Ribbons.MSprintExRibbon.channelMapping Is Nothing) Then
    '            Dim frmSelectChannel As New frmFilterChannels

    '            Dim CurrentChannel As System.Data.DataRowView = Globals.Ribbons.MSprintExRibbon.channelMapping.PlanChannelsBindingSource.Current
    '            frmSelectChannel.CurrentChannelCode = CurrentChannel.Row.Item("ChannelCode")
    '            frmSelectChannel.CurrentChannelName = CurrentChannel.Row.Item("ChannelName")
    '            frmSelectChannel.PlanChannelsBindingSource = Globals.Ribbons.MSprintExRibbon.channelMapping.PlanChannelsBindingSource
    '            If frmSelectChannel.ShowDialog() = Windows.Forms.DialogResult.OK Then
    '                Dim SelectedValue As String
    '                SelectedValue = frmSelectChannel.lbChannelMaster.SelectedValue
    '                CurrentChannel.Row.Item("ChannelCode") = SelectedValue
    '            End If
    '            frmSelectChannel.Dispose()
    '        End If
    '    Catch ex As Exception
    '        LogMpsrintExException("Exception occured while showing more channels." + ex.Message)
    '    End Try
    'End Sub
    Friend Sub CheckMate()
        'Dim conn As String
        'Dim appPath As String = System.AppDomain.CurrentDomain.BaseDirectory()
        'conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & appPath & ";Extended Properties=&qout;text;HDR=YES;FMT=FixedLength&quot;"
        Dim myComputer As New clsComputerInfo
        MsgBox(myComputer.GetProcessorId())
    End Sub
    Friend Function GetMachineInfo(ByVal byteLicense As Byte(), ByRef strLicenseType As String) As Boolean
        Dim strTrialText, strFullText As String
        Dim textbytes As Byte()
        Dim encoder As New UTF8Encoding()
        Dim myComputer As New clsComputerInfo
        Dim strProcessorID, strMACAddress, strVolumeSerial As String
        strProcessorID = myComputer.GetProcessorId
        strMACAddress = myComputer.GetMACAddress
        strVolumeSerial = myComputer.GetVolumeSerial
        strTrialText = strProcessorID & "badnansc" & strVolumeSerial
        strFullText = strProcessorID & "fullversions" & strVolumeSerial

        Dim strStoredLicense As String
        Dim cp As New CspParameters()
        cp.KeyContainerName = "MediaMatrix"
        Using RSA As New RSACryptoServiceProvider(cp)
            textbytes = RSADecrypt(byteLicense, RSA.ExportParameters(True), False)
            'textbytes = RSADecrypt(My.Settings.LicenseKey.Data, RSA.ExportParameters(True), False)
        End Using
        'textbytes = rsa.Decrypt(My.Settings.LicenseKey.Data, True)
        Try
            strStoredLicense = Convert.ToBase64String(textbytes)
        Catch ex As Exception
            Return False
        End Try
        If strTrialText = strStoredLicense Then
            strLicenseType = "Trial Version"
            Return True
        End If
        If strFullText = strStoredLicense Then
            strLicenseType = "Full Version"
            Return True
        End If

        Return False
    End Function
    Public Function RSAEncrypt(ByVal DataToEncrypt() As Byte, ByVal RSAKeyInfo As RSAParameters, ByVal DoOAEPPadding As Boolean) As Byte()
        Try
            Dim encryptedData() As Byte
            'Create a new instance of RSACryptoServiceProvider.
            Using RSA As New RSACryptoServiceProvider

                'Import the RSA Key information. This only needs
                'toinclude the public key information.
                RSA.ImportParameters(RSAKeyInfo)

                'Encrypt the passed byte array and specify OAEP padding.  
                'OAEP padding is only available on Microsoft Windows XP or
                'later.  
                encryptedData = RSA.Encrypt(DataToEncrypt, DoOAEPPadding)
            End Using
            Return encryptedData
            'Catch and display a CryptographicException  
            'to the console.
        Catch e As CryptographicException
            Console.WriteLine(e.Message)

            Return Nothing
        End Try
    End Function


    Public Function RSADecrypt(ByVal DataToDecrypt() As Byte, ByVal RSAKeyInfo As RSAParameters, ByVal DoOAEPPadding As Boolean) As Byte()
        Try
            Dim decryptedData() As Byte
            'Create a new instance of RSACryptoServiceProvider.
            Using RSA As New RSACryptoServiceProvider
                'Import the RSA Key information. This needs
                'to include the private key information.
                RSA.ImportParameters(RSAKeyInfo)

                'Decrypt the passed byte array and specify OAEP padding.  
                'OAEP padding is only available on Microsoft Windows XP or
                'later.  
                decryptedData = RSA.Decrypt(DataToDecrypt, DoOAEPPadding)
                'Catch and display a CryptographicException  
                'to the console.
            End Using
            Return decryptedData
        Catch e As CryptographicException
            Console.WriteLine(e.ToString())

            Return Nothing
        End Try
    End Function

    Public Class clsComputerInfo

        Friend Function GetProcessorId() As String
            Dim strProcessorId As String = String.Empty
            Dim query As New SelectQuery("Win32_processor")
            Dim search As New ManagementObjectSearcher(query)
            Dim info As ManagementObject

            For Each info In search.Get()
                strProcessorId = info("processorId").ToString()
            Next
            Return strProcessorId

        End Function

        Friend Function GetMACAddress() As String

            Dim mc As ManagementClass = New ManagementClass("Win32_NetworkAdapterConfiguration")
            Dim moc As ManagementObjectCollection = mc.GetInstances()
            Dim MACAddress As String = String.Empty
            For Each mo As ManagementObject In moc

                If (MACAddress.Equals(String.Empty)) Then
                    If CBool(mo("IPEnabled")) Then MACAddress = mo("MacAddress").ToString()

                    mo.Dispose()
                End If
                MACAddress = MACAddress.Replace(":", String.Empty)

            Next
            Return MACAddress
        End Function

        Friend Function GetVolumeSerial(Optional ByVal strDriveLetter As String = "C") As String

            Dim disk As ManagementObject = New ManagementObject(String.Format("win32_logicaldisk.deviceid=""{0}:""", strDriveLetter))
            disk.Get()
            Return disk("VolumeSerialNumber").ToString()
        End Function

        Friend Function GetMotherBoardID() As String

            Dim strMotherBoardID As String = String.Empty
            Dim query As New SelectQuery("Win32_BaseBoard")
            Dim search As New ManagementObjectSearcher(query)
            Dim info As ManagementObject
            For Each info In search.Get()

                strMotherBoardID = info("SerialNumber").ToString()

            Next
            Return strMotherBoardID

        End Function

    End Class

    Private Sub myCustomTaskPane_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MSprintExTaskPane.VisibleChanged
        'Globals.Ribbons.MSprintExRibbon.btnShowHide.Checked = CType(sender, Microsoft.Office.Tools.CustomTaskPane).Visible
    End Sub
    Friend Sub getTopPrograms(ByVal ReportType As String)
        If MSprintExTAMTaskPane Is Nothing Then
            mpTopPrograms = New ucTopPrograms
            MSprintExTAMTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(mpTopPrograms, "Top programs")
            MSprintExTAMTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
        End If
        MSprintExTAMTaskPane.Height = 226
        MSprintExTAMTaskPane.Width = 271
        MSprintExTAMTaskPane.Visible = True

        Select Case ReportType
            Case "Break TVR"
                mpTopPrograms.getTopProgramsBreakTVR()
            Case "Program TVR"

        End Select

    End Sub
    Public Function CheckSheetExists(ByVal name As String) As Boolean
        Dim exists As Boolean = False
        '        foreach (Sheet sheet in workbook.Sheets)
        '{
        '        If (sheet.Name.equals("sheetName")) Then
        '    {
        '        //do something
        '    }
        '}
        For Each sheet As Microsoft.Office.Interop.Excel.Worksheet In Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets


            If sheet.Name.Equals(name) Then
                exists = True
                Return exists
                Exit Function

            End If

        Next

        Return exists
    End Function
    Public Function ReturnActualSheet(ByVal name As String) As Microsoft.Office.Interop.Excel.Worksheet
        Dim sheet1 As Microsoft.Office.Interop.Excel.Worksheet
        Try
            For Each sheet As Microsoft.Office.Interop.Excel.Worksheet In Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets
                If sheet.Name.Contains(name) Then
                    sheet1 = sheet
                    Return sheet1
                    Exit Function
                    '  count += 1

                End If
            Next

        Catch ex As Exception

        End Try
        Return sheet1
    End Function
    Public Function CheckAndReturnSheet(ByVal name As String) As Integer
        'Dim sheet1 As Microsoft.Office.Interop.Excel.Worksheet
        Dim count As Integer = 0
        For Each sheet As Microsoft.Office.Interop.Excel.Worksheet In Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets


            If sheet.Name.Contains(name) Then
                'sheet1 = sheet
                'Return sheet1
                'Exit Function
                count += 1

            End If

        Next
        Return count
    End Function
End Module
