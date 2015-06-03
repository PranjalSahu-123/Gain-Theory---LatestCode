Imports System
Imports System.Data
Imports System.Linq
Imports System.Windows.Forms

Public Class ThisAddIn
    ' Public ribbonref As MSprintExRibbon
    Dim WithEvents cb, cb1, cb2 As Office.CommandBarButton
    Friend tpSelections, selectionsObject As ucPlanSelections
    Dim _ContextMenu, AvaiSpotsContextMenu, DeleteContextMenu As Office.CommandBar
    Private _isOpen As Boolean

    Public Property IsOpen() As Boolean
        Get
            Return _isOpen
        End Get
        Set(ByVal value As Boolean)
            _isOpen = value
        End Set
    End Property

    Private Sub MyWorkbookOpenEvent(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook)
        If Globals.ThisAddIn.IsOpen Then
            Globals.Ribbons.MSprintExRibbon.DisplayTVBuilder()
        End If
    End Sub

    Private Sub MyNewWorkbookEvent(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook)
        If Globals.ThisAddIn.IsOpen Then
            Globals.Ribbons.MSprintExRibbon.DisplayTVBuilder()
        End If
    End Sub

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        ''System.Threading.SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext())
        '' AddHandler Globals.ThisAddIn.Application.SheetBeforeDoubleClick, AddressOf Application_SheetRowSelected
        ' Globals.ThisAddIn.Application.Workbooks.Add()
        Try

            'AddHandler Globals.ThisAddIn.Application.WorkbookOpen, AddressOf MyWorkbookOpenEvent
            'AddHandler Globals.ThisAddIn.Application.NewWorkbook, AddressOf MyNewWorkbookEvent
            AddHandler Globals.ThisAddIn.Application.SheetChange, AddressOf Application_SheetSelectionChange
            AddHandler Globals.ThisAddIn.Application.SheetBeforeRightClick, AddressOf Application_SheetBeforeRightClick
            Try
                _ContextMenu = Me.Application.CommandBars.Add("ContextMenu", Office.MsoBarPosition.msoBarPopup, Type.Missing, True)
                AvaiSpotsContextMenu = Me.Application.CommandBars.Add("AvailableSpotsContextMenu", Office.MsoBarPosition.msoBarPopup, Type.Missing, True)
                DeleteContextMenu = Me.Application.CommandBars.Add("DeleteContextMenu", Office.MsoBarPosition.msoBarPopup, Type.Missing, True)
            Catch ex As Exception
                LogMpsrintExException("Exception occured while loading MsprintX addin.Message :" + ex.Message)

            End Try
            If _ContextMenu IsNot Nothing Then
                cb = DirectCast(_ContextMenu.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, True), Office.CommandBarButton)
                cb.Caption = "Replace"
                ' cb.BeginGroup = False
            End If
            Try

                If DeleteContextMenu IsNot Nothing Then
                    cb2 = DirectCast(_ContextMenu.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, True), Office.CommandBarButton)
                    cb2.Caption = "Delete"
                End If

            Catch ex As Exception
                LogMpsrintExException("Exception occured while loading MsprintX addin.Message :" + ex.Message)

            End Try

            If AvaiSpotsContextMenu IsNot Nothing Then
                cb1 = DirectCast(AvaiSpotsContextMenu.Controls.Add(Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, True), Office.CommandBarButton)
                cb1.Caption = "Select"
            End If
        Catch ex As Exception
            LogMpsrintExException("Exception occured while loading MsprintX addin.Message :" + ex.Message)
        End Try
        ' DisplayTVBuilder()
        'AddHandler Globals.ThisAddIn.Application.SheetBeforeDoubleClick, AddressOf Application_SheetRowDClicked
    End Sub
    Public Function DisplayTVBuilder()
        Try
            If Not tpSelections Is Nothing Then tpSelections.Dispose()
            tpSelections = New ucPlanSelections()
            ' Dim col As CustomTaskPaneCollection

            MSprintExTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(tpSelections, "TV Plan builder", Globals.ThisAddIn.Application.ActiveWindow)

            MSprintExTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight

            MSprintExTaskPane.Width = 450
            MSprintExTaskPane.Visible = False
        Catch ex As Exception
            LogMpsrintExException("Exception occured while displaying TV Builder pane")
            Throw ex
        End Try
    End Function

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        Dim workSheet As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet

        ' workSheet.QueryTables.Add(,,, workSheet.get_Range("$A$7"));
    End Sub
    Private Sub Application_SheetRowDClicked(ByVal sh As Object, ByVal Target As Excel.Range, ByRef Cancel As Boolean)
        'Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        'Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(sheet)
        'Dim row As Microsoft.Office.Interop.Excel.Range = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection
        ' [Something here is done]
        'Dim selection As Microsoft.Office.Interop.Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Microsoft.Office.Interop.Excel.Range)
        'For Each cell As Object In selection.Cells
        '    Try
        '        System.Windows.Forms.MessageBox.Show(DirectCast(cell, Microsoft.Office.Interop.Excel.Range).Value2.ToString())
        '    Catch
        '        System.Windows.Forms.MessageBox.Show("NULL VALUE")
        '    End Try
        'Next
        ' Try



    End Sub
    Private Sub Application_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Excel.Range, ByRef Cancel As Boolean)
        'Selected_Spots
        Try


            Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

            If sheet.Name.Contains("Selected_Spots") Then
                Cancel = True

                If Target.Rows.Count > 1 Then
                    AddContextMenu(False)
                Else
                    AddContextMenu(True)
                End If
            End If


            If sheet.Name.Contains("Available_Spots") Then
                Cancel = True

                If Target.Rows.Count > 1 Or Target.Row < 3 Then
                    AddAvaiSpotSContextMenu(False)
                Else
                    AddAvaiSpotSContextMenu(True)
                End If

            End If
        Catch ex As Exception
            LogMpsrintExException("Exception occured while show/hide of right click options.Message:" + ex.Message)
        End Try

    End Sub
    Private Sub Application_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)
        'Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        'Dim text As String = String.Empty
        'If sheet.Name.Equals("Genre Share") Then

        'ElseIf sheet.Name.Equals("Channel Share") Then
        'Else

        'End If

        ''NativeWorksheet = Application.ActiveWorkbook.Sheets("abc")
        ''worksheet = Globals.Factory.GetVstoObject(NativeWorksheet)
        '' [Something here is done]
        ''Dim selection As Microsoft.Office.Interop.Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Microsoft.Office.Interop.Excel.Range)
        ''For Each cell As Object In selection.Cells
        ''    Try
        ''        System.Windows.Forms.MessageBox.Show(DirectCast(cell, Microsoft.Office.Interop.Excel.Range).Value2.ToString())
        ''    Catch
        ''        System.Windows.Forms.MessageBox.Show("NULL VALUE")
        ''    End Try
        ''Next
        Globals.Ribbons.MSprintExRibbon.DisplayRelevantTGPane()
    End Sub
    Friend Sub AddContextMenu(ByVal actionVal As Boolean)


        '  If actionVal Then
        cb.Enabled = actionVal
        _ContextMenu.ShowPopup(Type.Missing, Type.Missing)
        cb2.Enabled = actionVal
        DeleteContextMenu.ShowPopup(Type.Missing, Type.Missing)
        '   _ContextMenu.Enabled = actionVal

        'Else
        '    _ContextMenu.
        'End If


        'End If

    End Sub
    Friend Sub AddAvaiSpotSContextMenu(ByVal actionVal As Boolean)
        cb1.Enabled = actionVal
        AvaiSpotsContextMenu.ShowPopup(Type.Missing, Type.Missing)

    End Sub
    Private Sub SelectButton_Click() Handles cb1.Click
        Try
            Dim range As Microsoft.Office.Interop.Excel.Range = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection
            Dim avaiRow As Data.DataRow = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Rows(range.Row - 2)
            Globals.Ribbons.MSprintExRibbon.currentLineItem = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Rows(range.Row - 2)("GUID").ToString()

            Dim cbWeeks As String = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Rows(range.Row - 2)("WeekNum").ToString()

            Dim selecpots As DataTable = New DataTable()
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

            Dim dt As DataTable = CType(loSpotSelection.DataSource, DataTable)
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
                Windows.Forms.MessageBox.Show(String.Format("Maximum of {0} spot(s) can be selected.Please reselect {0} spot(s)", count - selecpots.Rows.Count))
                ' ShowErrorLabel(String.Format("Maximum of {0} spot(s) can be selected.Please reselect {0} spot(s)", count - selecpots.Rows.Count))
            Else
                Windows.Forms.MessageBox.Show("Maximum number of spot(s) has been selected.Please remove unwanted and/or increase number of required spot(s) count and Try again")
                ' ShowErrorLabel("Maximum number of spot(s) has been selected.Please remove unwanted and/or increase number of required spot(s) count and Try again")
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub DeleteSpot() Handles cb2.Click
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
    Private Sub Button_Click() Handles cb.Click
        Try
            '  Dim active_sheet As Microsoft.Office.Interop.Excel.Worksheet = CType(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            Dim range As Microsoft.Office.Interop.Excel.Range = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection
            Globals.Ribbons.MSprintExRibbon.currentLineItem = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows(range.Row - 2)("GUID").ToString()
            '  Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.RemoveAt(range.Row - 2)
            Globals.Ribbons.MSprintExRibbon.Selectedrow = Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows(range.Row - 2)

            ' Globals.Ribbons.MSprintExRibbon.selectedRowIndex = range.Row - 2

            'Globals.Ribbons.MSprintExRibbon.selectedSpotsListObject.SetDataBinding(Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots)
            '  Dim drows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.xecelTable.Select(String.Format("GUID='{0}'", Globals.Ribbons.MSprintExRibbon.currentLineItem))
            GetAvailableSpots()
        Catch ex As Exception

        End Try

        'Private Sub Application_WindowBeforeRightClick()
        '    Throw New NotImplementedException
        'End Sub
        ' System.Windows.Forms.MessageBox.Show("Replace spots")
        'cb.
    End Sub
    Public Function GetAvailableSpots()
        Try
            Dim nativeSheet As Microsoft.Office.Interop.Excel.Worksheet
            System.Windows.Forms.Application.DoEvents()
            'HideErrorLabel()
            'btnGetAvailableSpots.Enabled = False
            Globals.ThisAddIn.Application.StatusBar = "Getting Available Spots..."
            System.Windows.Forms.Application.DoEvents()
            Dim input As XElement = ReachNFrequency.ConstructInputRnFXMLForAvailableSpots()
            Dim name As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
            input.Save(Globals.Ribbons.MSprintExRibbon.LogDirectoryPath + "AvailableSpots_Inp_" + name)
            '   Globals.Ribbons.MSprintExRibbon.rnfoutputXml = GetOpXMLFromWS(input, "http://ec2-54-255-217-55.ap-southeast-1.compute.amazonaws.com:8080/GroupM/spotselectionnew/getavailablespot")
            Globals.Ribbons.MSprintExRibbon.rnfoutputXml = GetOpXMLFromWS(input, Globals.Ribbons.MSprintExRibbon.GetURLForWS("AvaiSpotsWSURL_New"))
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
                    Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots = ConstructOpRnFTableForAvailableSpots(Globals.Ribbons.MSprintExRibbon.rnfoutputXml)
                    If Not CheckSheetExists("Available_Spots") Then
                        nativeSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                        nativeSheet.Name = "Available_Spots"
                    Else
                        nativeSheet = ReturnActualSheet("Available_Spots")
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
                    Dim listobject = vstoWorkbook.Controls.AddListObject(lrange, "Available_Spots" + Date.Now.Millisecond.ToString() + Date.Now.Second.ToString())
                    listobject.AutoSetDataBoundColumnHeaders = True

                    '  listobject.ListColumns(0).

                    'If reftg.Length > 0 Then
                    'grid.Columns("GUID").Visible = False
                    'grid.Columns("Spot").Visible = False
                    'grid.Columns("Start Date").Visible = False
                    'grid.Columns("End Date").Visible = False
                    'grid.Columns("WeekNum").Visible = False
                    Dim rnfselectedCopy As Data.DataTable = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Copy()



                    'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Channel")
                    'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Programme")
                    'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Creative")
                    'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Day")
                    'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Date", System.Type.GetType("System.DateTime"))
                    'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Start Time")
                    '' RnFSelectedSpots.Columns.Add("End Time")
                    'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
                    '' RnFSelectedSpots.Columns.Add("Commercial")
                    'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Cost")
                    'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("PA")
                    'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("TA")
                    'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("GUID")
                    'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Spot")
                    'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("Start Date", System.Type.GetType("System.DateTime"))
                    'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("End Date", System.Type.GetType("System.DateTime"))
                    'Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Columns.Add("WeekNum", System.Type.GetType("System.Int32"))

                    'Available spots table
                    'Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("GUID")
                    'Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("Channel")
                    'Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("Spot")
                    '' output.Columns.Add("AvaiSpotString")
                    ''Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("TG")
                    ''Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("MG")
                    ''output.Columns.Add("ReachVal")
                    '' Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("TVRVal")
                    'Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("Start Date", System.Type.GetType("System.DateTime"))
                    'Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("End Date", System.Type.GetType("System.DateTime"))
                    'Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("WeekNum", System.Type.GetType("System.Int32"))
                    'Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("Date", System.Type.GetType("System.DateTime"))
                    'Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("Start Time")
                    '' RnFSelectedSpots.Columns.Add("End Time")
                    'Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("Duration(Sec)", System.Type.GetType("System.Int32"))
                    '' RnFSelectedSpots.Columns.Add("Commercial")
                    'Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("Cost")
                    'Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("PA")
                    'Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Columns.Add("TA")
                    Dim newrow As Data.DataRow = rnfselectedCopy.NewRow()
                    newrow("GUID") = Globals.Ribbons.MSprintExRibbon.Selectedrow("GUID").ToString()
                    newrow("Channel") = Globals.Ribbons.MSprintExRibbon.Selectedrow("Channel").ToString()
                    newrow("Spot") = Globals.Ribbons.MSprintExRibbon.Selectedrow("Spot").ToString()
                    'newrow("TG") = Globals.Ribbons.MSprintExRibbon.Selectedrow("").ToString()
                    'newrow("MG") = Globals.Ribbons.MSprintExRibbon.Selectedrow("").ToString()
                    newrow("Start Date") = Globals.Ribbons.MSprintExRibbon.Selectedrow("Start Date").ToString()
                    newrow("End Date") = Globals.Ribbons.MSprintExRibbon.Selectedrow("End Date").ToString()
                    newrow("WeekNum") = Globals.Ribbons.MSprintExRibbon.Selectedrow("WeekNum").ToString()
                    newrow("Date") = Globals.Ribbons.MSprintExRibbon.Selectedrow("Date").ToString()
                    newrow("Start Time") = Globals.Ribbons.MSprintExRibbon.Selectedrow("Start Time").ToString()
                    newrow("Duration(Sec)") = Globals.Ribbons.MSprintExRibbon.Selectedrow("Duration(Sec)").ToString()
                    newrow("Cost") = Globals.Ribbons.MSprintExRibbon.Selectedrow("Cost").ToString()
                    newrow("PA") = Globals.Ribbons.MSprintExRibbon.Selectedrow("PA").ToString()
                    newrow("TA") = Globals.Ribbons.MSprintExRibbon.Selectedrow("TA").ToString()
                    rnfselectedCopy.Rows.InsertAt(newrow, 0)
                    rnfselectedCopy.Columns.Remove("GUID")
                    rnfselectedCopy.Columns.Remove("Spot")
                    rnfselectedCopy.Columns.Remove("Start Date")
                    rnfselectedCopy.Columns.Remove("End Date")
                    rnfselectedCopy.Columns.Remove("WeekNum")
                    listobject.DataSource = rnfselectedCopy
                    '  listobject.ListColumns(5).Range.NumberFormat = "dd/MM/yyyy"
                    Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Remove(Globals.Ribbons.MSprintExRibbon.Selectedrow)
                    listobject.ListRows(1).Range.Cells.Interior.Color = RGB(232, 10, 10)
                    listobject.ListColumns(2).Range.NumberFormat = "dd/MM/yyyy"
                    'listobject.ListRows(1).Range.Cells.Validation.ShowInput = True
                    'listobject.ListRows(1).Range.Cells.Validation.InputMessage = "Selected row tbe replaced."
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
                    'Dim filter As String = String.Format("GUID ='{0}' ", Globals.Ribbons.MSprintExRibbon.currentLineItem)

                    'If cbWeeks.Text = "All" Then
                    '    'filter = String.Format("GUID ='{0}' ", Globals.Ribbons.MSprintExRibbon.currentLineItem)
                    '    dgvAvailableSpotsGrid.DataSource = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots
                    'Else
                    '    filter = String.Format("GUID ='{0}' and WeekNum={1}", Globals.Ribbons.MSprintExRibbon.currentLineItem, Convert.ToInt32(cbWeeks.Text))
                    '    Dim rows As Data.DataRow() = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Select(filter)

                    '    If rows.Count > 0 Then
                    '        'Dim aspots As DataTable = New DataTable()
                    '        'aspots = Globals.Ribbons.MSprintExRibbon.RnFAvaiSpots.Clone()
                    '        ''Parallel.For(0, rows.Count - 1, Sub(i)
                    '        ''                                    aspots.ImportRow(rows(i))
                    '        ''                                End Sub)
                    '        'For Each row As DataRow In rows
                    '        '    aspots.ImportRow(row)
                    '        'Next
                    '        'dgvAvailableSpotsGrid.DataSource = rows.CopyToDataTable()
                    '    End If
                    'End If


                    '  Globals.Ribbons.MSprintExRibbon.HideSelectedSpotsGrid(dgvAvailableSpotsGrid)
                Else
                    MessageBox.Show("Unable to retreive Available Spots from Server.")
                End If

            Else
                System.Windows.Forms.MessageBox.Show("Unable to retreive Available Spots from Server.")
            End If
            'btnGetAvailableSpots.Enabled = True
            Globals.ThisAddIn.Application.StatusBar = String.Empty
        Catch ex As Exception
            'btnGetAvailableSpots.Enabled = True
            Globals.ThisAddIn.Application.StatusBar = String.Empty
            LogMpsrintExException("Exception occured while getting available spots" + ex.Message)
        End Try
    End Function
End Class
