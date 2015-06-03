Imports Microsoft.Office.Tools
Imports Microsoft.Office.Interop.Excel
Imports System
Imports System.Data
Module BreakTVR
    Friend listObject, listobject1 As Excel.ListObject
    Friend nativeSheet, newSheet, vstoWorkbook As Microsoft.Office.Interop.Excel.Worksheet
    Public Function ConstructInputXMLForBreakTVR()
        Dim input As XElement
        Try

            Dim plantgname As String = String.Empty
            Dim reftgname As String = String.Empty
            Dim dtable As System.Data.DataTable = DirectCast(Globals.Ribbons.MSprintExRibbon.tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            plantgname = dtable.Rows(0)(1).ToString().Trim()
            ' reftgname = dtable.Rows(1)(1).ToString()


            'chds.ReadXml(Path.GetTempPath() + "ds1.xml")
            'Dim tvrform As TVRForm = New TVRForm(chds, Globals.Ribbons.MSprintExRibbon.tpSelections)
            'tvrform.ShowDialog()
            'If TVRTaskPane Is Nothing Then
            '    mpUcTVRScreen = New ucTVRScreen(chds, Globals.Ribbons.MSprintExRibbon.tpSelections)
            '    TVRTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(mpUcTVRScreen, "Program TVR Selections")
            '    TVRTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating
            'End If
            'TVRTaskPane.Height = 226
            'TVRTaskPane.Width = 271
            'TVRTaskPane.Visible = True
            Globals.ThisAddIn.Application.StatusBar = "Getting requested Break TVR details..."
            'frm = New frmWait()
            'frm.Show()
            'frm.Panel1.Refresh()
            'frm.Refresh()
            'Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlWait
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

            input =
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
            Dim dayparts As XElement =
         <day_parts></day_parts>
            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.LbDayParts.Items.Count - 1
                '<day_part>0200-0200</day_part>
                '  <day_part>0200-0200</day_part>
                Dim dpart As XElement = New XElement("day_part")
                dpart.Value = Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.LbDayParts.Items(index)
                dayparts.Add(dpart)
            Next
            Input.Add(dayparts)
            For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items.Count - 1
                ' Dim doc2 As XmlDocument = New XmlDocument()
                '  doc2.Load()
                Dim TG_MGElement As XElement =
                  <TG_MG name=<%= String.Format("{0}~{1}", plantgname, Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim()) %> type="Planning">
                  </TG_MG>
                TG_MGElement.Add(XElement.Load(Globals.Ribbons.MSprintExRibbon.tgDirectoryPath + plantgname + ".xml"))
                TG_MGElement.Add(XElement.Load(Globals.Ribbons.MSprintExRibbon.mgDirectoryPath + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim() + ".xml"))

                '<calc logic = "ProgTVR" display = "week-wise"/>
                '<num_programs>20</num_programs>
                Dim disp As String = String.Empty
                If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
                    disp = "clubbed"
                ElseIf Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
                    disp = "week-wise"
                End If

                Dim calclogic As XElement = <calc logic="BreakTVR" display=<%= disp %>/>
                Dim numprog As XElement = <num_programs><%= Convert.ToInt32(Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.nudTopPrograms.Value) %></num_programs>
                ' Dim numprog As XElement = <num_programs>20</num_programs>

                TG_MGElement.Add(calclogic)
                TG_MGElement.Add(numprog)
                'Dim ds As DataSet = New DataSet()
                'ds.ReadXml(Path.GetTempPath() + "ds1.xml")
                'dr("ChannelCode") = channel.Attribute("code").Value
                'dr("Channel Name")
                For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.lbSelectedChannels.Items.Count - 1
                    Dim chnnel As XElement =
                      <channel code=<%= Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.lbSelectedChannels.Items(index1).ToString() + "'")(0).Item("ID") %> name=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.lbSelectedChannels.Items(index1).ToString() %>/>
                    ' <channel code='479' name='COLORS'/>
                    'Dim cn As XElement = <channel code='4' name='STAR PLUS'/>
                    TG_MGElement.Add(chnnel)
                    'TG_MGElement.Add(cn)
                Next
                Input.Add(TG_MGElement)
                '  markets.Add(XElement.Load(Path.GetTempPath() + "\\MGS\\" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbPlan.Items(index).ToString().Trim() + ".xml"))
            Next
            '
            'For index = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items.Count - 1
            '    Dim TG_MGElement As XElement =
            '     <TG_MG name=<%= String.Format("{0}~{1}", reftgname, Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index).ToString().Trim()) %> type="Reference">
            '     </TG_MG>
            '    TG_MGElement.Add(XElement.Load(Globals.Ribbons.MSprintExRibbon.tgDirectoryPath + reftgname + ".xml"))
            '    TG_MGElement.Add(XElement.Load(Globals.Ribbons.MSprintExRibbon.mgDirectoryPath + Globals.Ribbons.MSprintExRibbon.tpSelections.UcMarkets1.lbRef.Items(index).ToString().Trim() + ".xml"))
            '    Dim disp As String = String.Empty
            '    If Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbSingle.Checked Then
            '        disp = "clubbed"
            '    ElseIf Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.rbWeekWise.Checked Then
            '        disp = "week-wise"
            '    End If
            '    Dim calclogic As XElement = <calc logic="BreakTVR" display=<%= disp %>/>
            '    '  Dim numprog As XElement = <num_programs><%= Convert.ToInt32(tvrform.nudTopPrograms.Value) %></num_programs>
            '    Dim numprog As XElement = <num_programs><%= Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.nudTopPrograms.Value %></num_programs>
            '    TG_MGElement.Add(calclogic)
            '    TG_MGElement.Add(numprog)
            '    'Dim ds As DataSet = New DataSet()
            '    'ds.ReadXml(Path.GetTempPath() + "ds1.xml")
            '    'dr("ChannelCode") = channel.Attribute("code").Value
            '    'dr("Channel Name")
            '    For index1 = 0 To Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.lbSelectedChannels.Items.Count - 1
            '        Dim chnnel As XElement =
            '          <channel code=<%= Globals.Ribbons.MSprintExRibbon.dtchannels.Select("Name = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.lbSelectedChannels.Items(index1).ToString() + "'")(0).Item("ID") %> name=<%= Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.lbSelectedChannels.Items(index1).ToString() %>/>
            '        TG_MGElement.Add(chnnel)
            '        'Dim chnnel As XElement = <channel code='479' name='COLORS'/>
            '        'Dim cn As XElement = <channel code='4' name='STAR PLUS'/>
            '        ' TG_MGElement.Add(chnnel)
            '        '  TG_MGElement.Add(cn)
            '    Next
            '    Input.Add(TG_MGElement)
            'Next
            Globals.Ribbons.MSprintExRibbon.UpdateUsageReport("BreakTVR WS", Globals.Ribbons.MSprintExRibbon.tpSelections.tbClientName.Text.Trim, Globals.Ribbons.MSprintExRibbon.tpSelections.tbBrandValue.Text.Trim, input, 0)

        Catch ex As Exception
            LogMpsrintExException("Exception occured while constructing input XML for Break TVR" + ex.Message)
            Throw ex
        End Try
        Return input
    End Function
    Public Function DisplayBreakTVRDetailsOnSheet(ByVal ds As DataSet, Optional ByVal fromPane As Boolean = False)
        Try
            Dim plantgname As String = String.Empty
            Dim reftgname As String = String.Empty
            Dim dtable As System.Data.DataTable = DirectCast(Globals.Ribbons.MSprintExRibbon.tpSelections.UcAudience1.DgPlanRefGrid.DataSource, System.Data.DataTable)
            plantgname = dtable.Rows(0)(1).ToString().Trim()
            ' reftgname = dtable.Rows(1)(1).ToString()
            If Not CheckSheetExists("Break TVR") Then
                newSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                newSheet.Name = "Break TVR"
            Else

                If fromPane Then
                    newSheet.UsedRange.Clear()
                    Globals.Ribbons.MSprintExRibbon.CleanSheet(newSheet)
                    newSheet.Activate()
                Else
                    '  newSheet = CheckAndReturnSheet("Break TVR")
                    Dim sheetcount As Integer = CheckAndReturnSheet("Break TVR")
                    If sheetcount > 0 Then
                        newSheet = DirectCast(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(, Globals.ThisAddIn.Application.ActiveSheet), Microsoft.Office.Interop.Excel.Worksheet)
                        Dim name As String = String.Format("Break TVR({0})", sheetcount)
                        newSheet.Name = name
                    End If
                End If

              

               
            End If


            Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(newSheet)
            ' Dim listObject As Excel.ListObject()
            Dim cell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.Range("$A$4", Type.Missing)
            cell.Value2 = String.Format("Period : {0} to {1}", Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToString("dd/MM/yyyy"), Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.ToString("dd/MM/yyyy"))
            '  cell.Interior.Color = System.Drawing.Color.Yellow
            ' cell.ColumnWidth = 15
            Dim cell1 As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.Range("$A$6", "$A$6")
            cell1.Value2 = String.Format("TG: {0}", plantgname)
            Dim ii As Integer = 1
            Dim rocount As Integer = 0
            Dim channel = String.Empty
            Dim lrange As Microsoft.Office.Interop.Excel.Range
            Dim count As Integer = 9
            For index = 0 To Globals.Ribbons.MSprintExRibbon.channels.Rows.Count - 1

                Dim cel As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count + Convert.ToInt32(Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.nudTopPrograms.Value + 4) * index, 1), vstoWorkbook.Cells(count + Convert.ToInt32(Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.nudTopPrograms.Value + 4) * index, 7 * ds.Tables.Count + 3))
                cel.Merge(True)
                cel.Value2 = Globals.Ribbons.MSprintExRibbon.channels.Rows(index)("CName").ToString()

                ' Dim periodCell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count + 1 + (24 * index), 1), vstoWorkbook.Cells(count + 1 + (24 * index), 7 * planningDataSet.Tables.Count + 3))
                Dim periodCell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count + 1 + Convert.ToInt32(Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.nudTopPrograms.Value + 4) * index, 1), vstoWorkbook.Cells(count + 1 + Convert.ToInt32(Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.nudTopPrograms.Value + 4) * index, 7 * ds.Tables.Count + 3))

                periodCell.Merge(True)
                periodCell.Value2 = String.Format("Period : {0} to {1}", Globals.Ribbons.MSprintExRibbon.channels.Rows(index)("StartDate").ToString(), Globals.Ribbons.MSprintExRibbon.channels.Rows(index)("EndDate").ToString())

                For index1 = 0 To ds.Tables.Count - 1
                    ' Dim marketcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(count + 2 + (24 * index), 7 * index1 + 1), vstoWorkbook.Cells(count + 1 + (24 * index), (7 * index1 + 1) * index1 + 1))
                    Dim value1 As Integer = count + 2 + Convert.ToInt32(Globals.Ribbons.MSprintExRibbon.tpSelections.UcChannels.nudTopPrograms.Value + 4) * index
                    Dim marketcell As Microsoft.Office.Interop.Excel.Range = vstoWorkbook.get_Range(vstoWorkbook.Cells(value1, 7 * index1 + 1), vstoWorkbook.Cells(value1, 7 * index1 + 1))

                    marketcell.Merge(True)
                    marketcell.Value2 = ds.Tables(index1).TableName.Split({"~"c}, StringSplitOptions.None)(1)

                    Dim lorange As Microsoft.Office.Interop.Excel.Range = marketcell.Offset(1, 0)
                    Dim listobject = vstoWorkbook.Controls.AddListObject(lorange, "list1C" + index.ToString() + Date.Now.Second.ToString() + Date.Now.Minute.ToString() + Date.Now.Millisecond.ToString())
                    listobject.AutoSetDataBoundColumnHeaders = True
                    ' 04/08/2013 ,10/08/2013
                    '  Dim dt As System.Data.DataTable = planningDataSet.Tables(index1).Select("ChannelName = '" + tvrform.cbChannels.CheckedItems(0).ToString() + "' and PeriodStartDate = '" + Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtFromDate.Value.ToShortDateString("DD/MM/YYYY") + "' and PeriodEndDate ='" + Globals.Ribbons.MSprintExRibbon.tpSelections.TaskPaneLogFile1.dtToDate.Value.ToShortDateString("DD/MM/YYYY") + "'").CopyToDataTable()
                    'channels.Columns.Add("CName")
                    'channels.Columns.Add("StartDate")
                    '  channels.Columns.Add("EndDate")
                    Dim dtrows As System.Data.DataRow() = ds.Tables(index1).Select("ChannelName = '" + Globals.Ribbons.MSprintExRibbon.channels.Rows(index)("CName").ToString() + "' and PeriodStartDate = '" + Globals.Ribbons.MSprintExRibbon.channels.Rows(index)("StartDate").ToString() + "' and PeriodEndDate = '" + Globals.Ribbons.MSprintExRibbon.channels.Rows(index)("EndDate").ToString() + "' ")
                    Dim dt As Data.DataTable
                    If dtrows.Count > 0 Then
                        dt = dtrows.CopyToDataTable()
                        dt.Columns.RemoveAt(0)
                        dt.Columns.RemoveAt(0)
                        dt.Columns.RemoveAt(0)
                        dt.Columns.RemoveAt(0)
                        dt.AcceptChanges()
                        listobject.DataSource = dt
                    End If

                    '.


                Next

                'If channel.Equals(channels.Rows(index)("CName").ToString()) Or index = 0 Then
                '    lrange = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(4, ii), vstoWorkbook.Cells(4, ii + 3)), Microsoft.Office.Interop.Excel.Range)
                'Else
                '    lrange = DirectCast(vstoWorkbook.get_Range(vstoWorkbook.Cells(7 + rocount, 1), vstoWorkbook.Cells(7 + rocount, 1)), Microsoft.Office.Interop.Excel.Range)

                'End If


                'ii += 10
                'rocount = listobject.ListRows.Count
                'channel = channels.Rows(index)("CName").ToString()
                count += 1
            Next
        Catch ex As Exception
            Globals.ThisAddIn.Application.StatusBar = String.Empty
            LogMpsrintExException("Exception occured while displaying Break TVR details." + ex.Message)
            Throw ex
        End Try
    End Function
End Module
