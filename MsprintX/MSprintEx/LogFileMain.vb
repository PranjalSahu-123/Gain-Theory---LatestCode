Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms
Imports System.Reflection
Imports System.Data
Imports System.Runtime.CompilerServices
Module LogFileMain
    Dim frmProgress As frmWait
    Dim SaveFile As SaveFileDialog
    Friend Sub CreateLog()
        If Not (Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots Is Nothing) Then ' And BreakPerformance.isBrkPerfFile Then
            Try


                TotalSpotsWritten = 0
                TotalPlanSpots = 0
                SaveFile = Globals.Ribbons.MSprintExRibbon.logFileSavePath
                SaveFile = New SaveFileDialog()
                If SaveFile.ShowDialog() = DialogResult.OK Then
                    Globals.ThisAddIn.Application.StatusBar = "Creating log File.."
                    Globals.ThisAddIn.Application.Cursor = Excel.XlMousePointer.xlWait
                    Dim fl As New FileInfo(SaveFile.FileName)
                    'frmProgress = New frmWait
                    'frmProgress.Show()
                    'frmProgress.imgWait.Visible = True
                    Dim input As XElement = New XElement("req_for_log")

                    For index = 0 To Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows.Count - 1
                        Dim spot As XElement =
                            <spot SeqNumber=<%= (index + 1).ToString() %> log=<%= Globals.Ribbons.MSprintExRibbon.RnFSelectedSpots.Rows(index)("Spot").ToString() %>></spot>
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
                    Dim logfile As String = GetOpForLogWS(input, "http://ec2-54-254-193-184.ap-southeast-1.compute.amazonaws.com:8080/GroupM/spotselectionnew/getlogformat")
                    Using PlanLogFile As StreamWriter = New StreamWriter(SaveFile.FileName)
                        PlanLogFile.WriteLine(logfile)
                    End Using
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
                    Globals.ThisAddIn.Application.Cursor = Excel.XlMousePointer.xlDefault
                End If
            Catch ex As Exception
                Globals.ThisAddIn.Application.StatusBar = String.Empty
                Globals.ThisAddIn.Application.Cursor = Excel.XlMousePointer.xlDefault
            End Try
        End If
    End Sub

    'Private Sub CreateLog(ByVal FileName As String, ByVal WeekNumber As Integer, ByVal ColumnNumber As Integer, ByVal DateFrom As Date, ByVal DateTo As Date)
    '    Dim rowChannelName As String
    '    Dim rowProgram As String
    '    Dim rowDays As String
    '    Dim rowStartTime, rowEndTime As Date
    '    Dim rowSpots As Integer

    '    Dim ChannelCode As String = "000"
    '    Dim dtPlanChannels As Plandata.PlanChannelsDataTable
    '    Dim arrPlanChannels() As Plandata.PlanChannelsRow
    '    Dim daBrks As New BrkPerfDataSetTableAdapters.ProgRFMtTableAdapter
    '    Dim dtBrks As BrkPerfDataSet.ProgRFMtDataTable
    '    Dim currChannel As String = ""

    '    Dim SearchStartTime, SearchEndTime As Date
    '    Dim SearchSpotDate As Date

    '    Dim arrDayBreaks() As BrkPerfDataSet.ProgRFMtRow
    '    Dim arrDayBreaksNoBreak1() As BrkPerfDataSet.ProgRFMtRow
    '    Dim arrDayBreaksBreak1() As BrkPerfDataSet.ProgRFMtRow
    '    Dim RandomSpot As New Random

    '    Dim rowSpotsWritten As Integer
    '    Dim StatusCell As Excel.Range
    '    Dim HeaderStatusCell As Excel.Range = loPlanData.Range.Cells(1, loPlanData.ListColumns.Count + 1)
    '    If WeekNumber = 0 Then
    '        HeaderStatusCell.Value = "Total spots written in Log"
    '    Else
    '        HeaderStatusCell.Value = "Week " & WeekNumber & ": spots written in Log"
    '    End If

    '    Using PlanLogFile As StreamWriter = New StreamWriter(FileName)
    '        If (PlanLogFile IsNot Nothing) Then
    '            Try
    '                'frmProgress.pbUpload.Value = 0
    '                'frmProgress.pbUpload.Style = ProgressBarStyle.Marquee

    '                'Write header into logfile
    '                PlanLogFile.WriteLine("$Ch|--Date--|TimeFr|TimeTo|------ Caption Name --------------------|---Cost---|AP |TA |Durn |---- Programme Name --------------------|Channel| City Name|---------Product------------------------|----------Brand-------------------------|---------Variant------------------------|---------Advertiser-------------------------------Genre-------------------------|------Language------")

    '                'Loop through each data row in excel and analyze to write logfile
    '                For Each row As Excel.Range In ExcelPlan.loPlanData.DataBodyRange.Rows
    '                    'Skip subtotal rows in excel if found
    '                    If Not SubtotalRows Is Nothing Then
    '                        If Not ExcelPlan.loPlanData.Application.Intersect(row, SubtotalRows) Is Nothing Then
    '                            System.Diagnostics.Debug.Print(DirectCast(row.Value, System.Object)(1, 1) & "~" & row.AddressLocal)
    '                            Continue For
    '                        End If
    '                    End If

    '                    rowSpotsWritten = 0

    '                    'Get values in current row of excel
    '                    rowChannelName = DirectCast(row.Value, System.Object)(1, 1)
    '                    rowProgram = DirectCast(row.Value, System.Object)(1, 2)
    '                    rowDays = DirectCast(row.Value, System.Object)(1, 3)
    '                    Dim oldDate As New Date(1900, 1, 1)
    '                    Dim dblHour, dblMins As Double
    '                    Dim tmStartTime, tmEndTime As Date
    '                    Dim currTime As String

    '                    'Adjust for excel date storage formats
    '                    If TypeOf DirectCast(row.Value, System.Object)(1, 4) Is Double Then
    '                        rowStartTime = DateTime.FromOADate(DirectCast(row.Value, System.Object)(1, 4))
    '                    ElseIf TypeOf DirectCast(row.Value, System.Object)(1, 4) Is String Then
    '                        currTime = DirectCast(row.Value, System.Object)(1, 4).ToString.Trim.PadLeft(5, "0")
    '                        dblHour = currTime.Substring(0, 2)
    '                        dblMins = currTime.Substring(3, 2)
    '                        tmStartTime = DateAdd(DateInterval.Hour, dblHour, oldDate)
    '                        tmStartTime = DateAdd(DateInterval.Minute, dblMins, tmStartTime)
    '                        rowStartTime = tmStartTime
    '                    End If

    '                    If TypeOf DirectCast(row.Value, System.Object)(1, 5) Is Double Then
    '                        rowEndTime = DateTime.FromOADate(DirectCast(row.Value, System.Object)(1, 5))
    '                    ElseIf TypeOf DirectCast(row.Value, System.Object)(1, 5) Is String Then
    '                        currTime = DirectCast(row.Value, System.Object)(1, 5).ToString.Trim.PadLeft(5, "0")
    '                        dblHour = currTime.Substring(0, 2)
    '                        dblMins = currTime.Substring(3, 2)
    '                        tmEndTime = DateAdd(DateInterval.Hour, dblHour, oldDate)
    '                        tmEndTime = DateAdd(DateInterval.Minute, dblMins, tmEndTime)
    '                        rowEndTime = tmEndTime
    '                    End If

    '                    'Set endtime one day ahead in case endtime is less than start time
    '                    If rowEndTime < rowStartTime Then rowEndTime = DateAdd(DateInterval.Day, 1, rowEndTime)

    '                    'Get ChannelCode for current channel
    '                    If currChannel <> rowChannelName Then
    '                        dtPlanChannels = CType(logTaskPane.ucChannelsMapping.PlanChannelsBindingSource.DataSource, System.Data.DataView).Table
    '                        arrPlanChannels = dtPlanChannels.Select("ChannelName = '" & rowChannelName & "'")
    '                        If arrPlanChannels.Length > 0 Then
    '                            ChannelCode = arrPlanChannels(0).ChannelCode
    '                        End If
    '                        currChannel = rowChannelName
    '                    End If

    '                    'Dim arrRowDays(), inRowDays As String
    '                    'arrRowDays = rowDays.Split(",")
    '                    'For i As Integer = 0 To arrRowDays.Length - 1
    '                    '    arrRowDays(i) = "'" & arrRowDays(i) & "'"
    '                    'Next


    '                    'Get reformatted days in the "days" column of current row for filtering the brkperformance dataset using the "IN" clause
    '                    Dim inRowDays As String = "'" & String.Join("','", rowDays.Split(",")) & "'"

    '                    'Get the number of spots to be selected from brkperformance for current excel row
    '                    rowSpots = DirectCast(row.Value, System.Object)(1, ColumnNumber)


    '                    TotalPlanSpots += rowSpots

    '                    'If zero spots in current row then update status and skip row
    '                    If rowSpots = 0 Then
    '                        StatusCell = row.Cells(1, row.Cells.Count)
    '                        StatusCell.Value = rowSpots
    '                        System.Diagnostics.Debug.Print("Spots Added:Plan spots:0~" & ChannelCode & "~" & rowChannelName & "~0~" & inRowDays & "~" & rowStartTime.ToString() & "~" & rowEndTime.ToString())
    '                        Continue For
    '                    End If

    '                    SearchStartTime = DateAdd(DateInterval.Minute, 5, rowStartTime) 'Add 5 minutes to row start time to accomodate for program starting late
    '                    SearchEndTime = DateAdd(DateInterval.Minute, -2, rowEndTime) 'Reduce 2 minutes to the row end time to accomodate for program ending before time

    '                    Dim minTime As New Date(SearchStartTime.Year, SearchStartTime.Month, SearchStartTime.Day, 7, 0, 0) 'Define min time to search for spots as 7am in case of RODPs
    '                    Dim maxTime As New Date(SearchStartTime.Year, SearchStartTime.Month, SearchStartTime.Day, 23, 59, 59) 'Define max time to search for spots as 11:59pm in case of RODPs

    '                    'Assume any time band more than 4 hours as RODP and adjust search start time / end time accordingly
    '                    If DateDiff(DateInterval.Hour, SearchStartTime, SearchEndTime) > 4 Then
    '                        If SearchStartTime < minTime Then SearchStartTime = minTime
    '                        If SearchEndTime > maxTime Then SearchEndTime = maxTime
    '                    End If

    '                    'Get breaks for current row channel, days, search start time and search end time
    '                    dtBrks = daBrks.GetBreaksForChannelAndDays(ChannelCode, inRowDays, SearchStartTime, SearchEndTime, DateFrom, DateTo)

    '                    Dim strFilter As String
    '                    Dim strFilterNoBreak1 As String
    '                    Dim strFilterBreak1 As String

    '                    'There are 3 possibilities:
    '                    '1) brkperformance does not have any spots available for current row
    '                    '2) No. of brks available are greater than required no. of spots
    '                    '3) There aren't enough breaks available

    '                    '1) brkperformance does not have any spots available for current row
    '                    'If brkperformance does not have any spots available for current row, update status and skip row
    '                    If dtBrks.Count = 0 Then
    '                        System.Diagnostics.Debug.Print("Spots Added:No spots found~" & "~" & rowSpots & ChannelCode & "~" & rowChannelName & "~" & inRowDays & "~" & rowStartTime.ToString() & "~" & rowEndTime.ToString())
    '                        Continue For
    '                    End If

    '                    '2) No. of brks available are greater than required no. of spots
    '                    If dtBrks.Count > rowSpots Then
    '                        strFilterNoBreak1 = "CommercialName <> '---- End of Break 1 ----                '"
    '                        strFilterBreak1 = "CommercialName = '---- End of Break 1 ----                '"

    '                        'Avoid "End of Break1" if possible

    '                        'Select only those breaks which are not "End of Break1"
    '                        arrDayBreaksNoBreak1 = dtBrks.Select(strFilterNoBreak1)

    '                        If arrDayBreaksNoBreak1.Length >= rowSpots Then
    '                            Dim NoOfDistinctDays As Integer
    '                            Dim distinctDates = (From sDate As BrkPerfDataSet.ProgRFMtRow In dtBrks Where sDate.CommercialName <> "---- End of Break 1 ----                " Select sDate._Date).Distinct()
    '                            NoOfDistinctDays = distinctDates.Count
    '                            Dim SkipDays As Integer = Math.Floor(NoOfDistinctDays / rowSpots)
    '                            If SkipDays < 1 Then SkipDays = 1

    '                            Dim MaxPerDaySpots, RemainingRowSpots, currDaySpots, pendingDaySpots As Integer
    '                            MaxPerDaySpots = Math.DivRem(rowSpots, NoOfDistinctDays, RemainingRowSpots)
    '                            For iDay As Integer = 0 To distinctDates.Count - 1 Step SkipDays
    '                                SearchSpotDate = distinctDates(iDay)
    '                                strFilter = "Date = #" & SearchSpotDate.Month & "-" & SearchSpotDate.Day & "-" & SearchSpotDate.Year & "# and IsSelected = False and CommercialName <> '---- End of Break 1 ----                '"
    '                                arrDayBreaks = dtBrks.Select(strFilter)

    '                                If RemainingRowSpots > 0 Then
    '                                    currDaySpots = MaxPerDaySpots + 1 + pendingDaySpots
    '                                    RemainingRowSpots += -1
    '                                Else
    '                                    currDaySpots = MaxPerDaySpots + pendingDaySpots
    '                                End If
    '                                pendingDaySpots = 0
    '                                For i As Integer = 1 To currDaySpots
    '                                    Try
    '                                        Dim RandomSpotPos As Integer = RandomSpot.Next(arrDayBreaks.Length)
    '                                        Dim currSpot As BrkPerfDataSet.ProgRFMtRow = arrDayBreaks(RandomSpotPos)
    '                                        AddSpotRow(currSpot)
    '                                        PlanLogFile.WriteLine(currSpot.FullString)
    '                                        currSpot.IsSelected = True
    '                                        rowSpotsWritten += 1
    '                                        TotalSpotsWritten += 1
    '                                    Catch ex As Exception
    '                                        pendingDaySpots += 1
    '                                    End Try
    '                                    frmProgress.lblPercentage.Text = rowChannelName & ": " & rowProgram & vbCrLf & rowSpotsWritten & " of " & rowSpots & " spots"
    '                                    Application.DoEvents()
    '                                    arrDayBreaks = dtBrks.Select(strFilter)
    '                                Next
    '                            Next
    '                            For i As Integer = 1 To pendingDaySpots
    '                                strFilter = "IsSelected = False and CommercialName <> '---- End of Break 1 ----                '"
    '                                arrDayBreaks = dtBrks.Select(strFilter)
    '                                Dim RandomSpotPos As Integer = RandomSpot.Next(arrDayBreaks.Length)
    '                                Dim currSpot As BrkPerfDataSet.ProgRFMtRow = arrDayBreaks(RandomSpotPos)
    '                                AddSpotRow(currSpot)
    '                                PlanLogFile.WriteLine(currSpot.FullString)
    '                                currSpot.IsSelected = True
    '                                rowSpotsWritten += 1
    '                                TotalSpotsWritten += 1
    '                                frmProgress.lblPercentage.Text = rowChannelName & ": " & rowProgram & vbCrLf & rowSpotsWritten & " of " & rowSpots & " spots"
    '                                Application.DoEvents()
    '                            Next
    '                            pendingDaySpots = 0
    '                        Else
    '                            For Each spotRow As BrkPerfDataSet.ProgRFMtRow In arrDayBreaksNoBreak1
    '                                AddSpotRow(spotRow)
    '                                PlanLogFile.WriteLine(spotRow.FullString)
    '                                spotRow.IsSelected = True
    '                                rowSpotsWritten += 1
    '                                TotalSpotsWritten += 1
    '                                frmProgress.lblPercentage.Text = rowChannelName & ": " & rowProgram & vbCrLf & rowSpotsWritten & " of " & rowSpots & " spots"
    '                                Application.DoEvents()
    '                            Next

    '                            arrDayBreaksBreak1 = dtBrks.Select(strFilterBreak1)
    '                            Dim NoOfDistinctDays As Integer
    '                            Dim distinctDates = (From sDate As BrkPerfDataSet.ProgRFMtRow In dtBrks Where sDate.CommercialName = "---- End of Break 1 ----                " Select sDate._Date).Distinct()
    '                            NoOfDistinctDays = distinctDates.Count
    '                            Dim SkipDays As Integer = Math.Floor(NoOfDistinctDays / rowSpots)
    '                            If SkipDays < 1 Then SkipDays = 1

    '                            Dim MaxPerDaySpots, RemainingRowSpots, currDaySpots, pendingDaySpots As Integer
    '                            MaxPerDaySpots = Math.DivRem(rowSpots - arrDayBreaksNoBreak1.Length, NoOfDistinctDays, RemainingRowSpots)

    '                            For iDay As Integer = 0 To distinctDates.Count - 1 Step SkipDays
    '                                SearchSpotDate = distinctDates(iDay)

    '                                If RemainingRowSpots > 0 Then
    '                                    currDaySpots = MaxPerDaySpots + 1 + pendingDaySpots
    '                                    RemainingRowSpots += -1
    '                                Else
    '                                    currDaySpots = MaxPerDaySpots + pendingDaySpots
    '                                End If
    '                                pendingDaySpots = 0

    '                                strFilter = "Date = #" & SearchSpotDate.Month & "-" & SearchSpotDate.Day & "-" & SearchSpotDate.Year & "# and IsSelected = False and CommercialName = '---- End of Break 1 ----                '"
    '                                arrDayBreaks = dtBrks.Select(strFilter)
    '                                For i As Integer = 1 To currDaySpots
    '                                    Try
    '                                        Dim RandomSpotPos As Integer = RandomSpot.Next(arrDayBreaks.Length)
    '                                        Dim currSpot As BrkPerfDataSet.ProgRFMtRow = arrDayBreaks(RandomSpotPos)
    '                                        AddSpotRow(currSpot)
    '                                        PlanLogFile.WriteLine(currSpot.FullString)
    '                                        currSpot.IsSelected = True
    '                                        rowSpotsWritten += 1
    '                                        TotalSpotsWritten += 1
    '                                    Catch ex As Exception
    '                                        pendingDaySpots += 1
    '                                    End Try

    '                                    frmProgress.lblPercentage.Text = rowChannelName & ": " & rowProgram & vbCrLf & rowSpotsWritten & " of " & rowSpots & " spots"
    '                                    Application.DoEvents()
    '                                    If rowSpotsWritten = rowSpots Then Exit For
    '                                    arrDayBreaks = dtBrks.Select(strFilter)
    '                                Next
    '                            Next
    '                            For i As Integer = 1 To pendingDaySpots
    '                                strFilter = "IsSelected = False and CommercialName = '---- End of Break 1 ----                '"
    '                                arrDayBreaks = dtBrks.Select(strFilter)
    '                                Dim RandomSpotPos As Integer = RandomSpot.Next(arrDayBreaks.Length)
    '                                Dim currSpot As BrkPerfDataSet.ProgRFMtRow = arrDayBreaks(RandomSpotPos)
    '                                AddSpotRow(currSpot)
    '                                PlanLogFile.WriteLine(currSpot.FullString)
    '                                currSpot.IsSelected = True
    '                                rowSpotsWritten += 1
    '                                TotalSpotsWritten += 1
    '                                frmProgress.lblPercentage.Text = rowChannelName & ": " & rowProgram & vbCrLf & rowSpotsWritten & " of " & rowSpots & " spots"
    '                                Application.DoEvents()
    '                            Next
    '                            pendingDaySpots = 0

    '                        End If
    '                    End If
    '                    '3) There aren't enough breaks available
    '                    '   Repeat all spots available
    '                    If dtBrks.Count <= rowSpots Then
    '                        Do While rowSpotsWritten < rowSpots
    '                            For Each spotRow As BrkPerfDataSet.ProgRFMtRow In dtBrks.Rows
    '                                AddSpotRow(spotRow)
    '                                PlanLogFile.WriteLine(spotRow.FullString)
    '                                spotRow.IsSelected = True
    '                                rowSpotsWritten += 1
    '                                TotalSpotsWritten += 1
    '                                frmProgress.lblPercentage.Text = rowChannelName & ": " & rowProgram & vbCrLf & rowSpotsWritten & " of " & rowSpots & " spots"
    '                                Application.DoEvents()
    '                                If rowSpotsWritten = rowSpots Then Exit For
    '                            Next
    '                        Loop
    '                    End If


    '                    StatusCell = row.Cells(1, row.Cells.Count)
    '                    StatusCell.Value = rowSpotsWritten
    '                    System.Diagnostics.Debug.Print("Spots Added:" & rowSpotsWritten & "~" & rowSpots & "~" & ChannelCode & "~" & rowChannelName & "~" & dtBrks.Count & "~" & inRowDays & "~" & rowStartTime.ToString() & "~" & rowEndTime.ToString())
    '                Next
    '            Catch ex As Exception
    '                'frmProgress.pbUpload.Style = ProgressBarStyle.Continuous
    '                'frmProgress.pbUpload.Value = 0
    '                frmProgress.lblPercentage.ForeColor = Drawing.Color.Red
    '                frmProgress.lblPercentage.Text = "Log file creation failed:" & vbCrLf & ex.Message
    '                frmProgress.imgWait.Visible = False
    '            End Try
    '        End If
    '    End Using

    'End Sub
    'For iDay As Integer = 0 To distinctDates.Count - 1 Step SkipDays
    '    SearchSpotDate = distinctDates(iDay)
    '    strFilterNoBreak1 = "Date = #" & SearchSpotDate.Month & "-" & SearchSpotDate.Day & "-" & SearchSpotDate.Year & "# and IsSelected = False and  CommercialName <> '---- End of Break 1 ----                '"
    '    strFilterBreak1 = "Date = #" & SearchSpotDate.Month & "-" & SearchSpotDate.Day & "-" & SearchSpotDate.Year & "# and IsSelected = False and  CommercialName = '---- End of Break 1 ----                '"
    '    arrDayBreaksNoBreak1 = dtBrks.Select(strFilterNoBreak1)
    '    arrDayBreaksBreak1 = dtBrks.Select(strFilterBreak1)

    '    If RemainingRowSpots > 0 Then
    '        currDaySpots = MaxPerDaySpots + 1
    '        RemainingRowSpots += -1
    '    Else
    '        currDaySpots = MaxPerDaySpots
    '    End If
    'Select Case arrDayBreaks.Length
    '    Case Is = 0
    '    Case Is >= currDaySpots
    '        For i As Integer = 1 To currDaySpots
    '            Dim RandomSpotPos As Integer = RandomSpot.Next(arrDayBreaks.Length)
    '            Dim currSpot As BrkPerfDataSet.ProgRFMtRow = arrDayBreaks(RandomSpotPos)
    '            currSpot.IsSelected = True

    '            PlanLogFile.WriteLine(currSpot.FullString)
    '            rowSpotsWritten += 1
    '            TotalSpotsWritten += 1

    '            frmProgress.lblPercentage.Text = rowChannelName & ": " & rowProgram & vbCrLf & rowSpotsWritten & " of " & rowSpots & " spots"
    '            Application.DoEvents()

    '            arrDayBreaks = dtBrks.Select(strFilter)
    '        Next
    '    Case Is < currDaySpots
    '        Dim CountSpots As Integer = 0
    '        Do While CountSpots < currDaySpots
    '            For Each sDayBreak As BrkPerfDataSet.ProgRFMtRow In arrDayBreaks
    '                PlanLogFile.WriteLine(sDayBreak.FullString)
    '                sDayBreak.IsSelected = True
    '                rowSpotsWritten += 1
    '                TotalSpotsWritten += 1

    '                frmProgress.lblPercentage.Text = rowChannelName & ": " & rowProgram & vbCrLf & rowSpotsWritten & " of " & rowSpots & " spots"
    '                Application.DoEvents()

    '                CountSpots += 1
    '                If CountSpots = currDaySpots Then Exit Do
    '            Next
    '        Loop
    'End Select
    'Next
    'Sub getDates()
    '    Dim rowChannelName As String
    '    Dim rowProgram As String
    '    Dim rowDays As String
    '    Dim rowStartTime, rowEndTime As Date
    '    Dim rowSpots As Integer
    '    Dim ChannelCode As String = "000"
    '    Dim dtPlanChannels As Plandata.PlanChannelsDataTable
    '    Dim arrPlanChannels() As Plandata.PlanChannelsRow
    '    Dim daBrks As New BrkPerfDataSetTableAdapters.ProgRFMtTableAdapter
    '    Dim dtBrks As BrkPerfDataSet.ProgRFMtDataTable
    '    Dim currChannel As String = ""

    '    Dim SearchStartTime, SearchEndTime As Date
    '    Dim SearchSpotDate As Date
    '    Dim arrDayBreaks() As BrkPerfDataSet.ProgRFMtRow
    '    Dim arrDayBreaksNoBreak1() As BrkPerfDataSet.ProgRFMtRow
    '    Dim arrDayBreaksBreak1() As BrkPerfDataSet.ProgRFMtRow

    '    Dim RandomSpot As New Random
    '    Dim rowSpotsWritten As Integer
    '    Dim dtBrkDist As New Plandata.BrkDistributionDataTable

    '    For Each row As Excel.Range In ExcelPlan.loPlanData.DataBodyRange.Rows
    '        'Skip subtotal rows in excel if found
    '        If Not SubtotalRows Is Nothing Then
    '            If Not ExcelPlan.loPlanData.Application.Intersect(row, SubtotalRows) Is Nothing Then Continue For
    '        End If

    '        rowSpotsWritten = 0

    '        'Get values in current row of excel
    '        rowChannelName = DirectCast(row.Value, System.Object)(1, 1)
    '        rowProgram = DirectCast(row.Value, System.Object)(1, 2)
    '        rowDays = DirectCast(row.Value, System.Object)(1, 3)
    '        Dim oldDate As New Date(1900, 1, 1)
    '        Dim dblHour, dblMins As Double
    '        Dim tmStartTime, tmEndTime As Date
    '        Dim currTime As String

    '        'Adjust for excel date storage formats
    '        If TypeOf DirectCast(row.Value, System.Object)(1, 4) Is Double Then
    '            rowStartTime = DateTime.FromOADate(DirectCast(row.Value, System.Object)(1, 4))
    '        ElseIf TypeOf DirectCast(row.Value, System.Object)(1, 4) Is String Then
    '            currTime = DirectCast(row.Value, System.Object)(1, 4).ToString.Trim.PadLeft(5, "0")
    '            dblHour = currTime.Substring(0, 2)
    '            dblMins = currTime.Substring(3, 2)
    '            tmStartTime = DateAdd(DateInterval.Hour, dblHour, oldDate)
    '            tmStartTime = DateAdd(DateInterval.Minute, dblMins, tmStartTime)
    '            rowStartTime = tmStartTime
    '        End If

    '        If TypeOf DirectCast(row.Value, System.Object)(1, 5) Is Double Then
    '            rowEndTime = DateTime.FromOADate(DirectCast(row.Value, System.Object)(1, 5))
    '        ElseIf TypeOf DirectCast(row.Value, System.Object)(1, 5) Is String Then
    '            currTime = DirectCast(row.Value, System.Object)(1, 5).ToString.Trim.PadLeft(5, "0")
    '            dblHour = currTime.Substring(0, 2)
    '            dblMins = currTime.Substring(3, 2)
    '            tmEndTime = DateAdd(DateInterval.Hour, dblHour, oldDate)
    '            tmEndTime = DateAdd(DateInterval.Minute, dblMins, tmEndTime)
    '            rowEndTime = tmEndTime
    '        End If

    '        'Set endtime one day ahead in case endtime is less than start time
    '        If rowEndTime < rowStartTime Then rowEndTime = DateAdd(DateInterval.Day, 1, rowEndTime)

    '        'Get ChannelCode for current channel
    '        If currChannel <> rowChannelName Then
    '            dtPlanChannels = CType(logTaskPane.ucChannelsMapping.PlanChannelsBindingSource.DataSource, System.Data.DataView).Table
    '            arrPlanChannels = dtPlanChannels.Select("ChannelName = '" & rowChannelName & "'")
    '            If arrPlanChannels.Length > 0 Then
    '                ChannelCode = arrPlanChannels(0).ChannelCode
    '            End If
    '            currChannel = rowChannelName
    '        End If

    '        'Dim arrRowDays(), inRowDays As String
    '        'arrRowDays = rowDays.Split(",")
    '        'For i As Integer = 0 To arrRowDays.Length - 1
    '        '    arrRowDays(i) = "'" & arrRowDays(i) & "'"
    '        'Next


    '        'Get reformatted days in the "days" column of current row for filtering the brkperformance dataset using the "IN" clause
    '        Dim inRowDays As String = "'" & String.Join("','", rowDays.Split(",")) & "'"

    '        'Get the number of spots to be selected from brkperformance for current excel row
    '        rowSpots = DirectCast(row.Value, System.Object)(1, loPlanData.ListColumns("Total Spots").Index)


    '        TotalPlanSpots += rowSpots

    '        'If zero spots in current row then update status and skip row
    '        If rowSpots = 0 Then Continue For


    '        SearchStartTime = DateAdd(DateInterval.Minute, 5, rowStartTime) 'Add 5 minutes to row start time to accomodate for program starting late
    '        SearchEndTime = DateAdd(DateInterval.Minute, -2, rowEndTime) 'Reduce 2 minutes to the row end time to accomodate for program ending before time

    '        Dim minTime As New Date(SearchStartTime.Year, SearchStartTime.Month, SearchStartTime.Day, 7, 0, 0) 'Define min time to search for spots as 7am in case of RODPs
    '        Dim maxTime As New Date(SearchStartTime.Year, SearchStartTime.Month, SearchStartTime.Day, 23, 59, 59) 'Define max time to search for spots as 11:59pm in case of RODPs

    '        'Assume any time band more than 4 hours as RODP and adjust search start time / end time accordingly
    '        If DateDiff(DateInterval.Hour, SearchStartTime, SearchEndTime) > 4 Then
    '            If SearchStartTime < minTime Then SearchStartTime = minTime
    '            If SearchEndTime > maxTime Then SearchEndTime = maxTime
    '        End If

    '        'Get breaks for current row channel, days, search start time and search end time
    '        dtBrks = daBrks.GetBreaksForChannelAndDays(ChannelCode, inRowDays, SearchStartTime, SearchEndTime, logTaskPane.dtFromDate.Value, logTaskPane.dtToDate.Value)

    '        Dim strFilter As String
    '        Dim strFilterNoBreak1 As String
    '        Dim strFilterBreak1 As String

    '        'There are 3 possibilities:
    '        '1) brkperformance does not have any spots available for current row
    '        '2) No. of brks available are greater than required no. of spots
    '        '3) There aren't enough breaks available

    '        '1) brkperformance does not have any spots available for current row
    '        'If brkperformance does not have any spots available for current row, update status and skip row
    '        If dtBrks.Count = 0 Then Continue For

    '        '2) No. of brks available are greater than required no. of spots
    '        'If dtBrks.Count > rowSpots Then
    '        strFilterNoBreak1 = "CommercialName <> '---- End of Break 1 ----                '"
    '        strFilterBreak1 = "CommercialName = '---- End of Break 1 ----                '"
    '        Dim reqdBreaks, avblBreaks, noOfRepeats, avblEOB1, avblNonEOB1, reqdEOB1, reqdNonEOB1, avblDays, nonEOB1SpotsPerDay, remSpots, skipDays, reverseSkipDays As Integer

    '        reqdBreaks = rowSpots
    '        avblBreaks = dtBrks.Count
    '        noOfRepeats = IIf(avblBreaks < reqdBreaks, (reqdBreaks \ avblBreaks) - 1, 0)

    '        arrDayBreaksNoBreak1 = dtBrks.Select(strFilterNoBreak1)

    '        avblNonEOB1 = arrDayBreaksNoBreak1.Length
    '        avblEOB1 = avblBreaks - avblNonEOB1

    '        reqdNonEOB1 = IIf(avblNonEOB1 < reqdBreaks, avblNonEOB1, reqdBreaks)
    '        reqdEOB1 = IIf(reqdBreaks - reqdNonEOB1 > avblEOB1, avblEOB1, reqdBreaks - reqdNonEOB1)

    '        Dim distinctDays = From sBreaks In dtBrks _
    '                           Group sBreaks By BreakType = IIf(Trim(sBreaks.CommercialName) = Trim("---- End of Break 1 ----                "), "Break1", "Other Breaks") Into grpBreaks = Group _
    '                           Select Break = BreakType, TotalBrks = grpBreaks.Count, sDays = (From sDays In grpBreaks _
    '                           Group sDays By BrkDay = sDays._Date.ToString("ddd") Into grpDays = Group _
    '                           Select BrkDay = BrkDay, TotalBrks = grpDays.Count, sDates = (From sDate In grpDays Group sDate By sDate._Date Into grpDates = Group _
    '                                                                                        Select BrkDate = _Date, TotalBrks = grpDates.Count, Spots = grpDates))

    '        If distinctDays.Count > 0 Then

    '            Dim dsRow As New dsBPBrks ' DataSet("Breaks")
    '            Dim dtRowBreakTypes As dsBPBrks.BreakTypesDataTable = dsRow.BreakTypes 'As New dsBPBrks.BreakTypesDataTable 'DataTable = dsRow.Tables.Add("BreakTypes")
    '            Dim dtRowDayBreaks As dsBPBrks.DayBreaksDataTable = dsRow.DayBreaks 'As New dsBPBrks.DayBreaksDataTable 'DataTable = dsRow.Tables.Add("DayBreaks")
    '            Dim dtRowDateBreaks As dsBPBrks.DateBreaksDataTable = dsRow.DateBreaks ' As New dsBPBrks.DateBreaksDataTable 'DataTable = dsRow.Tables.Add("DateBreaks")
    '            Dim drRowBreakTypes As dsBPBrks.BreakTypesRow
    '            Dim drRowDayBreaks As dsBPBrks.DayBreaksRow
    '            Dim drRowDateBreaks As dsBPBrks.DateBreaksRow
    '            'Dim keysBrkType(1), keysDayBrks(2), keysDateBrks(3) As DataColumn
    '            'Dim dc As DataColumn

    '            'With dtRowBreakTypes
    '            '    dc = New DataColumn("BreakType", Type.GetType("System.String"))
    '            '    .Columns.Add(dc)
    '            '    keysBrkType(0) = dc
    '            '    dc = New DataColumn("TotalBrks", Type.GetType("System.Int16"))
    '            '    .Columns.Add(dc)
    '            '    .PrimaryKey = keysBrkType

    '            'End With
    '            'With dtRowDayBreaks
    '            '    dc = New DataColumn("BreakType", Type.GetType("System.String"))
    '            '    dc.Unique = False
    '            '    .Columns.Add(dc)
    '            '    keysDayBrks(0) = dc

    '            '    dc = New DataColumn("BrkDay", Type.GetType("System.String"))
    '            '    dc.Unique = False
    '            '    .Columns.Add(dc)
    '            '    keysDayBrks(1) = dc

    '            '    .Columns.Add("TotalBrks", Type.GetType("System.Int16"))
    '            '    .PrimaryKey = keysDayBrks

    '            'End With
    '            'With dtRowDateBreaks
    '            '    dc = New DataColumn("BreakType", Type.GetType("System.String"))
    '            '    dc.Unique = False
    '            '    .Columns.Add(dc)
    '            '    keysDateBrks(0) = dc

    '            '    dc = New DataColumn("BrkDay", Type.GetType("System.String"))
    '            '    dc.Unique = False
    '            '    .Columns.Add(dc)
    '            '    keysDateBrks(1) = dc

    '            '    dc = New DataColumn("BrkDate", Type.GetType("System.String"))
    '            '    dc.Unique = False
    '            '    .Columns.Add(dc)
    '            '    keysDateBrks(2) = dc

    '            '    .Columns.Add("TotalBrks", Type.GetType("System.Int16"))
    '            '    .Columns.Add("BreakSpots", distinctDays(0).sDays(0).sDates(0).Spots.GetType)
    '            '    .PrimaryKey = keysDateBrks
    '            'End With
    '            'Dim relationDayBreaks_DateBreaks As Global.System.Data.DataRelation
    '            'relationDayBreaks_DateBreaks = New Global.System.Data.DataRelation("DayBreaks_DateBreaks", New Global.System.Data.DataColumn() {dtRowDayBreaks.Columns("BreakType"), dtRowDayBreaks.Columns("BrkDay")}, New Global.System.Data.DataColumn() {dtRowDateBreaks.Columns("BreakType"), dtRowDateBreaks.Columns("BrkDay")}, False)
    '            'dsRow.Relations.Add("rel_days_type", dtRowBreakTypes.Columns("BreakType"), dtRowDayBreaks.Columns("BreakType"))
    '            'dsRow.Relations.Add(relationDayBreaks_DateBreaks)

    '            For Each dBreak In distinctDays
    '                drRowBreakTypes = dtRowBreakTypes.NewBreakTypesRow
    '                drRowBreakTypes("BreakType") = dBreak.Break
    '                drRowBreakTypes("TotalBrks") = dBreak.TotalBrks
    '                dtRowBreakTypes.Rows.Add(drRowBreakTypes)
    '                System.Diagnostics.Debug.Print(drRowBreakTypes("BreakType"))

    '                For Each dDays In dBreak.sDays
    '                    drRowDayBreaks = dtRowDayBreaks.NewRow()
    '                    drRowDayBreaks("BreakType") = dBreak.Break
    '                    drRowDayBreaks("BrkDay") = dDays.BrkDay
    '                    drRowDayBreaks("TotalBrks") = dDays.TotalBrks
    '                    dtRowDayBreaks.Rows.Add(drRowDayBreaks)
    '                    System.Diagnostics.Debug.Print(vbTab & drRowDayBreaks("BrkDay"))
    '                    For Each dDate In dDays.sDates
    '                        drRowDateBreaks = dtRowDateBreaks.NewRow()
    '                        drRowDateBreaks("BreakType") = dBreak.Break
    '                        drRowDateBreaks("BrkDay") = dDays.BrkDay
    '                        drRowDateBreaks("BrkDate") = dDate.BrkDate
    '                        drRowDateBreaks("TotalBrks") = dDate.TotalBrks
    '                        drRowDateBreaks("BreakSpots") = dDate.Spots
    '                        dtRowDateBreaks.Rows.Add(drRowDateBreaks)
    '                        System.Diagnostics.Debug.Print(vbTab & vbTab & drRowDateBreaks("BrkDate"))
    '                        For Each dSpot In dDate.Spots
    '                            System.Diagnostics.Debug.Print(vbTab & vbTab & vbTab & dSpot.FullString)
    '                        Next
    '                    Next
    '                Next
    '            Next

    '            Dim x = dsRow.BreakTypes.FindByBreakType("Other Breaks")

    '        End If
    '        'With dtRowBreaks
    '        '    .Columns.Add("BreakType", Type.GetType("System.String"))
    '        '    .Columns.Add("BrkDay", Type.GetType("System.String"))
    '        '    .Columns.Add("BrkDate", Type.GetType("System.String"))
    '        '    .Columns.Add("TotalBrks", Type.GetType("System.Int16"))
    '        'End With
    '        'Dim myData As System.Data.DataTable
    '        'myData = distinctDays.CopyToDataTable
    '        'For Each dr As DataRow In myData.Rows
    '        '    System.Diagnostics.Debug.Print(dr.Item("Break"))
    '        'Next

    '        System.Diagnostics.Debug.Print("The end")

    '        'For Each distinctDay In distinctDays
    '        '    System.Diagnostics.Debug.Print(distinctDay.BrkDay & " :: " & distinctDay.TotalBrks)
    '        '    For Each distinctDate In distinctDay
    '        '        System.Diagnostics.Debug.Print("\t" & distinctDate.BrkDate & " :: " & distinctDay.TotalBrks)
    '        '    Next
    '        'Next
    '        'Select BrkDay = BrkDay, TotalBrks = grouping1.Count
    '        'From grouping2 In _
    '        '(From sDay In grouping1 Group sDay By _Date Into grouping3 = Group)

    '        'avblDays = distinctDates.Count
    '        'nonEOB1SpotsPerDay = reqdNonEOB1 \ avblDays
    '        'remSpots = reqdNonEOB1 Mod avblDays
    '        'skipDays = Math.Floor(avblDays / IIf(remSpots = 0, avblDays, remSpots))
    '        'reverseSkipDays = IIf(remSpots > 0 And Math.Floor(avblDays / IIf(remSpots = 0, avblDays, remSpots)) = 1, Math.Floor(avblDays / (avblDays - remSpots)), 1)
    '        'For Each arrDate In distinctDates
    '        '    Dim drBrkDist As Plandata.BrkDistributionRow
    '        '    drBrkDist = dtBrkDist.NewBrkDistributionRow
    '        '    drBrkDist.Channel = ChannelCode
    '        '    drBrkDist.Days = rowDays
    '        '    drBrkDist.StartTime = rowStartTime
    '        '    drBrkDist.EndTime = rowEndTime
    '        '    drBrkDist.BrkDate = arrDate.BrkDate
    '        '    drBrkDist.AvblNonEOB1Brks = arrDate.TotalBreaks
    '        '    dtBrkDist.AllocatedNonEOB1BrksColumn.DefaultValue = 0

    '        '    drBrkDist.AllocatedNonEOB1Brks += nonEOB1SpotsPerDay
    '        '    dtBrkDist.AddBrkDistributionRow(drBrkDist)
    '        '    System.Diagnostics.Debug.Print(arrDate.BrkDate & " --" & arrDate.TotalBreaks)
    '        'Next
    '        'Dim i As Integer

    '        'Do

    '        '    i += 1
    '        'Loop While noOfRepeats >= i
    '        'Avoid "End of Break1" if possible


    '        'Select only those breaks which are not "End of Break1"

    '        'If arrDayBreaksNoBreak1.Length >= rowSpots Then
    '        '    Dim NoOfDistinctDays As Integer
    '        '    Dim distinctDates = (From sDate As BrkPerfDataSet.ProgRFMtRow In dtBrks Where sDate.CommercialName <> "---- End of Break 1 ----                " Select sDate._Date).Distinct()
    '        '    NoOfDistinctDays = distinctDates.Count
    '        '    Dim SkipDays As Integer = Math.Floor(NoOfDistinctDays / rowSpots)
    '        '    If SkipDays < 1 Then SkipDays = 1

    '        '    Dim MaxPerDaySpots, RemainingRowSpots, currDaySpots, pendingDaySpots As Integer
    '        '    MaxPerDaySpots = Math.DivRem(rowSpots, NoOfDistinctDays, RemainingRowSpots)
    '        '    For iDay As Integer = 0 To distinctDates.Count - 1 Step SkipDays
    '        '        SearchSpotDate = distinctDates(iDay)
    '        '        strFilter = "Date = #" & SearchSpotDate.Month & "-" & SearchSpotDate.Day & "-" & SearchSpotDate.Year & "# and IsSelected = False and CommercialName <> '---- End of Break 1 ----                '"
    '        '        arrDayBreaks = dtBrks.Select(strFilter)

    '        '        If RemainingRowSpots > 0 Then
    '        '            currDaySpots = MaxPerDaySpots + 1 + pendingDaySpots
    '        '            RemainingRowSpots += -1
    '        '        Else
    '        '            currDaySpots = MaxPerDaySpots + pendingDaySpots
    '        '        End If
    '        '        pendingDaySpots = 0
    '        '        For i As Integer = 1 To currDaySpots
    '        '            Try
    '        '                Dim RandomSpotPos As Integer = RandomSpot.Next(arrDayBreaks.Length)
    '        '                Dim currSpot As BrkPerfDataSet.ProgRFMtRow = arrDayBreaks(RandomSpotPos)
    '        '                AddSpotRow(currSpot)
    '        '                currSpot.IsSelected = True
    '        '                rowSpotsWritten += 1
    '        '                TotalSpotsWritten += 1
    '        '            Catch ex As Exception
    '        '                pendingDaySpots += 1
    '        '            End Try
    '        '            frmProgress.lblPercentage.Text = rowChannelName & ": " & rowProgram & vbCrLf & rowSpotsWritten & " of " & rowSpots & " spots"
    '        '            Application.DoEvents()
    '        '            arrDayBreaks = dtBrks.Select(strFilter)
    '        '        Next
    '        '    Next
    '        '    For i As Integer = 1 To pendingDaySpots
    '        '        strFilter = "IsSelected = False and CommercialName <> '---- End of Break 1 ----                '"
    '        '        arrDayBreaks = dtBrks.Select(strFilter)
    '        '        Dim RandomSpotPos As Integer = RandomSpot.Next(arrDayBreaks.Length)
    '        '        Dim currSpot As BrkPerfDataSet.ProgRFMtRow = arrDayBreaks(RandomSpotPos)
    '        '        AddSpotRow(currSpot)
    '        '        currSpot.IsSelected = True
    '        '        rowSpotsWritten += 1
    '        '        TotalSpotsWritten += 1
    '        '        frmProgress.lblPercentage.Text = rowChannelName & ": " & rowProgram & vbCrLf & rowSpotsWritten & " of " & rowSpots & " spots"
    '        '        Application.DoEvents()
    '        '    Next
    '        '    pendingDaySpots = 0
    '        'Else
    '        '    For Each spotRow As BrkPerfDataSet.ProgRFMtRow In arrDayBreaksNoBreak1
    '        '        AddSpotRow(spotRow)
    '        '        spotRow.IsSelected = True
    '        '        rowSpotsWritten += 1
    '        '        TotalSpotsWritten += 1
    '        '        frmProgress.lblPercentage.Text = rowChannelName & ": " & rowProgram & vbCrLf & rowSpotsWritten & " of " & rowSpots & " spots"
    '        '        Application.DoEvents()
    '        '    Next

    '        '    arrDayBreaksBreak1 = dtBrks.Select(strFilterBreak1)
    '        '    Dim NoOfDistinctDays As Integer
    '        '    Dim distinctDates = (From sDate As BrkPerfDataSet.ProgRFMtRow In dtBrks Where sDate.CommercialName = "---- End of Break 1 ----                " Select sDate._Date).Distinct()
    '        '    NoOfDistinctDays = distinctDates.Count
    '        '    Dim SkipDays As Integer = Math.Floor(NoOfDistinctDays / rowSpots)
    '        '    If SkipDays < 1 Then SkipDays = 1

    '        '    Dim MaxPerDaySpots, RemainingRowSpots, currDaySpots, pendingDaySpots As Integer
    '        '    MaxPerDaySpots = Math.DivRem(rowSpots - arrDayBreaksNoBreak1.Length, NoOfDistinctDays, RemainingRowSpots)

    '        '    For iDay As Integer = 0 To distinctDates.Count - 1 Step SkipDays
    '        '        SearchSpotDate = distinctDates(iDay)

    '        '        If RemainingRowSpots > 0 Then
    '        '            currDaySpots = MaxPerDaySpots + 1 + pendingDaySpots
    '        '            RemainingRowSpots += -1
    '        '        Else
    '        '            currDaySpots = MaxPerDaySpots + pendingDaySpots
    '        '        End If
    '        '        pendingDaySpots = 0

    '        '        strFilter = "Date = #" & SearchSpotDate.Month & "-" & SearchSpotDate.Day & "-" & SearchSpotDate.Year & "# and IsSelected = False and CommercialName = '---- End of Break 1 ----                '"
    '        '        arrDayBreaks = dtBrks.Select(strFilter)
    '        '        For i As Integer = 1 To currDaySpots
    '        '            Try
    '        '                Dim RandomSpotPos As Integer = RandomSpot.Next(arrDayBreaks.Length)
    '        '                Dim currSpot As BrkPerfDataSet.ProgRFMtRow = arrDayBreaks(RandomSpotPos)
    '        '                AddSpotRow(currSpot)
    '        '                currSpot.IsSelected = True
    '        '                rowSpotsWritten += 1
    '        '                TotalSpotsWritten += 1
    '        '            Catch ex As Exception
    '        '                pendingDaySpots += 1
    '        '            End Try

    '        '            frmProgress.lblPercentage.Text = rowChannelName & ": " & rowProgram & vbCrLf & rowSpotsWritten & " of " & rowSpots & " spots"
    '        '            Application.DoEvents()
    '        '            If rowSpotsWritten = rowSpots Then Exit For
    '        '            arrDayBreaks = dtBrks.Select(strFilter)
    '        '        Next
    '        '    Next
    '        '    For i As Integer = 1 To pendingDaySpots
    '        '        strFilter = "IsSelected = False and CommercialName = '---- End of Break 1 ----                '"
    '        '        arrDayBreaks = dtBrks.Select(strFilter)
    '        '        Dim RandomSpotPos As Integer = RandomSpot.Next(arrDayBreaks.Length)
    '        '        Dim currSpot As BrkPerfDataSet.ProgRFMtRow = arrDayBreaks(RandomSpotPos)
    '        '        AddSpotRow(currSpot)
    '        '        currSpot.IsSelected = True
    '        '        rowSpotsWritten += 1
    '        '        TotalSpotsWritten += 1
    '        '        frmProgress.lblPercentage.Text = rowChannelName & ": " & rowProgram & vbCrLf & rowSpotsWritten & " of " & rowSpots & " spots"
    '        '        Application.DoEvents()
    '        '    Next
    '        '    pendingDaySpots = 0

    '        'End If
    '        'End If
    '        '3) There aren't enough breaks available
    '        '   Repeat all spots available
    '        If dtBrks.Count <= rowSpots Then
    '            'Do While rowSpotsWritten < rowSpots
    '            '    For Each spotRow As BrkPerfDataSet.ProgRFMtRow In dtBrks.Rows
    '            '        AddSpotRow(spotRow)
    '            '        spotRow.IsSelected = True
    '            '        rowSpotsWritten += 1
    '            '        TotalSpotsWritten += 1
    '            '        frmProgress.lblPercentage.Text = rowChannelName & ": " & rowProgram & vbCrLf & rowSpotsWritten & " of " & rowSpots & " spots"
    '            '        Application.DoEvents()
    '            '        If rowSpotsWritten = rowSpots Then Exit For
    '            '    Next
    '            'Loop
    '        End If


    '        System.Diagnostics.Debug.Print("Spots Added:" & rowSpotsWritten & "~" & rowSpots & "~" & ChannelCode & "~" & rowChannelName & "~" & dtBrks.Count & "~" & inRowDays & "~" & rowStartTime.ToString() & "~" & rowEndTime.ToString())
    '    Next
    'End Sub
End Module
Public Class ObjectShredder(Of T)
    ' Fields
    Private _fi As FieldInfo()
    Private _ordinalMap As Dictionary(Of String, Integer)
    Private _pi As PropertyInfo()
    Private _type As Type

    ' Constructor 
    Public Sub New()
        Me._type = GetType(T)
        Me._fi = Me._type.GetFields
        Me._pi = Me._type.GetProperties
        Me._ordinalMap = New Dictionary(Of String, Integer)
    End Sub

    Public Function ShredObject(ByVal table As DataTable, ByVal instance As T) As Object()
        Dim fi As FieldInfo() = Me._fi
        Dim pi As PropertyInfo() = Me._pi
        If (Not instance.GetType Is GetType(T)) Then
            ' If the instance is derived from T, extend the table schema
            ' and get the properties and fields.
            Me.ExtendTable(table, instance.GetType)
            fi = instance.GetType.GetFields
            pi = instance.GetType.GetProperties
        End If

        ' Add the property and field values of the instance to an array.
        Dim values As Object() = New Object(table.Columns.Count - 1) {}
        Dim f As FieldInfo
        For Each f In fi
            values(Me._ordinalMap.Item(f.Name)) = f.GetValue(instance)
        Next
        Dim p As PropertyInfo
        For Each p In pi
            values(Me._ordinalMap.Item(p.Name)) = p.GetValue(instance, Nothing)
        Next

        ' Return the property and field values of the instance.
        Return values
    End Function


    ' Summary:           Loads a DataTable from a sequence of objects.
    ' source parameter:  The sequence of objects to load into the DataTable.</param>
    ' table parameter:   The input table. The schema of the table must match that 
    '                    the type T.  If the table is null, a new table is created  
    '                    with a schema created from the public properties and fields 
    '                    of the type T.
    ' options parameter: Specifies how values from the source sequence will be applied to 
    '                    existing rows in the table.
    ' Returns:           A DataTable created from the source sequence.

    Public Function Shred(ByVal source As IEnumerable(Of T), ByVal table As DataTable, ByVal options As LoadOption?) As DataTable

        ' Load the table from the scalar sequence if T is a primitive type.
        If GetType(T).IsPrimitive Then
            Return Me.ShredPrimitive(source, table, options)
        End If

        ' Create a new table if the input table is null.
        If (table Is Nothing) Then
            table = New DataTable(GetType(T).Name)
        End If

        ' Initialize the ordinal map and extend the table schema based on type T.
        table = Me.ExtendTable(table, GetType(T))

        ' Enumerate the source sequence and load the object values into rows.
        table.BeginLoadData()
        Using e As IEnumerator(Of T) = source.GetEnumerator
            Do While e.MoveNext
                If options.HasValue Then
                    table.LoadDataRow(Me.ShredObject(table, e.Current), options.Value)
                Else
                    table.LoadDataRow(Me.ShredObject(table, e.Current), True)
                End If
            Loop
        End Using
        table.EndLoadData()

        ' Return the table.
        Return table
    End Function


    Public Function ShredPrimitive(ByVal source As IEnumerable(Of T), ByVal table As DataTable, ByVal options As LoadOption?) As DataTable
        ' Create a new table if the input table is null.
        If (table Is Nothing) Then
            table = New DataTable(GetType(T).Name)
        End If
        If Not table.Columns.Contains("Value") Then
            table.Columns.Add("Value", GetType(T))
        End If

        ' Enumerate the source sequence and load the scalar values into rows.
        table.BeginLoadData()
        Using e As IEnumerator(Of T) = source.GetEnumerator
            Dim values As Object() = New Object(table.Columns.Count - 1) {}
            Do While e.MoveNext
                values(table.Columns.Item("Value").Ordinal) = e.Current
                If options.HasValue Then
                    table.LoadDataRow(values, options.Value)
                Else
                    table.LoadDataRow(values, True)
                End If
            Loop
        End Using
        table.EndLoadData()

        ' Return the table.
        Return table
    End Function

    Public Function ExtendTable(ByVal table As DataTable, ByVal type As Type) As DataTable
        ' Extend the table schema if the input table was null or if the value 
        ' in the sequence is derived from type T.
        Dim f As FieldInfo
        Dim p As PropertyInfo

        For Each f In type.GetFields
            If Not Me._ordinalMap.ContainsKey(f.Name) Then
                Dim dc As DataColumn

                ' Add the field as a column in the table if it doesn't exist
                ' already.
                dc = IIf(table.Columns.Contains(f.Name), table.Columns.Item(f.Name), table.Columns.Add(f.Name, f.FieldType))

                ' Add the field to the ordinal map.
                Me._ordinalMap.Add(f.Name, dc.Ordinal)
            End If

        Next

        For Each p In type.GetProperties
            If Not Me._ordinalMap.ContainsKey(p.Name) Then
                ' Add the property as a column in the table if it doesn't exist
                ' already.
                Dim dc As DataColumn
                dc = IIf(table.Columns.Contains(p.Name), table.Columns.Item(p.Name), table.Columns.Add(p.Name, p.PropertyType))

                ' Add the property to the ordinal map.
                Me._ordinalMap.Add(p.Name, dc.Ordinal)
            End If
        Next

        ' Return the table.
        Return table
    End Function

End Class


Public Module CustomLINQtoDataSetMethods
    <Extension()> _
    Public Function CopyToDataTable(Of T)(ByVal source As IEnumerable(Of T)) As DataTable
        Return New ObjectShredder(Of T)().Shred(source, Nothing, Nothing)
    End Function

    <Extension()> _
    Public Function CopyToDataTable(Of T)(ByVal source As IEnumerable(Of T), ByVal table As DataTable, ByVal options As LoadOption?) As DataTable
        Return New ObjectShredder(Of T)().Shred(source, table, options)
    End Function

End Module