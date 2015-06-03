Imports System.Globalization
Imports System.Windows.Forms
Imports System.Data.SqlServerCe
Imports System.Timers
Imports System.Data.SqlClient
Imports System.Data
Public Class frmPrepareServer
    Dim ciCurr As CultureInfo = CultureInfo.CurrentCulture
    Friend dtWeeks As New Plandata.WeeksDataTable
    Friend showingErrors As Boolean = False
    Friend showingChannels As Boolean = False
    Dim isSetToDate As Boolean = False
    Private Shared isWeekNoChanged As Boolean = False
    Delegate Sub SetTextCallback(ByVal source As Object, ByVal e As ElapsedEventArgs)
    Public Enum CreateTempTableStatus
        Completed
        InProgress
    End Enum
    'Private Sub txtWeeks_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWeeks.ValueChanged
    '    Try

    '    Catch ex As Exception

    '    End Try
    'End Sub

    Private Sub frmPrepareServer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'End If
            dtFromDate.Value = Globals.Ribbons.MSprintExRibbon.db_ToDate.AddDays(-Globals.Ribbons.MSprintExRibbon.db_ToDate.DayOfWeek)
            ' dtFromDate.Value = Globals.Ribbons.MSprintExRibbon.db_FromDate
            'dtToDate.Value = Globals.Ribbons.MSprintExRibbon.db_ToDate

            '  lbDbText.Visible = False
            dtToDate.Value = DateAdd(DateInterval.Day, -1, ciCurr.Calendar.AddWeeks(dtFromDate.Value, txtWeeks.Value))
            isLoaded = True
            Label3.Text = "Data in Database is available only from " + Globals.Ribbons.MSprintExRibbon.db_FromDate.ToString("dd/MM/yyyy") + " to " + Globals.Ribbons.MSprintExRibbon.db_ToDate.ToString("dd/MM/yyyy")
            setWeekGrid() ''Added By Alok for Date picker
            txtWeeks.Value = dgvWeeks.Rows.Count ''Added By Alok for Date picker
            setToDate()
        Catch ex As Exception
            LogMpsrintExException("Exception occured while loading PrepareServer form.Message :" + ex.Message)
        End Try
    End Sub

    'Private Sub btnPrepare_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrepare.Click
    '    Try

    '    Catch ex As Exception

    '    End Try
    'End Sub

    'Private Sub dtFromDate_CloseUp(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtFromDate.CloseUp
    '    Try

    '    Catch ex As Exception

    '    End Try
    'End Sub

    'Private Sub dtFromDate_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtFromDate.Validating
    '    Try

    '    Catch ex As Exception

    '    End Try
    'End Sub

    'Private Sub dtToDate_CloseUp(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtToDate.CloseUp
    '    Try

    '    Catch ex As Exception

    '    End Try
    'End Sub

    'Private Sub dtToDate_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtToDate.Validating
    '    Try

    '    Catch ex As Exception

    '    End Try
    'End Sub
    Private Function isWeekStartDate(ByVal currDate As Date) As Boolean
        isWeekStartDate = ciCurr.Calendar.GetDayOfWeek(currDate) = DayOfWeek.Sunday
    End Function

    Private Function getWeekStartDate(ByVal currDate As Date) As Date
        getWeekStartDate = ciCurr.Calendar.AddDays(currDate, ciCurr.Calendar.GetDayOfWeek(currDate) * -1)
    End Function

    Private Function getWeekEndDate(ByVal currDate As Date) As Date
        getWeekEndDate = ciCurr.Calendar.AddDays(currDate, 6 - ciCurr.Calendar.GetDayOfWeek(currDate))

    End Function

    Private Sub setToDate()
        If (isWeekNoChanged) Then
            isSetToDate = True
            ' dtToDate.Value = ciCurr.Calendar.AddWeeks(dtFromDate.Value, -(txtWeeks.Value - 1))
            '  dtFromDate.Value = ciCurr.Calendar.AddWeeks(dtToDate.Value, -(txtWeeks.Value))
            ' dtToDate.Value = getWeekEndDate(dtToDate.Value)
            dtFromDate.Value = ciCurr.Calendar.AddDays(ciCurr.Calendar.AddWeeks(dtToDate.Value, -txtWeeks.Value), 1)
            'dtToDate.Value = getWeekEndDate(dtToDate.Value)
            '  dtFromDate.Value = getWeekStartDate(dtFromDate.Value)
            isSetToDate = False
            dtWeeks.Clear()
            dgvWeeks.DataSource = Nothing
            setWeekGrid()
        End If

    End Sub
    Private Sub setToDate1()
        isSetToDate = True
        dtToDate.Value = ciCurr.Calendar.AddWeeks(dtFromDate.Value, txtWeeks.Value - 1)
        '  dtFromDate.Value = ciCurr.Calendar.AddWeeks(dtToDate.Value, -(txtWeeks.Value))
        ' dtToDate.Value = getWeekEndDate(dtToDate.Value)
        '   dtFromDate.Value = ciCurr.Calendar.AddWeeks(dtFromDate.Value, -(txtWeeks.Value - 1))
        dtToDate.Value = getWeekEndDate(dtToDate.Value)
        '  dtFromDate.Value = getWeekStartDate(dtFromDate.Value)
        isSetToDate = False
        setWeekGrid()
    End Sub

    Private Sub setWeekGrid()
        Try
            ' If isLoaded Then
            dtWeeks.Clear()
            Dim date_to As Date = createWeeks(dtFromDate.Value, dtToDate.Value)

            Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

            If sheet.Name.Equals("Plan Selection") Then
                CreateWeekColumns()
            End If
            dgvWeeks.DataSource = dtWeeks
            ' dtToDate.Value = date_to
            'mcExclude.MinDate = dtFromDate.Value
            'mcExclude.MaxDate = dtToDate.Value
            'Dim ts As TimeSpan
            'ts = dtToDate.Value.Subtract(dtFromDate.Value)
            'If ts.Days > 0 Then mcExclude.MaxSelectionCount = ts.Days
            '  End If
        Catch ex As Exception

        End Try
    End Sub
    Private Function createWeeks(ByVal weekStartDate As Date, ByVal lastDate As Date) As Date
        '   lbDbText.Visible = False
        Dim drWeek As Plandata.WeeksRow
        drWeek = dtWeeks.NewWeeksRow
        If ciCurr.Calendar.GetYear(weekStartDate) < ciCurr.Calendar.GetYear(getWeekEndDate(weekStartDate)) Then
            drWeek.WeekNumber = 1
            drWeek.Year = ciCurr.Calendar.GetYear(weekStartDate) + 1
        Else
            drWeek.WeekNumber = ciCurr.Calendar.GetWeekOfYear(weekStartDate, CalendarWeekRule.FirstDay, DayOfWeek.Sunday)
            drWeek.Year = ciCurr.Calendar.GetYear(weekStartDate)
        End If
        drWeek.StartDate = New Date(weekStartDate.Year, weekStartDate.Month, weekStartDate.Day) 'CDate(weekStartDate.ToString("dd/MM/yyyy 00:00:00"))
        ' dtWeeks.AddWeeksRow(drWeek)
        Dim nextWeekStartDate As Date = ciCurr.Calendar.AddDays(getWeekEndDate(weekStartDate), 1)
        If nextWeekStartDate > lastDate Then
            drWeek.EndDate = New Date(lastDate.Year, lastDate.Month, lastDate.Day) 'CDate(lastDate.ToString("dd/MM/yyyy 23:59:59"))
            ' If Not (drWeek.StartDate.Equals(drWeek.EndDate)) Then
            dtWeeks.AddWeeksRow(drWeek)

            'If isLoaded Then
            '    dtToDate.Value = drWeek.EndDate
            'End If


            'End If
            createWeeks = lastDate
        Else
            Dim dd As Date = getWeekEndDate(weekStartDate)

            'If nextWeekStartDate.DayOfWeek = DayOfWeek.Sunday Then
            '    dd = nextWeekStartDate
            'End If

            drWeek.EndDate = New Date(dd.Year, dd.Month, dd.Day) 'CDate(getWeekEndDate(weekStartDate).ToString("dd/MM/yyyy 23:59:59"))

            ' If Not (drWeek.StartDate.Equals(drWeek.EndDate)) Then
            dtWeeks.AddWeeksRow(drWeek)

            'If isLoaded Then
            '    dtToDate.Value = drWeek.EndDate
            'End If


            'End If
            createWeeks = createWeeks(nextWeekStartDate, lastDate)
        End If
    End Function
    Private Function AddWeekColumn(ByVal WeekNumber As Integer) As Data.DataColumn
        Dim col As Data.DataColumn
        Try
            Dim loSpotSelection As Microsoft.Office.Tools.Excel.ListObject = GetSpotSelecListObject()
            Dim listObjectdt As Data.DataTable = CType(loSpotSelection.DataSource, Data.DataTable)

            '  If Not (loSpotSelection Is Nothing) Then
            'AddWeekColumn = loSpotSelection.ListColumns.Add()
            'AddWeekColumn.Name = "Week " & WeekNumber
            col = listObjectdt.Columns.Add()
            col.ColumnName = "Week " & WeekNumber
            col.DataType = Type.GetType("System.Int32")
            loSpotSelection.SetDataBinding(listObjectdt)


        Catch ex As Exception

        End Try
        ' End If
        '  loSpotSelection.Refresh()
        Return col
    End Function

    Private Sub CreateWeekColumns()
        '  lbDbText.Visible = False
        Dim loSpotSelection As Microsoft.Office.Tools.Excel.ListObject = GetSpotSelecListObject()
        Dim listObjectDt As Data.DataTable = CType(loSpotSelection.DataSource, Data.DataTable)
        ' If rbSingle.Checked Then
        Dim TotalWeeks As Integer = dtWeeks.Count
        Dim WeekNumber As Integer
        If Not loSpotSelection Is Nothing Then

            Do While WeekColumns.Count > 0
                Dim kp As KeyValuePair(Of WeekYear, Data.DataColumn) = WeekColumns.Last()
                '  kp.Value..Value.Delete()
                listObjectDt.Columns.Remove(WeekColumns(kp.Key))
                loSpotSelection.SetDataBinding(listObjectDt)
                WeekColumns.Remove(kp.Key)

            Loop
            If Not (listObjectDt.Columns.Contains("Total Spots")) Then
                Dim TotalSpots As Data.DataColumn = listObjectDt.Columns.Add("Total Spots")
                loSpotSelection.SetDataBinding(listObjectDt)
            End If
            'Catch ex As Exception

            '  TotalSpots.Name = "Total Spots"
            ' End Try
        End If
        ' Else
        'Dim TotalWeeks As Integer = dtWeeks.Count
        'Dim WeekNumber As Integer
        'Try
        '    listObjectDt.Columns.Remove("Total Spots")
        '    loSpotSelection.SetDataBinding(listObjectDt)
        'Catch ex As Exception
        'End Try
        'Dim DelWeeks As New Dictionary(Of WeekYear, Data.DataColumn)
        'For Each kp As KeyValuePair(Of WeekYear, Data.DataColumn) In WeekColumns
        '    If dtWeeks.FindByWeekNumberYear(kp.Key.WeekNumber, kp.Key.WeekYear) Is Nothing Then
        '        DelWeeks.Add(kp.Key, kp.Value)
        '    End If
        'Next
        'For Each wk As KeyValuePair(Of WeekYear, Data.DataColumn) In DelWeeks

        '    'If Not wk Is Nothing Then
        '    ' wk.Value.Delete()
        '    listObjectDt.Columns.Remove(WeekColumns(wk.Key))
        '    loSpotSelection.SetDataBinding(listObjectDt)
        '    WeekColumns.Remove(wk.Key)

        '    '  End If

        'Next
        'For i As Integer = 1 To TotalWeeks
        '    Dim CurrWeek As New WeekYear(dtWeeks(i - 1).WeekNumber, dtWeeks(i - 1).Year)
        '    If Not WeekColumns.ContainsKey(CurrWeek) Then
        '        WeekColumns.Add(CurrWeek, AddWeekColumn(CurrWeek.WeekNumber))
        '    End If
        'Next
        '  End If
    End Sub

    Private Sub txtWeeks_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWeeks.ValueChanged
        setToDate()
    End Sub
    Private Sub dtFromDate_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtFromDate.Validating
        If CType(sender, System.Windows.Forms.DateTimePicker).Value < Globals.Ribbons.MSprintExRibbon.db_FromDate Then
            System.Windows.Forms.MessageBox.Show("Invalid date chosen.Please choose dates between " + Globals.Ribbons.MSprintExRibbon.db_FromDate.ToString("dd/MM/yyyy") + " and " + Globals.Ribbons.MSprintExRibbon.db_ToDate.ToString("dd/MM/yyyy"))
            'dtWeeks.Clear()
            'dgvWeeks.DataSource = Nothing
            'dtFromDate.Value = Globals.Ribbons.MSprintExRibbon.db_FromDate
            'dtToDate.Value = Globals.Ribbons.MSprintExRibbon.db_ToDate
            e.Cancel = True
            'Else


        End If
    End Sub

    Private Sub dtFromDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'If dtFromDate.Value < Globals.Ribbons.MSprintExRibbon.db_FromDate Then
        '    ' System.Windows.Forms.MessageBox.Show("Invalid date chosen.Please choose dates between " + Globals.Ribbons.MSprintExRibbon.db_FromDate + " and " + Globals.Ribbons.MSprintExRibbon.db_ToDate)

        'Else
        '   lbDbText.Visible = False
        '    setToDate()
        ' End If

        'If Not (dtFromDate.Value.Equals(TempTableStatus.from_date)) Then
        '    TempTableStatus.from_date = dtFromDate.Value
        'End If

    End Sub

    Private Sub dtToDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtToDate.Validating
        ' lbDbText.Visible = False
        If CType(sender, System.Windows.Forms.DateTimePicker).Value < dtFromDate.Value Then
            MsgBox("End date cannot be less than start date.  Please correct the entry", MsgBoxStyle.Exclamation, "Resolve date")
            e.Cancel = True
        End If

        If CType(sender, System.Windows.Forms.DateTimePicker).Value > Globals.Ribbons.MSprintExRibbon.db_ToDate Then
            System.Windows.Forms.MessageBox.Show("Invalid date chosen.Please choose dates between " + Globals.Ribbons.MSprintExRibbon.db_FromDate.ToString("dd/MM/yyyy") + " and " + Globals.Ribbons.MSprintExRibbon.db_ToDate.ToString("dd/MM/yyyy"))
            'dtWeeks.Clear()
            'dgvWeeks.DataSource = Nothing
            'dtFromDate.Value = Globals.Ribbons.MSprintExRibbon.db_FromDate
            'dtToDate.Value = Globals.Ribbons.MSprintExRibbon.db_ToDate
            e.Cancel = True
            ' lbDbText.Visible = True

            '   e.Cancel = True
        End If


    End Sub

    Private Sub dtToDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '   lbDbText.Visible = False
        If CType(sender, System.Windows.Forms.DateTimePicker).Value < dtFromDate.Value Then
            MsgBox("End date cannot be less than start date.  Please correct the entry", MsgBoxStyle.Exclamation, "Resolve date")
            Exit Sub
        End If
        If Not isSetToDate Then

            'If dtToDate.Value > Globals.Ribbons.MSprintExRibbon.db_ToDate Then
            '    System.Windows.Forms.MessageBox.Show("Invalid date chosen.Please choose dates between " + Globals.Ribbons.MSprintExRibbon.db_FromDate + " and " + Globals.Ribbons.MSprintExRibbon.db_ToDate)
            'Else
            setWeekGrid()
            ' End If

        End If

    End Sub
    Private Sub dtFromDate_CloseUp(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtFromDate.CloseUp
        isWeekNoChanged = False
        setWeekGrid()
        txtWeeks.Value = dgvWeeks.Rows.Count
        isWeekNoChanged = True
    End Sub

    Private Sub dtToDate_CloseUp(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtToDate.CloseUp
        If CType(sender, System.Windows.Forms.DateTimePicker).Value < dtFromDate.Value Then
            MsgBox("End date cannot be less than start date.  Please correct the entry", MsgBoxStyle.Exclamation, "Resolve date")
            Exit Sub
        End If
        If Not isSetToDate Then
            If dtToDate.Value > Globals.Ribbons.MSprintExRibbon.db_ToDate Then
                System.Windows.Forms.MessageBox.Show("Invalid date chosen.Please choose dates between " + Globals.Ribbons.MSprintExRibbon.db_FromDate + " and " + Globals.Ribbons.MSprintExRibbon.db_ToDate)
            Else
                isWeekNoChanged = False
                setWeekGrid()
                txtWeeks.Value = dgvWeeks.Rows.Count
                isWeekNoChanged = True
            End If
        End If
    End Sub

    'Added By Alok for date selection from Calender, This will fire after user closes the calaender--End
    Public Function TempTableIsCreatedAndCompleted(ByVal from_Date As Date, ByVal to_Date As Date) As Boolean
        Dim created As Boolean = False
        Dim opXML As XElement = New XElement("mediaplan")
        Try
            opXML = GetTempTableList(Globals.Ribbons.MSprintExRibbon.GetURLForWS("GetTempTableList"), "GET")
            For Each metaTableEntry As XElement In opXML.Elements
                Dim from_Date_XML As Date = Convert.ToDateTime(metaTableEntry.Element("StartDate").Value)
                Dim to_Date_XML As Date = Convert.ToDateTime(metaTableEntry.Element("EndDate").Value)
                Dim status_XML As String = Convert.ToString(metaTableEntry.Element("status").Value)
                'And status_XML.Trim().ToLower().Equals("completed")
                If from_Date_XML.Equals(from_Date) And to_Date_XML.Equals(to_Date) And status_XML.Trim().ToLower().Equals("completed") Then
                    created = True
                    Exit For
                Else
                    created = False
                End If
            Next
        Catch ex As Exception
            LogMpsrintExException(String.Format("Exception occured while checking if temp table is created for dates: FromDate:{0};ToDate:{1}.Message:{2}", from_Date, to_Date, ex.Message))
            Throw ex
        End Try
        Return created
    End Function
    Public Function TempTableIsCreated(ByVal from_Date As Date, ByVal to_Date As Date) As TempTableStatus
        Dim ttsObject As TempTableStatus = New TempTableStatus()

        Dim opXML As XElement = New XElement("mediaplan")
        Try
            opXML = GetTempTableList(Globals.Ribbons.MSprintExRibbon.GetURLForWS("GetTempTableList"), "GET")
            For Each metaTableEntry As XElement In opXML.Elements
                Dim from_Date_XML As Date = Convert.ToDateTime(metaTableEntry.Element("StartDate").Value)
                Dim to_Date_XML As Date = Convert.ToDateTime(metaTableEntry.Element("EndDate").Value)
                Dim status_XML As String = Convert.ToString(metaTableEntry.Element("status").Value)
                'And status_XML.Trim().ToLower().Equals("completed")
                If from_Date >= from_Date_XML And to_Date <= to_Date_XML Then

                    If status_XML.Trim().ToLower().Equals("completed") Then
                        ttsObject.initiated = True
                        ttsObject.status = CreateTempTableStatus.Completed
                    Else
                        ttsObject.initiated = True
                        ttsObject.status = CreateTempTableStatus.InProgress
                    End If

                    '  Created = True
                    Exit For
                Else
                    ttsObject.initiated = False
                End If
            Next
        Catch ex As Exception
            LogMpsrintExException(String.Format("Exception occured while checking if temp table is created for dates: FromDate:{0};ToDate:{1}.Message:{2}", from_Date, to_Date, ex.Message))
            Throw ex
        End Try
        Return ttsObject
    End Function
    Private Sub btnPrepare_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrepare.Click
        Dim month, month1 As String
        Dim day, day1 As String
        Try
            If Not (TempTableIsCreated(dtFromDate.Value, dtToDate.Value).initiated) Then
                Globals.Ribbons.MSprintExRibbon.EnableDisableButtons(False)

                MessageBox.Show("MsprintX is getting the server ready for chosen date range and faster processing.This would take some time.Please await for email notification to restart using MsprintX .")
                Globals.ThisAddIn.Application.StatusBar = "Preparing MsprintX server..."
                If dtFromDate.Value.Month < 10 Then
                    month = "0" + dtFromDate.Value.Month.ToString()
                Else
                    month = dtFromDate.Value.Month.ToString()
                End If
                If dtFromDate.Value.Day < 10 Then
                    day = "0" + dtFromDate.Value.Day.ToString()
                Else
                    day = dtFromDate.Value.Day.ToString()
                End If
                If dtToDate.Value.Month < 10 Then
                    month1 = "0" + dtToDate.Value.Month.ToString()
                Else
                    month1 = dtToDate.Value.Month.ToString()
                End If
                If dtToDate.Value.Day < 10 Then
                    day1 = "0" + dtToDate.Value.Day.ToString()
                Else
                    day1 = dtToDate.Value.Day.ToString()
                End If
                Dim mediaplan As XElement = <mediaplan>
                                                <PreEvalPeriod>
                                                    <StartDate><%= dtFromDate.Value.Year.ToString() + month + day %></StartDate>
                                                    <EndDate><%= dtToDate.Value.Year.ToString() + month1 + day1 %></EndDate>
                                                </PreEvalPeriod>
                                            </mediaplan>
                Dim name As String = Date.Now.ToString("ddMMyyyy") & "_" & Date.Now.ToString("HHmmss") + ".xml"
                mediaplan.Save(Globals.Ribbons.MSprintExRibbon.LogDirectoryPath + "CreateTempTable_Inp_" + name)
                CreateTempTable(mediaplan, Globals.Ribbons.MSprintExRibbon.GetURLForWS("CreateTempTable_New"))
                Globals.Ribbons.MSprintExRibbon.TempTableTimer = New System.Timers.Timer(600000)
                AddHandler Globals.Ribbons.MSprintExRibbon.TempTableTimer.Elapsed, AddressOf OnTimedEvent
                Globals.Ribbons.MSprintExRibbon.TempTableTimer.Enabled = True
            ElseIf TempTableIsCreated(dtFromDate.Value, dtToDate.Value).initiated And TempTableIsCreated(dtFromDate.Value, dtToDate.Value).status = CreateTempTableStatus.InProgress Then
                Globals.Ribbons.MSprintExRibbon.EnableDisableButtons(False)

                MessageBox.Show("MsprintX is getting the server ready for chosen date range and faster processing.This would take some time.Please await for email notification to restart using MsprintX .")
                Globals.ThisAddIn.Application.StatusBar = "Preparing MsprintX server..."
                Globals.Ribbons.MSprintExRibbon.TempTableTimer = New System.Timers.Timer(600000)
                AddHandler Globals.Ribbons.MSprintExRibbon.TempTableTimer.Elapsed, AddressOf OnTimedEvent
                Globals.Ribbons.MSprintExRibbon.TempTableTimer.Enabled = True
            Else
                Globals.Ribbons.MSprintExRibbon.EnableDisableButtons(True)
                Me.Close()
            End If
        Catch ex As Exception
            LogMpsrintExException("Exception occured while preparing Temp Tables for chosen date range.Message:" + ex.Message)
            MessageBox.Show("Exception occured while preparing server for chosen date range.Please refer to Error log for more details.")
            Me.Close()
        Finally
            Globals.ThisAddIn.Application.StatusBar = String.Empty
        End Try
    End Sub
    Private Sub OnTimedEvent(ByVal source As Object, ByVal e As ElapsedEventArgs)
        Try

            If TempTableIsCreatedAndCompleted(dtFromDate.Value, dtToDate.Value) Then
                If Me.InvokeRequired Then
                    Dim d As New SetTextCallback(AddressOf OnTimedEvent)
                    Me.Invoke(d, New Object() {source, e})
                Else
                    Me.Close()
                End If
                Globals.Ribbons.MSprintExRibbon.TempTableTimer.Enabled = False
                Globals.Ribbons.MSprintExRibbon.TempTableTimer.Dispose()
                SendReadyEmail()
                Globals.Ribbons.MSprintExRibbon.EnableDisableButtons(True)

            End If
        Catch ex As Exception
            LogMpsrintExException("Exception occured while checking temp table status during timer elapsed event.Message :" + ex.Message)
        End Try
    End Sub
    Public Function SendReadyEmail() As Boolean
        Dim inserted As Boolean = False
        Try
            Dim sqlConnection1 As New System.Data.SqlClient.SqlConnection("Server= MUMSQLP01107\GRMINDSQL01;Database=MsprintXTracker;User Id=MSXAdmin;Password=MSXAdmin@123;")
            Dim cmd As New System.Data.SqlClient.SqlCommand
            cmd.CommandType = System.Data.CommandType.StoredProcedure
            ' Dim commandText As String = String.Format("INSERT UsageReport (NTUserName,MSprintX_Method_Invoked,Date,Client,Brand,InputXML,No_Of_Spots,SysDate) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}',{6},getdate())", loggedInUserName, methodName, Date.Now.ToString, clientValue, brandValue, xml.ToString(), no_spots)
            Dim commandText As String = "msdb.dbo.sp_send_dbmail"
            '  LogMpsrintExException(commandText)
            cmd.CommandText = commandText
            cmd.Connection = sqlConnection1
            cmd.Parameters.Add(AddNVarcharParameter("@profile_name", "GroupmCRI"))
            cmd.Parameters.Add(AddNVarcharParameter("@recipients", Globals.Ribbons.MSprintExRibbon.loggedInUserName))
            cmd.Parameters.Add(AddNVarcharParameter("@blind_copy_recipients", "Badri.Narayanan@groupm.com;Rohit.Sule@groupm.com"))
            cmd.Parameters.Add(AddNVarcharParameter("@body", "Hi,<br/>MsprintX is now all set for use.<br/><br/>Regards,<br/>MsprintXUITeam<br/><br/><br/><b>P.S -</b><br/> This is an auto generated e-mail and any reply goes to an unmonitored mailbox.Please reach out to Badri.Narayanan@groupm.com and/or Rohit.Sule@groupm.com for any queries.<br/><br/>"))
            cmd.Parameters.Add(AddNVarcharParameter("@body_format", "HTML"))
            cmd.Parameters.Add(AddNVarcharParameter("@subject", "MsprintX server is available .!"))
            'Dim parameter As New SqlParameter()
            'parameter.ParameterName = "@profile_name"
            'parameter.SqlDbType = SqlDbType.NVarChar
            'parameter.Direction = ParameterDirection.Input
            'parameter.Value = "GroupmCRI"

            'Dim parameter1 As New SqlParameter()
            'parameter1.ParameterName = "@profile_name"
            'parameter.SqlDbType = SqlDbType.NVarChar
            'parameter.Direction = ParameterDirection.Input
            'parameter.Value = "GroupmCRI"


            ' Add the parameter to the Parameters collection. 

            sqlConnection1.Open()
            '  Dim adapter As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(commandText,
            cmd.ExecuteNonQuery()
            sqlConnection1.Close()
        Catch ex As Exception
            LogMpsrintExException("Exception occured while sending MsprintX ready email.Message: " + ex.Message)
        End Try
        Return inserted
    End Function
    Private Function AddNVarcharParameter(ByVal name As String, ByVal value As String) As SqlParameter
        Dim parameter As New SqlParameter()
        Try
            parameter.ParameterName = name
            parameter.SqlDbType = SqlDbType.NVarChar
            parameter.Direction = ParameterDirection.Input
            parameter.Value = value
        Catch ex As Exception

        End Try
        Return parameter
    End Function

    Private Sub dgvWeeks_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvWeeks.CellContentClick

    End Sub

    Private Sub frmPrepareServer_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        'Try
        '    MessageBox.Show(e.CloseReason.ToString())
        '    If e.CloseReason.FormOwnerClosing Then
        '        e.Cancel = True
        '    End If



        'Catch ex As Exception

        'End Try
    End Sub
End Class

Public Class TempTableStatus
    Friend status As TaskPaneLogFile.CreateTempTableStatus
    Friend initiated As Boolean
End Class