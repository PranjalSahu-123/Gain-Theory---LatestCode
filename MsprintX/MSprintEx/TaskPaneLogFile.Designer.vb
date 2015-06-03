<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskPaneLogFile
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.txtWeeks = New System.Windows.Forms.NumericUpDown()
        Me.dtFromDate = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dtToDate = New System.Windows.Forms.DateTimePicker()
        Me.dgvWeeks = New System.Windows.Forms.DataGridView()
        Me.WeekNumberDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.YearDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.StartDateDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EndDateDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.WeeksDataTableBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.Plandata = New MSprintEx.Plandata()
        Me.gbLogFileType = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.rbWeekWise = New System.Windows.Forms.RadioButton()
        Me.rbSingle = New System.Windows.Forms.RadioButton()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnPrepare = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btndaypartsadd = New System.Windows.Forms.Button()
        Me.LbDayParts = New System.Windows.Forms.ListBox()
        Me.nudttmin = New System.Windows.Forms.NumericUpDown()
        Me.nudttHrs = New System.Windows.Forms.NumericUpDown()
        Me.LbTimeTo = New System.Windows.Forms.Label()
        Me.LbtimeFromMins = New System.Windows.Forms.Label()
        Me.nudtfmin = New System.Windows.Forms.NumericUpDown()
        Me.lbTimeFromHrs = New System.Windows.Forms.Label()
        Me.nudtfhrs = New System.Windows.Forms.NumericUpDown()
        Me.lbTimeFrom = New System.Windows.Forms.Label()
        Me.GbDayParts = New System.Windows.Forms.GroupBox()
        Me.gbDays = New System.Windows.Forms.GroupBox()
        Me.cbAllDays = New System.Windows.Forms.CheckBox()
        Me.cbWeekEnds = New System.Windows.Forms.CheckBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.cbThursday = New System.Windows.Forms.CheckBox()
        Me.cbWednesday = New System.Windows.Forms.CheckBox()
        Me.cbTuesday = New System.Windows.Forms.CheckBox()
        Me.cbMonday = New System.Windows.Forms.CheckBox()
        Me.ErrorRangeBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        CType(Me.txtWeeks, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvWeeks, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WeeksDataTableBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Plandata, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbLogFileType.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.nudttmin, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudttHrs, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudtfmin, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudtfhrs, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GbDayParts.SuspendLayout()
        Me.gbDays.SuspendLayout()
        CType(Me.ErrorRangeBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtWeeks
        '
        Me.txtWeeks.Location = New System.Drawing.Point(159, 3)
        Me.txtWeeks.Maximum = New Decimal(New Integer() {52, 0, 0, 0})
        Me.txtWeeks.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.txtWeeks.Name = "txtWeeks"
        Me.txtWeeks.Size = New System.Drawing.Size(39, 20)
        Me.txtWeeks.TabIndex = 1
        Me.txtWeeks.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtWeeks.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'dtFromDate
        '
        Me.dtFromDate.CustomFormat = "dd/MM/yyyy"
        Me.dtFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtFromDate.Location = New System.Drawing.Point(6, 28)
        Me.dtFromDate.MaxDate = New Date(9998, 1, 1, 0, 0, 0, 0)
        Me.dtFromDate.MinDate = New Date(1753, 7, 28, 0, 0, 0, 0)
        Me.dtFromDate.Name = "dtFromDate"
        Me.dtFromDate.Size = New System.Drawing.Size(85, 20)
        Me.dtFromDate.TabIndex = 0
        Me.dtFromDate.Value = New Date(2014, 6, 24, 0, 0, 0, 0)
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(3, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(157, 20)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Enter the number of weeks:"
        '
        'dtToDate
        '
        Me.dtToDate.CustomFormat = "dd/MM/yyyy"
        Me.dtToDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtToDate.Location = New System.Drawing.Point(120, 28)
        Me.dtToDate.MaxDate = New Date(9998, 1, 1, 0, 0, 0, 0)
        Me.dtToDate.MinDate = New Date(1753, 7, 28, 0, 0, 0, 0)
        Me.dtToDate.Name = "dtToDate"
        Me.dtToDate.Size = New System.Drawing.Size(85, 20)
        Me.dtToDate.TabIndex = 2
        Me.dtToDate.Value = New Date(2014, 6, 24, 0, 0, 0, 0)
        '
        'dgvWeeks
        '
        Me.dgvWeeks.AllowUserToAddRows = False
        Me.dgvWeeks.AllowUserToDeleteRows = False
        Me.dgvWeeks.AutoGenerateColumns = False
        Me.dgvWeeks.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvWeeks.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvWeeks.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvWeeks.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.WeekNumberDataGridViewTextBoxColumn, Me.YearDataGridViewTextBoxColumn, Me.StartDateDataGridViewTextBoxColumn, Me.EndDateDataGridViewTextBoxColumn})
        Me.dgvWeeks.DataSource = Me.WeeksDataTableBindingSource
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvWeeks.DefaultCellStyle = DataGridViewCellStyle4
        Me.dgvWeeks.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dgvWeeks.Location = New System.Drawing.Point(0, 54)
        Me.dgvWeeks.MultiSelect = False
        Me.dgvWeeks.Name = "dgvWeeks"
        Me.dgvWeeks.ReadOnly = True
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvWeeks.RowHeadersDefaultCellStyle = DataGridViewCellStyle5
        Me.dgvWeeks.RowHeadersVisible = False
        Me.dgvWeeks.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvWeeks.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvWeeks.Size = New System.Drawing.Size(314, 160)
        Me.dgvWeeks.StandardTab = True
        Me.dgvWeeks.TabIndex = 0
        '
        'WeekNumberDataGridViewTextBoxColumn
        '
        Me.WeekNumberDataGridViewTextBoxColumn.DataPropertyName = "WeekNumber"
        Me.WeekNumberDataGridViewTextBoxColumn.HeaderText = "Wk#"
        Me.WeekNumberDataGridViewTextBoxColumn.Name = "WeekNumberDataGridViewTextBoxColumn"
        Me.WeekNumberDataGridViewTextBoxColumn.ReadOnly = True
        Me.WeekNumberDataGridViewTextBoxColumn.Width = 35
        '
        'YearDataGridViewTextBoxColumn
        '
        Me.YearDataGridViewTextBoxColumn.DataPropertyName = "Year"
        Me.YearDataGridViewTextBoxColumn.HeaderText = "Year"
        Me.YearDataGridViewTextBoxColumn.Name = "YearDataGridViewTextBoxColumn"
        Me.YearDataGridViewTextBoxColumn.ReadOnly = True
        Me.YearDataGridViewTextBoxColumn.Width = 40
        '
        'StartDateDataGridViewTextBoxColumn
        '
        Me.StartDateDataGridViewTextBoxColumn.DataPropertyName = "StartDate"
        DataGridViewCellStyle2.Format = "dd/MM/yyyy"
        DataGridViewCellStyle2.NullValue = Nothing
        Me.StartDateDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle2
        Me.StartDateDataGridViewTextBoxColumn.HeaderText = "StartDate"
        Me.StartDateDataGridViewTextBoxColumn.Name = "StartDateDataGridViewTextBoxColumn"
        Me.StartDateDataGridViewTextBoxColumn.ReadOnly = True
        Me.StartDateDataGridViewTextBoxColumn.Width = 80
        '
        'EndDateDataGridViewTextBoxColumn
        '
        Me.EndDateDataGridViewTextBoxColumn.DataPropertyName = "EndDate"
        DataGridViewCellStyle3.Format = "dd/MM/yyyy"
        DataGridViewCellStyle3.NullValue = Nothing
        Me.EndDateDataGridViewTextBoxColumn.DefaultCellStyle = DataGridViewCellStyle3
        Me.EndDateDataGridViewTextBoxColumn.HeaderText = "EndDate"
        Me.EndDateDataGridViewTextBoxColumn.Name = "EndDateDataGridViewTextBoxColumn"
        Me.EndDateDataGridViewTextBoxColumn.ReadOnly = True
        Me.EndDateDataGridViewTextBoxColumn.Width = 80
        '
        'WeeksDataTableBindingSource
        '
        Me.WeeksDataTableBindingSource.DataMember = "Weeks"
        Me.WeeksDataTableBindingSource.DataSource = Me.Plandata
        '
        'Plandata
        '
        Me.Plandata.DataSetName = "Plandata"
        Me.Plandata.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'gbLogFileType
        '
        Me.gbLogFileType.Controls.Add(Me.Label3)
        Me.gbLogFileType.Controls.Add(Me.rbWeekWise)
        Me.gbLogFileType.Controls.Add(Me.rbSingle)
        Me.gbLogFileType.Location = New System.Drawing.Point(3, 0)
        Me.gbLogFileType.Name = "gbLogFileType"
        Me.gbLogFileType.Size = New System.Drawing.Size(314, 72)
        Me.gbLogFileType.TabIndex = 1
        Me.gbLogFileType.TabStop = False
        Me.gbLogFileType.Text = "Type of Report"
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Red
        Me.Label3.Location = New System.Drawing.Point(8, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(298, 27)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Data in Database is available only from 28/7/2014 to 28/7/2014"
        '
        'rbWeekWise
        '
        Me.rbWeekWise.AutoSize = True
        Me.rbWeekWise.Location = New System.Drawing.Point(120, 20)
        Me.rbWeekWise.Name = "rbWeekWise"
        Me.rbWeekWise.Size = New System.Drawing.Size(78, 17)
        Me.rbWeekWise.TabIndex = 1
        Me.rbWeekWise.Text = "Week-wise"
        Me.rbWeekWise.UseVisualStyleBackColor = True
        '
        'rbSingle
        '
        Me.rbSingle.AutoSize = True
        Me.rbSingle.Checked = True
        Me.rbSingle.Location = New System.Drawing.Point(27, 20)
        Me.rbSingle.Name = "rbSingle"
        Me.rbSingle.Size = New System.Drawing.Size(64, 17)
        Me.rbSingle.TabIndex = 0
        Me.rbSingle.TabStop = True
        Me.rbSingle.Text = "Clubbed"
        Me.rbSingle.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnPrepare)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.dtFromDate)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.txtWeeks)
        Me.Panel1.Controls.Add(Me.dgvWeeks)
        Me.Panel1.Controls.Add(Me.dtToDate)
        Me.Panel1.Location = New System.Drawing.Point(3, 70)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(314, 211)
        Me.Panel1.TabIndex = 2
        '
        'btnPrepare
        '
        Me.btnPrepare.BackColor = System.Drawing.SystemColors.Control
        Me.btnPrepare.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnPrepare.Location = New System.Drawing.Point(211, 27)
        Me.btnPrepare.Name = "btnPrepare"
        Me.btnPrepare.Size = New System.Drawing.Size(95, 23)
        Me.btnPrepare.TabIndex = 4
        Me.btnPrepare.Text = "Prepare"
        Me.btnPrepare.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(97, 28)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(17, 20)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "to"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btndaypartsadd
        '
        Me.btndaypartsadd.Location = New System.Drawing.Point(5, 153)
        Me.btndaypartsadd.Name = "btndaypartsadd"
        Me.btndaypartsadd.Size = New System.Drawing.Size(45, 23)
        Me.btndaypartsadd.TabIndex = 20
        Me.btndaypartsadd.Text = "Add"
        Me.btndaypartsadd.UseVisualStyleBackColor = True
        '
        'LbDayParts
        '
        Me.LbDayParts.FormattingEnabled = True
        Me.LbDayParts.Location = New System.Drawing.Point(10, 91)
        Me.LbDayParts.Name = "LbDayParts"
        Me.LbDayParts.Size = New System.Drawing.Size(151, 56)
        Me.LbDayParts.TabIndex = 19
        '
        'nudttmin
        '
        Me.nudttmin.Location = New System.Drawing.Point(112, 65)
        Me.nudttmin.Maximum = New Decimal(New Integer() {59, 0, 0, 0})
        Me.nudttmin.Name = "nudttmin"
        Me.nudttmin.Size = New System.Drawing.Size(42, 20)
        Me.nudttmin.TabIndex = 17
        Me.nudttmin.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.nudttmin.Value = New Decimal(New Integer() {59, 0, 0, 0})
        '
        'nudttHrs
        '
        Me.nudttHrs.Location = New System.Drawing.Point(63, 65)
        Me.nudttHrs.Maximum = New Decimal(New Integer() {25, 0, 0, 0})
        Me.nudttHrs.Minimum = New Decimal(New Integer() {2, 0, 0, 0})
        Me.nudttHrs.Name = "nudttHrs"
        Me.nudttHrs.Size = New System.Drawing.Size(43, 20)
        Me.nudttHrs.TabIndex = 14
        Me.nudttHrs.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.nudttHrs.Value = New Decimal(New Integer() {25, 0, 0, 0})
        '
        'LbTimeTo
        '
        Me.LbTimeTo.AutoSize = True
        Me.LbTimeTo.Location = New System.Drawing.Point(6, 72)
        Me.LbTimeTo.Name = "LbTimeTo"
        Me.LbTimeTo.Size = New System.Drawing.Size(52, 13)
        Me.LbTimeTo.TabIndex = 13
        Me.LbTimeTo.Text = "Time To :"
        '
        'LbtimeFromMins
        '
        Me.LbtimeFromMins.AutoSize = True
        Me.LbtimeFromMins.Location = New System.Drawing.Point(111, 19)
        Me.LbtimeFromMins.Name = "LbtimeFromMins"
        Me.LbtimeFromMins.Size = New System.Drawing.Size(24, 13)
        Me.LbtimeFromMins.TabIndex = 12
        Me.LbtimeFromMins.Text = "Min"
        '
        'nudtfmin
        '
        Me.nudtfmin.Location = New System.Drawing.Point(112, 38)
        Me.nudtfmin.Maximum = New Decimal(New Integer() {59, 0, 0, 0})
        Me.nudtfmin.Name = "nudtfmin"
        Me.nudtfmin.Size = New System.Drawing.Size(42, 20)
        Me.nudtfmin.TabIndex = 11
        Me.nudtfmin.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lbTimeFromHrs
        '
        Me.lbTimeFromHrs.AutoSize = True
        Me.lbTimeFromHrs.Location = New System.Drawing.Point(60, 16)
        Me.lbTimeFromHrs.Name = "lbTimeFromHrs"
        Me.lbTimeFromHrs.Size = New System.Drawing.Size(18, 13)
        Me.lbTimeFromHrs.TabIndex = 9
        Me.lbTimeFromHrs.Text = "Hr"
        '
        'nudtfhrs
        '
        Me.nudtfhrs.Location = New System.Drawing.Point(63, 39)
        Me.nudtfhrs.Maximum = New Decimal(New Integer() {25, 0, 0, 0})
        Me.nudtfhrs.Minimum = New Decimal(New Integer() {2, 0, 0, 0})
        Me.nudtfhrs.Name = "nudtfhrs"
        Me.nudtfhrs.Size = New System.Drawing.Size(43, 20)
        Me.nudtfhrs.TabIndex = 8
        Me.nudtfhrs.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.nudtfhrs.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'lbTimeFrom
        '
        Me.lbTimeFrom.AutoSize = True
        Me.lbTimeFrom.Location = New System.Drawing.Point(-3, 45)
        Me.lbTimeFrom.Name = "lbTimeFrom"
        Me.lbTimeFrom.Size = New System.Drawing.Size(62, 13)
        Me.lbTimeFrom.TabIndex = 7
        Me.lbTimeFrom.Text = "Time From :"
        '
        'GbDayParts
        '
        Me.GbDayParts.Controls.Add(Me.gbDays)
        Me.GbDayParts.Controls.Add(Me.LbDayParts)
        Me.GbDayParts.Controls.Add(Me.lbTimeFromHrs)
        Me.GbDayParts.Controls.Add(Me.nudtfmin)
        Me.GbDayParts.Controls.Add(Me.btndaypartsadd)
        Me.GbDayParts.Controls.Add(Me.nudtfhrs)
        Me.GbDayParts.Controls.Add(Me.nudttmin)
        Me.GbDayParts.Controls.Add(Me.lbTimeFrom)
        Me.GbDayParts.Controls.Add(Me.nudttHrs)
        Me.GbDayParts.Controls.Add(Me.LbTimeTo)
        Me.GbDayParts.Controls.Add(Me.LbtimeFromMins)
        Me.GbDayParts.Location = New System.Drawing.Point(9, 287)
        Me.GbDayParts.Name = "GbDayParts"
        Me.GbDayParts.Size = New System.Drawing.Size(308, 182)
        Me.GbDayParts.TabIndex = 21
        Me.GbDayParts.TabStop = False
        Me.GbDayParts.Text = "Choose Day Parts"
        '
        'gbDays
        '
        Me.gbDays.Controls.Add(Me.cbAllDays)
        Me.gbDays.Controls.Add(Me.cbWeekEnds)
        Me.gbDays.Controls.Add(Me.CheckBox1)
        Me.gbDays.Controls.Add(Me.cbThursday)
        Me.gbDays.Controls.Add(Me.cbWednesday)
        Me.gbDays.Controls.Add(Me.cbTuesday)
        Me.gbDays.Controls.Add(Me.cbMonday)
        Me.gbDays.Location = New System.Drawing.Point(167, 19)
        Me.gbDays.Name = "gbDays"
        Me.gbDays.Size = New System.Drawing.Size(135, 151)
        Me.gbDays.TabIndex = 21
        Me.gbDays.TabStop = False
        Me.gbDays.Text = "Choose Days"
        Me.gbDays.Visible = False
        '
        'cbAllDays
        '
        Me.cbAllDays.AutoSize = True
        Me.cbAllDays.Location = New System.Drawing.Point(6, 128)
        Me.cbAllDays.Name = "cbAllDays"
        Me.cbAllDays.Size = New System.Drawing.Size(64, 17)
        Me.cbAllDays.TabIndex = 6
        Me.cbAllDays.Text = "All Days"
        Me.cbAllDays.UseVisualStyleBackColor = True
        '
        'cbWeekEnds
        '
        Me.cbWeekEnds.AutoSize = True
        Me.cbWeekEnds.Location = New System.Drawing.Point(59, 92)
        Me.cbWeekEnds.Name = "cbWeekEnds"
        Me.cbWeekEnds.Size = New System.Drawing.Size(74, 17)
        Me.cbWeekEnds.TabIndex = 5
        Me.cbWeekEnds.Text = "WeekEnd"
        Me.cbWeekEnds.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(6, 92)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(37, 17)
        Me.CheckBox1.TabIndex = 4
        Me.CheckBox1.Text = "Fri"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'cbThursday
        '
        Me.cbThursday.AutoSize = True
        Me.cbThursday.Location = New System.Drawing.Point(59, 60)
        Me.cbThursday.Name = "cbThursday"
        Me.cbThursday.Size = New System.Drawing.Size(45, 17)
        Me.cbThursday.TabIndex = 3
        Me.cbThursday.Text = "Thu"
        Me.cbThursday.UseVisualStyleBackColor = True
        '
        'cbWednesday
        '
        Me.cbWednesday.AutoSize = True
        Me.cbWednesday.Location = New System.Drawing.Point(5, 60)
        Me.cbWednesday.Name = "cbWednesday"
        Me.cbWednesday.Size = New System.Drawing.Size(49, 17)
        Me.cbWednesday.TabIndex = 2
        Me.cbWednesday.Text = "Wed"
        Me.cbWednesday.UseVisualStyleBackColor = True
        '
        'cbTuesday
        '
        Me.cbTuesday.AutoSize = True
        Me.cbTuesday.Location = New System.Drawing.Point(59, 19)
        Me.cbTuesday.Name = "cbTuesday"
        Me.cbTuesday.Size = New System.Drawing.Size(45, 17)
        Me.cbTuesday.TabIndex = 1
        Me.cbTuesday.Text = "Tue"
        Me.cbTuesday.UseVisualStyleBackColor = True
        '
        'cbMonday
        '
        Me.cbMonday.AutoSize = True
        Me.cbMonday.Location = New System.Drawing.Point(6, 19)
        Me.cbMonday.Name = "cbMonday"
        Me.cbMonday.Size = New System.Drawing.Size(47, 17)
        Me.cbMonday.TabIndex = 0
        Me.cbMonday.Text = "Mon"
        Me.cbMonday.UseVisualStyleBackColor = True
        '
        'ErrorRangeBindingSource
        '
        Me.ErrorRangeBindingSource.DataMember = "ErrorRange"
        Me.ErrorRangeBindingSource.DataSource = Me.Plandata
        '
        'TaskPaneLogFile
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Controls.Add(Me.GbDayParts)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.gbLogFileType)
        Me.Name = "TaskPaneLogFile"
        Me.Size = New System.Drawing.Size(322, 483)
        CType(Me.txtWeeks, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvWeeks, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WeeksDataTableBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Plandata, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbLogFileType.ResumeLayout(False)
        Me.gbLogFileType.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        CType(Me.nudttmin, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudttHrs, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudtfmin, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudtfhrs, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GbDayParts.ResumeLayout(False)
        Me.GbDayParts.PerformLayout()
        Me.gbDays.ResumeLayout(False)
        Me.gbDays.PerformLayout()
        CType(Me.ErrorRangeBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txtWeeks As System.Windows.Forms.NumericUpDown
    Friend WithEvents dtFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Plandata As MSprintEx.Plandata
    Friend WithEvents WeeksDataTableBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents dgvWeeks As System.Windows.Forms.DataGridView
    Friend WithEvents WeekNumberDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents YearDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents StartDateDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EndDateDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents gbLogFileType As System.Windows.Forms.GroupBox
    Friend WithEvents rbWeekWise As System.Windows.Forms.RadioButton
    Friend WithEvents rbSingle As System.Windows.Forms.RadioButton
    Friend WithEvents ErrorRangeBindingSource As System.Windows.Forms.BindingSource
    'Friend WithEvents ucChannelsMapping As MSprintEx.ChannelMapping
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents nudtfhrs As System.Windows.Forms.NumericUpDown
    Friend WithEvents lbTimeFrom As System.Windows.Forms.Label
    Friend WithEvents lbTimeFromHrs As System.Windows.Forms.Label
    Friend WithEvents LbtimeFromMins As System.Windows.Forms.Label
    Friend WithEvents nudtfmin As System.Windows.Forms.NumericUpDown
    Friend WithEvents LbTimeTo As System.Windows.Forms.Label
    Friend WithEvents nudttHrs As System.Windows.Forms.NumericUpDown
    Friend WithEvents nudttmin As System.Windows.Forms.NumericUpDown
    Friend WithEvents LbDayParts As System.Windows.Forms.ListBox
    Friend WithEvents btndaypartsadd As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents GbDayParts As System.Windows.Forms.GroupBox
    Friend WithEvents gbDays As System.Windows.Forms.GroupBox
    Friend WithEvents cbWednesday As System.Windows.Forms.CheckBox
    Friend WithEvents cbTuesday As System.Windows.Forms.CheckBox
    Friend WithEvents cbMonday As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents cbThursday As System.Windows.Forms.CheckBox
    Friend WithEvents cbAllDays As System.Windows.Forms.CheckBox
    Friend WithEvents cbWeekEnds As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnPrepare As System.Windows.Forms.Button

End Class
