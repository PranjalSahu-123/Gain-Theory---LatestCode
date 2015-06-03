<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPrepareServer
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dtFromDate = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtWeeks = New System.Windows.Forms.NumericUpDown()
        Me.dgvWeeks = New System.Windows.Forms.DataGridView()
        Me.dtToDate = New System.Windows.Forms.DateTimePicker()
        Me.btnPrepare = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Plandata1 = New MSprintEx.Plandata()
        Me.Panel1.SuspendLayout()
        CType(Me.txtWeeks, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvWeeks, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Plandata1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.dtFromDate)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.txtWeeks)
        Me.Panel1.Controls.Add(Me.dgvWeeks)
        Me.Panel1.Controls.Add(Me.dtToDate)
        Me.Panel1.Location = New System.Drawing.Point(32, 39)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(323, 217)
        Me.Panel1.TabIndex = 3
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
        'dgvWeeks
        '
        Me.dgvWeeks.AllowUserToAddRows = False
        Me.dgvWeeks.AllowUserToDeleteRows = False
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
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvWeeks.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvWeeks.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dgvWeeks.Location = New System.Drawing.Point(0, 54)
        Me.dgvWeeks.MultiSelect = False
        Me.dgvWeeks.Name = "dgvWeeks"
        Me.dgvWeeks.ReadOnly = True
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvWeeks.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvWeeks.RowHeadersVisible = False
        Me.dgvWeeks.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvWeeks.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvWeeks.Size = New System.Drawing.Size(320, 160)
        Me.dgvWeeks.StandardTab = True
        Me.dgvWeeks.TabIndex = 0
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
        'btnPrepare
        '
        Me.btnPrepare.BackColor = System.Drawing.Color.OrangeRed
        Me.btnPrepare.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnPrepare.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnPrepare.Location = New System.Drawing.Point(113, 265)
        Me.btnPrepare.Name = "btnPrepare"
        Me.btnPrepare.Size = New System.Drawing.Size(133, 33)
        Me.btnPrepare.TabIndex = 4
        Me.btnPrepare.Text = "Prepare"
        Me.btnPrepare.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.Red
        Me.Label3.Location = New System.Drawing.Point(29, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(318, 27)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Data in Database is available only from 28/7/2014 to 28/7/2014"
        '
        'Plandata1
        '
        Me.Plandata1.DataSetName = "Plandata"
        Me.Plandata1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'frmPrepareServer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(397, 301)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnPrepare)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Panel1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPrepareServer"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Prepare MsprintX Server by choosing date range"
        Me.Panel1.ResumeLayout(False)
        CType(Me.txtWeeks, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvWeeks, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Plandata1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Plandata1 As MSprintEx.Plandata
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnPrepare As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtFromDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtWeeks As System.Windows.Forms.NumericUpDown
    Friend WithEvents dgvWeeks As System.Windows.Forms.DataGridView
    Friend WithEvents dtToDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
