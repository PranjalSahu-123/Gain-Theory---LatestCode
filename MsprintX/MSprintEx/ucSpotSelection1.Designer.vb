<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucSpotSelection
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
        Me.dgSelectedSpotsGrid = New System.Windows.Forms.DataGridView()
        Me.btnPushOne = New System.Windows.Forms.Button()
        Me.btnPushOneToSelected = New System.Windows.Forms.Button()
        Me.dgvAvailableSpotsGrid = New System.Windows.Forms.DataGridView()
        Me.btnGetAvailableSpots = New System.Windows.Forms.Button()
        Me.lbFromDate = New System.Windows.Forms.Label()
        Me.cbWeeks = New System.Windows.Forms.ComboBox()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lbRate = New System.Windows.Forms.Label()
        Me.lbEndtime = New System.Windows.Forms.Label()
        Me.lbStartTime = New System.Windows.Forms.Label()
        Me.lbDays = New System.Windows.Forms.Label()
        Me.lbProg = New System.Windows.Forms.Label()
        Me.lbChannel = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lbErrorLabel = New System.Windows.Forms.Label()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.SecTableAdapter1 = New MSprintEx.METISTableAdapters.SECTableAdapter()
        CType(Me.dgSelectedSpotsGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvAvailableSpotsGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgSelectedSpotsGrid
        '
        Me.dgSelectedSpotsGrid.AllowUserToAddRows = False
        Me.dgSelectedSpotsGrid.AllowUserToDeleteRows = False
        Me.dgSelectedSpotsGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells
        Me.dgSelectedSpotsGrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgSelectedSpotsGrid.Location = New System.Drawing.Point(3, 31)
        Me.dgSelectedSpotsGrid.Name = "dgSelectedSpotsGrid"
        Me.dgSelectedSpotsGrid.ReadOnly = True
        Me.dgSelectedSpotsGrid.RowHeadersVisible = False
        Me.dgSelectedSpotsGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgSelectedSpotsGrid.Size = New System.Drawing.Size(389, 228)
        Me.dgSelectedSpotsGrid.TabIndex = 0
        '
        'btnPushOne
        '
        Me.btnPushOne.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnPushOne.Location = New System.Drawing.Point(335, 265)
        Me.btnPushOne.Name = "btnPushOne"
        Me.btnPushOne.Size = New System.Drawing.Size(57, 22)
        Me.btnPushOne.TabIndex = 1
        Me.btnPushOne.Text = "Remove"
        Me.btnPushOne.UseVisualStyleBackColor = True
        '
        'btnPushOneToSelected
        '
        Me.btnPushOneToSelected.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnPushOneToSelected.Location = New System.Drawing.Point(282, 265)
        Me.btnPushOneToSelected.Name = "btnPushOneToSelected"
        Me.btnPushOneToSelected.Size = New System.Drawing.Size(83, 22)
        Me.btnPushOneToSelected.TabIndex = 3
        Me.btnPushOneToSelected.Text = "Select spots"
        Me.btnPushOneToSelected.UseVisualStyleBackColor = True
        '
        'dgvAvailableSpotsGrid
        '
        Me.dgvAvailableSpotsGrid.AllowUserToAddRows = False
        Me.dgvAvailableSpotsGrid.AllowUserToDeleteRows = False
        Me.dgvAvailableSpotsGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells
        Me.dgvAvailableSpotsGrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvAvailableSpotsGrid.Location = New System.Drawing.Point(3, 31)
        Me.dgvAvailableSpotsGrid.Name = "dgvAvailableSpotsGrid"
        Me.dgvAvailableSpotsGrid.ReadOnly = True
        Me.dgvAvailableSpotsGrid.RowHeadersVisible = False
        Me.dgvAvailableSpotsGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvAvailableSpotsGrid.Size = New System.Drawing.Size(362, 228)
        Me.dgvAvailableSpotsGrid.TabIndex = 7
        '
        'btnGetAvailableSpots
        '
        Me.btnGetAvailableSpots.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnGetAvailableSpots.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGetAvailableSpots.Location = New System.Drawing.Point(3, 3)
        Me.btnGetAvailableSpots.Name = "btnGetAvailableSpots"
        Me.btnGetAvailableSpots.Size = New System.Drawing.Size(116, 22)
        Me.btnGetAvailableSpots.TabIndex = 8
        Me.btnGetAvailableSpots.Text = "Get Available Spots"
        Me.btnGetAvailableSpots.UseVisualStyleBackColor = True
        '
        'lbFromDate
        '
        Me.lbFromDate.Location = New System.Drawing.Point(5, 11)
        Me.lbFromDate.Name = "lbFromDate"
        Me.lbFromDate.Size = New System.Drawing.Size(81, 13)
        Me.lbFromDate.TabIndex = 9
        Me.lbFromDate.Text = "Choose Week :"
        '
        'cbWeeks
        '
        Me.cbWeeks.FormattingEnabled = True
        Me.cbWeeks.Items.AddRange(New Object() {"All"})
        Me.cbWeeks.Location = New System.Drawing.Point(90, 11)
        Me.cbWeeks.Name = "cbWeeks"
        Me.cbWeeks.Size = New System.Drawing.Size(65, 21)
        Me.cbWeeks.TabIndex = 10
        Me.cbWeeks.Text = "All"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Location = New System.Drawing.Point(-73, 68)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Size = New System.Drawing.Size(150, 100)
        Me.SplitContainer1.TabIndex = 11
        '
        'SplitContainer2
        '
        Me.SplitContainer2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(3, 99)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.TableLayoutPanel1)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.TableLayoutPanel2)
        Me.SplitContainer2.Size = New System.Drawing.Size(775, 294)
        Me.SplitContainer2.SplitterDistance = 399
        Me.SplitContainer2.TabIndex = 11
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.dgSelectedSpotsGrid, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.btnPushOne, 0, 2)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 3
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(395, 290)
        Me.TableLayoutPanel1.TabIndex = 14
        '
        'Label1
        '
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label1.Location = New System.Drawing.Point(3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(389, 28)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Selected spots"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 1
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.dgvAvailableSpotsGrid, 0, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.btnGetAvailableSpots, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.btnPushOneToSelected, 0, 2)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 3
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(368, 290)
        Me.TableLayoutPanel2.TabIndex = 15
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.lbErrorLabel)
        Me.Panel1.Controls.Add(Me.lbFromDate)
        Me.Panel1.Controls.Add(Me.cbWeeks)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(3, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(775, 90)
        Me.Panel1.TabIndex = 13
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lbRate)
        Me.GroupBox1.Controls.Add(Me.lbEndtime)
        Me.GroupBox1.Controls.Add(Me.lbStartTime)
        Me.GroupBox1.Controls.Add(Me.lbDays)
        Me.GroupBox1.Controls.Add(Me.lbProg)
        Me.GroupBox1.Controls.Add(Me.lbChannel)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(161, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(609, 58)
        Me.GroupBox1.TabIndex = 14
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Current Plan Item"
        '
        'lbRate
        '
        Me.lbRate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbRate.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.lbRate.Location = New System.Drawing.Point(510, 32)
        Me.lbRate.Name = "lbRate"
        Me.lbRate.Size = New System.Drawing.Size(93, 23)
        Me.lbRate.TabIndex = 20
        '
        'lbEndtime
        '
        Me.lbEndtime.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbEndtime.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.lbEndtime.Location = New System.Drawing.Point(455, 32)
        Me.lbEndtime.Name = "lbEndtime"
        Me.lbEndtime.Size = New System.Drawing.Size(49, 23)
        Me.lbEndtime.TabIndex = 19
        '
        'lbStartTime
        '
        Me.lbStartTime.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbStartTime.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.lbStartTime.Location = New System.Drawing.Point(397, 32)
        Me.lbStartTime.Name = "lbStartTime"
        Me.lbStartTime.Size = New System.Drawing.Size(52, 23)
        Me.lbStartTime.TabIndex = 18
        '
        'lbDays
        '
        Me.lbDays.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbDays.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.lbDays.Location = New System.Drawing.Point(273, 32)
        Me.lbDays.Name = "lbDays"
        Me.lbDays.Size = New System.Drawing.Size(118, 23)
        Me.lbDays.TabIndex = 17
        '
        'lbProg
        '
        Me.lbProg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbProg.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.lbProg.Location = New System.Drawing.Point(112, 32)
        Me.lbProg.Name = "lbProg"
        Me.lbProg.Size = New System.Drawing.Size(152, 23)
        Me.lbProg.TabIndex = 16
        '
        'lbChannel
        '
        Me.lbChannel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbChannel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.lbChannel.Location = New System.Drawing.Point(9, 32)
        Me.lbChannel.Name = "lbChannel"
        Me.lbChannel.Size = New System.Drawing.Size(100, 18)
        Me.lbChannel.TabIndex = 15
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(270, 16)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(26, 13)
        Me.Label7.TabIndex = 10
        Me.Label7.Text = "Day"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(507, 16)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(30, 13)
        Me.Label6.TabIndex = 8
        Me.Label6.Text = "Rate"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(452, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(52, 13)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "End Time"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(394, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(55, 13)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Start Time"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(109, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Programme"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(46, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Channel"
        '
        'lbErrorLabel
        '
        Me.lbErrorLabel.AutoSize = True
        Me.lbErrorLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbErrorLabel.ForeColor = System.Drawing.Color.Red
        Me.lbErrorLabel.Location = New System.Drawing.Point(52, 66)
        Me.lbErrorLabel.Name = "lbErrorLabel"
        Me.lbErrorLabel.Size = New System.Drawing.Size(34, 13)
        Me.lbErrorLabel.TabIndex = 12
        Me.lbErrorLabel.Text = "Error"
        Me.lbErrorLabel.Visible = False
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.ColumnCount = 1
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel3.Controls.Add(Me.Panel1, 0, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.SplitContainer2, 0, 1)
        Me.TableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 2
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 96.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(781, 396)
        Me.TableLayoutPanel3.TabIndex = 14
        '
        'SecTableAdapter1
        '
        Me.SecTableAdapter1.ClearBeforeFill = True
        '
        'ucSpotSelection
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.Controls.Add(Me.TableLayoutPanel3)
        Me.Name = "ucSpotSelection"
        Me.Size = New System.Drawing.Size(781, 396)
        CType(Me.dgSelectedSpotsGrid, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvAvailableSpotsGrid, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgSelectedSpotsGrid As System.Windows.Forms.DataGridView
    Friend WithEvents btnPushOne As System.Windows.Forms.Button
    Friend WithEvents btnPushOneToSelected As System.Windows.Forms.Button
    Friend WithEvents dgvAvailableSpotsGrid As System.Windows.Forms.DataGridView
    Friend WithEvents btnGetAvailableSpots As System.Windows.Forms.Button
    Friend WithEvents lbFromDate As System.Windows.Forms.Label
    Friend WithEvents cbWeeks As System.Windows.Forms.ComboBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents TableLayoutPanel3 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents lbErrorLabel As System.Windows.Forms.Label
    Friend WithEvents SecTableAdapter1 As MSprintEx.METISTableAdapters.SECTableAdapter
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lbChannel As System.Windows.Forms.Label
    Friend WithEvents lbProg As System.Windows.Forms.Label
    Friend WithEvents lbDays As System.Windows.Forms.Label
    Friend WithEvents lbStartTime As System.Windows.Forms.Label
    Friend WithEvents lbEndtime As System.Windows.Forms.Label
    Friend WithEvents lbRate As System.Windows.Forms.Label

End Class
