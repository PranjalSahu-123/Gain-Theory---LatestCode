<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucMkets
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.chkSelectAll = New System.Windows.Forms.CheckBox()
        Me.dgvMarkets = New System.Windows.Forms.DataGridView()
        Me.lbMarketGroup = New System.Windows.Forms.ListBox()
        Me.dgvSelectedMarkets = New System.Windows.Forms.DataGridView()
        Me.btnSetPlanMG = New System.Windows.Forms.Button()
        Me.lbPlan = New System.Windows.Forms.ListBox()
        Me.lbPlanning = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.tpMarkets = New System.Windows.Forms.TabPage()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.tpGroups = New System.Windows.Forms.TabPage()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnCreateGroup = New System.Windows.Forms.Button()
        Me.txtGroupName = New System.Windows.Forms.TextBox()
        Me.dgvForNewMG = New System.Windows.Forms.DataGridView()
        Me.btnNewGroup = New System.Windows.Forms.Button()
        Me.tlpNewGroup = New System.Windows.Forms.TableLayoutPanel()
        Me.btnAddtoGroup = New System.Windows.Forms.Button()
        Me.chkSelectallMarkets = New System.Windows.Forms.CheckBox()
        Me.flpNewGroup = New System.Windows.Forms.Panel()
        CType(Me.dgvMarkets, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSelectedMarkets, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.tpMarkets.SuspendLayout()
        Me.tpGroups.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.dgvForNewMG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tlpNewGroup.SuspendLayout()
        Me.flpNewGroup.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkSelectAll
        '
        Me.chkSelectAll.Dock = System.Windows.Forms.DockStyle.Top
        Me.chkSelectAll.Location = New System.Drawing.Point(3, 3)
        Me.chkSelectAll.Name = "chkSelectAll"
        Me.chkSelectAll.Size = New System.Drawing.Size(304, 31)
        Me.chkSelectAll.TabIndex = 1
        Me.chkSelectAll.Text = "Select all"
        Me.chkSelectAll.UseVisualStyleBackColor = True
        '
        'dgvMarkets
        '
        Me.dgvMarkets.AllowUserToAddRows = False
        Me.dgvMarkets.AllowUserToDeleteRows = False
        Me.dgvMarkets.AllowUserToResizeRows = False
        Me.dgvMarkets.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvMarkets.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvMarkets.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvMarkets.ColumnHeadersVisible = False
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvMarkets.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvMarkets.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvMarkets.Location = New System.Drawing.Point(3, 34)
        Me.dgvMarkets.Margin = New System.Windows.Forms.Padding(0)
        Me.dgvMarkets.Name = "dgvMarkets"
        Me.dgvMarkets.ReadOnly = True
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvMarkets.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvMarkets.RowHeadersVisible = False
        Me.dgvMarkets.RowTemplate.Height = 20
        Me.dgvMarkets.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvMarkets.Size = New System.Drawing.Size(304, 136)
        Me.dgvMarkets.TabIndex = 2
        '
        'lbMarketGroup
        '
        Me.lbMarketGroup.FormattingEnabled = True
        Me.lbMarketGroup.Location = New System.Drawing.Point(3, 3)
        Me.lbMarketGroup.Margin = New System.Windows.Forms.Padding(0)
        Me.lbMarketGroup.Name = "lbMarketGroup"
        Me.lbMarketGroup.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lbMarketGroup.Size = New System.Drawing.Size(152, 134)
        Me.lbMarketGroup.TabIndex = 8
        '
        'dgvSelectedMarkets
        '
        Me.dgvSelectedMarkets.AllowUserToAddRows = False
        Me.dgvSelectedMarkets.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSelectedMarkets.ColumnHeadersVisible = False
        Me.dgvSelectedMarkets.Location = New System.Drawing.Point(155, 3)
        Me.dgvSelectedMarkets.Name = "dgvSelectedMarkets"
        Me.dgvSelectedMarkets.ReadOnly = True
        Me.dgvSelectedMarkets.RowHeadersVisible = False
        Me.dgvSelectedMarkets.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvSelectedMarkets.Size = New System.Drawing.Size(152, 134)
        Me.dgvSelectedMarkets.TabIndex = 12
        '
        'btnSetPlanMG
        '
        Me.btnSetPlanMG.Location = New System.Drawing.Point(254, 224)
        Me.btnSetPlanMG.Name = "btnSetPlanMG"
        Me.btnSetPlanMG.Size = New System.Drawing.Size(64, 95)
        Me.btnSetPlanMG.TabIndex = 13
        Me.btnSetPlanMG.Text = "Add to Plan"
        Me.btnSetPlanMG.UseVisualStyleBackColor = True
        '
        'lbPlan
        '
        Me.lbPlan.FormattingEnabled = True
        Me.lbPlan.Location = New System.Drawing.Point(7, 224)
        Me.lbPlan.Name = "lbPlan"
        Me.lbPlan.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lbPlan.Size = New System.Drawing.Size(246, 95)
        Me.lbPlan.TabIndex = 17
        '
        'lbPlanning
        '
        Me.lbPlanning.Location = New System.Drawing.Point(7, 198)
        Me.lbPlanning.Name = "lbPlanning"
        Me.lbPlanning.Size = New System.Drawing.Size(308, 23)
        Me.lbPlanning.TabIndex = 19
        Me.lbPlanning.Text = "Planning MGs"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.tpMarkets)
        Me.TabControl1.Controls.Add(Me.tpGroups)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Top
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(318, 199)
        Me.TabControl1.TabIndex = 24
        '
        'tpMarkets
        '
        Me.tpMarkets.Controls.Add(Me.Label1)
        Me.tpMarkets.Controls.Add(Me.TextBox1)
        Me.tpMarkets.Controls.Add(Me.dgvMarkets)
        Me.tpMarkets.Controls.Add(Me.chkSelectAll)
        Me.tpMarkets.Location = New System.Drawing.Point(4, 22)
        Me.tpMarkets.Name = "tpMarkets"
        Me.tpMarkets.Padding = New System.Windows.Forms.Padding(3)
        Me.tpMarkets.Size = New System.Drawing.Size(310, 173)
        Me.tpMarkets.TabIndex = 0
        Me.tpMarkets.Text = "Markets"
        Me.tpMarkets.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(87, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(47, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Search :"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(134, 8)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(100, 20)
        Me.TextBox1.TabIndex = 3
        '
        'tpGroups
        '
        Me.tpGroups.Controls.Add(Me.lbMarketGroup)
        Me.tpGroups.Controls.Add(Me.dgvSelectedMarkets)
        Me.tpGroups.Location = New System.Drawing.Point(4, 22)
        Me.tpGroups.Name = "tpGroups"
        Me.tpGroups.Padding = New System.Windows.Forms.Padding(3)
        Me.tpGroups.Size = New System.Drawing.Size(310, 173)
        Me.tpGroups.TabIndex = 1
        Me.tpGroups.Text = "Groups"
        Me.tpGroups.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.TabControl1)
        Me.Panel1.Controls.Add(Me.btnSetPlanMG)
        Me.Panel1.Controls.Add(Me.lbPlan)
        Me.Panel1.Controls.Add(Me.lbPlanning)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(318, 322)
        Me.Panel1.TabIndex = 28
        '
        'btnCreateGroup
        '
        Me.btnCreateGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnCreateGroup.Location = New System.Drawing.Point(223, 285)
        Me.btnCreateGroup.Name = "btnCreateGroup"
        Me.btnCreateGroup.Size = New System.Drawing.Size(95, 28)
        Me.btnCreateGroup.TabIndex = 11
        Me.btnCreateGroup.Text = "Create Group"
        Me.btnCreateGroup.UseVisualStyleBackColor = True
        '
        'txtGroupName
        '
        Me.txtGroupName.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtGroupName.Location = New System.Drawing.Point(3, 285)
        Me.txtGroupName.Name = "txtGroupName"
        Me.txtGroupName.Size = New System.Drawing.Size(214, 20)
        Me.txtGroupName.TabIndex = 10
        Me.txtGroupName.Text = "Enter group name here"
        '
        'dgvForNewMG
        '
        Me.dgvForNewMG.AllowUserToAddRows = False
        Me.dgvForNewMG.AllowUserToDeleteRows = False
        Me.dgvForNewMG.AllowUserToResizeRows = False
        Me.dgvForNewMG.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvForNewMG.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvForNewMG.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvForNewMG.ColumnHeadersVisible = False
        Me.tlpNewGroup.SetColumnSpan(Me.dgvForNewMG, 2)
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvForNewMG.DefaultCellStyle = DataGridViewCellStyle5
        Me.dgvForNewMG.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvForNewMG.Location = New System.Drawing.Point(3, 59)
        Me.dgvForNewMG.Name = "dgvForNewMG"
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvForNewMG.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.dgvForNewMG.RowHeadersVisible = False
        Me.dgvForNewMG.RowTemplate.Height = 20
        Me.dgvForNewMG.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvForNewMG.Size = New System.Drawing.Size(315, 220)
        Me.dgvForNewMG.TabIndex = 26
        '
        'btnNewGroup
        '
        Me.tlpNewGroup.SetColumnSpan(Me.btnNewGroup, 2)
        Me.btnNewGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnNewGroup.Location = New System.Drawing.Point(3, 3)
        Me.btnNewGroup.Name = "btnNewGroup"
        Me.btnNewGroup.Size = New System.Drawing.Size(315, 23)
        Me.btnNewGroup.TabIndex = 25
        Me.btnNewGroup.Text = "Click here to create additional Market Group"
        Me.btnNewGroup.UseVisualStyleBackColor = True
        '
        'tlpNewGroup
        '
        Me.tlpNewGroup.ColumnCount = 2
        Me.tlpNewGroup.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 68.84735!))
        Me.tlpNewGroup.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 31.15265!))
        Me.tlpNewGroup.Controls.Add(Me.btnNewGroup, 0, 0)
        Me.tlpNewGroup.Controls.Add(Me.btnCreateGroup, 1, 3)
        Me.tlpNewGroup.Controls.Add(Me.txtGroupName, 0, 3)
        Me.tlpNewGroup.Controls.Add(Me.dgvForNewMG, 0, 2)
        Me.tlpNewGroup.Controls.Add(Me.btnAddtoGroup, 1, 1)
        Me.tlpNewGroup.Controls.Add(Me.chkSelectallMarkets, 0, 1)
        Me.tlpNewGroup.Location = New System.Drawing.Point(0, 0)
        Me.tlpNewGroup.Name = "tlpNewGroup"
        Me.tlpNewGroup.RowCount = 4
        Me.tlpNewGroup.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 29.0!))
        Me.tlpNewGroup.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tlpNewGroup.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 226.0!))
        Me.tlpNewGroup.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34.0!))
        Me.tlpNewGroup.Size = New System.Drawing.Size(321, 316)
        Me.tlpNewGroup.TabIndex = 29
        '
        'btnAddtoGroup
        '
        Me.btnAddtoGroup.Location = New System.Drawing.Point(223, 32)
        Me.btnAddtoGroup.Name = "btnAddtoGroup"
        Me.btnAddtoGroup.Size = New System.Drawing.Size(95, 21)
        Me.btnAddtoGroup.TabIndex = 28
        Me.btnAddtoGroup.Text = "Add to Group"
        Me.btnAddtoGroup.UseVisualStyleBackColor = True
        '
        'chkSelectallMarkets
        '
        Me.chkSelectallMarkets.AutoSize = True
        Me.chkSelectallMarkets.Location = New System.Drawing.Point(3, 32)
        Me.chkSelectallMarkets.Name = "chkSelectallMarkets"
        Me.chkSelectallMarkets.Size = New System.Drawing.Size(70, 17)
        Me.chkSelectallMarkets.TabIndex = 27
        Me.chkSelectallMarkets.Text = "Select All"
        Me.chkSelectallMarkets.UseVisualStyleBackColor = True
        '
        'flpNewGroup
        '
        Me.flpNewGroup.Controls.Add(Me.tlpNewGroup)
        Me.flpNewGroup.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.flpNewGroup.Location = New System.Drawing.Point(0, 461)
        Me.flpNewGroup.Margin = New System.Windows.Forms.Padding(0)
        Me.flpNewGroup.Name = "flpNewGroup"
        Me.flpNewGroup.Size = New System.Drawing.Size(321, 29)
        Me.flpNewGroup.TabIndex = 30
        '
        'ucMkets
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.flpNewGroup)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "ucMkets"
        Me.Size = New System.Drawing.Size(321, 490)
        CType(Me.dgvMarkets, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSelectedMarkets, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.tpMarkets.ResumeLayout(False)
        Me.tpMarkets.PerformLayout()
        Me.tpGroups.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        CType(Me.dgvForNewMG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tlpNewGroup.ResumeLayout(False)
        Me.tlpNewGroup.PerformLayout()
        Me.flpNewGroup.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents chkSelectAll As System.Windows.Forms.CheckBox
    Friend WithEvents dgvMarkets As System.Windows.Forms.DataGridView
    Friend WithEvents lbMarketGroup As System.Windows.Forms.ListBox
    Friend WithEvents dgvSelectedMarkets As System.Windows.Forms.DataGridView
    Friend WithEvents btnSetPlanMG As System.Windows.Forms.Button
    Friend WithEvents lbPlan As System.Windows.Forms.ListBox
    Friend WithEvents lbPlanning As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents tpMarkets As System.Windows.Forms.TabPage
    Friend WithEvents tpGroups As System.Windows.Forms.TabPage
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnCreateGroup As System.Windows.Forms.Button
    Friend WithEvents txtGroupName As System.Windows.Forms.TextBox
    Friend WithEvents dgvForNewMG As System.Windows.Forms.DataGridView
    Friend WithEvents btnNewGroup As System.Windows.Forms.Button
    Friend WithEvents tlpNewGroup As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents flpNewGroup As System.Windows.Forms.Panel
    Friend WithEvents chkSelectallMarkets As System.Windows.Forms.CheckBox
    Friend WithEvents btnAddtoGroup As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox

End Class
