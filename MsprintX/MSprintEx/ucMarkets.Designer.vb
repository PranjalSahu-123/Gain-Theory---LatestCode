<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucMarkets
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.dgvMarkets = New System.Windows.Forms.DataGridView()
        Me.MarketsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.btnCreateGroup = New System.Windows.Forms.Button()
        Me.dgvSelectedMarkets = New System.Windows.Forms.DataGridView()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.lbMarketGroup = New System.Windows.Forms.ListBox()
        Me.txtGroupName = New System.Windows.Forms.TextBox()
        Me.chkSelectAll = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()

        CType(Me.dgvMarkets, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MarketsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSelectedMarkets, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvMarkets
        '
        Me.dgvMarkets.AllowUserToAddRows = False
        Me.dgvMarkets.AllowUserToDeleteRows = False
        Me.dgvMarkets.AllowUserToOrderColumns = True
        Me.dgvMarkets.AllowUserToResizeRows = False
        Me.dgvMarkets.AutoGenerateColumns = False
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
        Me.dgvMarkets.DataSource = Me.MarketsBindingSource
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvMarkets.DefaultCellStyle = DataGridViewCellStyle2
        '  Me.dgvMarkets.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvMarkets.Location = New System.Drawing.Point(0, 23)
        Me.dgvMarkets.Margin = New System.Windows.Forms.Padding(0)
        Me.dgvMarkets.Name = "dgvMarkets"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvMarkets.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvMarkets.RowHeadersVisible = False
        Me.TableLayoutPanel1.SetRowSpan(Me.dgvMarkets, 3)
        Me.dgvMarkets.RowTemplate.Height = 20
        Me.dgvMarkets.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvMarkets.Size = New System.Drawing.Size(220, 367)
        Me.dgvMarkets.TabIndex = 1
        '
        'MarketsBindingSource
        '
        Me.MarketsBindingSource.DataMember = "Markets"
        '
        'btnCreateGroup
        '
        '   Me.btnCreateGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnCreateGroup.Location = New System.Drawing.Point(489, 195)
        Me.btnCreateGroup.Margin = New System.Windows.Forms.Padding(0)
        Me.btnCreateGroup.Name = "btnCreateGroup"
        Me.btnCreateGroup.Size = New System.Drawing.Size(151, 23)
        Me.btnCreateGroup.TabIndex = 4
        Me.btnCreateGroup.Text = "Create Group"
        Me.btnCreateGroup.UseVisualStyleBackColor = True
        '
        'dgvSelectedMarkets
        '
        Me.dgvSelectedMarkets.AllowUserToAddRows = False
        Me.dgvSelectedMarkets.AllowUserToDeleteRows = False
        Me.dgvSelectedMarkets.AutoGenerateColumns = False
        Me.dgvSelectedMarkets.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSelectedMarkets.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvSelectedMarkets.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.TableLayoutPanel1.SetColumnSpan(Me.dgvSelectedMarkets, 3)
        Me.dgvSelectedMarkets.DataSource = Me.MarketsBindingSource
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvSelectedMarkets.DefaultCellStyle = DataGridViewCellStyle5
        Me.dgvSelectedMarkets.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSelectedMarkets.Location = New System.Drawing.Point(220, 0)
        Me.dgvSelectedMarkets.Margin = New System.Windows.Forms.Padding(0)
        Me.dgvSelectedMarkets.MultiSelect = False
        Me.dgvSelectedMarkets.Name = "dgvSelectedMarkets"
        Me.dgvSelectedMarkets.ReadOnly = True
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvSelectedMarkets.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.dgvSelectedMarkets.RowHeadersVisible = False
        Me.TableLayoutPanel1.SetRowSpan(Me.dgvSelectedMarkets, 2)
        Me.dgvSelectedMarkets.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvSelectedMarkets.Size = New System.Drawing.Size(420, 195)
        Me.dgvSelectedMarkets.TabIndex = 2
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 4
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 220.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 89.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 151.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.lbMarketGroup, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.btnCreateGroup, 3, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.txtGroupName, 2, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.dgvMarkets, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.chkSelectAll, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.dgvSelectedMarkets, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 1, 2)
        '  Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.AllowDrop = True
        Me.TableLayoutPanel1.RowCount = 5
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 23.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 23.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 37.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(640, 428)
        Me.TableLayoutPanel1.TabIndex = 5
        '
        'lbMarketGroup
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.lbMarketGroup, 3)
        Me.lbMarketGroup.FormattingEnabled = True
        Me.lbMarketGroup.Location = New System.Drawing.Point(220, 218)
        Me.lbMarketGroup.Margin = New System.Windows.Forms.Padding(0)
        Me.lbMarketGroup.Name = "lbMarketGroup"
        Me.lbMarketGroup.Size = New System.Drawing.Size(420, 121)
        Me.lbMarketGroup.TabIndex = 7
        '
        'txtGroupName
        '
        '   Me.txtGroupName.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtGroupName.Location = New System.Drawing.Point(309, 197)
        Me.txtGroupName.Margin = New System.Windows.Forms.Padding(0, 2, 0, 0)
        Me.txtGroupName.Name = "txtGroupName"
        Me.txtGroupName.Size = New System.Drawing.Size(180, 20)
        Me.txtGroupName.TabIndex = 3
        '
        'chkSelectAll
        '
        '   Me.chkSelectAll.Dock = System.Windows.Forms.DockStyle.Fill
        Me.chkSelectAll.Location = New System.Drawing.Point(3, 3)
        Me.chkSelectAll.Name = "chkSelectAll"
        Me.chkSelectAll.Size = New System.Drawing.Size(214, 17)
        Me.chkSelectAll.TabIndex = 0
        Me.chkSelectAll.Text = "Select all"
        Me.chkSelectAll.UseVisualStyleBackColor = True
        '
        'Label1
        '
        ' Me.Label1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label1.Location = New System.Drawing.Point(223, 195)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(83, 23)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Group Name"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ucMarkets
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "ucMarkets"
        Me.Size = New System.Drawing.Size(640, 428)
        CType(Me.dgvMarkets, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MarketsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSelectedMarkets, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnCreateGroup As System.Windows.Forms.Button
    Friend WithEvents dgvMarkets As System.Windows.Forms.DataGridView
    Friend WithEvents MarketsBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents METIS As MSprintEx.METIS
    Friend WithEvents MarketsTableAdapter As MSprintEx.METISTableAdapters.MarketsTableAdapter
    Friend WithEvents dgvSelectedMarkets As System.Windows.Forms.DataGridView
    Friend WithEvents MarketdescDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents txtGroupName As System.Windows.Forms.TextBox
    Friend WithEvents SelectedDataGridViewCheckBoxColumn As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents MarketdescDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents chkSelectAll As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lbMarketGroup As System.Windows.Forms.ListBox

End Class
