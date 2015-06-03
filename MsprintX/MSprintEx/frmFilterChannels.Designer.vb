<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFilterChannels
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
        Me.components = New System.ComponentModel.Container()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.dgvChannels = New System.Windows.Forms.DataGridView()
        Me.ChannelNameDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ChannelCodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.ChannelMasterDDBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.Plandata = New MSprintEx.Plandata()
        Me.PlanChannelsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.txtSearch = New System.Windows.Forms.TextBox()
        Me.lbChannelMaster = New System.Windows.Forms.ListBox()
        Me.ChannelMasterBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.dgvChannels, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ChannelMasterDDBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Plandata, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PlanChannelsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ChannelMasterBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 324.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.dgvChannels, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.txtSearch, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.lbChannelMaster, 1, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 3
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(689, 287)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'dgvChannels
        '
        Me.dgvChannels.AllowUserToAddRows = False
        Me.dgvChannels.AllowUserToDeleteRows = False
        Me.dgvChannels.AutoGenerateColumns = False
        Me.dgvChannels.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvChannels.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ChannelNameDataGridViewTextBoxColumn, Me.ChannelCodeDataGridViewTextBoxColumn})
        Me.dgvChannels.DataSource = Me.PlanChannelsBindingSource
        Me.dgvChannels.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvChannels.Location = New System.Drawing.Point(3, 3)
        Me.dgvChannels.Name = "dgvChannels"
        Me.dgvChannels.RowHeadersVisible = False
        Me.TableLayoutPanel1.SetRowSpan(Me.dgvChannels, 2)
        Me.dgvChannels.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvChannels.Size = New System.Drawing.Size(318, 281)
        Me.dgvChannels.StandardTab = True
        Me.dgvChannels.TabIndex = 0
        '
        'ChannelNameDataGridViewTextBoxColumn
        '
        Me.ChannelNameDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.ChannelNameDataGridViewTextBoxColumn.DataPropertyName = "ChannelName"
        Me.ChannelNameDataGridViewTextBoxColumn.HeaderText = "Name in Plan"
        Me.ChannelNameDataGridViewTextBoxColumn.Name = "ChannelNameDataGridViewTextBoxColumn"
        Me.ChannelNameDataGridViewTextBoxColumn.ReadOnly = True
        '
        'ChannelCodeDataGridViewTextBoxColumn
        '
        Me.ChannelCodeDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.ChannelCodeDataGridViewTextBoxColumn.DataPropertyName = "ChannelCode"
        Me.ChannelCodeDataGridViewTextBoxColumn.DataSource = Me.ChannelMasterDDBindingSource
        Me.ChannelCodeDataGridViewTextBoxColumn.DisplayMember = "ChannelName"
        Me.ChannelCodeDataGridViewTextBoxColumn.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.[Nothing]
        Me.ChannelCodeDataGridViewTextBoxColumn.DropDownWidth = 200
        Me.ChannelCodeDataGridViewTextBoxColumn.HeaderText = "Name in Master"
        Me.ChannelCodeDataGridViewTextBoxColumn.Name = "ChannelCodeDataGridViewTextBoxColumn"
        Me.ChannelCodeDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ChannelCodeDataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ChannelCodeDataGridViewTextBoxColumn.ValueMember = "ChannelCode"
        '
        'ChannelMasterDDBindingSource
        '
        Me.ChannelMasterDDBindingSource.DataMember = "ChannelMaster"
        Me.ChannelMasterDDBindingSource.DataSource = Me.Plandata
        '
        'Plandata
        '
        Me.Plandata.DataSetName = "Plandata"
        Me.Plandata.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'PlanChannelsBindingSource
        '
        Me.PlanChannelsBindingSource.DataMember = "PlanChannels"
        Me.PlanChannelsBindingSource.DataSource = Me.Plandata
        Me.PlanChannelsBindingSource.Sort = "ChannelName"
        '
        'txtSearch
        '
        Me.txtSearch.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.txtSearch.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.txtSearch.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtSearch.Location = New System.Drawing.Point(324, 0)
        Me.txtSearch.Margin = New System.Windows.Forms.Padding(0)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(365, 20)
        Me.txtSearch.TabIndex = 1
        '
        'lbChannelMaster
        '
        Me.lbChannelMaster.DataSource = Me.ChannelMasterBindingSource
        Me.lbChannelMaster.DisplayMember = "ChannelName"
        Me.lbChannelMaster.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lbChannelMaster.FormattingEnabled = True
        Me.lbChannelMaster.Location = New System.Drawing.Point(324, 20)
        Me.lbChannelMaster.Margin = New System.Windows.Forms.Padding(0)
        Me.lbChannelMaster.Name = "lbChannelMaster"
        Me.lbChannelMaster.Size = New System.Drawing.Size(365, 267)
        Me.lbChannelMaster.TabIndex = 2
        Me.lbChannelMaster.ValueMember = "ChannelCode"
        '
        'ChannelMasterBindingSource
        '
        Me.ChannelMasterBindingSource.DataMember = "ChannelMaster"
        Me.ChannelMasterBindingSource.DataSource = Me.Plandata
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(278, 317)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'frmFilterChannels
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(689, 287)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmFilterChannels"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Select Channel"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        CType(Me.dgvChannels, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ChannelMasterDDBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Plandata, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PlanChannelsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ChannelMasterBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents txtSearch As System.Windows.Forms.TextBox
    Friend WithEvents lbChannelMaster As System.Windows.Forms.ListBox
    Friend WithEvents ChannelMasterBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Plandata As MSprintEx.Plandata
    Friend WithEvents PlanChannelsBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents dgvChannels As System.Windows.Forms.DataGridView
    Friend WithEvents ChannelMasterDDBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents ChannelNameDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ChannelCodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents btnCancel As System.Windows.Forms.Button
End Class
