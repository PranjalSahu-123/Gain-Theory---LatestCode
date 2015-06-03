<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ChannelMapping
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
        Me.btnChannels = New System.Windows.Forms.Button()
        Me.dgvChannels = New System.Windows.Forms.DataGridView()
        Me.ChannelNameDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ChannelCodeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.ChannelMasterBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.Plandata = New MSprintEx.Plandata()
        Me.More = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.PlanChannelsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        CType(Me.dgvChannels, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ChannelMasterBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Plandata, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PlanChannelsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnChannels
        '
        Me.btnChannels.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.btnChannels.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnChannels.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnChannels.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnChannels.Location = New System.Drawing.Point(0, 0)
        Me.btnChannels.Margin = New System.Windows.Forms.Padding(0)
        Me.btnChannels.Name = "btnChannels"
        Me.btnChannels.Size = New System.Drawing.Size(292, 23)
        Me.btnChannels.TabIndex = 3
        Me.btnChannels.Text = "Channel Mapping"
        Me.btnChannels.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnChannels.UseVisualStyleBackColor = False
        Me.btnChannels.Visible = False
        '
        'dgvChannels
        '
        Me.dgvChannels.AllowUserToAddRows = False
        Me.dgvChannels.AllowUserToDeleteRows = False
        Me.dgvChannels.AutoGenerateColumns = False
        Me.dgvChannels.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvChannels.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ChannelNameDataGridViewTextBoxColumn, Me.ChannelCodeDataGridViewTextBoxColumn, Me.More})
        Me.dgvChannels.DataSource = Me.PlanChannelsBindingSource
        Me.dgvChannels.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvChannels.Location = New System.Drawing.Point(0, 23)
        Me.dgvChannels.Margin = New System.Windows.Forms.Padding(0)
        Me.dgvChannels.Name = "dgvChannels"
        Me.dgvChannels.RowHeadersVisible = False
        Me.dgvChannels.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvChannels.Size = New System.Drawing.Size(292, 293)
        Me.dgvChannels.StandardTab = True
        Me.dgvChannels.TabIndex = 4
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
        Me.ChannelCodeDataGridViewTextBoxColumn.DataPropertyName = "ChannelCode"
        Me.ChannelCodeDataGridViewTextBoxColumn.DataSource = Me.ChannelMasterBindingSource
        Me.ChannelCodeDataGridViewTextBoxColumn.DisplayMember = "ChannelName"
        Me.ChannelCodeDataGridViewTextBoxColumn.DropDownWidth = 200
        Me.ChannelCodeDataGridViewTextBoxColumn.HeaderText = "Name in Master"
        Me.ChannelCodeDataGridViewTextBoxColumn.Name = "ChannelCodeDataGridViewTextBoxColumn"
        Me.ChannelCodeDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ChannelCodeDataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ChannelCodeDataGridViewTextBoxColumn.ValueMember = "ChannelCode"
        '
        'ChannelMasterBindingSource
        '
        Me.ChannelMasterBindingSource.DataMember = "ChannelMaster"
        Me.ChannelMasterBindingSource.DataSource = Me.Plandata
        '
        'Plandata
        '
        Me.Plandata.DataSetName = "Plandata"
        Me.Plandata.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'More
        '
        Me.More.HeaderText = ""
        Me.More.Name = "More"
        Me.More.Text = "..."
        Me.More.Width = 20
        '
        'PlanChannelsBindingSource
        '
        Me.PlanChannelsBindingSource.DataMember = "PlanChannels"
        Me.PlanChannelsBindingSource.DataSource = Me.Plandata
        Me.PlanChannelsBindingSource.Sort = "ChannelName"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.btnChannels, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.dgvChannels, 0, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(292, 316)
        Me.TableLayoutPanel1.TabIndex = 6
        '
        'ChannelMapping
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Margin = New System.Windows.Forms.Padding(0)
        Me.Name = "ChannelMapping"
        Me.Size = New System.Drawing.Size(292, 316)
        CType(Me.dgvChannels, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ChannelMasterBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Plandata, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PlanChannelsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnChannels As System.Windows.Forms.Button
    Friend WithEvents dgvChannels As System.Windows.Forms.DataGridView
    Friend WithEvents ChannelMasterBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Plandata As MSprintEx.Plandata
    Friend WithEvents PlanChannelsBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents ChannelNameDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ChannelCodeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents More As System.Windows.Forms.DataGridViewButtonColumn

End Class
