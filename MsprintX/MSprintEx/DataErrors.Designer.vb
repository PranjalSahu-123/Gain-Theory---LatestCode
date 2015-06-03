<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DataErrors
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
        Me.btnErrors = New System.Windows.Forms.Button()
        Me.ErrorRangeBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.Plandata = New MSprintEx.Plandata()
        Me.dgvErrors = New System.Windows.Forms.DataGridView()
        Me.AddressDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewLinkColumn()
        Me.Value = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.ErrorRangeBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Plandata, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvErrors, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnErrors
        '
        Me.btnErrors.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.btnErrors.Dock = System.Windows.Forms.DockStyle.Top
        Me.btnErrors.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnErrors.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnErrors.Location = New System.Drawing.Point(0, 0)
        Me.btnErrors.Name = "btnErrors"
        Me.btnErrors.Size = New System.Drawing.Size(277, 23)
        Me.btnErrors.TabIndex = 2
        Me.btnErrors.Text = "Errors"
        Me.btnErrors.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnErrors.UseVisualStyleBackColor = False
        Me.btnErrors.Visible = False
        '
        'ErrorRangeBindingSource
        '
        Me.ErrorRangeBindingSource.DataMember = "ErrorRange"
        Me.ErrorRangeBindingSource.DataSource = Me.Plandata
        '
        'Plandata
        '
        Me.Plandata.DataSetName = "Plandata"
        Me.Plandata.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'dgvErrors
        '
        Me.dgvErrors.AllowUserToAddRows = False
        Me.dgvErrors.AllowUserToDeleteRows = False
        Me.dgvErrors.AutoGenerateColumns = False
        Me.dgvErrors.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvErrors.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.AddressDataGridViewTextBoxColumn, Me.Value})
        Me.dgvErrors.DataSource = Me.ErrorRangeBindingSource
        Me.dgvErrors.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvErrors.Location = New System.Drawing.Point(0, 23)
        Me.dgvErrors.MultiSelect = False
        Me.dgvErrors.Name = "dgvErrors"
        Me.dgvErrors.ReadOnly = True
        Me.dgvErrors.RowHeadersVisible = False
        Me.dgvErrors.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvErrors.ShowEditingIcon = False
        Me.dgvErrors.Size = New System.Drawing.Size(277, 275)
        Me.dgvErrors.TabIndex = 1
        '
        'AddressDataGridViewTextBoxColumn
        '
        Me.AddressDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.AddressDataGridViewTextBoxColumn.DataPropertyName = "Address"
        Me.AddressDataGridViewTextBoxColumn.HeaderText = "Address"
        Me.AddressDataGridViewTextBoxColumn.Name = "AddressDataGridViewTextBoxColumn"
        Me.AddressDataGridViewTextBoxColumn.ReadOnly = True
        Me.AddressDataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.AddressDataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'Value
        '
        Me.Value.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Value.DataPropertyName = "Value"
        Me.Value.HeaderText = "Value"
        Me.Value.Name = "Value"
        Me.Value.ReadOnly = True
        '
        'DataErrors
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.dgvErrors)
        Me.Controls.Add(Me.btnErrors)
        Me.Name = "DataErrors"
        Me.Size = New System.Drawing.Size(277, 298)
        CType(Me.ErrorRangeBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Plandata, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvErrors, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnErrors As System.Windows.Forms.Button
    Friend WithEvents ErrorRangeBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents dgvErrors As System.Windows.Forms.DataGridView
    Friend WithEvents AddressDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents Value As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Plandata As MSprintEx.Plandata

End Class
