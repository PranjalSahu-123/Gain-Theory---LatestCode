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
        Me.btnPushOneToSelected = New System.Windows.Forms.Button()
        Me.dgvAvailableSpotsGrid = New System.Windows.Forms.DataGridView()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.SecTableAdapter1 = New MSprintEx.METISTableAdapters.SECTableAdapter()
        CType(Me.dgvAvailableSpotsGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnPushOneToSelected
        '
        Me.btnPushOneToSelected.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnPushOneToSelected.Location = New System.Drawing.Point(656, 347)
        Me.btnPushOneToSelected.Name = "btnPushOneToSelected"
        Me.btnPushOneToSelected.Size = New System.Drawing.Size(83, 22)
        Me.btnPushOneToSelected.TabIndex = 3
        Me.btnPushOneToSelected.Text = "Replace"
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
        Me.dgvAvailableSpotsGrid.Size = New System.Drawing.Size(736, 310)
        Me.dgvAvailableSpotsGrid.TabIndex = 7
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
        Me.SplitContainer2.Location = New System.Drawing.Point(3, 17)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.TableLayoutPanel1)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.TableLayoutPanel2)
        Me.SplitContainer2.Size = New System.Drawing.Size(775, 376)
        Me.SplitContainer2.SplitterDistance = 25
        Me.SplitContainer2.TabIndex = 11
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 3
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(21, 372)
        Me.TableLayoutPanel1.TabIndex = 14
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 1
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.dgvAvailableSpotsGrid, 0, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.btnPushOneToSelected, 0, 2)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 3
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(742, 372)
        Me.TableLayoutPanel2.TabIndex = 15
        '
        'Panel1
        '
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(3, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(775, 8)
        Me.Panel1.TabIndex = 13
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
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 14.0!))
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
        CType(Me.dgvAvailableSpotsGrid, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnPushOneToSelected As System.Windows.Forms.Button
    Friend WithEvents dgvAvailableSpotsGrid As System.Windows.Forms.DataGridView
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents TableLayoutPanel3 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents SecTableAdapter1 As MSprintEx.METISTableAdapters.SECTableAdapter

End Class
