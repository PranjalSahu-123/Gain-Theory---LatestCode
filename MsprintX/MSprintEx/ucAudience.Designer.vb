<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucAudience
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.clbGender = New System.Windows.Forms.CheckedListBox()
        Me.clbSEC = New System.Windows.Forms.CheckedListBox()
        Me.clbHH = New System.Windows.Forms.CheckedListBox()
        Me.clbAge = New System.Windows.Forms.CheckedListBox()
        Me.chkSECAll = New System.Windows.Forms.CheckBox()
        Me.chkAgeAll = New System.Windows.Forms.CheckBox()
        Me.btnCreateTG = New System.Windows.Forms.Button()
        Me.txtTGInput = New System.Windows.Forms.TextBox()
        Me.lbTGDefs = New System.Windows.Forms.ListBox()
        Me.scMain = New System.Windows.Forms.SplitContainer()
        Me.lbPredefined = New System.Windows.Forms.Label()
        Me.btnSetasPlan = New System.Windows.Forms.Button()
        Me.DgPlanRefGrid = New System.Windows.Forms.DataGridView()
        Me.Del = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.scMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.scMain.SuspendLayout()
        CType(Me.DgPlanRefGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'clbGender
        '
        Me.clbGender.CheckOnClick = True
        Me.clbGender.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbGender.FormattingEnabled = True
        Me.clbGender.Items.AddRange(New Object() {"MALE", "FEMALE"})
        Me.clbGender.Location = New System.Drawing.Point(3, 160)
        Me.clbGender.Name = "clbGender"
        Me.clbGender.Size = New System.Drawing.Size(124, 78)
        Me.clbGender.TabIndex = 0
        '
        'clbSEC
        '
        Me.clbSEC.CheckOnClick = True
        Me.clbSEC.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbSEC.FormattingEnabled = True
        Me.clbSEC.Items.AddRange(New Object() {"A", "B", "C", "D/E"})
        Me.clbSEC.Location = New System.Drawing.Point(3, 269)
        Me.clbSEC.Name = "clbSEC"
        Me.clbSEC.Size = New System.Drawing.Size(124, 64)
        Me.clbSEC.TabIndex = 1
        '
        'clbHH
        '
        Me.clbHH.CheckOnClick = True
        Me.clbHH.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbHH.FormattingEnabled = True
        Me.clbHH.Items.AddRange(New Object() {"CABLE & SATELLITE", "NON CABLE & SATELLITE", "ANALOG", "DIGITAL"})
        Me.clbHH.Location = New System.Drawing.Point(133, 160)
        Me.clbHH.Name = "clbHH"
        Me.clbHH.Size = New System.Drawing.Size(175, 78)
        Me.clbHH.TabIndex = 2
        '
        'clbAge
        '
        Me.clbAge.CheckOnClick = True
        Me.clbAge.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbAge.FormattingEnabled = True
        Me.clbAge.Items.AddRange(New Object() {"4-9", "10-14", "15-24", "25-34", "35-44", "45-54", "55+"})
        Me.clbAge.Location = New System.Drawing.Point(133, 269)
        Me.clbAge.Name = "clbAge"
        Me.clbAge.Size = New System.Drawing.Size(175, 64)
        Me.clbAge.TabIndex = 3
        '
        'chkSECAll
        '
        Me.chkSECAll.AutoSize = True
        Me.chkSECAll.Location = New System.Drawing.Point(3, 244)
        Me.chkSECAll.Name = "chkSECAll"
        Me.chkSECAll.Size = New System.Drawing.Size(98, 17)
        Me.chkSECAll.TabIndex = 4
        Me.chkSECAll.Text = "Select all SECs"
        Me.chkSECAll.UseVisualStyleBackColor = True
        '
        'chkAgeAll
        '
        Me.chkAgeAll.AutoSize = True
        Me.chkAgeAll.Location = New System.Drawing.Point(133, 244)
        Me.chkAgeAll.Name = "chkAgeAll"
        Me.chkAgeAll.Size = New System.Drawing.Size(125, 17)
        Me.chkAgeAll.TabIndex = 5
        Me.chkAgeAll.Text = "Select all age groups"
        Me.chkAgeAll.UseVisualStyleBackColor = True
        '
        'btnCreateTG
        '
        Me.btnCreateTG.Location = New System.Drawing.Point(188, 0)
        Me.btnCreateTG.Name = "btnCreateTG"
        Me.btnCreateTG.Size = New System.Drawing.Size(114, 22)
        Me.btnCreateTG.TabIndex = 6
        Me.btnCreateTG.Text = "Create TG"
        Me.btnCreateTG.UseVisualStyleBackColor = True
        '
        'txtTGInput
        '
        Me.txtTGInput.Location = New System.Drawing.Point(0, 0)
        Me.txtTGInput.Name = "txtTGInput"
        Me.txtTGInput.Size = New System.Drawing.Size(187, 20)
        Me.txtTGInput.TabIndex = 7
        '
        'lbTGDefs
        '
        Me.lbTGDefs.Dock = System.Windows.Forms.DockStyle.Left
        Me.lbTGDefs.FormattingEnabled = True
        Me.lbTGDefs.Location = New System.Drawing.Point(0, 0)
        Me.lbTGDefs.Name = "lbTGDefs"
        Me.lbTGDefs.Size = New System.Drawing.Size(124, 120)
        Me.lbTGDefs.TabIndex = 8
        '
        'scMain
        '
        Me.scMain.Location = New System.Drawing.Point(0, 0)
        Me.scMain.Name = "scMain"
        Me.scMain.Size = New System.Drawing.Size(150, 100)
        Me.scMain.TabIndex = 0
        '
        'lbPredefined
        '
        Me.lbPredefined.AutoSize = True
        Me.TableLayoutPanel1.SetColumnSpan(Me.lbPredefined, 2)
        Me.lbPredefined.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lbPredefined.Location = New System.Drawing.Point(3, 0)
        Me.lbPredefined.Name = "lbPredefined"
        Me.lbPredefined.Size = New System.Drawing.Size(305, 20)
        Me.lbPredefined.TabIndex = 9
        Me.lbPredefined.Text = "Predefined TGs"
        '
        'btnSetasPlan
        '
        Me.btnSetasPlan.Location = New System.Drawing.Point(129, 1)
        Me.btnSetasPlan.Name = "btnSetasPlan"
        Me.btnSetasPlan.Size = New System.Drawing.Size(172, 42)
        Me.btnSetasPlan.TabIndex = 10
        Me.btnSetasPlan.Text = "Set as Planning TG"
        Me.btnSetasPlan.UseVisualStyleBackColor = True
        '
        'DgPlanRefGrid
        '
        Me.DgPlanRefGrid.AllowUserToAddRows = False
        Me.DgPlanRefGrid.AllowUserToDeleteRows = False
        Me.DgPlanRefGrid.AllowUserToResizeRows = False
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DgPlanRefGrid.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.DgPlanRefGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgPlanRefGrid.ColumnHeadersVisible = False
        Me.DgPlanRefGrid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Del})
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DgPlanRefGrid.DefaultCellStyle = DataGridViewCellStyle5
        Me.DgPlanRefGrid.Location = New System.Drawing.Point(129, 49)
        Me.DgPlanRefGrid.Name = "DgPlanRefGrid"
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DgPlanRefGrid.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.DgPlanRefGrid.RowHeadersVisible = False
        Me.DgPlanRefGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgPlanRefGrid.Size = New System.Drawing.Size(172, 60)
        Me.DgPlanRefGrid.TabIndex = 12
        '
        'Del
        '
        Me.Del.HeaderText = "Del"
        Me.Del.Name = "Del"
        Me.Del.Text = "x"
        Me.Del.Width = 20
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Location = New System.Drawing.Point(249, 145)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Size = New System.Drawing.Size(150, 100)
        Me.SplitContainer1.TabIndex = 16
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 42.12218!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 57.87782!))
        Me.TableLayoutPanel1.Controls.Add(Me.lbPredefined, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.clbGender, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.clbHH, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.chkSECAll, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.chkAgeAll, 1, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.clbSEC, 0, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.clbAge, 1, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 0, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel2, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 2)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(3, 3)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 7
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 126.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 11.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 84.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 25.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 70.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(311, 372)
        Me.TableLayoutPanel1.TabIndex = 16
        '
        'Panel1
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.Panel1, 2)
        Me.Panel1.Controls.Add(Me.txtTGInput)
        Me.Panel1.Controls.Add(Me.btnCreateTG)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(3, 339)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(305, 30)
        Me.Panel1.TabIndex = 10
        '
        'Panel2
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.Panel2, 2)
        Me.Panel2.Controls.Add(Me.DgPlanRefGrid)
        Me.Panel2.Controls.Add(Me.lbTGDefs)
        Me.Panel2.Controls.Add(Me.btnSetasPlan)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(3, 23)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(305, 120)
        Me.Panel2.TabIndex = 11
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.TableLayoutPanel1.SetColumnSpan(Me.Label1, 2)
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label1.Location = New System.Drawing.Point(0, 146)
        Me.Label1.Margin = New System.Windows.Forms.Padding(0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(311, 2)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Label1"
        '
        'ucAudience
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "ucAudience"
        Me.Size = New System.Drawing.Size(318, 378)
        CType(Me.scMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.scMain.ResumeLayout(False)
        CType(Me.DgPlanRefGrid, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents clbGender As System.Windows.Forms.CheckedListBox
    Friend WithEvents clbSEC As System.Windows.Forms.CheckedListBox
    Friend WithEvents clbHH As System.Windows.Forms.CheckedListBox
    Friend WithEvents clbAge As System.Windows.Forms.CheckedListBox
    Friend WithEvents chkSECAll As System.Windows.Forms.CheckBox
    Friend WithEvents chkAgeAll As System.Windows.Forms.CheckBox
    Friend WithEvents btnCreateTG As System.Windows.Forms.Button
    Friend WithEvents txtTGInput As System.Windows.Forms.TextBox
    Friend WithEvents lbTGDefs As System.Windows.Forms.ListBox
    Friend WithEvents lbPredefined As System.Windows.Forms.Label
    Friend WithEvents btnSetasPlan As System.Windows.Forms.Button
    Friend WithEvents DgPlanRefGrid As System.Windows.Forms.DataGridView
    Friend WithEvents scMain As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Del As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
