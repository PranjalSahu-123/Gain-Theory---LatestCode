<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucChannels
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
        Me.lbSelectChannels = New System.Windows.Forms.Label()
        Me.lbChosenChannels = New System.Windows.Forms.Label()
        Me.clbSelectChannels = New System.Windows.Forms.ListBox()
        Me.lbSelectedChannels = New System.Windows.Forms.ListBox()
        Me.ClbSelectAll = New System.Windows.Forms.CheckBox()
        Me.btnClearAll = New System.Windows.Forms.Button()
        Me.LbTopPrograms = New System.Windows.Forms.Label()
        Me.nudTopPrograms = New System.Windows.Forms.NumericUpDown()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.tbFilterMasterGenres = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbGenre = New System.Windows.Forms.ComboBox()
        CType(Me.nudTopPrograms, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lbSelectChannels
        '
        Me.lbSelectChannels.AutoSize = True
        Me.lbSelectChannels.Location = New System.Drawing.Point(43, 83)
        Me.lbSelectChannels.Name = "lbSelectChannels"
        Me.lbSelectChannels.Size = New System.Drawing.Size(96, 13)
        Me.lbSelectChannels.TabIndex = 17
        Me.lbSelectChannels.Text = "Choose Channel(s)"
        '
        'lbChosenChannels
        '
        Me.lbChosenChannels.AutoSize = True
        Me.lbChosenChannels.Location = New System.Drawing.Point(177, 81)
        Me.lbChosenChannels.Name = "lbChosenChannels"
        Me.lbChosenChannels.Size = New System.Drawing.Size(102, 13)
        Me.lbChosenChannels.TabIndex = 22
        Me.lbChosenChannels.Text = "Selected Channel(s)"
        '
        'clbSelectChannels
        '
        Me.clbSelectChannels.FormattingEnabled = True
        Me.clbSelectChannels.Location = New System.Drawing.Point(6, 101)
        Me.clbSelectChannels.Name = "clbSelectChannels"
        Me.clbSelectChannels.Size = New System.Drawing.Size(177, 277)
        Me.clbSelectChannels.TabIndex = 24
        '
        'lbSelectedChannels
        '
        Me.lbSelectedChannels.FormattingEnabled = True
        Me.lbSelectedChannels.Location = New System.Drawing.Point(186, 101)
        Me.lbSelectedChannels.Name = "lbSelectedChannels"
        Me.lbSelectedChannels.Size = New System.Drawing.Size(129, 199)
        Me.lbSelectedChannels.TabIndex = 23
        '
        'ClbSelectAll
        '
        Me.ClbSelectAll.AutoSize = True
        Me.ClbSelectAll.Location = New System.Drawing.Point(3, 82)
        Me.ClbSelectAll.Name = "ClbSelectAll"
        Me.ClbSelectAll.Size = New System.Drawing.Size(40, 17)
        Me.ClbSelectAll.TabIndex = 26
        Me.ClbSelectAll.Text = "All "
        Me.ClbSelectAll.UseVisualStyleBackColor = True
        '
        'btnClearAll
        '
        Me.btnClearAll.Location = New System.Drawing.Point(285, 76)
        Me.btnClearAll.Name = "btnClearAll"
        Me.btnClearAll.Size = New System.Drawing.Size(39, 23)
        Me.btnClearAll.TabIndex = 27
        Me.btnClearAll.Text = "Clear All"
        Me.btnClearAll.UseVisualStyleBackColor = True
        '
        'LbTopPrograms
        '
        Me.LbTopPrograms.AutoSize = True
        Me.LbTopPrograms.Location = New System.Drawing.Point(189, 314)
        Me.LbTopPrograms.Name = "LbTopPrograms"
        Me.LbTopPrograms.Size = New System.Drawing.Size(117, 13)
        Me.LbTopPrograms.TabIndex = 28
        Me.LbTopPrograms.Text = "Top Programs required:"
        '
        'nudTopPrograms
        '
        Me.nudTopPrograms.Location = New System.Drawing.Point(192, 330)
        Me.nudTopPrograms.Name = "nudTopPrograms"
        Me.nudTopPrograms.Size = New System.Drawing.Size(45, 20)
        Me.nudTopPrograms.TabIndex = 29
        Me.nudTopPrograms.Value = New Decimal(New Integer() {20, 0, 0, 0})
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.clbSelectChannels)
        Me.Panel1.Controls.Add(Me.nudTopPrograms)
        Me.Panel1.Controls.Add(Me.ClbSelectAll)
        Me.Panel1.Controls.Add(Me.LbTopPrograms)
        Me.Panel1.Controls.Add(Me.lbChosenChannels)
        Me.Panel1.Controls.Add(Me.lbSelectedChannels)
        Me.Panel1.Controls.Add(Me.btnClearAll)
        Me.Panel1.Controls.Add(Me.lbSelectChannels)
        Me.Panel1.Location = New System.Drawing.Point(3, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(333, 514)
        Me.Panel1.TabIndex = 30
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.tbFilterMasterGenres)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.cbGenre)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(321, 73)
        Me.GroupBox1.TabIndex = 34
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Search"
        '
        'tbFilterMasterGenres
        '
        Me.tbFilterMasterGenres.Location = New System.Drawing.Point(3, 46)
        Me.tbFilterMasterGenres.Name = "tbFilterMasterGenres"
        Me.tbFilterMasterGenres.Size = New System.Drawing.Size(130, 20)
        Me.tbFilterMasterGenres.TabIndex = 33
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(136, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(42, 13)
        Me.Label1.TabIndex = 30
        Me.Label1.Text = "Genre :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(-1, 18)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(41, 13)
        Me.Label2.TabIndex = 32
        Me.Label2.Text = "Name :"
        '
        'cbGenre
        '
        Me.cbGenre.FormattingEnabled = True
        Me.cbGenre.Location = New System.Drawing.Point(139, 46)
        Me.cbGenre.Name = "cbGenre"
        Me.cbGenre.Size = New System.Drawing.Size(172, 21)
        Me.cbGenre.TabIndex = 31
        Me.cbGenre.Text = "--Select--"
        '
        'ucChannels
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Panel1)
        Me.Name = "ucChannels"
        Me.Size = New System.Drawing.Size(446, 426)
        CType(Me.nudTopPrograms, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lbSelectChannels As System.Windows.Forms.Label
    '   Friend WithEvents tbFilterMasterGenres As System.Windows.Forms.TextBox
    Friend WithEvents lbChosenChannels As System.Windows.Forms.Label
    Friend WithEvents clbSelectChannels As System.Windows.Forms.ListBox
    Friend WithEvents lbSelectedChannels As System.Windows.Forms.ListBox
    Friend WithEvents ClbSelectAll As System.Windows.Forms.CheckBox
    Friend WithEvents btnClearAll As System.Windows.Forms.Button
    Friend WithEvents LbTopPrograms As System.Windows.Forms.Label
    Friend WithEvents nudTopPrograms As System.Windows.Forms.NumericUpDown
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbGenre As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbFilterMasterGenres As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox

End Class
