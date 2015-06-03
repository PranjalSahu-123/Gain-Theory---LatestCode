<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UcGenres
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
        Me.lbSelectGenres = New System.Windows.Forms.Label()
        Me.lbChosenGenres = New System.Windows.Forms.Label()
        Me.lbSelectedGenres = New System.Windows.Forms.ListBox()
        Me.tbFilterMasterGenres = New System.Windows.Forms.TextBox()
        Me.clbSelectGenres = New System.Windows.Forms.ListBox()
        Me.chbclearall = New System.Windows.Forms.CheckBox()
        Me.chbSelectAll = New System.Windows.Forms.CheckBox()
        Me.dgshowchannels = New System.Windows.Forms.DataGridView()
        Me.Panel1 = New System.Windows.Forms.Panel()
        CType(Me.dgshowchannels, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lbSelectGenres
        '
        Me.lbSelectGenres.AutoSize = True
        Me.lbSelectGenres.Location = New System.Drawing.Point(3, 0)
        Me.lbSelectGenres.Name = "lbSelectGenres"
        Me.lbSelectGenres.Size = New System.Drawing.Size(144, 13)
        Me.lbSelectGenres.TabIndex = 16
        Me.lbSelectGenres.Text = "Choose Genre(s) from Master"
        '
        'lbChosenGenres
        '
        Me.lbChosenGenres.AutoSize = True
        Me.lbChosenGenres.Location = New System.Drawing.Point(222, 1)
        Me.lbChosenGenres.Name = "lbChosenGenres"
        Me.lbChosenGenres.Size = New System.Drawing.Size(92, 13)
        Me.lbChosenGenres.TabIndex = 17
        Me.lbChosenGenres.Text = "Selected Genre(s)"
        '
        'lbSelectedGenres
        '
        Me.lbSelectedGenres.FormattingEnabled = True
        Me.lbSelectedGenres.Location = New System.Drawing.Point(172, 40)
        Me.lbSelectedGenres.Name = "lbSelectedGenres"
        Me.lbSelectedGenres.Size = New System.Drawing.Size(142, 212)
        Me.lbSelectedGenres.TabIndex = 18
        '
        'tbFilterMasterGenres
        '
        Me.tbFilterMasterGenres.Location = New System.Drawing.Point(51, 16)
        Me.tbFilterMasterGenres.Name = "tbFilterMasterGenres"
        Me.tbFilterMasterGenres.Size = New System.Drawing.Size(143, 20)
        Me.tbFilterMasterGenres.TabIndex = 19
        '
        'clbSelectGenres
        '
        Me.clbSelectGenres.FormattingEnabled = True
        Me.clbSelectGenres.Location = New System.Drawing.Point(6, 40)
        Me.clbSelectGenres.Name = "clbSelectGenres"
        Me.clbSelectGenres.Size = New System.Drawing.Size(160, 212)
        Me.clbSelectGenres.TabIndex = 20
        '
        'chbclearall
        '
        Me.chbclearall.AutoSize = True
        Me.chbclearall.Location = New System.Drawing.Point(200, 16)
        Me.chbclearall.Name = "chbclearall"
        Me.chbclearall.Size = New System.Drawing.Size(67, 17)
        Me.chbclearall.TabIndex = 21
        Me.chbclearall.Text = "Clear All "
        Me.chbclearall.UseVisualStyleBackColor = True
        '
        'chbSelectAll
        '
        Me.chbSelectAll.AutoSize = True
        Me.chbSelectAll.Location = New System.Drawing.Point(8, 20)
        Me.chbSelectAll.Name = "chbSelectAll"
        Me.chbSelectAll.Size = New System.Drawing.Size(37, 17)
        Me.chbSelectAll.TabIndex = 22
        Me.chbSelectAll.Text = "All"
        Me.chbSelectAll.UseVisualStyleBackColor = True
        '
        'dgshowchannels
        '
        Me.dgshowchannels.AllowUserToAddRows = False
        Me.dgshowchannels.AllowUserToDeleteRows = False
        Me.dgshowchannels.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgshowchannels.Location = New System.Drawing.Point(3, 256)
        Me.dgshowchannels.Name = "dgshowchannels"
        Me.dgshowchannels.ReadOnly = True
        Me.dgshowchannels.RowHeadersVisible = False
        Me.dgshowchannels.Size = New System.Drawing.Size(311, 124)
        Me.dgshowchannels.TabIndex = 23
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.clbSelectGenres)
        Me.Panel1.Controls.Add(Me.dgshowchannels)
        Me.Panel1.Controls.Add(Me.chbSelectAll)
        Me.Panel1.Controls.Add(Me.lbSelectedGenres)
        Me.Panel1.Controls.Add(Me.chbclearall)
        Me.Panel1.Controls.Add(Me.lbChosenGenres)
        Me.Panel1.Controls.Add(Me.tbFilterMasterGenres)
        Me.Panel1.Controls.Add(Me.lbSelectGenres)
        Me.Panel1.Location = New System.Drawing.Point(0, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(320, 514)
        Me.Panel1.TabIndex = 24
        '
        'UcGenres
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Panel1)
        Me.Name = "UcGenres"
        Me.Size = New System.Drawing.Size(384, 386)
        CType(Me.dgshowchannels, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lbSelectGenres As System.Windows.Forms.Label
    Friend WithEvents lbChosenGenres As System.Windows.Forms.Label
    Friend WithEvents lbSelectedGenres As System.Windows.Forms.ListBox
    Friend WithEvents tbFilterMasterGenres As System.Windows.Forms.TextBox
    Friend WithEvents clbSelectGenres As System.Windows.Forms.ListBox
    Friend WithEvents chbclearall As System.Windows.Forms.CheckBox
    Friend WithEvents chbSelectAll As System.Windows.Forms.CheckBox
    Friend WithEvents dgshowchannels As System.Windows.Forms.DataGridView
    Friend WithEvents Panel1 As System.Windows.Forms.Panel

End Class
