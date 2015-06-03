<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucBrand
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
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbFilterMasterGenres = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ClbSelectAll = New System.Windows.Forms.CheckBox()
        Me.lbSelectBrands = New System.Windows.Forms.Label()
        Me.lbChosenBrands = New System.Windows.Forms.Label()
        Me.clbSelectBrands = New System.Windows.Forms.ListBox()
        Me.lbSelectedBrands = New System.Windows.Forms.ListBox()
        Me.btnClearAll = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(32, 25)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(41, 13)
        Me.Label2.TabIndex = 33
        Me.Label2.Text = "Name :"
        '
        'tbFilterMasterGenres
        '
        Me.tbFilterMasterGenres.Location = New System.Drawing.Point(79, 22)
        Me.tbFilterMasterGenres.Name = "tbFilterMasterGenres"
        Me.tbFilterMasterGenres.Size = New System.Drawing.Size(130, 20)
        Me.tbFilterMasterGenres.TabIndex = 34
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.tbFilterMasterGenres)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(248, 58)
        Me.GroupBox1.TabIndex = 35
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Search/Filter"
        '
        'ClbSelectAll
        '
        Me.ClbSelectAll.AutoSize = True
        Me.ClbSelectAll.Location = New System.Drawing.Point(3, 67)
        Me.ClbSelectAll.Name = "ClbSelectAll"
        Me.ClbSelectAll.Size = New System.Drawing.Size(40, 17)
        Me.ClbSelectAll.TabIndex = 36
        Me.ClbSelectAll.Text = "All "
        Me.ClbSelectAll.UseVisualStyleBackColor = True
        '
        'lbSelectBrands
        '
        Me.lbSelectBrands.AutoSize = True
        Me.lbSelectBrands.Location = New System.Drawing.Point(49, 68)
        Me.lbSelectBrands.Name = "lbSelectBrands"
        Me.lbSelectBrands.Size = New System.Drawing.Size(85, 13)
        Me.lbSelectBrands.TabIndex = 37
        Me.lbSelectBrands.Text = "Choose Brand(s)"
        '
        'lbChosenBrands
        '
        Me.lbChosenBrands.AutoSize = True
        Me.lbChosenBrands.Location = New System.Drawing.Point(189, 68)
        Me.lbChosenBrands.Name = "lbChosenBrands"
        Me.lbChosenBrands.Size = New System.Drawing.Size(91, 13)
        Me.lbChosenBrands.TabIndex = 38
        Me.lbChosenBrands.Text = "Selected Brand(s)"
        '
        'clbSelectBrands
        '
        Me.clbSelectBrands.FormattingEnabled = True
        Me.clbSelectBrands.HorizontalScrollbar = True
        Me.clbSelectBrands.Location = New System.Drawing.Point(3, 87)
        Me.clbSelectBrands.Name = "clbSelectBrands"
        Me.clbSelectBrands.Size = New System.Drawing.Size(177, 277)
        Me.clbSelectBrands.Sorted = True
        Me.clbSelectBrands.TabIndex = 39
        '
        'lbSelectedBrands
        '
        Me.lbSelectedBrands.FormattingEnabled = True
        Me.lbSelectedBrands.HorizontalScrollbar = True
        Me.lbSelectedBrands.Location = New System.Drawing.Point(192, 87)
        Me.lbSelectedBrands.Name = "lbSelectedBrands"
        Me.lbSelectedBrands.Size = New System.Drawing.Size(156, 277)
        Me.lbSelectedBrands.Sorted = True
        Me.lbSelectedBrands.TabIndex = 40
        '
        'btnClearAll
        '
        Me.btnClearAll.Location = New System.Drawing.Point(295, 63)
        Me.btnClearAll.Name = "btnClearAll"
        Me.btnClearAll.Size = New System.Drawing.Size(39, 23)
        Me.btnClearAll.TabIndex = 41
        Me.btnClearAll.Text = "Clear All"
        Me.btnClearAll.UseVisualStyleBackColor = True
        '
        'ucBrand
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.btnClearAll)
        Me.Controls.Add(Me.lbSelectedBrands)
        Me.Controls.Add(Me.clbSelectBrands)
        Me.Controls.Add(Me.lbChosenBrands)
        Me.Controls.Add(Me.lbSelectBrands)
        Me.Controls.Add(Me.ClbSelectAll)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "ucBrand"
        Me.Size = New System.Drawing.Size(359, 374)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbFilterMasterGenres As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ClbSelectAll As System.Windows.Forms.CheckBox
    Friend WithEvents lbSelectBrands As System.Windows.Forms.Label
    Friend WithEvents lbChosenBrands As System.Windows.Forms.Label
    Friend WithEvents clbSelectBrands As System.Windows.Forms.ListBox
    Friend WithEvents lbSelectedBrands As System.Windows.Forms.ListBox
    Friend WithEvents btnClearAll As System.Windows.Forms.Button

End Class
