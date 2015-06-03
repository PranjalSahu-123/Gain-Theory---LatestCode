<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucVariant
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbFilterMasterGenres = New System.Windows.Forms.TextBox()
        Me.ClbSelectAll = New System.Windows.Forms.CheckBox()
        Me.lbSelectVariant = New System.Windows.Forms.Label()
        Me.clbSelectVariant = New System.Windows.Forms.ListBox()
        Me.lbChosenVariants = New System.Windows.Forms.Label()
        Me.btnClearAll = New System.Windows.Forms.Button()
        Me.lbSelectedVariants = New System.Windows.Forms.ListBox()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.tbFilterMasterGenres)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(248, 58)
        Me.GroupBox1.TabIndex = 37
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Search/Filter"
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
        'ClbSelectAll
        '
        Me.ClbSelectAll.AutoSize = True
        Me.ClbSelectAll.Location = New System.Drawing.Point(3, 64)
        Me.ClbSelectAll.Name = "ClbSelectAll"
        Me.ClbSelectAll.Size = New System.Drawing.Size(40, 17)
        Me.ClbSelectAll.TabIndex = 41
        Me.ClbSelectAll.Text = "All "
        Me.ClbSelectAll.UseVisualStyleBackColor = True
        '
        'lbSelectVariant
        '
        Me.lbSelectVariant.AutoSize = True
        Me.lbSelectVariant.Location = New System.Drawing.Point(49, 65)
        Me.lbSelectVariant.Name = "lbSelectVariant"
        Me.lbSelectVariant.Size = New System.Drawing.Size(90, 13)
        Me.lbSelectVariant.TabIndex = 42
        Me.lbSelectVariant.Text = "Choose Variant(s)"
        '
        'clbSelectVariant
        '
        Me.clbSelectVariant.FormattingEnabled = True
        Me.clbSelectVariant.HorizontalScrollbar = True
        Me.clbSelectVariant.Location = New System.Drawing.Point(3, 81)
        Me.clbSelectVariant.Name = "clbSelectVariant"
        Me.clbSelectVariant.Size = New System.Drawing.Size(167, 290)
        Me.clbSelectVariant.Sorted = True
        Me.clbSelectVariant.TabIndex = 43
        '
        'lbChosenVariants
        '
        Me.lbChosenVariants.AutoSize = True
        Me.lbChosenVariants.Location = New System.Drawing.Point(183, 65)
        Me.lbChosenVariants.Name = "lbChosenVariants"
        Me.lbChosenVariants.Size = New System.Drawing.Size(96, 13)
        Me.lbChosenVariants.TabIndex = 44
        Me.lbChosenVariants.Text = "Selected Variant(s)"
        '
        'btnClearAll
        '
        Me.btnClearAll.Location = New System.Drawing.Point(285, 55)
        Me.btnClearAll.Name = "btnClearAll"
        Me.btnClearAll.Size = New System.Drawing.Size(39, 23)
        Me.btnClearAll.TabIndex = 47
        Me.btnClearAll.Text = "Clear All"
        Me.btnClearAll.UseVisualStyleBackColor = True
        '
        'lbSelectedVariants
        '
        Me.lbSelectedVariants.FormattingEnabled = True
        Me.lbSelectedVariants.HorizontalScrollbar = True
        Me.lbSelectedVariants.Location = New System.Drawing.Point(176, 81)
        Me.lbSelectedVariants.Name = "lbSelectedVariants"
        Me.lbSelectedVariants.Size = New System.Drawing.Size(152, 277)
        Me.lbSelectedVariants.Sorted = True
        Me.lbSelectedVariants.TabIndex = 46
        '
        'ucVariant
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.btnClearAll)
        Me.Controls.Add(Me.lbSelectedVariants)
        Me.Controls.Add(Me.lbChosenVariants)
        Me.Controls.Add(Me.clbSelectVariant)
        Me.Controls.Add(Me.lbSelectVariant)
        Me.Controls.Add(Me.ClbSelectAll)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "ucVariant"
        Me.Size = New System.Drawing.Size(344, 383)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbFilterMasterGenres As System.Windows.Forms.TextBox
    Friend WithEvents ClbSelectAll As System.Windows.Forms.CheckBox
    Friend WithEvents lbSelectVariant As System.Windows.Forms.Label
    Friend WithEvents clbSelectVariant As System.Windows.Forms.ListBox
    Friend WithEvents lbChosenVariants As System.Windows.Forms.Label
    Friend WithEvents btnClearAll As System.Windows.Forms.Button
    Friend WithEvents lbSelectedVariants As System.Windows.Forms.ListBox

End Class
