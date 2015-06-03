<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucCategory
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
        Me.clbSelectCategories = New System.Windows.Forms.ListBox()
        Me.lbSelectCategories = New System.Windows.Forms.Label()
        Me.ClbSelectAll = New System.Windows.Forms.CheckBox()
        Me.lbSelectedCategories = New System.Windows.Forms.ListBox()
        Me.lbSelectCategory = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.btnClearAll = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.tbFilterMasterGenres)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 13)
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
        'clbSelectCategories
        '
        Me.clbSelectCategories.FormattingEnabled = True
        Me.clbSelectCategories.HorizontalScrollbar = True
        Me.clbSelectCategories.Location = New System.Drawing.Point(3, 99)
        Me.clbSelectCategories.Name = "clbSelectCategories"
        Me.clbSelectCategories.Size = New System.Drawing.Size(177, 290)
        Me.clbSelectCategories.Sorted = True
        Me.clbSelectCategories.TabIndex = 45
        '
        'lbSelectCategories
        '
        Me.lbSelectCategories.AutoSize = True
        Me.lbSelectCategories.Location = New System.Drawing.Point(59, 78)
        Me.lbSelectCategories.Name = "lbSelectCategories"
        Me.lbSelectCategories.Size = New System.Drawing.Size(99, 13)
        Me.lbSelectCategories.TabIndex = 44
        Me.lbSelectCategories.Text = "Choose Category(s)"
        '
        'ClbSelectAll
        '
        Me.ClbSelectAll.AutoSize = True
        Me.ClbSelectAll.Location = New System.Drawing.Point(13, 77)
        Me.ClbSelectAll.Name = "ClbSelectAll"
        Me.ClbSelectAll.Size = New System.Drawing.Size(40, 17)
        Me.ClbSelectAll.TabIndex = 43
        Me.ClbSelectAll.Text = "All "
        Me.ClbSelectAll.UseVisualStyleBackColor = True
        '
        'lbSelectedCategories
        '
        Me.lbSelectedCategories.FormattingEnabled = True
        Me.lbSelectedCategories.HorizontalScrollbar = True
        Me.lbSelectedCategories.Location = New System.Drawing.Point(186, 99)
        Me.lbSelectedCategories.Name = "lbSelectedCategories"
        Me.lbSelectedCategories.Size = New System.Drawing.Size(177, 290)
        Me.lbSelectedCategories.Sorted = True
        Me.lbSelectedCategories.TabIndex = 48
        '
        'lbSelectCategory
        '
        Me.lbSelectCategory.AutoSize = True
        Me.lbSelectCategory.Location = New System.Drawing.Point(217, 76)
        Me.lbSelectCategory.Name = "lbSelectCategory"
        Me.lbSelectCategory.Size = New System.Drawing.Size(105, 13)
        Me.lbSelectCategory.TabIndex = 47
        Me.lbSelectCategory.Text = "Selected Category(s)"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(186, 76)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(40, 17)
        Me.CheckBox1.TabIndex = 46
        Me.CheckBox1.Text = "All "
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'btnClearAll
        '
        Me.btnClearAll.Location = New System.Drawing.Point(324, 70)
        Me.btnClearAll.Name = "btnClearAll"
        Me.btnClearAll.Size = New System.Drawing.Size(39, 23)
        Me.btnClearAll.TabIndex = 49
        Me.btnClearAll.Text = "Clear All"
        Me.btnClearAll.UseVisualStyleBackColor = True
        '
        'ucCategory
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.btnClearAll)
        Me.Controls.Add(Me.lbSelectedCategories)
        Me.Controls.Add(Me.lbSelectCategory)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.clbSelectCategories)
        Me.Controls.Add(Me.lbSelectCategories)
        Me.Controls.Add(Me.ClbSelectAll)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "ucCategory"
        Me.Size = New System.Drawing.Size(369, 396)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbFilterMasterGenres As System.Windows.Forms.TextBox
    Friend WithEvents clbSelectCategories As System.Windows.Forms.ListBox
    Friend WithEvents lbSelectCategories As System.Windows.Forms.Label
    Friend WithEvents ClbSelectAll As System.Windows.Forms.CheckBox
    Friend WithEvents lbSelectedCategories As System.Windows.Forms.ListBox
    Friend WithEvents lbSelectCategory As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents btnClearAll As System.Windows.Forms.Button

End Class
