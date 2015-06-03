<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucAdvertiser
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
        Me.clbSelectAdvertisers = New System.Windows.Forms.ListBox()
        Me.lbSelectAdvertiser = New System.Windows.Forms.Label()
        Me.ClbSelectAll = New System.Windows.Forms.CheckBox()
        Me.lbSelectedAdvertisers = New System.Windows.Forms.ListBox()
        Me.lbChosenAdvertisers = New System.Windows.Forms.Label()
        Me.btnClearAll = New System.Windows.Forms.Button()
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
        Me.GroupBox1.TabIndex = 36
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
        'clbSelectAdvertisers
        '
        Me.clbSelectAdvertisers.FormattingEnabled = True
        Me.clbSelectAdvertisers.HorizontalScrollbar = True
        Me.clbSelectAdvertisers.Location = New System.Drawing.Point(3, 95)
        Me.clbSelectAdvertisers.Name = "clbSelectAdvertisers"
        Me.clbSelectAdvertisers.Size = New System.Drawing.Size(177, 277)
        Me.clbSelectAdvertisers.Sorted = True
        Me.clbSelectAdvertisers.TabIndex = 42
        '
        'lbSelectAdvertiser
        '
        Me.lbSelectAdvertiser.AutoSize = True
        Me.lbSelectAdvertiser.Location = New System.Drawing.Point(49, 74)
        Me.lbSelectAdvertiser.Name = "lbSelectAdvertiser"
        Me.lbSelectAdvertiser.Size = New System.Drawing.Size(104, 13)
        Me.lbSelectAdvertiser.TabIndex = 41
        Me.lbSelectAdvertiser.Text = "Choose Advertiser(s)"
        '
        'ClbSelectAll
        '
        Me.ClbSelectAll.AutoSize = True
        Me.ClbSelectAll.Location = New System.Drawing.Point(3, 72)
        Me.ClbSelectAll.Name = "ClbSelectAll"
        Me.ClbSelectAll.Size = New System.Drawing.Size(40, 17)
        Me.ClbSelectAll.TabIndex = 40
        Me.ClbSelectAll.Text = "All "
        Me.ClbSelectAll.UseVisualStyleBackColor = True
        '
        'lbSelectedAdvertisers
        '
        Me.lbSelectedAdvertisers.FormattingEnabled = True
        Me.lbSelectedAdvertisers.HorizontalScrollbar = True
        Me.lbSelectedAdvertisers.Location = New System.Drawing.Point(186, 95)
        Me.lbSelectedAdvertisers.Name = "lbSelectedAdvertisers"
        Me.lbSelectedAdvertisers.Size = New System.Drawing.Size(152, 277)
        Me.lbSelectedAdvertisers.Sorted = True
        Me.lbSelectedAdvertisers.TabIndex = 44
        '
        'lbChosenAdvertisers
        '
        Me.lbChosenAdvertisers.AutoSize = True
        Me.lbChosenAdvertisers.Location = New System.Drawing.Point(183, 74)
        Me.lbChosenAdvertisers.Name = "lbChosenAdvertisers"
        Me.lbChosenAdvertisers.Size = New System.Drawing.Size(110, 13)
        Me.lbChosenAdvertisers.TabIndex = 43
        Me.lbChosenAdvertisers.Text = "Selected Advertiser(s)"
        '
        'btnClearAll
        '
        Me.btnClearAll.Location = New System.Drawing.Point(299, 69)
        Me.btnClearAll.Name = "btnClearAll"
        Me.btnClearAll.Size = New System.Drawing.Size(39, 23)
        Me.btnClearAll.TabIndex = 45
        Me.btnClearAll.Text = "Clear All"
        Me.btnClearAll.UseVisualStyleBackColor = True
        '
        'ucAdvertiser
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.btnClearAll)
        Me.Controls.Add(Me.lbSelectedAdvertisers)
        Me.Controls.Add(Me.lbChosenAdvertisers)
        Me.Controls.Add(Me.clbSelectAdvertisers)
        Me.Controls.Add(Me.lbSelectAdvertiser)
        Me.Controls.Add(Me.ClbSelectAll)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "ucAdvertiser"
        Me.Size = New System.Drawing.Size(344, 378)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbFilterMasterGenres As System.Windows.Forms.TextBox
    Friend WithEvents clbSelectAdvertisers As System.Windows.Forms.ListBox
    Friend WithEvents lbSelectAdvertiser As System.Windows.Forms.Label
    Friend WithEvents ClbSelectAll As System.Windows.Forms.CheckBox
    Friend WithEvents lbSelectedAdvertisers As System.Windows.Forms.ListBox
    Friend WithEvents lbChosenAdvertisers As System.Windows.Forms.Label
    Friend WithEvents btnClearAll As System.Windows.Forms.Button

End Class
