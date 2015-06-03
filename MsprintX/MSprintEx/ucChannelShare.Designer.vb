<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucChannelShare
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
        Me.btnView = New System.Windows.Forms.Button()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.lbmgs = New System.Windows.Forms.Label()
        Me.lbdatepicker = New System.Windows.Forms.Label()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(99, 134)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 4
        Me.btnView.Text = "Display"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(114, 16)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(158, 21)
        Me.ComboBox1.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(105, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Select TG for display"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(115, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(166, 52)
        Me.Label2.TabIndex = 11
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Location = New System.Drawing.Point(153, 93)
        Me.NumericUpDown1.Minimum = New Decimal(New Integer() {5, 0, 0, 0})
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(43, 20)
        Me.NumericUpDown1.TabIndex = 13
        Me.NumericUpDown1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.NumericUpDown1.Value = New Decimal(New Integer() {10, 0, 0, 0})
        '
        'lbmgs
        '
        Me.lbmgs.AutoSize = True
        Me.lbmgs.Location = New System.Drawing.Point(3, 56)
        Me.lbmgs.Name = "lbmgs"
        Me.lbmgs.Size = New System.Drawing.Size(89, 13)
        Me.lbmgs.TabIndex = 14
        Me.lbmgs.Text = "Market Group(s) :"
        '
        'lbdatepicker
        '
        Me.lbdatepicker.AutoSize = True
        Me.lbdatepicker.Location = New System.Drawing.Point(3, 95)
        Me.lbdatepicker.Name = "lbdatepicker"
        Me.lbdatepicker.Size = New System.Drawing.Size(144, 13)
        Me.lbdatepicker.TabIndex = 15
        Me.lbdatepicker.Text = "Choose number of items reqd"
        '
        'ucChannelShare
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.lbdatepicker)
        Me.Controls.Add(Me.lbmgs)
        Me.Controls.Add(Me.NumericUpDown1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.btnView)
        Me.Name = "ucChannelShare"
        Me.Size = New System.Drawing.Size(317, 215)
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents lbmgs As System.Windows.Forms.Label
    Friend WithEvents lbdatepicker As System.Windows.Forms.Label

End Class
