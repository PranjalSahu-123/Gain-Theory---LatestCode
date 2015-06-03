<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucAvgTVRMGSelection
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbSelectMG = New System.Windows.Forms.ComboBox()
        Me.btnView = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rbAdCount = New System.Windows.Forms.RadioButton()
        Me.rbBreakCount = New System.Windows.Forms.RadioButton()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 89)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(111, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Select Market Group :"
        '
        'cbSelectMG
        '
        Me.cbSelectMG.FormattingEnabled = True
        Me.cbSelectMG.Location = New System.Drawing.Point(119, 81)
        Me.cbSelectMG.Name = "cbSelectMG"
        Me.cbSelectMG.Size = New System.Drawing.Size(200, 21)
        Me.cbSelectMG.TabIndex = 1
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(119, 115)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 2
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rbAdCount)
        Me.GroupBox1.Controls.Add(Me.rbBreakCount)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(296, 71)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Average TVR View"
        '
        'rbAdCount
        '
        Me.rbAdCount.AutoSize = True
        Me.rbAdCount.Location = New System.Drawing.Point(142, 33)
        Me.rbAdCount.Name = "rbAdCount"
        Me.rbAdCount.Size = New System.Drawing.Size(96, 17)
        Me.rbAdCount.TabIndex = 1
        Me.rbAdCount.TabStop = True
        Me.rbAdCount.Text = "Ad/Spot Count"
        Me.rbAdCount.UseVisualStyleBackColor = True
        '
        'rbBreakCount
        '
        Me.rbBreakCount.AutoSize = True
        Me.rbBreakCount.Location = New System.Drawing.Point(24, 33)
        Me.rbBreakCount.Name = "rbBreakCount"
        Me.rbBreakCount.Size = New System.Drawing.Size(84, 17)
        Me.rbBreakCount.TabIndex = 0
        Me.rbBreakCount.TabStop = True
        Me.rbBreakCount.Text = "Break Count"
        Me.rbBreakCount.UseVisualStyleBackColor = True
        '
        'ucAvgTVRMGSelection
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnView)
        Me.Controls.Add(Me.cbSelectMG)
        Me.Controls.Add(Me.Label1)
        Me.Name = "ucAvgTVRMGSelection"
        Me.Size = New System.Drawing.Size(322, 150)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbSelectMG As System.Windows.Forms.ComboBox
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rbAdCount As System.Windows.Forms.RadioButton
    Friend WithEvents rbBreakCount As System.Windows.Forms.RadioButton

End Class
