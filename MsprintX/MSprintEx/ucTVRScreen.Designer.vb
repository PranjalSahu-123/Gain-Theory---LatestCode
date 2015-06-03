<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucTVRScreen
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
        Me.lbTTpChoose = New System.Windows.Forms.Label()
        Me.nudTopPrograms = New System.Windows.Forms.NumericUpDown()
        Me.cbChannels = New System.Windows.Forms.CheckedListBox()
        Me.lbchooseChannels = New System.Windows.Forms.Label()
        Me.btnView = New System.Windows.Forms.Button()
        CType(Me.nudTopPrograms, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lbTTpChoose
        '
        Me.lbTTpChoose.AutoSize = True
        Me.lbTTpChoose.Location = New System.Drawing.Point(3, 11)
        Me.lbTTpChoose.Name = "lbTTpChoose"
        Me.lbTTpChoose.Size = New System.Drawing.Size(189, 13)
        Me.lbTTpChoose.TabIndex = 0
        Me.lbTTpChoose.Text = "Choose number of Top Programs reqd:"
        '
        'nudTopPrograms
        '
        Me.nudTopPrograms.Location = New System.Drawing.Point(198, 9)
        Me.nudTopPrograms.Name = "nudTopPrograms"
        Me.nudTopPrograms.Size = New System.Drawing.Size(45, 20)
        Me.nudTopPrograms.TabIndex = 1
        '
        'cbChannels
        '
        Me.cbChannels.FormattingEnabled = True
        Me.cbChannels.Location = New System.Drawing.Point(58, 76)
        Me.cbChannels.Name = "cbChannels"
        Me.cbChannels.Size = New System.Drawing.Size(212, 94)
        Me.cbChannels.TabIndex = 2
        '
        'lbchooseChannels
        '
        Me.lbchooseChannels.AutoSize = True
        Me.lbchooseChannels.Location = New System.Drawing.Point(102, 47)
        Me.lbchooseChannels.Name = "lbchooseChannels"
        Me.lbchooseChannels.Size = New System.Drawing.Size(90, 13)
        Me.lbchooseChannels.TabIndex = 3
        Me.lbchooseChannels.Text = "Choose Channels"
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(119, 187)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 4
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'ucTVRScreen
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.btnView)
        Me.Controls.Add(Me.lbchooseChannels)
        Me.Controls.Add(Me.cbChannels)
        Me.Controls.Add(Me.nudTopPrograms)
        Me.Controls.Add(Me.lbTTpChoose)
        Me.Name = "ucTVRScreen"
        Me.Size = New System.Drawing.Size(315, 222)
        CType(Me.nudTopPrograms, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbTTpChoose As System.Windows.Forms.Label
    Friend WithEvents nudTopPrograms As System.Windows.Forms.NumericUpDown
    Friend WithEvents cbChannels As System.Windows.Forms.CheckedListBox
    Friend WithEvents lbchooseChannels As System.Windows.Forms.Label
    Friend WithEvents btnView As System.Windows.Forms.Button

End Class
