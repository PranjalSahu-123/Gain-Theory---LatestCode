<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TVRForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Me.lbchooseChannels = New System.Windows.Forms.Label()
        Me.cbChannels = New System.Windows.Forms.CheckedListBox()
        Me.btnView = New System.Windows.Forms.Button()
        CType(Me.nudTopPrograms, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lbTTpChoose
        '
        Me.lbTTpChoose.AutoSize = True
        Me.lbTTpChoose.Location = New System.Drawing.Point(1, 20)
        Me.lbTTpChoose.Name = "lbTTpChoose"
        Me.lbTTpChoose.Size = New System.Drawing.Size(189, 13)
        Me.lbTTpChoose.TabIndex = 1
        Me.lbTTpChoose.Text = "Choose number of Top Programs reqd:"
        '
        'nudTopPrograms
        '
        Me.nudTopPrograms.Location = New System.Drawing.Point(211, 18)
        Me.nudTopPrograms.Name = "nudTopPrograms"
        Me.nudTopPrograms.Size = New System.Drawing.Size(45, 20)
        Me.nudTopPrograms.TabIndex = 2
        '
        'lbchooseChannels
        '
        Me.lbchooseChannels.AutoSize = True
        Me.lbchooseChannels.Location = New System.Drawing.Point(127, 50)
        Me.lbchooseChannels.Name = "lbchooseChannels"
        Me.lbchooseChannels.Size = New System.Drawing.Size(90, 13)
        Me.lbchooseChannels.TabIndex = 4
        Me.lbchooseChannels.Text = "Choose Channels"
        '
        'cbChannels
        '
        Me.cbChannels.FormattingEnabled = True
        Me.cbChannels.Location = New System.Drawing.Point(93, 66)
        Me.cbChannels.Name = "cbChannels"
        Me.cbChannels.Size = New System.Drawing.Size(212, 139)
        Me.cbChannels.TabIndex = 5
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(158, 227)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 6
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'TVRForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(422, 262)
        Me.Controls.Add(Me.btnView)
        Me.Controls.Add(Me.cbChannels)
        Me.Controls.Add(Me.lbchooseChannels)
        Me.Controls.Add(Me.nudTopPrograms)
        Me.Controls.Add(Me.lbTTpChoose)
        Me.Name = "TVRForm"
        Me.Text = "TVRForm"
        CType(Me.nudTopPrograms, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbTTpChoose As System.Windows.Forms.Label
    Friend WithEvents nudTopPrograms As System.Windows.Forms.NumericUpDown
    Friend WithEvents lbchooseChannels As System.Windows.Forms.Label
    Friend WithEvents cbChannels As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnView As System.Windows.Forms.Button
End Class
