<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWait
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
        Me.lblPercentage = New System.Windows.Forms.Label()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.imgWait = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        CType(Me.imgWait, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblPercentage
        '
        Me.lblPercentage.BackColor = System.Drawing.Color.White
        Me.lblPercentage.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPercentage.Location = New System.Drawing.Point(0, 120)
        Me.lblPercentage.Name = "lblPercentage"
        Me.lblPercentage.Size = New System.Drawing.Size(183, 45)
        Me.lblPercentage.TabIndex = 1
        Me.lblPercentage.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblPercentage.UseWaitCursor = True
        '
        'lblTitle
        '
        Me.lblTitle.BackColor = System.Drawing.Color.White
        Me.lblTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTitle.Location = New System.Drawing.Point(0, 0)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(199, 33)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.UseWaitCursor = True
        '
        'imgWait
        '
        Me.imgWait.Image = Global.MSprintEx.My.Resources.Resources.ajax_loader
        Me.imgWait.Location = New System.Drawing.Point(-1, 24)
        Me.imgWait.Name = "imgWait"
        Me.imgWait.Size = New System.Drawing.Size(200, 100)
        Me.imgWait.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.imgWait.TabIndex = 4
        Me.imgWait.TabStop = False
        Me.imgWait.UseWaitCursor = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.Controls.Add(Me.imgWait)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(200, 165)
        Me.Panel1.TabIndex = 5
        Me.Panel1.UseWaitCursor = True
        '
        'frmWait
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(199, 165)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.lblPercentage)
        Me.Controls.Add(Me.Panel1)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmWait"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.TopMost = True
        Me.UseWaitCursor = True
        CType(Me.imgWait, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblPercentage As System.Windows.Forms.Label
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents imgWait As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
End Class
