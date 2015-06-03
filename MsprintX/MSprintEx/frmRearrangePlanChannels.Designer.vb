<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRearrangePlanChannels
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
        Me.lstChannelSortOrder = New System.Windows.Forms.ListBox()
        Me.btnMvChnlUP = New System.Windows.Forms.Button()
        Me.btnMvChDown = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnRearrange = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lstChannelSortOrder
        '
        Me.lstChannelSortOrder.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstChannelSortOrder.Location = New System.Drawing.Point(32, 38)
        Me.lstChannelSortOrder.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.lstChannelSortOrder.Name = "lstChannelSortOrder"
        Me.lstChannelSortOrder.Size = New System.Drawing.Size(195, 225)
        Me.lstChannelSortOrder.TabIndex = 23
        '
        'btnMvChnlUP
        '
        Me.btnMvChnlUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnMvChnlUP.Location = New System.Drawing.Point(241, 91)
        Me.btnMvChnlUP.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.btnMvChnlUP.Name = "btnMvChnlUP"
        Me.btnMvChnlUP.Size = New System.Drawing.Size(84, 54)
        Me.btnMvChnlUP.TabIndex = 24
        Me.btnMvChnlUP.Text = "UP"
        '
        'btnMvChDown
        '
        Me.btnMvChDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnMvChDown.Location = New System.Drawing.Point(241, 173)
        Me.btnMvChDown.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.btnMvChDown.Name = "btnMvChDown"
        Me.btnMvChDown.Size = New System.Drawing.Size(84, 58)
        Me.btnMvChDown.TabIndex = 25
        Me.btnMvChDown.Text = "DOWN"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(32, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(162, 13)
        Me.Label1.TabIndex = 26
        Me.Label1.Text = "Plan Channels to be rearranged :"
        '
        'btnRearrange
        '
        Me.btnRearrange.Location = New System.Drawing.Point(35, 281)
        Me.btnRearrange.Name = "btnRearrange"
        Me.btnRearrange.Size = New System.Drawing.Size(75, 23)
        Me.btnRearrange.TabIndex = 27
        Me.btnRearrange.Text = "Rearrange"
        Me.btnRearrange.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(133, 281)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 28
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'frmRearrangePlanChannels
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(422, 325)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnRearrange)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnMvChDown)
        Me.Controls.Add(Me.btnMvChnlUP)
        Me.Controls.Add(Me.lstChannelSortOrder)
        Me.Name = "frmRearrangePlanChannels"
        Me.Text = "frmRearrangePlanChannels"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents lstChannelSortOrder As System.Windows.Forms.ListBox
    Private WithEvents btnMvChnlUP As System.Windows.Forms.Button
    Private WithEvents btnMvChDown As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnRearrange As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
End Class
