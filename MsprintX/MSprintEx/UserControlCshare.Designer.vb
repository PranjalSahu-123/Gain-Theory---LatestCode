<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UserControlCshare
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
        Me.lblRefMG = New System.Windows.Forms.Label()
        Me.lblPlanMg = New System.Windows.Forms.Label()
        Me.cbrefmgs = New System.Windows.Forms.ComboBox()
        Me.cbpmgs = New System.Windows.Forms.ComboBox()
        Me.btnView = New System.Windows.Forms.Button()
        Me.CbRef = New System.Windows.Forms.ComboBox()
        Me.lblRefPlan = New System.Windows.Forms.Label()
        Me.cbPlan = New System.Windows.Forms.ComboBox()
        Me.lbChoosePlan = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblRefMG
        '
        Me.lblRefMG.AutoSize = True
        Me.lblRefMG.Location = New System.Drawing.Point(25, 104)
        Me.lblRefMG.Name = "lblRefMG"
        Me.lblRefMG.Size = New System.Drawing.Size(116, 13)
        Me.lblRefMG.TabIndex = 17
        Me.lblRefMG.Text = "Choose Reference MG"
        '
        'lblPlanMg
        '
        Me.lblPlanMg.AutoSize = True
        Me.lblPlanMg.Location = New System.Drawing.Point(54, 33)
        Me.lblPlanMg.Name = "lblPlanMg"
        Me.lblPlanMg.Size = New System.Drawing.Size(87, 13)
        Me.lblPlanMg.TabIndex = 16
        Me.lblPlanMg.Text = "Choose Plan MG"
        '
        'cbrefmgs
        '
        Me.cbrefmgs.FormattingEnabled = True
        Me.cbrefmgs.Location = New System.Drawing.Point(147, 96)
        Me.cbrefmgs.Name = "cbrefmgs"
        Me.cbrefmgs.Size = New System.Drawing.Size(121, 21)
        Me.cbrefmgs.TabIndex = 15
        '
        'cbpmgs
        '
        Me.cbpmgs.FormattingEnabled = True
        Me.cbpmgs.Location = New System.Drawing.Point(147, 30)
        Me.cbpmgs.Name = "cbpmgs"
        Me.cbpmgs.Size = New System.Drawing.Size(121, 21)
        Me.cbpmgs.TabIndex = 14
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(103, 147)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(75, 23)
        Me.btnView.TabIndex = 13
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = True
        '
        'CbRef
        '
        Me.CbRef.FormattingEnabled = True
        Me.CbRef.Location = New System.Drawing.Point(147, 57)
        Me.CbRef.Name = "CbRef"
        Me.CbRef.Size = New System.Drawing.Size(121, 21)
        Me.CbRef.TabIndex = 12
        '
        'lblRefPlan
        '
        Me.lblRefPlan.AutoSize = True
        Me.lblRefPlan.Location = New System.Drawing.Point(27, 65)
        Me.lblRefPlan.Name = "lblRefPlan"
        Me.lblRefPlan.Size = New System.Drawing.Size(114, 13)
        Me.lblRefPlan.TabIndex = 11
        Me.lblRefPlan.Text = "Choose Reference TG"
        '
        'cbPlan
        '
        Me.cbPlan.FormattingEnabled = True
        Me.cbPlan.Location = New System.Drawing.Point(147, 3)
        Me.cbPlan.Name = "cbPlan"
        Me.cbPlan.Size = New System.Drawing.Size(121, 21)
        Me.cbPlan.TabIndex = 10
        '
        'lbChoosePlan
        '
        Me.lbChoosePlan.AutoSize = True
        Me.lbChoosePlan.Location = New System.Drawing.Point(56, 11)
        Me.lbChoosePlan.Name = "lbChoosePlan"
        Me.lbChoosePlan.Size = New System.Drawing.Size(85, 13)
        Me.lbChoosePlan.TabIndex = 9
        Me.lbChoosePlan.Text = "Choose Plan TG"
        '
        'UserControlCshare
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.lblRefMG)
        Me.Controls.Add(Me.lblPlanMg)
        Me.Controls.Add(Me.cbrefmgs)
        Me.Controls.Add(Me.cbpmgs)
        Me.Controls.Add(Me.btnView)
        Me.Controls.Add(Me.CbRef)
        Me.Controls.Add(Me.lblRefPlan)
        Me.Controls.Add(Me.cbPlan)
        Me.Controls.Add(Me.lbChoosePlan)
        Me.Name = "UserControlCshare"
        Me.Size = New System.Drawing.Size(322, 205)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblRefMG As System.Windows.Forms.Label
    Friend WithEvents lblPlanMg As System.Windows.Forms.Label
    Friend WithEvents cbrefmgs As System.Windows.Forms.ComboBox
    Friend WithEvents cbpmgs As System.Windows.Forms.ComboBox
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents CbRef As System.Windows.Forms.ComboBox
    Friend WithEvents lblRefPlan As System.Windows.Forms.Label
    Friend WithEvents cbPlan As System.Windows.Forms.ComboBox
    Friend WithEvents lbChoosePlan As System.Windows.Forms.Label

End Class
