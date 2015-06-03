<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucPlanSelections
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
        Me.TabAudience = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.UcAudience1 = New MSprintEx.ucAudience()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.UcMarkets1 = New MSprintEx.ucMkets()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.TaskPaneLogFile1 = New MSprintEx.TaskPaneLogFile()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.UcGenres = New MSprintEx.UcGenres()
        Me.TabPage5 = New System.Windows.Forms.TabPage()
        Me.UcChannels = New MSprintEx.ucChannels()
        Me.TabAudience.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.TabPage5.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabAudience
        '
        Me.TabAudience.Controls.Add(Me.TabPage1)
        Me.TabAudience.Controls.Add(Me.TabPage2)
        Me.TabAudience.Controls.Add(Me.TabPage3)
        Me.TabAudience.Controls.Add(Me.TabPage4)
        Me.TabAudience.Controls.Add(Me.TabPage5)
        Me.TabAudience.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabAudience.Location = New System.Drawing.Point(0, 0)
        Me.TabAudience.Name = "TabAudience"
        Me.TabAudience.SelectedIndex = 0
        Me.TabAudience.Size = New System.Drawing.Size(334, 505)
        Me.TabAudience.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.UcAudience1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(326, 479)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Audience"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'UcAudience1
        '
        Me.UcAudience1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UcAudience1.Location = New System.Drawing.Point(3, 3)
        Me.UcAudience1.Name = "UcAudience1"
        Me.UcAudience1.Size = New System.Drawing.Size(320, 473)
        Me.UcAudience1.TabIndex = 0
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.UcMarkets1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(326, 479)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Markets"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'UcMarkets1
        '
        Me.UcMarkets1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UcMarkets1.Location = New System.Drawing.Point(3, 3)
        Me.UcMarkets1.Name = "UcMarkets1"
        Me.UcMarkets1.Size = New System.Drawing.Size(320, 473)
        Me.UcMarkets1.TabIndex = 0
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.TaskPaneLogFile1)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(326, 479)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Pre-Eval Dates"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'TaskPaneLogFile1
        '
        Me.TaskPaneLogFile1.AutoScroll = True
        Me.TaskPaneLogFile1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TaskPaneLogFile1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TaskPaneLogFile1.Location = New System.Drawing.Point(3, 3)
        Me.TaskPaneLogFile1.Name = "TaskPaneLogFile1"
        Me.TaskPaneLogFile1.Size = New System.Drawing.Size(320, 473)
        Me.TaskPaneLogFile1.TabIndex = 0
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.UcGenres)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(326, 479)
        Me.TabPage4.TabIndex = 2
        Me.TabPage4.Text = "Genres"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'UcGenres
        '
        Me.UcGenres.Location = New System.Drawing.Point(0, 0)
        Me.UcGenres.Name = "UcGenres"
        Me.UcGenres.Size = New System.Drawing.Size(320, 514)
        Me.UcGenres.TabIndex = 0
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.UcChannels)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage5.Size = New System.Drawing.Size(326, 479)
        Me.TabPage5.TabIndex = 2
        Me.TabPage5.Text = "Channels"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'UcChannels
        '
        Me.UcChannels.Location = New System.Drawing.Point(0, 0)
        Me.UcChannels.Name = "UcChannels"
        Me.UcChannels.Size = New System.Drawing.Size(446, 582)
        Me.UcChannels.TabIndex = 0
        '
        'ucPlanSelections
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.TabAudience)
        Me.Name = "ucPlanSelections"
        Me.Size = New System.Drawing.Size(334, 505)
        Me.TabAudience.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage5.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabAudience As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents UcMarkets1 As MSprintEx.ucMkets
    Friend WithEvents UcGenres As MSprintEx.UcGenres
    Friend WithEvents UcChannels As MSprintEx.ucChannels
    Friend WithEvents TaskPaneLogFile1 As MSprintEx.TaskPaneLogFile
    Friend WithEvents UcAudience1 As MSprintEx.ucAudience

End Class
