<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucSelections
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
        Me.tabSelections = New System.Windows.Forms.TabControl()
        Me.tpAudience = New System.Windows.Forms.TabPage()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.lucAudience = New MSprintEx.ucAudience()
        Me.clbGender = New System.Windows.Forms.CheckedListBox()
        Me.clbHH = New System.Windows.Forms.CheckedListBox()
        Me.tpMarkets = New System.Windows.Forms.TabPage()
        Me.lucMarkets = New MSprintEx.ucMarkets()
        Me.tpPlanDates = New System.Windows.Forms.TabPage()
        Me.tabSelections.SuspendLayout()
        Me.tpAudience.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.tpMarkets.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabSelections
        '
        Me.tabSelections.Controls.Add(Me.tpAudience)
        Me.tabSelections.Controls.Add(Me.tpMarkets)
        Me.tabSelections.Controls.Add(Me.tpPlanDates)
        Me.tabSelections.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabSelections.Location = New System.Drawing.Point(0, 0)
        Me.tabSelections.Margin = New System.Windows.Forms.Padding(0)
        Me.tabSelections.Name = "tabSelections"
        Me.tabSelections.SelectedIndex = 0
        Me.tabSelections.Size = New System.Drawing.Size(555, 559)
        Me.tabSelections.TabIndex = 0
        '
        'tpAudience
        '
        Me.tpAudience.Controls.Add(Me.TableLayoutPanel1)
        Me.tpAudience.Location = New System.Drawing.Point(4, 22)
        Me.tpAudience.Name = "tpAudience"
        Me.tpAudience.Padding = New System.Windows.Forms.Padding(3)
        Me.tpAudience.Size = New System.Drawing.Size(547, 533)
        Me.tpAudience.TabIndex = 1
        Me.tpAudience.Text = "Audience"
        Me.tpAudience.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 251.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.lucAudience, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.clbHH, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.clbGender, 1, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(3, 3)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 55.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(541, 527)
        Me.TableLayoutPanel1.TabIndex = 1
        '
        'lucAudience
        '
        Me.lucAudience.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lucAudience.Location = New System.Drawing.Point(0, 0)
        Me.lucAudience.Margin = New System.Windows.Forms.Padding(0)
        Me.lucAudience.Name = "lucAudience"
        Me.lucAudience.Size = New System.Drawing.Size(251, 472)
        Me.lucAudience.TabIndex = 0
        '
        'clbGender
        '
        Me.clbGender.CheckOnClick = True
        Me.clbGender.FormattingEnabled = True
        Me.clbGender.Location = New System.Drawing.Point(254, 3)
        Me.clbGender.Name = "clbGender"
        Me.clbGender.Size = New System.Drawing.Size(120, 109)
        Me.clbGender.TabIndex = 1
        '
        'clbHH
        '
        Me.clbHH.CheckOnClick = True
        Me.clbHH.FormattingEnabled = True
        Me.clbHH.Location = New System.Drawing.Point(254, 475)
        Me.clbHH.Name = "clbHH"
        Me.clbHH.Size = New System.Drawing.Size(120, 49)
        Me.clbHH.TabIndex = 3
        '
        'tpMarkets
        '
        Me.tpMarkets.Controls.Add(Me.lucMarkets)
        Me.tpMarkets.Location = New System.Drawing.Point(4, 22)
        Me.tpMarkets.Name = "tpMarkets"
        Me.tpMarkets.Padding = New System.Windows.Forms.Padding(3)
        Me.tpMarkets.Size = New System.Drawing.Size(547, 423)
        Me.tpMarkets.TabIndex = 0
        Me.tpMarkets.Text = "Markets"
        Me.tpMarkets.UseVisualStyleBackColor = True
        '
        'lucMarkets
        '
        Me.lucMarkets.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lucMarkets.Location = New System.Drawing.Point(3, 3)
        Me.lucMarkets.Name = "lucMarkets"
        Me.lucMarkets.Size = New System.Drawing.Size(541, 417)
        Me.lucMarkets.TabIndex = 0
        '
        'tpPlanDates
        '
        Me.tpPlanDates.Location = New System.Drawing.Point(4, 22)
        Me.tpPlanDates.Name = "tpPlanDates"
        Me.tpPlanDates.Padding = New System.Windows.Forms.Padding(3)
        Me.tpPlanDates.Size = New System.Drawing.Size(547, 423)
        Me.tpPlanDates.TabIndex = 2
        Me.tpPlanDates.Text = "Plan Dates"
        Me.tpPlanDates.UseVisualStyleBackColor = True
        '
        'ucSelections
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.tabSelections)
        Me.Name = "ucSelections"
        Me.Size = New System.Drawing.Size(555, 559)
        Me.tabSelections.ResumeLayout(False)
        Me.tpAudience.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.tpMarkets.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tabSelections As System.Windows.Forms.TabControl
    Friend WithEvents tpMarkets As System.Windows.Forms.TabPage
    Friend WithEvents tpAudience As System.Windows.Forms.TabPage
    Friend WithEvents lucMarkets As MSprintEx.ucMarkets
    Friend WithEvents lucAudience As MSprintEx.ucAudience
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents tpPlanDates As System.Windows.Forms.TabPage
    Friend WithEvents TaskPaneLogFile1 As MSprintEx.TaskPaneLogFile
    Friend WithEvents clbGender As System.Windows.Forms.CheckedListBox
    Friend WithEvents clbHH As System.Windows.Forms.CheckedListBox
    'Friend WithEvents lucPeriod As MSprintEx.TaskPaneLogFile

End Class
