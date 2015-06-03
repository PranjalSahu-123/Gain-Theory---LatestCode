<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucTopPrograms
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
        Me.cbTG = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cmbTopNumber = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmbWeeks = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmbMinutes = New System.Windows.Forms.ComboBox()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cbTG
        '
        Me.cbTG.Dock = System.Windows.Forms.DockStyle.Fill
        Me.cbTG.FormattingEnabled = True
        Me.cbTG.Location = New System.Drawing.Point(0, 21)
        Me.cbTG.Margin = New System.Windows.Forms.Padding(0)
        Me.cbTG.Name = "cbTG"
        Me.cbTG.Size = New System.Drawing.Size(248, 21)
        Me.cbTG.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label1.Location = New System.Drawing.Point(0, 0)
        Me.Label1.Margin = New System.Windows.Forms.Padding(0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(248, 21)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Please select TG"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label2.Location = New System.Drawing.Point(0, 42)
        Me.Label2.Margin = New System.Windows.Forms.Padding(0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(248, 21)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Number of programs"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbTopNumber
        '
        Me.cmbTopNumber.Dock = System.Windows.Forms.DockStyle.Left
        Me.cmbTopNumber.FormattingEnabled = True
        Me.cmbTopNumber.Location = New System.Drawing.Point(0, 63)
        Me.cmbTopNumber.Margin = New System.Windows.Forms.Padding(0)
        Me.cmbTopNumber.Name = "cmbTopNumber"
        Me.cmbTopNumber.Size = New System.Drawing.Size(59, 21)
        Me.cmbTopNumber.TabIndex = 2
        Me.cmbTopNumber.Text = "20"
        '
        'Label3
        '
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label3.Location = New System.Drawing.Point(0, 84)
        Me.Label3.Margin = New System.Windows.Forms.Padding(0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(248, 21)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Weeks"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbWeeks
        '
        Me.cmbWeeks.Dock = System.Windows.Forms.DockStyle.Left
        Me.cmbWeeks.FormattingEnabled = True
        Me.cmbWeeks.Location = New System.Drawing.Point(0, 105)
        Me.cmbWeeks.Margin = New System.Windows.Forms.Padding(0)
        Me.cmbWeeks.Name = "cmbWeeks"
        Me.cmbWeeks.Size = New System.Drawing.Size(59, 21)
        Me.cmbWeeks.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label4.Location = New System.Drawing.Point(0, 126)
        Me.Label4.Margin = New System.Windows.Forms.Padding(0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(248, 21)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Dayparts"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbMinutes
        '
        Me.cmbMinutes.Dock = System.Windows.Forms.DockStyle.Left
        Me.cmbMinutes.FormattingEnabled = True
        Me.cmbMinutes.Location = New System.Drawing.Point(0, 147)
        Me.cmbMinutes.Margin = New System.Windows.Forms.Padding(0)
        Me.cmbMinutes.Name = "cmbMinutes"
        Me.cmbMinutes.Size = New System.Drawing.Size(59, 21)
        Me.cmbMinutes.TabIndex = 6
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnOK.Location = New System.Drawing.Point(86, 172)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 8
        Me.btnOK.Text = "Ok"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.cbTG, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.btnOK, 0, 8)
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.cmbMinutes, 0, 7)
        Me.TableLayoutPanel1.Controls.Add(Me.Label4, 0, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.cmbTopNumber, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.cmbWeeks, 0, 5)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(11, 3)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 9
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 21.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 21.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 21.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 21.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 21.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 21.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 21.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 21.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 21.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(248, 200)
        Me.TableLayoutPanel1.TabIndex = 10
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(184, 183)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 11
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ucTopPrograms
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "ucTopPrograms"
        Me.Size = New System.Drawing.Size(271, 206)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cbTG As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmbTopNumber As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmbWeeks As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmbMinutes As System.Windows.Forms.ComboBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Button1 As System.Windows.Forms.Button

End Class
