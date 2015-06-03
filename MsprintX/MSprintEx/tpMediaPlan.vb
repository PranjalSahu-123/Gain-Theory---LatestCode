Imports System.Windows.Forms

Public Class tpMediaPlan


    Private Sub tpMediaPlan_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Panel1.Height = 25
        Panel2.Height = 25
        Panel3.Height = 25
    End Sub
    Private Sub ButtonClick _
    (ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles Label1.Click, Label2.Click, Label3.Click

        'find out which label was clicked
        Dim lbl As Label = CType(sender, Label)
        'find the panel containing the label
        Dim pnl As Panel = lbl.Parent

        '### this assumes that no other controls are present in the main table ###
        'the code loops through the panels in the table and expands/collapses
        'each panel according to whether it contains the clicked label. The label
        'images are also swapped depending on the height of the panel.
        Dim accHeight As Integer = Me.Height - 110
        For Each p As Panel In TableLayoutPanel1.Controls
            'Dim l As Label = CType(p.Controls(0), Label)
            If p.Equals(pnl) Then
                'expand or collapse the panel
                If p.Height = accHeight Then
                    p.Height = 25
                    'Change the image name to YOUR image
                    'l.Image = My.Resources.Expander_Collapsed16
                Else
                    p.Height = accHeight
                    'Change the image name to YOUR image
                    'l.Image = My.Resources.Expander_Expanded16
                End If

            Else
                p.Height = 25
                'Change the image name to YOUR image
                'l.Image = My.Resources.Expander_Collapsed16
            End If
        Next

    End Sub
End Class
