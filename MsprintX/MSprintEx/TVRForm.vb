Imports System.Data
Public Class TVRForm
    Friend channels As DataSet
    Friend tpSelObject As ucPlanSelections
    Public Sub New(ByVal ds As DataSet, ByVal tpSObject As ucPlanSelections)
        channels = ds
        tpSelObject = tpSObject
        InitializeComponent()
    End Sub
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub TVRForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        nudTopPrograms.Value = 20
        For Each dr As DataRow In channels.Tables(0).Rows

            If Not (cbChannels.Items.Contains(dr("Channel Name").ToString())) Then
                cbChannels.Items.Add(dr("Channel Name").ToString())
            End If

        Next
    End Sub

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Me.Close()
    End Sub
End Class