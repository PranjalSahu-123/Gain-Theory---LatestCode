Imports System
Imports System.Windows.Forms
Public Class ucGenEndTime
    Dim table As Data.DataTable
    Private Sub btnGenerate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenerate.Click
        Try
            Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(Globals.ThisAddIn.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            Dim vstoWorkbook As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(sheet)
            ' vstoWorkbook.Name = "Plan Selection"
            loSpotSelection = DirectCast(vstoWorkbook.Controls.Item("InputSpotSelection"), Microsoft.Office.Tools.Excel.ListObject)
            table = CType(loSpotSelection.DataSource, Data.DataTable)
            For Each row As Data.DataRow In table.Rows
                Dim et As String = GenerateEndTime(row("Start Time"), Double.Parse(NumericUpDown1.Value)).ToString()
                row("End Time") = et
            Next
            loSpotSelection.SetDataBinding(table)

        Catch ex As Exception
            LogMpsrintExException("Exception occured while binding datatable with End Time with plan selection listobject.Message :" + ex.Message)
            Throw ex
        End Try
    End Sub
    Public Function GenerateEndTime(ByVal timespanDoubleVal As String, ByVal duration As Double) As String
        'Dim starTime As String = String.Empty
        'For Each cell As Microsoft.Office.Interop.Excel.Range In loSpotSelection.ListRows(rIndex + 1).Range.Cells

        '    If cell.Column = 6 Then
        '        starTime = cell.Text

        '    End If

        'Next
        Dim timespanString = String.Empty
        Dim dateVal As Date
        Try
            Dim doubleVal As Double = Convert.ToDouble(timespanDoubleVal)
            dateVal = Date.FromOADate(doubleVal)
        Catch ex As Exception
            dateVal = Date.Parse(timespanDoubleVal)
        End Try
 
        Dim endDate As Date = dateVal.AddMinutes(duration)
        Return endDate.TimeOfDay.ToString("hh\:mm")
        ' Return starTime
    End Function
End Class
