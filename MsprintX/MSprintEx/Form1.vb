Imports System
Imports System.Xml
Imports System.IO
Public Class Form1
    Dim table1 As New DataTable("Items")
    Dim table2 As DataTable = table1.Clone()
    Dim ds As DataSet = New DataSet()

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim doc As XmlDocument = New XmlDocument()
        'Dim input As XmlElement = doc.CreateElement("input")
        'doc.AppendChild(input)
        'Dim markets As XmlElement = doc.CreateElement("markets")
        table1.Merge(table2, False, MissingSchemaAction.Add)
        DataGridView3.DataSource = table1
        'doc.AppendChild(markets)
        Dim mgrp As MarketGroup = New MarketGroup()
        ' mgrp.
    End Sub
    Public Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AddHandler table1.RowChanged, AddressOf Row_Changed
        ' Add columns 
        Dim idColumn As New DataColumn("id", GetType(System.Int32))
        Dim itemColumn As New DataColumn("item", GetType(System.Int32))
        table1.Columns.Add(idColumn)
        table1.Columns.Add(itemColumn)

        ' Set the primary key column.
        table1.PrimaryKey = New DataColumn() {idColumn}
        Dim a As StreamReader = New StreamReader("D:\\r.txt")
        ' Add RowChanged event handler for the table. 
        Dim sw As StreamWriter = New StreamWriter("D:\\csv1.csv")
        sw.Write(a.ReadToEnd())
        a.Close()
        sw.Close()
        a = New StreamReader("D:\\csv1.csv")
        ds.Tables.Add("GenreViewerShip")
        ds.Tables(0).Columns.Add("Genre")
        ds.Tables(0).Columns.Add("Viewership", System.Type.GetType("System.Int32"))
        ds.Tables.Add("TopTenPrograms")
        ds.Tables(1).Columns.Add("Rank", System.Type.GetType("System.Int32"))
        ds.Tables(1).Columns.Add("Channel Code", System.Type.GetType("System.Int32"))
        ds.Tables(1).Columns.Add("Viewership", System.Type.GetType("System.Int32"))
        While Not (a.EndOfStream)

            Dim s As String = a.ReadLine()
            Dim values As String()
            If Not (s.Contains("Market:") Or s.Contains("TG:") Or s.Contains("Top 10 Channels")) Or s.Equals(" ") Or s.Contains("Rank") Then
                values = s.Split(New [String]() {","c}, StringSplitOptions.None)

                If Not (values Is Nothing) And values.Length.Equals(2) Then
                    ds.Tables(0).Rows.Add(values)
                ElseIf Not (values Is Nothing) And values.Length.Equals(3) And Not (values(0).Contains("Rank")) Then
                    ds.Tables(1).Rows.Add(values)
                End If


            End If

           


        End While
        Dim dtt As DataTable = New DataTable()
        dtt.Columns.Add("Genre")
        ' dtt.Columns.Add("PLan mg1-"
        ' Add some rows. 
        Dim row As DataRow
        For i As Integer = 0 To 3
            row = table1.NewRow()
            row("id") = i
            row("item") = i
            table1.Rows.Add(row)
        Next i

        ' Accept changes.
        table1.AcceptChanges()
        PrintValues(table1, "Original values")
        DataGridView1.DataSource = table1
        ' Create a second DataTable identical to the first. 


        ' Add column to the second column, so that the  
        ' schemas no longer match.
        table2.Columns.Add("newColumn", GetType(System.String))

        ' Add three rows. Note that the id column can't be the  
        ' same as existing rows in the original table.
        row = table2.NewRow()
        row("id") = 14
        row("item") = 774
        row("newColumn") = "new column 1"
        table2.Rows.Add(row)

        row = table2.NewRow()
        row("id") = 12
        row("item") = 555
        row("newColumn") = "new column 2"
        table2.Rows.Add(row)

        row = table2.NewRow()
        row("id") = 13
        row("item") = 665
        row("newColumn") = "new column 3"
        table2.Rows.Add(row)
        DataGridView2.DataSource = table2
        ' Merge table2 into the table1.
        Console.WriteLine("Merging")

        PrintValues(table1, "Merged With table1, Schema added")
    End Sub
    Private Sub Row_Changed(ByVal sender As Object, _
      ByVal e As DataRowChangeEventArgs)
        Console.WriteLine("Row changed {0}{1}{2}", _
          e.Action, ControlChars.Tab, e.Row.ItemArray(0))
    End Sub

    Private Sub PrintValues(ByVal table As DataTable, _
          ByVal label As String)
        ' Display the values in the supplied DataTable:
        Console.WriteLine(label)
        For Each row As DataRow In table.Rows
            For Each col As DataColumn In table.Columns
                Console.Write(ControlChars.Tab + " " + row(col).ToString())
            Next col
            Console.WriteLine()
        Next row
    End Sub
End Class
Public Class MarketGroup
    Friend mgroupname As String
    Friend marketnames As [String]()
End Class
