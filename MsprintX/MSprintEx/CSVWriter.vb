Imports System.Collections.Generic
Imports System.IO
Imports System.Text
''' <summary>
''' Class to write data to a CSV file
''' </summary>
Public Class CsvRow
    Inherits List(Of String)
    Public Property LineText() As String
        Get
            Return m_LineText
        End Get
        Set(ByVal value As String)
            m_LineText = Value
        End Set
    End Property
    Private m_LineText As String
End Class
Public Class CsvFileWriter
    Inherits StreamWriter
    Public Sub New(ByVal stream As Stream)
        MyBase.New(stream)
    End Sub

    Public Sub New(ByVal filename As String)
        MyBase.New(filename)
    End Sub

    ''' <summary>
    ''' Writes a single row to a CSV file.
    ''' </summary>
    ''' <param name="row">The row to be written</param>
    Public Sub WriteRow(ByVal row As CsvRow)
        Dim builder As New StringBuilder()
        Dim firstColumn As Boolean = True
        For Each value As String In row
            ' Add separator if this isn't the first value
            If Not firstColumn Then
                builder.Append(","c)
            End If
            ' Implement special handling for values that contain comma or quote
            ' Enclose in quotes and double up any double quotes
            If value.IndexOfAny(New Char() {""""c, ","c}) <> -1 Then
                builder.AppendFormat("""{0}""", value.Replace("""", """"""))
            Else
                builder.Append(value)
            End If
            firstColumn = False
        Next
        row.LineText = builder.ToString()
        WriteLine(row.LineText)
    End Sub
End Class
''' <summary>
''' Class to read data from a CSV file
''' </summary>
Public Class CsvFileReader
    Inherits StreamReader
    Public Sub New(ByVal stream As Stream)
        MyBase.New(stream)
    End Sub

    Public Sub New(ByVal filename As String)
        MyBase.New(filename)
    End Sub

    ''' <summary>
    ''' Reads a row of data from a CSV file
    ''' </summary>
    ''' <param name="row"></param>
    ''' <returns></returns>
    Public Function ReadRow(ByVal row As CsvRow) As Boolean
        row.LineText = ReadLine()
        'If [String].IsNullOrEmpty(row.LineText) Then
        '    Return False
        'End If

        Dim pos As Integer = 0
        Dim rows As Integer = 0

        If Not (row.LineText Is Nothing) Then



            While pos < row.LineText.Length
                Dim value As String

                ' Special handling for quoted field
                If row.LineText(pos) = """"c Then
                    ' Skip initial quote
                    pos += 1

                    ' Parse quoted value
                    Dim start As Integer = pos
                    While pos < row.LineText.Length
                        ' Test for quote character
                        If row.LineText(pos) = """"c Then
                            ' Found one
                            pos += 1

                            ' If two quotes together, keep one
                            ' Otherwise, indicates end of value
                            If pos >= row.LineText.Length OrElse row.LineText(pos) <> """"c Then
                                pos -= 1
                                Exit While
                            End If
                        End If
                        pos += 1
                    End While
                    value = row.LineText.Substring(start, pos - start)
                    value = value.Replace("""""", """")
                Else
                    ' Parse unquoted value
                    Dim start As Integer = pos
                    While pos < row.LineText.Length AndAlso row.LineText(pos) <> ","c
                        pos += 1
                    End While
                    value = row.LineText.Substring(start, pos - start)
                End If

                ' Add field to list
                If rows < row.Count Then
                    row(rows) = value
                Else
                    row.Add(value)
                End If
                rows += 1

                ' Eat up to and including next comma
                While pos < row.LineText.Length AndAlso row.LineText(pos) <> ","c
                    pos += 1
                End While
                If pos < row.LineText.Length Then
                    pos += 1
                End If
            End While
        End If
        ' Delete any unused items
        While row.Count > rows
            row.RemoveAt(rows)
        End While

        ' Return true if any columns read
        Return (row.Count > 0)
    End Function
End Class
