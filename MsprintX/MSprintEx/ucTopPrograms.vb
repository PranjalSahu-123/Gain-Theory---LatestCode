Imports Microsoft.Office.Interop.Excel
Public Class ucTopPrograms
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
    Dim exlObj As Excel.Application
    Dim wbBrkPerf As Excel.Workbook
    Dim wsBrkSht, wsCalcSht As Excel.Worksheet
    Dim DataRows As Range
    Dim BrkTGs As List(Of TAMTG)
    Dim ReportTGs As Range
    Dim RowHeader, NewRowHeader As Range
    Dim colChannel, colProgramme, colDate, colDay, colStartTime, colPA, colTA, colProgStartTime, colCommercial As Range
    Dim calRange As Range

    Friend Sub getTopProgramsBreakTVR()
        exlObj = Globals.ThisAddIn.Application
        reInitialize()
    End Sub
    Public Sub reInitialize()
        wbBrkPerf = exlObj.ActiveWorkbook
        wsBrkSht = wbBrkPerf.ActiveSheet
        wsCalcSht = wbBrkPerf.Worksheets.Add()

        ReportTGs = wsBrkSht.Rows("1:1")
        RowHeader = wsBrkSht.Rows("2:2")
        With wsBrkSht
            'RowHeader.Cut()
            '.Range("A6").Insert(Shift:=XlInsertShiftDirection.xlShiftDown)
            '.Rows("2:4").Delete(Shift:=XlDeleteShiftDirection.xlShiftUp)

            DataRows = .Range(.Range("A6"), RowHeader.CurrentRegion.Cells(RowHeader.CurrentRegion.Count).addresslocal)
            'DataRows.Name = "BreakData"
        End With
        cmbTopNumber.Items.Clear()
        cmbTopNumber.Items.Add(5)
        cmbTopNumber.Items.Add(10)
        cmbTopNumber.Items.Add(20)
        cmbTopNumber.Items.Add(50)
        cmbTopNumber.Items.Add(75)
        cmbTopNumber.Items.Add(100)
        BrkTGs = New List(Of TAMTG)
        For Each cell As Range In ReportTGs.SpecialCells(XlCellType.xlCellTypeConstants)
            If Trim(cell.Value) <> "" Then
                BrkTGs.Add(New TAMTG(cell.Value, cell))
            End If
        Next
        cbTG.DataSource = BrkTGs
        cbTG.DisplayMember = "TGName"
        cbTG.ValueMember = "TGName"

        cmbWeeks.Items.Clear()
        For i = 1 To 52
            cmbWeeks.Items.Add(i)
        Next
        cmbWeeks.Text = 1

        cmbMinutes.Items.Clear()
        cmbMinutes.Items.Add("0")
        cmbMinutes.Items.Add("30")
        cmbMinutes.Items.Add("60")
        cmbMinutes.Text = "30"

    End Sub
    Private Sub btnOK_Click(sender As System.Object, e As System.EventArgs) Handles btnOK.Click
        colProgStartTime = Nothing
        wsCalcSht.Cells.Clear()
        RowHeader.Copy()
        wsCalcSht.Range("A1").PasteSpecial(XlPasteType.xlPasteAll)
        calRange = wsCalcSht.Range("2:2")
        Dim selectedTG As TAMTG = cbTG.SelectedItem
        'DataRows.Sort(Key1:=colDate, Order1:=XlSortOrder.xlAscending, Key2:=colProgramme, Order2:=XlSortOrder.xlAscending, Key3:=colProgStartTime, Order3:=XlSortOrder.xlAscending, Header:=XlYesNoGuess.xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=XlSortOrientation.xlSortColumns)
        'exlObj.CutCopyMode = False
        'exlObj.ScreenUpdating = False
        Dim hdrCommercial As Integer = RowHeader.Find("Commercial").Column
        DataRows.AutoFilter(Field:=hdrCommercial, Criteria1:= _
        "---- End of Break 1 ----")
        DataRows.SpecialCells(XlCellType.xlCellTypeVisible).EntireRow.Copy()
        calRange.PasteSpecial(XlPasteType.xlPasteAll)
        wsBrkSht.AutoFilterMode = False
        exlObj.CutCopyMode = False
        calRange = calRange.CurrentRegion

        NewRowHeader = wsCalcSht.Range("1:1")
        With wsCalcSht
            For Each cell As Range In RowHeader.SpecialCells(XlCellType.xlCellTypeConstants)

                Select Case Trim(cell.Value)
                    Case "Channel"
                        colChannel = .Range(calRange.Cells(1, cell.Column), calRange.Cells(calRange.Rows.Count, cell.Column))
                    Case "Programme"
                        colProgramme = .Range(calRange.Cells(1, cell.Column), calRange.Cells(calRange.Rows.Count, cell.Column))
                    Case "Date"
                        colDate = .Range(calRange.Cells(1, cell.Column), calRange.Cells(calRange.Rows.Count, cell.Column))
                    Case "Day"
                        colDay = .Range(calRange.Cells(1, cell.Column), calRange.Cells(calRange.Rows.Count, cell.Column))
                    Case "Start Time"
                        colStartTime = .Range(calRange.Cells(1, cell.Column), calRange.Cells(calRange.Rows.Count, cell.Column))
                    Case "PA"
                        colPA = .Range(calRange.Cells(1, cell.Column), calRange.Cells(calRange.Rows.Count, cell.Column))
                    Case "TA"
                        colTA = .Range(calRange.Cells(1, cell.Column), calRange.Cells(calRange.Rows.Count, cell.Column))
                    Case "Commercial"
                        colCommercial = .Range(calRange.Cells(1, cell.Column), calRange.Cells(calRange.Rows.Count, cell.Column))
                End Select
            Next
        End With

        colStartTime.EntireColumn.Insert(Shift:=XlInsertShiftDirection.xlShiftToRight)
        colProgStartTime = colStartTime.Offset(0, -1)
        colProgStartTime.Formula = "=TIME(HOUR( RC[1]),FLOOR( MINUTE( RC[1]), 30), 0)"
        NewRowHeader.Cells(1, colProgStartTime.Column).value = "Program Start Time"
        'Dim newPivotCache As PivotCache = wbBrkPerf.PivotCaches.Create(SourceType:=XlPivotTableSourceType.xlDatabase, SourceData:= _
        'calRange.AddressLocal, Version:=XlPivotTableVersionList.xlPivotTableVersion14)
        'Dim newPivot As PivotTable = newPivotCache.CreatePivotTable( _
        'TableDestination:="", TableName:="PivotTable1", DefaultVersion:= _
        'XlPivotTableVersionList.xlPivotTableVersion14)

        'ActiveSheet.PivotTableWizard(TableDestination:=ActiveSheet.Cells(3, 1))

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim newPivotRange As String = wsCalcSht.Name & "!R1C1:R" & calRange.Rows.Count & "C" & calRange.Columns.Count
        Dim newPivotCache As PivotCache = wbBrkPerf.PivotCaches.Create(SourceType:=XlPivotTableSourceType.xlDatabase, SourceData:= _
                newPivotRange, Version:=XlPivotTableVersionList.xlPivotTableVersion14)
        Dim newPivot As PivotTable = newPivotCache.CreatePivotTable( _
        TableDestination:="", TableName:="PivotTable1", DefaultVersion:= _
        XlPivotTableVersionList.xlPivotTableVersion14)
        Dim newPivotField As PivotField
        newPivotField = newPivot.PivotFields("Channel")
        With newPivotField
            .Orientation = XlPivotFieldOrientation.xlRowField
            .Position = 1
            .Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        End With
        newPivotField = newPivot.PivotFields(" Programme")
        With newPivotField
            .Orientation = XlPivotFieldOrientation.xlRowField
            .Position = 2
            .Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        End With
        newPivotField = newPivot.PivotFields("Program Start Time")
        With newPivotField
            .Orientation = XlPivotFieldOrientation.xlRowField
            .Position = 3
            .Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        End With
        newPivotField = newPivot.PivotFields(" Day")
        With newPivotField
            .Orientation = XlPivotFieldOrientation.xlRowField
            .Position = 4
            .Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        End With
        newPivot.RowAxisLayout(XlLayoutRowType.xlTabularRow)
        newPivot.RepeatAllLabels(XlPivotFieldRepeatLabels.xlRepeatLabels)
        Dim uniqueRange As Range = newPivot.TableRange1
        uniqueRange.Copy()
        uniqueRange.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats)
        Dim formulaRange As Range = uniqueRange.Columns(1).offset(1, 4)
        formulaRange.FormulaR1C1 = _
        "=IF(RC[-4]&RC[-3]&RC[-2]=R[-1]C[-4]&R[-1]C[-3]&R[-1]C[-2],R[-1]C&"",""&RC[-1],RC[-1])"
        formulaRange.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight)
        Dim formulaRange2 As Range = formulaRange.Offset(0, -1)
        formulaRange2.FormulaR1C1 = "=RC[-4]&RC[-3]&RC[-2]"
    End Sub
End Class

Class TAMTG
    Private _tgName As String
    Private _column As Range

    Sub New()
        _tgName = ""
        _column = Nothing
    End Sub
    Sub New(TGName As String, Column As Range)
        _tgName = TGName
        _column = Column
    End Sub

    Public Property TGName As String
        Get
            Return _tgName
        End Get
        Set(value As String)
            _tgName = value
        End Set
    End Property

    Public Property Column As Range
        Get
            Return _column
        End Get
        Set(value As Range)
            _column = value
        End Set
    End Property

End Class
