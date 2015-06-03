Imports System.Windows.Forms
Imports MSprintEx.MSprintEx2014
Imports System.Diagnostics
Imports System.Xml
Imports System.IO

Public Class ucAudience
    Private tgArr As New List(Of TG)
    Dim tgCounter As Integer = 0
    Dim tgDirectoryPath As String = AppDomain.CurrentDomain.BaseDirectory & "\\Masters\TGS\\"
    Dim mgDirectoryPath As String = AppDomain.CurrentDomain.BaseDirectory & "\\Masters\MGS\\"
    Private Sub ucAudience_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim daGender As New METISTableAdapters.GENDERTableAdapter
        Dim daAgeBand As New METISTableAdapters.AGEBANDTableAdapter
        Dim daSEC As New METISTableAdapters.SECTableAdapter
        Dim daHH As New METISTableAdapters.HOUSEHOLDTableAdapter
        'Dim dtGender As METIS.GENDERDataTable = daGender.GetGenders
        'Dim dtAgeBand As METIS.AGEBANDDataTable = daAgeBand.GetAgeBands
        'Dim dtSEC As METIS.SECDataTable = daSEC.GetSEC
        'Dim dtHH As METIS.HOUSEHOLDDataTable = daHH.GetHouseholds
        'Dim fileList As List(Of String) = New List(Of String)()
        'clbGender.DataSource = dtGender
        'clbGender.DisplayMember = dtGender.Gender_DescColumn.ColumnName

        'clbAge.DataSource = dtAgeBand
        'clbAge.DisplayMember = dtAgeBand.AgeGroup_DescColumn.ColumnName

        'clbSEC.DataSource = dtSEC
        'clbSEC.DisplayMember = dtSEC.SEC_DescColumn.ColumnName

        'clbHH.DataSource = dtHH
        'clbHH.DisplayMember = dtHH.Household_DescColumn.ColumnName
        ' Dim di As DirectoryInfo = New DirectoryInfo()
        Try

       
        If Directory.Exists(tgDirectoryPath) Then
            For index = 0 To Directory.GetFiles(tgDirectoryPath, "*.xml").Count - 1
                'fileList.Add(Path.GetFileNameWithoutExtension(Directory.GetFiles(System.IO.Path.GetTempPath() + "\\TGS", "*.xml")(index)))
                lbTGDefs.Items.Add(Path.GetFileNameWithoutExtension(Directory.GetFiles(tgDirectoryPath, "*.xml")(index)))
            Next
        End If
        Dim dt As System.Data.DataTable = New System.Data.DataTable()
        dt.Columns.Add("Type")
        dt.Columns.Add("TG Name")
        dt.Rows.Add("Planning")
        dt.Rows.Add("Reference")
        DgPlanRefGrid.DataSource = dt
        DgPlanRefGrid.Columns(1).Width = 65
        DgPlanRefGrid.Columns(2).Width = 75
        DgPlanRefGrid.Columns(1).ReadOnly = True
            DgPlanRefGrid.Columns(0).DisplayIndex = 2
        Catch ex As Exception
            LogMpsrintExException("Exception occured while loading Audience tab." + ex.Message)
        End Try
    End Sub

    Private Sub chkSECAll_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkSECAll.CheckedChanged
        For intSec As Integer = 0 To clbSEC.Items.Count - 1
            clbSEC.SetItemChecked(intSec, chkSECAll.Checked)
        Next
    End Sub

    Private Sub chkAgeAll_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkAgeAll.CheckedChanged
        For intSec As Integer = 0 To clbAge.Items.Count - 1
            clbAge.SetItemChecked(intSec, chkAgeAll.Checked)
        Next
    End Sub

    Private Sub btnCreateTG_Click(sender As System.Object, e As System.EventArgs) Handles btnCreateTG.Click
        Dim csValue As String = String.Empty
        Dim secValue As String = String.Empty
        Dim sexValue As String = String.Empty
        Dim ageValue As String = String.Empty
        Try
            If tgArr.Count = 5 Then
                MsgBox("Permissible limit of 5 TGs has already been reached. Delete" & " some tg definition and then try again", MsgBoxStyle.Exclamation, "TG limits")
                Return
            End If
            ' setting the two static counters to false to enable estimation of weights 
            ' and loading of tgs again

            Dim tgidStr As String = ""
            Dim tgName As String = ""
            Dim tgDef As TG

            If Me.clbGender.CheckedItems.Count = 0 OrElse Me.clbAge.CheckedItems.Count = 0 OrElse Me.clbSEC.CheckedItems.Count = 0 OrElse Me.clbHH.CheckedItems.Count = 0 Then
                MsgBox("Each of the four parameters - Gender/Age/SEC/Household" & " has to be selected.", MsgBoxStyle.Information, "Insufficient TG parameters")
            Else
                ' all four sections have some databox checked. create the tg object
                tgDef = New TG()
                For i As Integer = 0 To Me.clbGender.Items.Count - 1
                    If Me.clbGender.GetItemCheckState(i) = CheckState.Checked Then
                        tgDef.setGenderIds(i)
                        If sexValue <> String.Empty Then
                            sexValue = sexValue + ","
                        End If
                        sexValue = sexValue + (i + 1).ToString()
                    End If
                Next
                For i As Integer = 0 To Me.clbAge.Items.Count - 1
                    If Me.clbAge.GetItemCheckState(i) = CheckState.Checked Then
                        tgDef.setAgeIds(i)
                        If ageValue <> String.Empty Then
                            ageValue = ageValue + ","
                        End If
                        ageValue = ageValue + (i + 1).ToString()
                    End If
                Next
                For i As Integer = 0 To Me.clbSEC.Items.Count - 1
                    If Me.clbSEC.GetItemCheckState(i) = CheckState.Checked Then
                        tgDef.setSECIds(i)
                        If secValue <> String.Empty Then
                            secValue = secValue + ","
                        End If
                        secValue = secValue + (i + 1).ToString()
                    End If
                Next
                For i As Integer = 0 To Me.clbHH.Items.Count - 1
                    If Me.clbHH.GetItemCheckState(i) = CheckState.Checked Then
                        tgDef.setHHIds(i)

                        If csValue <> String.Empty Then
                            csValue = csValue + ","
                        End If
                        csValue = csValue + (i + 1).ToString()
                    End If
                Next
                ' set the tgDef's unique idstring
                tgDef.setUniqueID()
                tgidStr = tgDef.getUniqueID()
                ' verify that the tgdef combination is unique
                If Not tgArr.Exists(Function(x) x.getUniqueID() = tgidStr) Then
                    If Trim(txtTGInput.Text) = "" Then
                        MsgBox("Please enter a valid name for the TG.", MsgBoxStyle.Information, "TG name is missing")
                        Return
                    End If
                    tgName = txtTGInput.Text.ToUpper().Trim()
                    tgDef.setTGName(tgName)
                    If Not tgArr.Exists(Function(x) x.getTGName() = tgName) Then
                        tgDef.setQueryString()
                        tgArr.Add(tgDef)
                        lbTGDefs.Items.Add(tgName)
                        Dim temppath As String = System.IO.Path.GetTempPath()

                        If Not (Directory.Exists(temppath + "\\TGS")) Then
                            Directory.CreateDirectory(temppath + "\\TGS")
                        End If
                        Dim c As XElement =
                        <tg name=<%= tgName %>>
                            <cs><%= csValue %></cs>
                            <sec><%= secValue %></sec>
                            <sex><%= sexValue %></sex>
                            <age><%= ageValue %></age>
                        </tg>

                     
                        c.Save(tgDirectoryPath + tgName + ".xml")
                     
                        '  Dim attribute As XmlAttribute = New XmlAttribute()
                        'Dim node As XmlNode = doc.CreateNode(XmlNodeType.Document, "tg", String.Empty)
                        'node.Attributes.Append(newAttr)
                        'Dim node1 As XmlNode = doc.CreateNode(XmlNodeType.Text, "cs", "cs")
                        '' doc.
                        'node1.Value = csValue

                        '' doc.AppendChild(node1)
                        'Dim secNode As XmlNode = doc.CreateNode(XmlNodeType.Text, "sec", "sec")
                        'secNode.Value = secValue
                        '' doc.AppendChild(secNode)
                        'Dim sexNode As XmlNode = doc.CreateNode(XmlNodeType.Text, "sex", "sex")
                        'sexNode.Value = sexValue
                        '' doc.AppendChild(sexNode)
                        'Dim ageNode As XmlNode = doc.CreateNode(XmlNodeType.Text, "age", "age")
                        'ageNode.Value = ageValue
                        ''doc.AppendChild(ageNode)
                        'doc.AppendChild(node)
                        'doc.Save()

                    Else
                        MsgBox("You have already assigned this name to an earlier selection. Please choose a distinct name", MsgBoxStyle.Exclamation, "Duplicate TG name")
                    End If
                Else
                    MsgBox("The same parameters are already selected for a different TG. Please modify your selections", MsgBoxStyle.Exclamation, "Duplicate TG selection")
                End If
            End If
            clearChecks()
        Catch ex As Exception
            LogMpsrintExException("Exception occured while creating TG" + ex.Message)
        End Try
    End Sub

    Private Sub lbTGDefs_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles lbTGDefs.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Delete Then
            Try
                Dim strDeleteTG As String = lbTGDefs.SelectedItem
                tgArr.Remove(tgArr.Find(Function(x) x.getTGName() = strDeleteTG))
                File.Delete(tgDirectoryPath & "\\" & lbTGDefs.SelectedItem.ToString() & ".xml")
                lbTGDefs.Items.Remove(lbTGDefs.SelectedItem)


            Catch ex As Exception
                '  Debug.WriteLine(ex.ToString)
                LogMpsrintExException("Exception occured while deleting selected TG" + ex.Message)
            End Try
        End If
    End Sub

    Private Sub lbTGDefs_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles lbTGDefs.SelectedIndexChanged

        Dim currTG As TG
        Try
            currTG = tgArr.Find(Function(x) x.getTGName() = lbTGDefs.SelectedItem)
            clearChecks()
            Dim tgelement As XElement = XElement.Load(tgDirectoryPath & "\\" & lbTGDefs.SelectedItem.ToString() & ".xml")
            Dim gender As String = tgelement.Element("sex").Value
            Dim age As String = tgelement.Element("age").Value
            Dim sec As String = tgelement.Element("sec").Value
            Dim cs As String = tgelement.Element("cs").Value
            Dim gendervalues As String() = gender.Split({","c}, StringSplitOptions.None)
            Dim agevalues As String() = age.Split({","c}, StringSplitOptions.None)
            Dim secvalues As String() = sec.Split({","c}, StringSplitOptions.None)
            Dim csvalues As String() = cs.Split({","c}, StringSplitOptions.None)
            For Each Str As String In gendervalues
                Me.clbGender.SetItemCheckState(Convert.ToInt32(Str) - 1, CheckState.Checked)
            Next
            For Each Str As String In agevalues
                Me.clbAge.SetItemCheckState(Convert.ToInt32(Str) - 1, CheckState.Checked)
            Next
            For Each Str As String In secvalues
                Me.clbSEC.SetItemCheckState(Convert.ToInt32(Str) - 1, CheckState.Checked)
            Next
            For Each Str As String In csvalues
                Me.clbHH.SetItemCheckState(Convert.ToInt32(Str) - 1, CheckState.Checked)
            Next
            'If tgelement.HasElements Then
            '    Dim elements As XElement() = tgelement.e
            'End If
        Catch ex As Exception
            LogMpsrintExException("Exception occured while setting properties of selected TG" + ex.Message)
        End Try
        Try
            'Dim cbxTG As System.Windows.Forms.CheckBox = DirectCast(sender, CheckBox)
            For Each tgdef As TG In tgArr
                'for1begin

                If tgdef.getTGName() = currTG.getTGName Then
                    'if1begin
                    For j As Integer = 0 To Me.clbGender.Items.Count - 1
                        If tgdef.getGenderIds(j) = 1 Then
                            Me.clbGender.SetItemCheckState(j, CheckState.Checked)
                        Else
                            Me.clbGender.SetItemCheckState(j, CheckState.Unchecked)
                        End If
                    Next

                    For k As Integer = 0 To Me.clbAge.Items.Count - 1
                        If tgdef.getAgeIds(k) = 1 Then
                            Me.clbAge.SetItemCheckState(k, CheckState.Checked)
                        Else
                            Me.clbAge.SetItemCheckState(k, CheckState.Unchecked)
                        End If
                    Next

                    For l As Integer = 0 To Me.clbSEC.Items.Count - 1
                        If tgdef.getSECIds(l) = 1 Then
                            Me.clbSEC.SetItemCheckState(l, CheckState.Checked)
                        Else
                            Me.clbSEC.SetItemCheckState(l, CheckState.Unchecked)
                        End If
                    Next

                    For k As Integer = 0 To Me.clbHH.Items.Count - 1
                        If tgdef.getHHIds(k) = 1 Then
                            Me.clbHH.SetItemCheckState(k, CheckState.Checked)
                        Else
                            Me.clbHH.SetItemCheckState(k, CheckState.Unchecked)
                        End If

                    Next
                End If
            Next
        Catch ex As Exception
            'writeExceptionIntoDB(ex)
            LogMpsrintExException("Exception occured while setting properties of selected TG" + ex.Message)
        End Try

    End Sub
    Private Sub clearChecks()
        For Each item In clbAge.CheckedIndices
            clbAge.SetItemChecked(item, False)
        Next
        clbAge.ClearSelected()
        For Each item In clbGender.CheckedIndices
            clbGender.SetItemChecked(item, False)
        Next
        clbGender.ClearSelected()
        For Each item In clbHH.CheckedIndices
            clbHH.SetItemChecked(item, False)
        Next
        clbHH.ClearSelected()
        For Each item In clbSEC.CheckedIndices
            clbSEC.SetItemChecked(item, False)
        Next
        clbSEC.ClearSelected()
        txtTGInput.Text = ""
    End Sub

    Private Sub btnSetasPlan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetasPlan.Click
        Dim dt As System.Data.DataTable = DirectCast(DgPlanRefGrid.DataSource, System.Data.DataTable)
        dt.Rows(0)(1) = lbTGDefs.SelectedItem.ToString()
        DgPlanRefGrid.DataSource = dt
    End Sub

    Private Sub btnRefTG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefTG.Click
        Dim dt As System.Data.DataTable = DirectCast(DgPlanRefGrid.DataSource, System.Data.DataTable)
        dt.Rows(1)(1) = lbTGDefs.SelectedItem.ToString()
        DgPlanRefGrid.DataSource = dt
    End Sub

    'Private Sub ChbDelRefTG_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChbDelRefTG.CheckedChanged
    '    Dim plandt As Data.DataTable = DirectCast(DgPlanRefGrid.DataSource, System.Data.DataTable)
    '    plandt.Rows(1)(1) = String.Empty
    '    DgPlanRefGrid.DataSource = plandt
    'End Sub
    Private Sub DgPlanRefGrid_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgPlanRefGrid.CellClick
        If e.ColumnIndex = 0 Then
            DgPlanRefGrid.Rows(e.RowIndex).Cells(2).Value = String.Empty
        End If
    End Sub
End Class
