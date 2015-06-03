Partial Class MSprintExRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MSprintExRibbon))
        Me.tabMsprintEx = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.menuTopPrograms = Me.Factory.CreateRibbonMenu
        Me.btnBreakTVR = Me.Factory.CreateRibbonButton
        Me.btnProgramTVR = Me.Factory.CreateRibbonButton
        Me.Group9 = Me.Factory.CreateRibbonGroup
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.Group6 = Me.Factory.CreateRibbonGroup
        Me.btnGenerateEndTime = Me.Factory.CreateRibbonButton
        Me.Group7 = Me.Factory.CreateRibbonGroup
        Me.btnGetReqSpots = Me.Factory.CreateRibbonButton
        Me.btnRnF = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Group10 = Me.Factory.CreateRibbonGroup
        Me.btnSpotSelect = Me.Factory.CreateRibbonButton
        Me.btnSpotReplace = Me.Factory.CreateRibbonButton
        Me.btnDeleteSpot = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.btnMarketSummary = Me.Factory.CreateRibbonButton
        Me.btnChannelSummary = Me.Factory.CreateRibbonButton
        Me.btnDurationSummary = Me.Factory.CreateRibbonButton
        Me.btnCreativeSummary = Me.Factory.CreateRibbonButton
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.btnAllSummary = Me.Factory.CreateRibbonButton
        Me.Group8 = Me.Factory.CreateRibbonGroup
        Me.btnLogFile = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.btnTGSelectionShowHide = Me.Factory.CreateRibbonButton
        Me.ToggleButton1 = Me.Factory.CreateRibbonToggleButton
        Me.btnShowHideAvgTVrPlan = Me.Factory.CreateRibbonButton
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.DirectorySearcher1 = New System.DirectoryServices.DirectorySearcher()
        Me.btnStartMsprint = Me.Factory.CreateRibbonButton
        Me.btnChangeLogDir = Me.Factory.CreateRibbonButton
        Me.menGenreShare = Me.Factory.CreateRibbonMenu
        Me.btnGenreShareAll = Me.Factory.CreateRibbonButton
        Me.btnGenreShareTopTen = Me.Factory.CreateRibbonButton
        Me.btnChannelShare = Me.Factory.CreateRibbonButton
        Me.btnProgAvgTVR = Me.Factory.CreateRibbonButton
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.btnOpen = Me.Factory.CreateRibbonButton
        Me.menubtnOpenPlan = Me.Factory.CreateRibbonButton
        Me.btnAddPlan = Me.Factory.CreateRibbonButton
        Me.btnSpotSplit = Me.Factory.CreateRibbonButton
        Me.menuBtnSavePlan = Me.Factory.CreateRibbonButton
        Me.btnCleanupplan = Me.Factory.CreateRibbonButton
        Me.btnMapChannels = Me.Factory.CreateRibbonButton
        Me.btnReorderPlanChannels = Me.Factory.CreateRibbonButton
        Me.tabMsprintEx.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group9.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.Group6.SuspendLayout()
        Me.Group7.SuspendLayout()
        Me.Group10.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group8.SuspendLayout()
        Me.Group4.SuspendLayout()
        '
        'tabMsprintEx
        '
        Me.tabMsprintEx.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.tabMsprintEx.Groups.Add(Me.Group1)
        Me.tabMsprintEx.Groups.Add(Me.Group2)
        Me.tabMsprintEx.Groups.Add(Me.Group9)
        Me.tabMsprintEx.Groups.Add(Me.Group5)
        Me.tabMsprintEx.Groups.Add(Me.Group6)
        Me.tabMsprintEx.Groups.Add(Me.Group7)
        Me.tabMsprintEx.Groups.Add(Me.Group10)
        Me.tabMsprintEx.Groups.Add(Me.Group3)
        Me.tabMsprintEx.Groups.Add(Me.Group8)
        Me.tabMsprintEx.Groups.Add(Me.Group4)
        Me.tabMsprintEx.Label = "MSprintX"
        Me.tabMsprintEx.Name = "tabMsprintEx"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.btnStartMsprint)
        Me.Group1.Items.Add(Me.btnChangeLogDir)
        Me.Group1.Label = "MSprintX"
        Me.Group1.Name = "Group1"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.menGenreShare)
        Me.Group2.Items.Add(Me.btnChannelShare)
        Me.Group2.Items.Add(Me.menuTopPrograms)
        Me.Group2.Label = "View"
        Me.Group2.Name = "Group2"
        '
        'menuTopPrograms
        '
        Me.menuTopPrograms.Enabled = False
        Me.menuTopPrograms.Items.Add(Me.btnBreakTVR)
        Me.menuTopPrograms.Items.Add(Me.btnProgramTVR)
        Me.menuTopPrograms.Label = "Top Programs"
        Me.menuTopPrograms.Name = "menuTopPrograms"
        '
        'btnBreakTVR
        '
        Me.btnBreakTVR.Label = "Break TVR"
        Me.btnBreakTVR.Name = "btnBreakTVR"
        Me.btnBreakTVR.ShowImage = True
        '
        'btnProgramTVR
        '
        Me.btnProgramTVR.Label = "Program TVR"
        Me.btnProgramTVR.Name = "btnProgramTVR"
        Me.btnProgramTVR.ShowImage = True
        '
        'Group9
        '
        Me.Group9.Items.Add(Me.btnProgAvgTVR)
        Me.Group9.Items.Add(Me.Button1)
        Me.Group9.Label = "Average TVR"
        Me.Group9.Name = "Group9"
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.btnOpen)
        Me.Group5.Items.Add(Me.menubtnOpenPlan)
        Me.Group5.Items.Add(Me.btnAddPlan)
        Me.Group5.Items.Add(Me.btnSpotSplit)
        Me.Group5.Items.Add(Me.menuBtnSavePlan)
        Me.Group5.Label = "Plan Options"
        Me.Group5.Name = "Group5"
        '
        'Group6
        '
        Me.Group6.Items.Add(Me.btnCleanupplan)
        Me.Group6.Items.Add(Me.btnMapChannels)
        Me.Group6.Items.Add(Me.btnReorderPlanChannels)
        Me.Group6.Items.Add(Me.btnGenerateEndTime)
        Me.Group6.Label = "Data"
        Me.Group6.Name = "Group6"
        '
        'btnGenerateEndTime
        '
        Me.btnGenerateEndTime.Enabled = False
        Me.btnGenerateEndTime.Label = "Generate End Time"
        Me.btnGenerateEndTime.Name = "btnGenerateEndTime"
        '
        'Group7
        '
        Me.Group7.Items.Add(Me.btnGetReqSpots)
        Me.Group7.Items.Add(Me.btnRnF)
        Me.Group7.Items.Add(Me.Button3)
        Me.Group7.Label = "Spots"
        Me.Group7.Name = "Group7"
        '
        'btnGetReqSpots
        '
        Me.btnGetReqSpots.Enabled = False
        Me.btnGetReqSpots.Label = "Get Spots"
        Me.btnGetReqSpots.Name = "btnGetReqSpots"
        '
        'btnRnF
        '
        Me.btnRnF.Enabled = False
        Me.btnRnF.Label = "Reach n Frequency"
        Me.btnRnF.Name = "btnRnF"
        '
        'Button3
        '
        Me.Button3.Label = "Generate Schedule"
        Me.Button3.Name = "Button3"
        Me.Button3.Visible = False
        '
        'Group10
        '
        Me.Group10.Items.Add(Me.btnSpotSelect)
        Me.Group10.Items.Add(Me.btnSpotReplace)
        Me.Group10.Items.Add(Me.btnDeleteSpot)
        Me.Group10.Label = "Spot Operations"
        Me.Group10.Name = "Group10"
        '
        'btnSpotSelect
        '
        Me.btnSpotSelect.Enabled = False
        Me.btnSpotSelect.Label = "Select"
        Me.btnSpotSelect.Name = "btnSpotSelect"
        '
        'btnSpotReplace
        '
        Me.btnSpotReplace.Enabled = False
        Me.btnSpotReplace.Label = "Replace"
        Me.btnSpotReplace.Name = "btnSpotReplace"
        '
        'btnDeleteSpot
        '
        Me.btnDeleteSpot.Enabled = False
        Me.btnDeleteSpot.Label = "Delete"
        Me.btnDeleteSpot.Name = "btnDeleteSpot"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.btnMarketSummary)
        Me.Group3.Items.Add(Me.btnChannelSummary)
        Me.Group3.Items.Add(Me.btnDurationSummary)
        Me.Group3.Items.Add(Me.btnCreativeSummary)
        Me.Group3.Items.Add(Me.Button2)
        Me.Group3.Items.Add(Me.btnAllSummary)
        Me.Group3.Label = "RnF Summary"
        Me.Group3.Name = "Group3"
        '
        'btnMarketSummary
        '
        Me.btnMarketSummary.Enabled = False
        Me.btnMarketSummary.Label = "Market"
        Me.btnMarketSummary.Name = "btnMarketSummary"
        '
        'btnChannelSummary
        '
        Me.btnChannelSummary.Enabled = False
        Me.btnChannelSummary.Label = "Channel"
        Me.btnChannelSummary.Name = "btnChannelSummary"
        '
        'btnDurationSummary
        '
        Me.btnDurationSummary.Enabled = False
        Me.btnDurationSummary.Label = "Duration"
        Me.btnDurationSummary.Name = "btnDurationSummary"
        '
        'btnCreativeSummary
        '
        Me.btnCreativeSummary.Enabled = False
        Me.btnCreativeSummary.Label = "Creative"
        Me.btnCreativeSummary.Name = "btnCreativeSummary"
        '
        'Button2
        '
        Me.Button2.Label = "BSL Summary"
        Me.Button2.Name = "Button2"
        Me.Button2.Visible = False
        '
        'btnAllSummary
        '
        Me.btnAllSummary.Enabled = False
        Me.btnAllSummary.Label = "All"
        Me.btnAllSummary.Name = "btnAllSummary"
        '
        'Group8
        '
        Me.Group8.Items.Add(Me.btnLogFile)
        Me.Group8.Label = "Log File"
        Me.Group8.Name = "Group8"
        '
        'btnLogFile
        '
        Me.btnLogFile.Enabled = False
        Me.btnLogFile.Label = "Write Logfile"
        Me.btnLogFile.Name = "btnLogFile"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.btnTGSelectionShowHide)
        Me.Group4.Items.Add(Me.ToggleButton1)
        Me.Group4.Items.Add(Me.btnShowHideAvgTVrPlan)
        Me.Group4.Label = "Show/Hide Panels"
        Me.Group4.Name = "Group4"
        '
        'btnTGSelectionShowHide
        '
        Me.btnTGSelectionShowHide.Label = "Genre / Channel Outputs"
        Me.btnTGSelectionShowHide.Name = "btnTGSelectionShowHide"
        '
        'ToggleButton1
        '
        Me.ToggleButton1.Label = "Audience / Markets"
        Me.ToggleButton1.Name = "ToggleButton1"
        '
        'btnShowHideAvgTVrPlan
        '
        Me.btnShowHideAvgTVrPlan.Label = "AvgTVR MG Selection"
        Me.btnShowHideAvgTVrPlan.Name = "btnShowHideAvgTVrPlan"
        '
        'BackgroundWorker1
        '
        '
        'DirectorySearcher1
        '
        Me.DirectorySearcher1.ClientTimeout = System.TimeSpan.Parse("-00:00:01")
        Me.DirectorySearcher1.ServerPageTimeLimit = System.TimeSpan.Parse("-00:00:01")
        Me.DirectorySearcher1.ServerTimeLimit = System.TimeSpan.Parse("-00:00:01")
        '
        'btnStartMsprint
        '
        Me.btnStartMsprint.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnStartMsprint.Image = Global.MSprintEx.My.Resources.Resources.Start_Button
        Me.btnStartMsprint.Label = "Start MSprintX"
        Me.btnStartMsprint.Name = "btnStartMsprint"
        Me.btnStartMsprint.ShowImage = True
        '
        'btnChangeLogDir
        '
        Me.btnChangeLogDir.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnChangeLogDir.Enabled = False
        Me.btnChangeLogDir.Image = Global.MSprintEx.My.Resources.Resources.analyze
        Me.btnChangeLogDir.Label = "View/Change Log Directory"
        Me.btnChangeLogDir.Name = "btnChangeLogDir"
        Me.btnChangeLogDir.ShowImage = True
        '
        'menGenreShare
        '
        Me.menGenreShare.Enabled = False
        Me.menGenreShare.Image = Global.MSprintEx.My.Resources.Resources.genre1
        Me.menGenreShare.Items.Add(Me.btnGenreShareAll)
        Me.menGenreShare.Items.Add(Me.btnGenreShareTopTen)
        Me.menGenreShare.Label = "Genre Share"
        Me.menGenreShare.Name = "menGenreShare"
        Me.menGenreShare.ShowImage = True
        '
        'btnGenreShareAll
        '
        Me.btnGenreShareAll.Enabled = False
        Me.btnGenreShareAll.Label = "All"
        Me.btnGenreShareAll.Name = "btnGenreShareAll"
        Me.btnGenreShareAll.ShowImage = True
        '
        'btnGenreShareTopTen
        '
        Me.btnGenreShareTopTen.Enabled = False
        Me.btnGenreShareTopTen.Label = "TopTen"
        Me.btnGenreShareTopTen.Name = "btnGenreShareTopTen"
        Me.btnGenreShareTopTen.ShowImage = True
        '
        'btnChannelShare
        '
        Me.btnChannelShare.Enabled = False
        Me.btnChannelShare.Image = Global.MSprintEx.My.Resources.Resources.channel
        Me.btnChannelShare.Label = "Channel Share"
        Me.btnChannelShare.Name = "btnChannelShare"
        Me.btnChannelShare.ShowImage = True
        '
        'btnProgAvgTVR
        '
        Me.btnProgAvgTVR.Enabled = False
        Me.btnProgAvgTVR.Image = Global.MSprintEx.My.Resources.Resources.images
        Me.btnProgAvgTVR.Label = " AvgTVR Sheet"
        Me.btnProgAvgTVR.Name = "btnProgAvgTVR"
        Me.btnProgAvgTVR.ShowImage = True
        '
        'Button1
        '
        Me.Button1.Enabled = False
        Me.Button1.Image = Global.MSprintEx.My.Resources.Resources.tvr1
        Me.Button1.Label = "Get AvgTVR"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'btnOpen
        '
        Me.btnOpen.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnOpen.Enabled = False
        Me.btnOpen.Image = Global.MSprintEx.My.Resources.Resources.Folder
        Me.btnOpen.Label = "New Plan"
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.ShowImage = True
        '
        'menubtnOpenPlan
        '
        Me.menubtnOpenPlan.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.menubtnOpenPlan.Enabled = False
        Me.menubtnOpenPlan.Image = Global.MSprintEx.My.Resources.Resources.Open_file
        Me.menubtnOpenPlan.Label = "Open Plan"
        Me.menubtnOpenPlan.Name = "menubtnOpenPlan"
        Me.menubtnOpenPlan.ShowImage = True
        '
        'btnAddPlan
        '
        Me.btnAddPlan.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnAddPlan.Enabled = False
        Me.btnAddPlan.Image = Global.MSprintEx.My.Resources.Resources.Upload1
        Me.btnAddPlan.Label = "Add Plan"
        Me.btnAddPlan.Name = "btnAddPlan"
        Me.btnAddPlan.ShowImage = True
        '
        'btnSpotSplit
        '
        Me.btnSpotSplit.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnSpotSplit.Image = Global.MSprintEx.My.Resources.Resources.split
        Me.btnSpotSplit.Label = "Split Spots "
        Me.btnSpotSplit.Name = "btnSpotSplit"
        Me.btnSpotSplit.ShowImage = True
        Me.btnSpotSplit.Visible = False
        '
        'menuBtnSavePlan
        '
        Me.menuBtnSavePlan.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.menuBtnSavePlan.Enabled = False
        Me.menuBtnSavePlan.Image = Global.MSprintEx.My.Resources.Resources.save
        Me.menuBtnSavePlan.Label = "Save Plan"
        Me.menuBtnSavePlan.Name = "menuBtnSavePlan"
        Me.menuBtnSavePlan.ShowImage = True
        '
        'btnCleanupplan
        '
        Me.btnCleanupplan.Enabled = False
        Me.btnCleanupplan.Image = Global.MSprintEx.My.Resources.Resources.Tasks
        Me.btnCleanupplan.Label = "Clean Up Plan"
        Me.btnCleanupplan.Name = "btnCleanupplan"
        Me.btnCleanupplan.ShowImage = True
        '
        'btnMapChannels
        '
        Me.btnMapChannels.Enabled = False
        Me.btnMapChannels.Image = Global.MSprintEx.My.Resources.Resources.table_link
        Me.btnMapChannels.Label = "Map Channels"
        Me.btnMapChannels.Name = "btnMapChannels"
        Me.btnMapChannels.ShowImage = True
        '
        'btnReorderPlanChannels
        '
        Me.btnReorderPlanChannels.Enabled = False
        Me.btnReorderPlanChannels.Image = CType(resources.GetObject("btnReorderPlanChannels.Image"), System.Drawing.Image)
        Me.btnReorderPlanChannels.Label = "Rearrange Plan Channels"
        Me.btnReorderPlanChannels.Name = "btnReorderPlanChannels"
        Me.btnReorderPlanChannels.ShowImage = True
        '
        'MSprintExRibbon
        '
        Me.Name = "MSprintExRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.tabMsprintEx)
        Me.tabMsprintEx.ResumeLayout(False)
        Me.tabMsprintEx.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group9.ResumeLayout(False)
        Me.Group9.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.Group6.ResumeLayout(False)
        Me.Group6.PerformLayout()
        Me.Group7.ResumeLayout(False)
        Me.Group7.PerformLayout()
        Me.Group10.ResumeLayout(False)
        Me.Group10.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group8.ResumeLayout(False)
        Me.Group8.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()

    End Sub

    Friend WithEvents tabMsprintEx As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnOpen As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnChannelShare As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents menuTopPrograms As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnCleanupplan As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnMapChannels As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnRnF As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnLogFile As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnBreakTVR As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnProgramTVR As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnStartMsprint As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnTGSelectionShowHide As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents menubtnOpenPlan As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents menuBtnSavePlan As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnAddPlan As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnMarketSummary As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnChannelSummary As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnDurationSummary As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnCreativeSummary As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnGetReqSpots As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnChangeLogDir As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group6 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group7 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnReorderPlanChannels As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents menGenreShare As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents btnGenreShareAll As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnGenreShareTopTen As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group8 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents DirectorySearcher1 As System.DirectoryServices.DirectorySearcher
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSpotSplit As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnShowHideAvgTVrPlan As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnAllSummary As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group9 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnProgAvgTVR As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group10 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnSpotSelect As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSpotReplace As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnDeleteSpot As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ToggleButton1 As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents btnGenerateEndTime As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property MSprintExRibbon() As MSprintExRibbon
        Get
            Return Me.GetRibbon(Of MSprintExRibbon)()
        End Get
    End Property
End Class
