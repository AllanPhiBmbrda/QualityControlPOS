<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainMenu
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainMenu))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UserToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RegisterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LogOutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EmployeesModuleToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewEmployee20ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Employee30FinalToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.InsertEmployeesWorkToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.InsertEmployeesWorkFastToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportModuleToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PeriodeReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FieldReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PPH21ReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CouponGenerateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UtilityModuleToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DateControlSetupToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DateHolidaySetupToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.IncentiveDateSetUpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.IncentivesControlToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EntityRemoverToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.KeluarKerjaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UploadEmployeeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UploadEmployeeToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.UploadIncentivesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UploadCutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UpDatePayNoRekToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ModificationOfEntityToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StandardToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SynchFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SalarySynchToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EmployeeSynchToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MainTotalSynchToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SuratDoktorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportExportCSVFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ClearCacheToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutQCCodexToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.QCStatus = New System.Windows.Forms.ToolStripStatusLabel()
        Me.QCStatus2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.SSTab1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.SSTab2 = New System.Windows.Forms.ToolStripProgressBar()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.MenuStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.Color.LightSkyBlue
        Me.MenuStrip1.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.EmployeesModuleToolStripMenuItem, Me.ReportModuleToolStripMenuItem, Me.UtilityModuleToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(867, 26)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UserToolStripMenuItem, Me.LogOutToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(43, 22)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'UserToolStripMenuItem
        '
        Me.UserToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RegisterToolStripMenuItem})
        Me.UserToolStripMenuItem.Image = CType(resources.GetObject("UserToolStripMenuItem.Image"), System.Drawing.Image)
        Me.UserToolStripMenuItem.Name = "UserToolStripMenuItem"
        Me.UserToolStripMenuItem.Size = New System.Drawing.Size(123, 22)
        Me.UserToolStripMenuItem.Text = "User"
        '
        'RegisterToolStripMenuItem
        '
        Me.RegisterToolStripMenuItem.Image = CType(resources.GetObject("RegisterToolStripMenuItem.Image"), System.Drawing.Image)
        Me.RegisterToolStripMenuItem.Name = "RegisterToolStripMenuItem"
        Me.RegisterToolStripMenuItem.Size = New System.Drawing.Size(165, 22)
        Me.RegisterToolStripMenuItem.Text = "Register/View"
        '
        'LogOutToolStripMenuItem
        '
        Me.LogOutToolStripMenuItem.Image = CType(resources.GetObject("LogOutToolStripMenuItem.Image"), System.Drawing.Image)
        Me.LogOutToolStripMenuItem.Name = "LogOutToolStripMenuItem"
        Me.LogOutToolStripMenuItem.Size = New System.Drawing.Size(123, 22)
        Me.LogOutToolStripMenuItem.Text = "Log Out"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Image = CType(resources.GetObject("ExitToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(123, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'EmployeesModuleToolStripMenuItem
        '
        Me.EmployeesModuleToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ViewEmployee20ToolStripMenuItem, Me.Employee30FinalToolStripMenuItem, Me.InsertEmployeesWorkToolStripMenuItem, Me.InsertEmployeesWorkFastToolStripMenuItem})
        Me.EmployeesModuleToolStripMenuItem.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EmployeesModuleToolStripMenuItem.Name = "EmployeesModuleToolStripMenuItem"
        Me.EmployeesModuleToolStripMenuItem.Size = New System.Drawing.Size(140, 22)
        Me.EmployeesModuleToolStripMenuItem.Text = "&Employees Module"
        '
        'ViewEmployee20ToolStripMenuItem
        '
        Me.ViewEmployee20ToolStripMenuItem.Name = "ViewEmployee20ToolStripMenuItem"
        Me.ViewEmployee20ToolStripMenuItem.Size = New System.Drawing.Size(261, 22)
        Me.ViewEmployee20ToolStripMenuItem.Text = "View Employee"
        '
        'Employee30FinalToolStripMenuItem
        '
        Me.Employee30FinalToolStripMenuItem.Image = CType(resources.GetObject("Employee30FinalToolStripMenuItem.Image"), System.Drawing.Image)
        Me.Employee30FinalToolStripMenuItem.Name = "Employee30FinalToolStripMenuItem"
        Me.Employee30FinalToolStripMenuItem.Size = New System.Drawing.Size(261, 22)
        Me.Employee30FinalToolStripMenuItem.Text = "Employee Control"
        '
        'InsertEmployeesWorkToolStripMenuItem
        '
        Me.InsertEmployeesWorkToolStripMenuItem.Name = "InsertEmployeesWorkToolStripMenuItem"
        Me.InsertEmployeesWorkToolStripMenuItem.Size = New System.Drawing.Size(261, 22)
        Me.InsertEmployeesWorkToolStripMenuItem.Text = "Insert Employee's Work"
        '
        'InsertEmployeesWorkFastToolStripMenuItem
        '
        Me.InsertEmployeesWorkFastToolStripMenuItem.Image = CType(resources.GetObject("InsertEmployeesWorkFastToolStripMenuItem.Image"), System.Drawing.Image)
        Me.InsertEmployeesWorkFastToolStripMenuItem.Name = "InsertEmployeesWorkFastToolStripMenuItem"
        Me.InsertEmployeesWorkFastToolStripMenuItem.Size = New System.Drawing.Size(261, 22)
        Me.InsertEmployeesWorkFastToolStripMenuItem.Text = "Insert Employee's Work (Fast)"
        '
        'ReportModuleToolStripMenuItem
        '
        Me.ReportModuleToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PeriodeReportToolStripMenuItem, Me.FieldReportToolStripMenuItem, Me.PPH21ReportToolStripMenuItem, Me.CouponGenerateToolStripMenuItem})
        Me.ReportModuleToolStripMenuItem.Name = "ReportModuleToolStripMenuItem"
        Me.ReportModuleToolStripMenuItem.Size = New System.Drawing.Size(114, 22)
        Me.ReportModuleToolStripMenuItem.Text = "&Report Module"
        '
        'PeriodeReportToolStripMenuItem
        '
        Me.PeriodeReportToolStripMenuItem.Name = "PeriodeReportToolStripMenuItem"
        Me.PeriodeReportToolStripMenuItem.Size = New System.Drawing.Size(186, 22)
        Me.PeriodeReportToolStripMenuItem.Text = "Periode Report"
        '
        'FieldReportToolStripMenuItem
        '
        Me.FieldReportToolStripMenuItem.Name = "FieldReportToolStripMenuItem"
        Me.FieldReportToolStripMenuItem.Size = New System.Drawing.Size(186, 22)
        Me.FieldReportToolStripMenuItem.Text = "Field Report"
        '
        'PPH21ReportToolStripMenuItem
        '
        Me.PPH21ReportToolStripMenuItem.Name = "PPH21ReportToolStripMenuItem"
        Me.PPH21ReportToolStripMenuItem.Size = New System.Drawing.Size(186, 22)
        Me.PPH21ReportToolStripMenuItem.Text = "PPH21 Report"
        '
        'CouponGenerateToolStripMenuItem
        '
        Me.CouponGenerateToolStripMenuItem.Name = "CouponGenerateToolStripMenuItem"
        Me.CouponGenerateToolStripMenuItem.Size = New System.Drawing.Size(186, 22)
        Me.CouponGenerateToolStripMenuItem.Text = "Coupon Generate"
        '
        'UtilityModuleToolStripMenuItem
        '
        Me.UtilityModuleToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DateControlSetupToolStripMenuItem, Me.DateHolidaySetupToolStripMenuItem, Me.IncentiveDateSetUpToolStripMenuItem, Me.IncentivesControlToolStripMenuItem, Me.EntityRemoverToolStripMenuItem, Me.KeluarKerjaToolStripMenuItem, Me.UploadEmployeeToolStripMenuItem, Me.ModificationOfEntityToolStripMenuItem, Me.SynchFileToolStripMenuItem, Me.SuratDoktorToolStripMenuItem, Me.ImportExportCSVFileToolStripMenuItem})
        Me.UtilityModuleToolStripMenuItem.Name = "UtilityModuleToolStripMenuItem"
        Me.UtilityModuleToolStripMenuItem.Size = New System.Drawing.Size(111, 22)
        Me.UtilityModuleToolStripMenuItem.Text = "&Utility Module"
        '
        'DateControlSetupToolStripMenuItem
        '
        Me.DateControlSetupToolStripMenuItem.Name = "DateControlSetupToolStripMenuItem"
        Me.DateControlSetupToolStripMenuItem.Size = New System.Drawing.Size(223, 22)
        Me.DateControlSetupToolStripMenuItem.Text = "Date Control Setup"
        '
        'DateHolidaySetupToolStripMenuItem
        '
        Me.DateHolidaySetupToolStripMenuItem.Name = "DateHolidaySetupToolStripMenuItem"
        Me.DateHolidaySetupToolStripMenuItem.Size = New System.Drawing.Size(223, 22)
        Me.DateHolidaySetupToolStripMenuItem.Text = "Date Holiday Setup"
        '
        'IncentiveDateSetUpToolStripMenuItem
        '
        Me.IncentiveDateSetUpToolStripMenuItem.Name = "IncentiveDateSetUpToolStripMenuItem"
        Me.IncentiveDateSetUpToolStripMenuItem.Size = New System.Drawing.Size(223, 22)
        Me.IncentiveDateSetUpToolStripMenuItem.Text = "Incentive Date Set Up"
        '
        'IncentivesControlToolStripMenuItem
        '
        Me.IncentivesControlToolStripMenuItem.Name = "IncentivesControlToolStripMenuItem"
        Me.IncentivesControlToolStripMenuItem.Size = New System.Drawing.Size(223, 22)
        Me.IncentivesControlToolStripMenuItem.Text = "Incentives Control"
        '
        'EntityRemoverToolStripMenuItem
        '
        Me.EntityRemoverToolStripMenuItem.Name = "EntityRemoverToolStripMenuItem"
        Me.EntityRemoverToolStripMenuItem.Size = New System.Drawing.Size(223, 22)
        Me.EntityRemoverToolStripMenuItem.Text = "Entity Remover"
        '
        'KeluarKerjaToolStripMenuItem
        '
        Me.KeluarKerjaToolStripMenuItem.Name = "KeluarKerjaToolStripMenuItem"
        Me.KeluarKerjaToolStripMenuItem.Size = New System.Drawing.Size(223, 22)
        Me.KeluarKerjaToolStripMenuItem.Text = "Keluar Kerja"
        '
        'UploadEmployeeToolStripMenuItem
        '
        Me.UploadEmployeeToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UploadEmployeeToolStripMenuItem1, Me.UploadIncentivesToolStripMenuItem, Me.UploadCutToolStripMenuItem, Me.UpDatePayNoRekToolStripMenuItem})
        Me.UploadEmployeeToolStripMenuItem.Name = "UploadEmployeeToolStripMenuItem"
        Me.UploadEmployeeToolStripMenuItem.Size = New System.Drawing.Size(223, 22)
        Me.UploadEmployeeToolStripMenuItem.Text = "Upload/Update Item"
        '
        'UploadEmployeeToolStripMenuItem1
        '
        Me.UploadEmployeeToolStripMenuItem1.Name = "UploadEmployeeToolStripMenuItem1"
        Me.UploadEmployeeToolStripMenuItem1.Size = New System.Drawing.Size(198, 22)
        Me.UploadEmployeeToolStripMenuItem1.Text = "Upload Employee"
        '
        'UploadIncentivesToolStripMenuItem
        '
        Me.UploadIncentivesToolStripMenuItem.Name = "UploadIncentivesToolStripMenuItem"
        Me.UploadIncentivesToolStripMenuItem.Size = New System.Drawing.Size(198, 22)
        Me.UploadIncentivesToolStripMenuItem.Text = "Upload Incentives"
        '
        'UploadCutToolStripMenuItem
        '
        Me.UploadCutToolStripMenuItem.Name = "UploadCutToolStripMenuItem"
        Me.UploadCutToolStripMenuItem.Size = New System.Drawing.Size(198, 22)
        Me.UploadCutToolStripMenuItem.Text = "Upload Potongan"
        '
        'UpDatePayNoRekToolStripMenuItem
        '
        Me.UpDatePayNoRekToolStripMenuItem.Name = "UpDatePayNoRekToolStripMenuItem"
        Me.UpDatePayNoRekToolStripMenuItem.Size = New System.Drawing.Size(198, 22)
        Me.UpDatePayNoRekToolStripMenuItem.Text = "UpDate Pay/No Rek"
        '
        'ModificationOfEntityToolStripMenuItem
        '
        Me.ModificationOfEntityToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StandardToolStripMenuItem})
        Me.ModificationOfEntityToolStripMenuItem.Name = "ModificationOfEntityToolStripMenuItem"
        Me.ModificationOfEntityToolStripMenuItem.Size = New System.Drawing.Size(223, 22)
        Me.ModificationOfEntityToolStripMenuItem.Text = "Modification of Entity"
        '
        'StandardToolStripMenuItem
        '
        Me.StandardToolStripMenuItem.Name = "StandardToolStripMenuItem"
        Me.StandardToolStripMenuItem.Size = New System.Drawing.Size(131, 22)
        Me.StandardToolStripMenuItem.Text = "Standard"
        '
        'SynchFileToolStripMenuItem
        '
        Me.SynchFileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SalarySynchToolStripMenuItem, Me.EmployeeSynchToolStripMenuItem, Me.MainTotalSynchToolStripMenuItem})
        Me.SynchFileToolStripMenuItem.Name = "SynchFileToolStripMenuItem"
        Me.SynchFileToolStripMenuItem.Size = New System.Drawing.Size(223, 22)
        Me.SynchFileToolStripMenuItem.Text = "Synch File"
        Me.SynchFileToolStripMenuItem.Visible = False
        '
        'SalarySynchToolStripMenuItem
        '
        Me.SalarySynchToolStripMenuItem.Name = "SalarySynchToolStripMenuItem"
        Me.SalarySynchToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.SalarySynchToolStripMenuItem.Text = "Salary Synch"
        '
        'EmployeeSynchToolStripMenuItem
        '
        Me.EmployeeSynchToolStripMenuItem.Name = "EmployeeSynchToolStripMenuItem"
        Me.EmployeeSynchToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.EmployeeSynchToolStripMenuItem.Text = "Employee Synch"
        '
        'MainTotalSynchToolStripMenuItem
        '
        Me.MainTotalSynchToolStripMenuItem.Name = "MainTotalSynchToolStripMenuItem"
        Me.MainTotalSynchToolStripMenuItem.Size = New System.Drawing.Size(180, 22)
        Me.MainTotalSynchToolStripMenuItem.Text = "Main Total Synch"
        '
        'SuratDoktorToolStripMenuItem
        '
        Me.SuratDoktorToolStripMenuItem.Name = "SuratDoktorToolStripMenuItem"
        Me.SuratDoktorToolStripMenuItem.Size = New System.Drawing.Size(223, 22)
        Me.SuratDoktorToolStripMenuItem.Text = "Surat Doktor"
        '
        'ImportExportCSVFileToolStripMenuItem
        '
        Me.ImportExportCSVFileToolStripMenuItem.Name = "ImportExportCSVFileToolStripMenuItem"
        Me.ImportExportCSVFileToolStripMenuItem.Size = New System.Drawing.Size(223, 22)
        Me.ImportExportCSVFileToolStripMenuItem.Text = "Import / Export CSV File"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ClearCacheToolStripMenuItem, Me.AboutQCCodexToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(49, 22)
        Me.HelpToolStripMenuItem.Text = "&Help"
        '
        'ClearCacheToolStripMenuItem
        '
        Me.ClearCacheToolStripMenuItem.Name = "ClearCacheToolStripMenuItem"
        Me.ClearCacheToolStripMenuItem.Size = New System.Drawing.Size(177, 22)
        Me.ClearCacheToolStripMenuItem.Text = "Clear Cache"
        '
        'AboutQCCodexToolStripMenuItem
        '
        Me.AboutQCCodexToolStripMenuItem.Name = "AboutQCCodexToolStripMenuItem"
        Me.AboutQCCodexToolStripMenuItem.Size = New System.Drawing.Size(177, 22)
        Me.AboutQCCodexToolStripMenuItem.Text = "About QC Codex"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.QCStatus, Me.QCStatus2, Me.SSTab1, Me.SSTab2})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 376)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(867, 22)
        Me.StatusStrip1.TabIndex = 2
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'QCStatus
        '
        Me.QCStatus.BackColor = System.Drawing.Color.Transparent
        Me.QCStatus.BorderStyle = System.Windows.Forms.Border3DStyle.Raised
        Me.QCStatus.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.QCStatus.Name = "QCStatus"
        Me.QCStatus.Size = New System.Drawing.Size(47, 17)
        Me.QCStatus.Text = "Name: "
        '
        'QCStatus2
        '
        Me.QCStatus2.BackColor = System.Drawing.Color.Transparent
        Me.QCStatus2.BorderStyle = System.Windows.Forms.Border3DStyle.Raised
        Me.QCStatus2.Name = "QCStatus2"
        Me.QCStatus2.Size = New System.Drawing.Size(38, 17)
        Me.QCStatus2.Text = "Time:"
        '
        'SSTab1
        '
        Me.SSTab1.Name = "SSTab1"
        Me.SSTab1.Size = New System.Drawing.Size(115, 17)
        Me.SSTab1.Text = "     ||  Data Progress"
        Me.SSTab1.Visible = False
        '
        'SSTab2
        '
        Me.SSTab2.Maximum = 3000
        Me.SSTab2.Name = "SSTab2"
        Me.SSTab2.Size = New System.Drawing.Size(200, 16)
        Me.SSTab2.Visible = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'MainMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(867, 398)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "MainMenu"
        Me.Text = "QC Main Menu"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UserToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LogOutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents QCStatus As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents RegisterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReportModuleToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents QCStatus2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents UtilityModuleToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DateControlSetupToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DateHolidaySetupToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents KeluarKerjaToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UploadEmployeeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ModificationOfEntityToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StandardToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SynchFileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PeriodeReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FieldReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents IncentiveDateSetUpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SalarySynchToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EmployeeSynchToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PPH21ReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CouponGenerateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SuratDoktorToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MainTotalSynchToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UploadEmployeeToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UploadIncentivesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UploadCutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UpDatePayNoRekToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SSTab1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents SSTab2 As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents AboutQCCodexToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EntityRemoverToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents IncentivesControlToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EmployeesModuleToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ViewEmployee20ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Employee30FinalToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents InsertEmployeesWorkToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents InsertEmployeesWorkFastToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportExportCSVFileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ClearCacheToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
