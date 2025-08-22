Public Class MainMenu


    Private Delegate Sub RepFrm1()

    Private Sub MainMenu_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        LoginDoor.Close()
    End Sub

    Private Sub MainMenu_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F3 Then
            If WorkBlock.Visible = True Then
                WorkBlock.PanelBtn1.PerformClick()

            End If
        ElseIf e.KeyCode = Keys.F4 Then
            If WorkBlock.Visible = True Then
                WorkBlock.PanelBtn2.PerformClick()

            End If

        ElseIf e.KeyCode = Keys.F5 Then

            If WorkBlock.Visible = True Then

                If WorkBlock.ConSave.Visible = True Then
                    WorkBlock.ConSave.PerformClick()

                ElseIf WorkBlock.WallSave.Visible = True Then
                    WorkBlock.WallSave.PerformClick()

                ElseIf WorkBlock.MutuSave.Visible = True Then
                    WorkBlock.MutuSave.PerformClick()

                ElseIf WorkBlock.PackSave.Visible = True Then
                    WorkBlock.PackSave.PerformClick()

                ElseIf WorkBlock.MiscSave.Visible = True Then
                    WorkBlock.MiscSave.PerformClick()

                ElseIf WorkBlock.SortSave.Visible = True Then
                    WorkBlock.SortSave.PerformClick()

                End If

            ElseIf WorkFastBlock.Visible = True Then

                If WorkFastBlock.WPSave.Visible = True Then

                    WorkFastBlock.WPSave.PerformClick()

                End If

            End If
        End If
    End Sub

    Sub LowLevelDisabler()

        If UserRec.QcLevel = "Low" Then
            UserToolStripMenuItem.Enabled = False
            ReportModuleToolStripMenuItem.Enabled = False
            KeluarKerjaToolStripMenuItem.Enabled = False
            ModificationOfEntityToolStripMenuItem.Enabled = False
            SynchFileToolStripMenuItem.Enabled = False
            UploadEmployeeToolStripMenuItem.Enabled = False
            PPH21ReportToolStripMenuItem.Enabled = False
            CouponGenerateToolStripMenuItem.Enabled = False
            SuratDoktorToolStripMenuItem.Enabled = False

        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LoadDB()
        QCStatus.Text = " Welcome: " & User & "          ||            "
        LowLevelDisabler()

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        QCStatus2.Text = " Time: " & TimeOfDay

    End Sub

    Private Sub RegisterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegisterToolStripMenuItem.Click

        Dim NewMDIChild As New UserBlock()
        UserBlock.MdiParent = Me
        UserBlock.Show()

    End Sub

    Private Sub AddEmployeeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim NewMDIChild As New EmployeeBlock()
        EmployeeBlock.MdiParent = Me
        EmployeeBlock.Show()

    End Sub

    Private Sub ViewEmployeeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim NewMDIChild As New Employee2Block()
        Employee2Block.MdiParent = Me
        Employee2Block.Show()

    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click

        Me.Dispose()
        LoginDoor.Dispose()

    End Sub

    Protected Overrides Sub Finalize()

        MyBase.Finalize()

    End Sub

    Private Sub DateHolidaySetupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateHolidaySetupToolStripMenuItem.Click

        Dim NewMDIChild As New DateCtrlBlock()
        DateCtrlBlock.MdiParent = Me
        DateCtrlBlock.Show()

    End Sub

    Private Sub InsertEmployeesWorkToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InsertEmployeesWorkToolStripMenuItem.Click

        Dim NewMDIChild As New WorkBlock()
        WorkBlock.MdiParent = Me
        WorkBlock.Show()
        Me.Refresh()

    End Sub

    Private Sub StandardToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StandardToolStripMenuItem.Click

        Dim NewMDIChild As New StandardBlock()
        StandardBlock.MdiParent = Me
        StandardBlock.Show()
        Me.Refresh()

    End Sub

    Private Sub DateControlSetupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateControlSetupToolStripMenuItem.Click

        Dim NewMDIChild As New DateBlock()
        DateBlock.MdiParent = Me
        DateBlock.Show()
        Me.Refresh()

    End Sub

    Private Sub PeriodeReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PeriodeReportToolStripMenuItem.Click

        Dim NewMDIChild As New ReportBlock()
        ReportBlock.MdiParent = Me
        ReportBlock.Show()
        Me.Refresh()

    End Sub

    Private Sub FieldReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FieldReportToolStripMenuItem.Click

        Me.Invoke(New RepFrm1(AddressOf OpenNew))

    End Sub
    Private Sub OpenNew()

        Dim NewMDIChild As New Report2Block
        Report2Block.MdiParent = Me
        Report2Block.Show()
        Me.Refresh()

    End Sub
    Private Sub IncentiveDateSetUpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IncentiveDateSetUpToolStripMenuItem.Click

        Dim NewMDIChild As New IncentivesBlock
        IncentivesBlock.MdiParent = Me
        IncentivesBlock.Show()
        Me.Refresh()

    End Sub
    Private Sub SalarySynchToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalarySynchToolStripMenuItem.Click

        Dim NewMDIChild As New SynchBlock
        SynchBlock.MdiParent = Me
        SynchBlock.Show()
        Me.Refresh()

    End Sub
    Private Sub EmployeeSynchToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmployeeSynchToolStripMenuItem.Click

        Dim NewMDIChild As New UserSynchBlock
        UserSynchBlock.MdiParent = Me
        UserSynchBlock.Show()
        Me.Refresh()

    End Sub
    Private Sub InsertEmployeesWorkFastToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InsertEmployeesWorkFastToolStripMenuItem.Click

        Dim NewMDIChild As New WorkFastBlock
        WorkFastBlock.MdiParent = Me
        WorkFastBlock.Show()
        Me.Refresh()

    End Sub
    Private Sub KeluarKerjaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeluarKerjaToolStripMenuItem.Click

        Dim NewMDIChild As New KeluarBlock()
        KeluarBlock.MdiParent = Me
        KeluarBlock.Show()
        Me.Refresh()

    End Sub
    Private Sub EmployeeData25ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim NewMDIChild As New EmpDatusBlock()
        EmpDatusBlock.MdiParent = Me
        EmpDatusBlock.Show()
        Me.Refresh()

    End Sub
    Private Sub ViewEmployee20ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewEmployee20ToolStripMenuItem.Click

        Dim NewMDIChild As New EmployeeOldBlock()
        EmployeeOldBlock.MdiParent = Me
        EmployeeOldBlock.Show()
        Me.Refresh()

    End Sub
    Private Sub PPH21ReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PPH21ReportToolStripMenuItem.Click

        Dim NewMDIChild As New PPH21Block()
        PPH21Block.MdiParent = Me
        PPH21Block.Show()
        Me.Refresh()

    End Sub
    Private Sub CouponGenerateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CouponGenerateToolStripMenuItem.Click

        Dim NewMDIChild As New CouponBlock()
        CouponBlock.MdiParent = Me
        CouponBlock.Show()
        Me.Refresh()

    End Sub
    Private Sub SuratDoktorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SuratDoktorToolStripMenuItem.Click

        Dim NewMDIChild As New PermitBlock()
        PermitBlock.MdiParent = Me
        PermitBlock.Show()
        Me.Refresh()

    End Sub
    Private Sub LogOutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogOutToolStripMenuItem.Click

        Select Case MsgBox("Do you want to logout", MsgBoxStyle.YesNo, "Logging Out")
            Case MsgBoxResult.Yes

                Me.Hide()
                LoginDoor.Show()
                LoginDoor.LoginTbx1.Text = ""
                LoginDoor.LoginTbx2.Text = ""

        End Select

    End Sub
    Private Sub MainTotalSynchToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MainTotalSynchToolStripMenuItem.Click

        Dim NewMDIChild As New MainTotBlock()
        MainTotBlock.MdiParent = Me
        MainTotBlock.Show()
        Me.Refresh()

    End Sub
    Private Sub UploadEmployeeToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UploadEmployeeToolStripMenuItem1.Click

        Dim NewMDIChild As New UpEmpBlock()
        UpEmpBlock.MdiParent = Me
        UpEmpBlock.Show()
        Me.Refresh()

    End Sub
    Private Sub UploadCutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UploadCutToolStripMenuItem.Click

        Dim NewMDIChild As New DedBlock()
        DedBlock.MdiParent = Me
        DedBlock.Show()
        Me.Refresh()

    End Sub
    Private Sub UpDatePayNoRekToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpDatePayNoRekToolStripMenuItem.Click

        Dim NewMDIChild As New NoRekBlock()
        NoRekBlock.MdiParent = Me
        NoRekBlock.Show()
        Me.Refresh()

    End Sub
    Private Sub AboutQCCodexToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutQCCodexToolStripMenuItem.Click

        Dim NewMDIChild As New UniAbout()
        UniAbout.MdiParent = Me
        UniAbout.Show()
        Me.Refresh()

    End Sub

    Private Sub EntityRemoverToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EntityRemoverToolStripMenuItem.Click

        Dim NewMDIChild As New RemoverBlock()
        RemoverBlock.MdiParent = Me
        RemoverBlock.Show()
        Me.Refresh()

    End Sub

    Private Sub TrialOpenCrystalReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim NewMDIChild As New KeluarReportView()
        KeluarReportView.MdiParent = Me
        KeluarReportView.Show()
        Me.Refresh()

    End Sub

    Private Sub IncentivesControlToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IncentivesControlToolStripMenuItem.Click

        QCIncentivesGEN.MdiParent = Me
        QCIncentivesGEN.Show()
        Me.Refresh()

    End Sub

    Private Sub Employee30FinalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Employee30FinalToolStripMenuItem.Click

        QCFinalEmpAdd.MdiParent = Me
        QCFinalEmpAdd.Show()
        Me.Refresh()

    End Sub

    Private Sub KeluarNewKerjaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        KeluarNewKerja02.MdiParent = Me
        KeluarNewKerja02.Show()
        Me.Refresh()
    End Sub

    Private Sub ImportExportCSVFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportExportCSVFileToolStripMenuItem.Click
        QCImpExp.MdiParent = Me
        QCImpExp.Show()
    End Sub

    Private Sub ClearCacheToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearCacheToolStripMenuItem.Click
        FlushMemory()
    End Sub

End Class