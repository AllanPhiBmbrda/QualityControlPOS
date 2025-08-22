Public Class WorkFastBlock

    Dim FLocMode As New System.Drawing.Point(173, 21)
    Dim DivSwitcher As String = "OverA"
    Dim ProcessNum As String
    Dim ProcessDig As String
    Dim m1 As String
    Dim y1 As String
    Dim m2 As String
    Dim tr3 As String
    Dim TMasuk As String
    Dim PID As String
    Dim DateYearDate As New System.DateTime
    Dim AstekLook As String
    Dim Incenload1 As String
    Dim Inceload2 As String
    Dim IncentiveCount As String
    Dim IncentiveLock As String
    Dim EnabTot As String
    Dim EnabSave As String
    Dim SortMTot As String
    Dim SalSub As String = ""
    Dim FixGram As String
    Dim SortSub1 As String
    Dim SortTotalPieces As String
    Dim SortMainSpc As String
    Dim ConCD As Double
    Dim MutuIICD As Double
    Dim WalletCD As Double
    Dim PackingCD As Double
    Dim MiscellaneousCD As Double
    Dim NewMiscellaneousCD As Double
    Dim StringCaller As String

    Dim DateGet As Date

    Dim AddSecond As String

    ' Pieces Counter for Each Department
    Dim ConPiecesRes As String
    Dim MutPiecesRes As String
    Dim WalPiecesRes As String
    Dim PackPieceRes As String
 
    Dim ProID As String
    Dim NikCtrl As String
    Dim PcsCtrl As String
    Dim TarCtrl As String
    Dim DateCtrl As String
    Dim TimeCtrl As String
    Dim CartCtrl As String
    Dim SalCtrl As String
    Dim CouCtrl As String
    Dim NoKgCtrl As String
    Dim NoGrCtrl As String
    Dim NoBagCtrl As String

    Dim GridVal0 As String
    Dim GridVal1 As Date
    Dim GridVal2 As String
    Dim GridVal3 As String
    Dim GridVal4 As String
    Dim GridVal5 As String
    Dim GridVal6 As String
    Dim GridPiPack As String
    Dim GridVal7 As String
    Dim GridVal8 As String
    Dim GridVal9 As String
    Dim Gridval10 As String


    'Sub Result per Dept.
    Dim ConHiddenRes As String
    Dim MutuHiddenRes As String
    Dim WalletHiddenRes As String
    Dim PackingHiddenRes As String
    Dim SortasiHiddenRes As String
    Dim MiscHiddenRes As String

    'Main Result per Dept.
    Dim ConMainResult As String
    Dim MutuMainResult As String
    Dim WallMainResult As String
    Dim PackMainResult As String
    Dim SortasiMainResult As String
    Dim MiscMainResult As String

    ' Main Pieces Total
    Dim SortPiecesTotal As String

#Region "Button Function"

    Private Sub PanelBtn4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PanelBtn4.Click
        If WorkSupFrame.Visible = False Then
            WorkSupFrame.Visible = True
        Else
            WorkSupFrame.Visible = False
        End If
    End Sub
    Private Sub WFBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WFBtn1.Click
        DivSwitcher = ""
        DivSwitcher = "Conv"
        IndiLabel.Text = "Conveyour"
        WPTbx5.Enabled = False
        FGenConCode()
        WPTbx3.Text = ProcessNum

        WFGrp1.Visible = True
        WFGrp1.Height = 46
        WFGrp1.Width = 784
        WFGrp1.Location = FLocMode

        WFGrp2.Visible = False
        WFGrp3.Visible = False
        WFGrp4.Visible = False
        WFGrp5.Visible = False
        WFGrp6.Visible = False
        Me.Refresh()

        WFGridHeader()
        ConCmb.Focus()
        WPTbx5.Text = ""

        LoadConveyour()

        'If Not WFBgWorker.IsBusy Then
        '    'WFBgWorker.RunWorkerAsync()
        'End If


    End Sub
    Private Sub WFBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WFBtn2.Click
        DivSwitcher = ""
        DivSwitcher = "Mutu"
        IndiLabel.Text = "Mutu II"
        WPTbx5.Enabled = True

        WFGrp2.Visible = True
        WFGrp2.Height = 46
        WFGrp2.Width = 784
        WFGrp2.Location = FLocMode

        WFGrp1.Visible = False
        WFGrp3.Visible = False
        WFGrp4.Visible = False
        WFGrp5.Visible = False
        WFGrp6.Visible = False

        Me.Refresh()

        FGenMutuIICode()
        WPTbx3.Text = ProcessNum

        WFGridHeader()
        WPTbx5.Text = ""


        LoadMutuII()
        'If Not WFBgWorker.IsBusy Then
        '    WFBgWorker.RunWorkerAsync()
        'End If

    End Sub
    Private Sub WFBtn3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WFBtn3.Click
        DivSwitcher = ""
        DivSwitcher = "Wallet"
        IndiLabel.Text = "Wallet"
        WPTbx5.Enabled = True

        WFGrp3.Visible = True
        WFGrp3.Height = 46
        WFGrp3.Width = 784
        WFGrp3.Location = FLocMode

        WFGrp1.Visible = False
        WFGrp2.Visible = False
        WFGrp4.Visible = False
        WFGrp5.Visible = False
        WFGrp6.Visible = False

        Me.Refresh()

        FGenWalletCode()
        WPTbx3.Text = ProcessNum
        WFGridHeader()
        WPTbx5.Text = ""

        LoadWallet()

        'If Not WFBgWorker.IsBusy Then
        '    WFBgWorker.RunWorkerAsync()
        'End If

    End Sub
    Private Sub WFBtn4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WFBtn4.Click
        DivSwitcher = ""
        DivSwitcher = "Pack"
        IndiLabel.Text = "Packing"
        WPTbx5.Enabled = True


        WFGrp4.Visible = True
        WFGrp4.Height = 46
        WFGrp4.Width = 784
        WFGrp4.Location = FLocMode

        WFGrp1.Visible = False
        WFGrp2.Visible = False
        WFGrp3.Visible = False
        WFGrp5.Visible = False
        WFGrp6.Visible = False

        Me.Refresh()

        FGenPackingCode()
        WPTbx3.Text = ProcessNum

        WFGridHeader()
        WPTbx5.Text = ""


        LoadPacking()

        'If Not WFBgWorker.IsBusy Then
        '    WFBgWorker.RunWorkerAsync()
        'End If
    End Sub
    Private Sub WFBtn5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WFBtn5.Click
        DivSwitcher = ""
        DivSwitcher = "Sort"
        IndiLabel.Text = "Sortasi"
        WPTbx5.Enabled = True

        WFGrp6.Visible = True
        WFGrp6.Height = 46
        WFGrp6.Width = 784
        WFGrp6.Location = FLocMode

        WFGrp1.Visible = False
        WFGrp2.Visible = False
        WFGrp3.Visible = False
        WFGrp4.Visible = False
        WFGrp5.Visible = False

        Me.Refresh()


        FGenNewMiscCode()
        WPTbx3.Text = ProcessNum

        WFGridHeader()
        WPTbx5.Text = ""

        PiecesCounter()
        ErrorSpecial()
        WPTot2.Text = Val(SortMTot).ToString("N0", CustomtoUS)

        LoadSortasi()
 
        'If Not WFBgWorker.IsBusy Then
        '    WFBgWorker.RunWorkerAsync()
        'End If


    End Sub
    Private Sub WFBtn6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WFBtn6.Click
        DivSwitcher = ""
        DivSwitcher = "Misc"
        IndiLabel.Text = "Miscellaneous"
        WPTbx5.Enabled = False

        WFGrp5.Visible = True
        WFGrp5.Height = 46
        WFGrp5.Width = 784
        WFGrp5.Location = FLocMode

        WFGrp1.Visible = False
        WFGrp2.Visible = False
        WFGrp3.Visible = False
        WFGrp4.Visible = False
        WFGrp6.Visible = False

        Me.Refresh()

        FGenMiscCode()
        WPTbx3.Text = ProcessNum

        WFGridHeader()
        WPTbx5.Text = ""

        LoadMisc()
        'If Not WFBgWorker.IsBusy Then
        '    WFBgWorker.RunWorkerAsync()
        'End If

    End Sub
    Private Sub WFBtn7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WFBtn7.Click
        DivSwitcher = ""
        DivSwitcher = "OverA"
        IndiLabel.Text = "Over All"
        WPTbx5.Enabled = False

        WFGrp1.Visible = False
        WFGrp2.Visible = False
        WFGrp3.Visible = False
        WFGrp4.Visible = False
        WFGrp5.Visible = False
        Me.Refresh()

        WPTbx3.Text = ""

        'IncentivesControlLoad()
        WFGridHeader()
        WPTbx5.Text = ""

        LoadAllSalary()

        'If Not WFBgWorker.IsBusy Then
        '    WFBgWorker.RunWorkerAsync()
        'End If

    End Sub

    Sub EmptyBox()

        WPTot1.Text = ""
        WPTot2.Text = ""

    End Sub

#End Region

#Region "Generate Number"

    Sub FGenConCode()

        SQL = ""
        SQL = SQL & "Select `Process_ID` From 03_Conveyour_Table "
        SQL = SQL & "Order by Process_ID Desc "
        OpenTbl(ADb, DbTbl6, SQL)
        If DbTbl6.RecordCount <> 0 Then
            ProcessDig = DbTbl6("Process_ID").Value
            ProcessNum = Format(ProcessDig + 1, "0000000000")
        Else
            ProcessNum = "0000000001"
        End If

    End Sub
    Sub FGenMutuIICode()

        SQL = ""
        SQL = SQL & "Select `Process_ID` From 04_MutuII_Table "
        SQL = SQL & "Order by Process_ID Desc "
        OpenTbl(ADb, DbTbl6, SQL)
        If DbTbl6.RecordCount <> 0 Then
            ProcessDig = DbTbl6("Process_ID").Value
            ProcessNum = Format(ProcessDig + 1, "0000000000")
        Else
            ProcessNum = "0000000001"
        End If

    End Sub
    Sub FGenPackingCode()

        SQL = ""
        SQL = SQL & "Select `Process_ID` From 05_Packing_Table "
        SQL = SQL & "Order by Process_ID Desc "
        OpenTbl(ADb, DbTbl6, SQL)
        If DbTbl6.RecordCount <> 0 Then
            ProcessDig = DbTbl6("Process_ID").Value
            ProcessNum = Format(ProcessDig + 1, "0000000000")
        Else
            ProcessNum = "0000000001"
        End If

    End Sub
    Sub FGenWalletCode()

        SQL = ""
        SQL = SQL & "Select `Process_ID` From 06_Wallet_Table "
        SQL = SQL & "Order by Process_ID Desc "
        OpenTbl(ADb, DbTbl6, SQL)
        If DbTbl6.RecordCount <> 0 Then
            ProcessDig = DbTbl6("Process_ID").Value
            ProcessNum = Format(ProcessDig + 1, "0000000000")
        Else
            ProcessNum = "0000000001"
        End If

    End Sub
    Sub FGenMiscCode()

        SQL = ""
        SQL = SQL & "Select `Process_ID` From 19_Miscellaneous_Table "
        SQL = SQL & "Order by Process_ID Desc "
        OpenTbl(ADb, DbTbl6, SQL)
        If DbTbl6.RecordCount <> 0 Then
            ProcessDig = DbTbl6("Process_ID").Value
            ProcessNum = Format(ProcessDig + 1, "0000000000")
        Else
            ProcessNum = "0000000001"
        End If

    End Sub
    Sub FGenNewMiscCode()

        SQL = ""
        SQL = SQL & "Select `Process_ID` From 21_NewMiscellaneous_Table "
        SQL = SQL & "Order by Process_ID Desc "
        OpenTbl(ADb, DbTbl6, SQL)
        If DbTbl6.RecordCount <> 0 Then
            ProcessDig = DbTbl6("Process_ID").Value
            ProcessNum = Format(ProcessDig + 1, "0000000000")
        Else
            ProcessNum = "0000000001"
        End If
    End Sub

#End Region

#Region "Loading Data"

    Sub FEmpLookup()
        SQL = ""
        SQL = SQL & "Select `Nik`, `Name`, `Pay` from 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "Order by Nik"
        OpenTbl(ADb, Atb3, SQL)

        If Atb3.RecordCount > 0 Then
            StringCaller = Atb3("Name").Value
            StringCaller = StringCaller.Replace("?", "'")
            WPTbx2.Text = StringCaller
            PayAsSetup = Atb3("Pay").Value
            WorkSupTbx4.Text = PayAsSetup

        Else
            MsgBox("Employee Not Found", MsgBoxStyle.Information, "Codex ~ QC Build " & BuildCounter & " Warning!!")

        End If


    End Sub

    Sub FMasaKerjaCtrl()

        tr3 = WorkSupTbx1.Text

        m1 = CInt(tr3) / 30

        If m1 <= 0 Then m1 = 0

        y1 = Int(CDbl(m1) / 12)

        m2 = Format((m1 - CDbl(y1) * 12), "#")

        WPTbx4.Text = (y1) + " Tahun " + (m2) + " Bulan"
    End Sub

    Sub YearHolMod()

        SQL = ""
        SQL = SQL & "Select `Nik`, `DateStart`, `Jamsostek` from 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        OpenTbl(ADb, Dbtb33, SQL)

        If Dbtb33.RecordCount > 0 Then
            YearDate = Dbtb33("DateStart").Value
            AstekLook = Dbtb33("Jamsostek").Value
            DateYearDate = WFCal.Text

        End If
        YearMod = DateYearDate.Subtract(YearDate).Days
        WorkSupTbx2.Text = YearDate
        WorkSupTbx3.Text = AstekLook
        WorkSupTbx1.Text = YearMod

    End Sub

#End Region

    Private Sub WPTbx2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles WPTbx2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub
    Private Sub WPTbx3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles WPTbx3.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub
    Private Sub WPTbx4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles WPTbx4.KeyPress

        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True

    End Sub


    Private Sub WPTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles WPTbx1.KeyPress

        If Not InValid4.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If

        If e.KeyChar.ToString = ChrW(Keys.Enter) Then

            FEmpLookup()
            WPTbx1.Select()
            ErrorPeriode()
            ErrorHoliday()
            'IncentivesControlLoad()
            e.Handled = True

        End If
    End Sub
    Sub TextToGrid()

        If DivSwitcher = "Conv" Then
            WFGrid01.Rows.Add()

        ElseIf DivSwitcher = "Mutu" Then
            WFGrid01.Rows.Add()

        ElseIf DivSwitcher = "Wallet" Then
            WFGrid01.Rows.Add()

        ElseIf DivSwitcher = "Pack" Then
            WFGrid01.Rows.Add()

        ElseIf DivSwitcher = "Sort" Then
            WFGrid01.Rows.Add()

        ElseIf DivSwitcher = "Misc" Then
            WFGrid01.Rows.Add()

        ElseIf DivSwitcher = "OverA" Then
            WFGrid01.Rows.Add()

        End If
        WFGrid01.Rows.Add()

    End Sub
    Sub WFGridHeader()

        WFGrid01.Rows.Clear()
        WFGrid01.Columns.Clear()

        With WFGrid01

            If DivSwitcher = "Conv" Then

                .Columns.Add("Col1", "ID")
                .Columns.Add("Col2", "Date")
                .Columns.Add("Col3", "Time")
                .Columns.Add("Col4", "Nik")
                .Columns.Add("Col5", "Pieces")
                .Columns.Add("Col6", "Target")
                .Columns.Add("Col7", "Salary")
                .Columns.Add("Col8", "Status")


            ElseIf DivSwitcher = "Mutu" Then

                .Columns.Add("Col1", "ID")
                .Columns.Add("Col2", "Date")
                .Columns.Add("Col3", "Time")
                .Columns.Add("Col4", "Nik")
                .Columns.Add("Col5", "Target")
                .Columns.Add("Col6", "Pieces")
                .Columns.Add("Col7", "Salary")
                .Columns.Add("Col8", "Coupon")
                .Columns.Add("Col9", "Status")



            ElseIf DivSwitcher = "Wallet" Then

                .Columns.Add("Col1", "ID")
                .Columns.Add("Col2", "Date")
                .Columns.Add("Col3", "Time")
                .Columns.Add("Col4", "Nik")
                .Columns.Add("Col5", "Target")
                .Columns.Add("Col6", "Pieces")
                .Columns.Add("Col7", "Salary")
                .Columns.Add("Col8", "Coupon")
                .Columns.Add("Col9", "Status")


            ElseIf DivSwitcher = "Pack" Then

                .Columns.Add("Col1", "ID")
                .Columns.Add("Col2", "Date")
                .Columns.Add("Col3", "Time")
                .Columns.Add("Col4", "Nik")
                .Columns.Add("Col5", "Target")
                .Columns.Add("Col6", "Carton")
                .Columns.Add("Col7", "Salary")
                .Columns.Add("Col8", "Coupon")
                .Columns.Add("Col9", "Status")

            ElseIf DivSwitcher = "Sort" Then

                .Columns.Add("Col1", "ID")
                .Columns.Add("Col2", "Date")
                .Columns.Add("Col3", "Time")
                .Columns.Add("Col4", "Nik")
                .Columns.Add("Col5", "No. of Kilogram")
                .Columns.Add("Col6", "No. of Bag")
                .Columns.Add("Col7", "No. of Gram")
                .Columns.Add("Col8", "Pieces")
                .Columns.Add("Col9", "Salary")
                .Columns.Add("Col10", "Coupon")
                .Columns.Add("Col11", "Status")
                .Columns(4).Width = 120


            ElseIf DivSwitcher = "Misc" Then

                .Columns.Add("Col1", "ID")
                .Columns.Add("Col2", "Date")
                .Columns.Add("Col3", "Time")
                .Columns.Add("Col4", "Nik")
                .Columns.Add("Col5", "Salary")
                .Columns.Add("Col6", "Status")

            ElseIf DivSwitcher = "OverA" Then
                .Columns.Add("Col1", "Date")
                .Columns.Add("Col2", "Conveyour")
                .Columns.Add("Col3", "MutuII")
                .Columns.Add("Col4", "Wallet")
                .Columns.Add("Col5", "Packing")
                .Columns.Add("Col6", "Sortasi")
                .Columns.Add("Col7", "Miscellaneous")
                .Columns.Add("Col8", "Total")
                .Columns.Add("Col9", "Status")

            End If

        End With
    End Sub
    Private Sub QCWorkFastBlock_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        MainMenu.Refresh()
    End Sub
    Private Sub QCWorkFastBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LoadDB()
        LoadDB2()
        LoadDB3()
        LoadHolidayMod()

        DirectSaveEnabler()
        DirectTotEnabler()
        WPTbx1.CharacterCasing = CharacterCasing.Upper
    End Sub

    Sub SpecialSortasi() ' Special Calculation in Sortasi Function

        If tr3 >= 46 And SortPiecesTotal >= 20001 Then
            SortSpecialTot = Val(((SortPiecesTotal - 20000) / 20000) * (1.2 * StandardsSalary) + StandardsSalary)
            SortMTot = Format(SortSpecialTot, "###.00")

        ElseIf tr3 >= 46 And SortPiecesTotal <= 20000 Then
            SortSpecialTot = Format((Val(SortPiecesTotal / 20000)) * (StandardsSalary), "###.00")
            SortMTot = Format(SortSpecialTot, "###.00")

            ' For Anak Baru

        ElseIf tr3 <= 45 And SortPiecesTotal <= 15000 Then
            SortMTot = SubsidiSalary

        ElseIf tr3 <= 45 And SortPiecesTotal >= 20001 Then

            SortSpecialTot = Val(((SortPiecesTotal - 20000) / 20000) * (1.2 * SubsidiSalary) + SubsidiSalary)
            SortMTot = Format(SortSpecialTot, "###.00")

        ElseIf tr3 <= 45 And SortPiecesTotal <= 20000 Then
            SortSpecialTot = Format((Val(SortPiecesTotal / 20000)) * (SubsidiSalary), "###.00")
            SortMTot = Format(SortSpecialTot, "###.00")

        Else

            SortMTot = Format(SortasiMainResult, "###.00")

        End If
    End Sub


    Sub SubsidyWork()

        If tr3 <= 60 Then

            SortMTot = SubsidiSalary

        End If

    End Sub

#Region "Incentives"


    Sub DirectTotEnabler()
        SQL = ""
        SQL = SQL & "Select * From 08_Standard_Table "
        SQL = SQL & "Where Original = ('" & "LockDirectTot" & "') "
        OpenTbl(ADb, Atbl41, SQL)

        If Atbl41.RecordCount > 0 Then

            EnabTot = Atbl41("Standard_Wage").Value

        End If
        Me.Refresh()


        If EnabTot = "True" Then
            DirChk1.Checked = True

        End If
    End Sub
    Sub DirectSaveEnabler()
        SQL = ""
        SQL = SQL & "Select * From 08_Standard_Table "
        SQL = SQL & "Where Original = ('" & "LockDirectSave" & "') "
        OpenTbl(ADb, Atbl42, SQL)

        If Atbl42.RecordCount > 0 Then

            EnabSave = Atbl42("Standard_Wage").Value

        End If
        Me.Refresh()


        If EnabSave = "True" Then
            DirChk2.Checked = True

        End If
    End Sub

#End Region

    Sub LoadHolidayMod()

        SQL = ""
        SQL = SQL & "Select * from 17_Holiday_Table "
        SQL = SQL & "Where Date = ('" & WFCal.Text & "') "
        OpenTbl(ADb, Dbtb29, SQL)

        If Dbtb29.RecordCount > 0 Then
            HolMod = Dbtb29("Salary_Mod").Value
            WorkSupTbx5.Text = HolMod
        Else
            HolMod = "1"
            WorkSupTbx5.Text = HolMod

        End If
    End Sub
    Sub AutoNum()

        WPTbx3.Text = Format(Val(WPTbx3.Text) + 1, "0000000000")

    End Sub
    Sub LoadDayPeriodeCtrl()
        Dim DateSet As Date = WFCal.Text
        SQL = ""
        SQL = SQL & "Select * from Periode_CounterTable "
        SQL = SQL & "Where Date = ('" & DateSet.ToString("yyyy-MM-dd") & "') "
        OpenTbl(CBb, Ctbl42, SQL)

        If Ctbl42.RecordCount <> 0 Then


            PeriodeMonthCtrl = IIf(IsDBNull(Ctbl42("PeriodeRange").Value), "", Ctbl42("PeriodeRange").Value)
            PeriodeCtrl = IIf(IsDBNull(Ctbl42("Periode").Value), "", Ctbl42("Periode").Value)
            PeriodeDayCtrl = IIf(IsDBNull(Ctbl42("Counter").Value), "", Ctbl42("Counter").Value)
        Else
            PeriodeMonthCtrl = ""
            PeriodeCtrl = ""
            PeriodeDayCtrl = ""

        End If

        WorkSupTbx8.Text = PeriodeDayCtrl
        WorkSupTbx6.Text = PeriodeCtrl
        WorkSupTbx7.Text = PeriodeMonthCtrl

    End Sub
    Sub ErrorPeriode()
        On Error GoTo Err
        LoadDayPeriodeCtrl()
        Exit Sub
Err:
    End Sub
    Sub ErrorHoliday()
        On Error GoTo ErrHol
        YearHolMod()
        Exit Sub
ErrHol:
    End Sub
    Private Sub WFCal_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WFCal.ValueChanged
        ErrorPeriode()
        ErrorHoliday()
        LoadHolidayMod()
    End Sub

    Private Sub WorkSupTbx1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WorkSupTbx1.TextChanged
        FMasaKerjaCtrl()
    End Sub

    ' Conveyour GUI
    Private Sub ConMskTb2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ConMskTb2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InValid2.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If

        If e.KeyChar.ToString = ChrW(Keys.Enter) Then

            If ConMskTb2.Text = "" Or ConCmb.Text = "" Then
                MsgBox("Please Insert the Required Data")

            ElseIf WPTbx2.Text = "" Then
                MsgBox("Please Look for Personnel/ Orang Karyawan")

            ElseIf WPTbx3.Text = "" Then
                MsgBox("Kindly Select Your Desire Dept/Jabatan")

            Else
                If HolMod = 2 And YearMod >= 365 Then
                    ConHiddenRes = Format((Val(ConMskTb2.Text) / Val(ConCmb.Text) * StandardsSalary) * 2, "0.")
                    WFGrid01.Rows.Add(WPTbx3.Text, WFCal.Text, TimeOfDay.ToString("T"), WPTbx1.Text, ConMskTb2.Text, ConCmb.Text, ConHiddenRes)
                    AutoNum()
                    ProcessMode()
                Else

                    If tr3 >= 46 Then
                        ConHiddenRes = Format((Val(ConMskTb2.Text) / Val(ConCmb.Text) * StandardsSalary), "0.")
                        WFGrid01.Rows.Add(WPTbx3.Text, WFCal.Text, TimeOfDay.ToString("T"), WPTbx1.Text, ConMskTb2.Text, ConCmb.Text, ConHiddenRes)
                        AutoNum()
                        ProcessMode()

                    ElseIf tr3 <= 45 Then

                        ConHiddenRes = Format((Val(ConMskTb2.Text) / Val(ConCmb.Text) * SubsidiSalary), "0.")
                        WFGrid01.Rows.Add(WPTbx3.Text, WFCal.Text, TimeOfDay.ToString("T"), WPTbx1.Text, ConMskTb2.Text, ConCmb.Text, ConHiddenRes)
                        AutoNum()
                        ProcessMode()

                    End If
                End If
            End If
        End If

    End Sub
    Private Sub ConCmb_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ConCmb.KeyPress
        If e.KeyChar.ToString = Chr(Keys.Back) Then Exit Sub
        If Not InValid2.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If

        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            ConMskTb2.Focus()
        End If

    End Sub
    Private Sub ConCmb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConCmb.SelectedIndexChanged
        ConMskTb2.Focus()
    End Sub

    ' Packing GUI
    Private Sub PackCmb_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PackCmb.KeyPress
        If Not InValid2.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If

        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            PackTbx1.Focus()
        End If
    End Sub
    Private Sub PackCmb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PackCmb.SelectedIndexChanged
        PackTbx1.Focus()
    End Sub
    Private Sub PackTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PackTbx1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InValid2.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If

        If e.KeyChar.ToString = ChrW(Keys.Enter) Then

            If PackCmb.Text = "" Or PackTbx1.Text = "" Then
                MsgBox("Please Insert the Required Data")

            ElseIf WPTbx2.Text = "" Then
                MsgBox("Please Look for Personnel/ Orang Karyawan")

            ElseIf WPTbx3.Text = "" Then
                MsgBox("Kindly Select Your Desire Dept/Jabatan")

            ElseIf WPTbx5.Text = "" Then
                MsgBox("Please Input the Coupon Number")

            Else
                If HolMod = 2 And YearMod >= 365 Then
                    PackingHiddenRes = Format((Val(PackTbx1.Text) / Val(PackCmb.Text) * StandardsSalary) * 2, "0.")
                    WFGrid01.Rows.Add(WPTbx3.Text, WFCal.Text, TimeOfDay.ToString("T"), WPTbx1.Text, PackCmb.Text, PackTbx1.Text, PackingHiddenRes, WPTbx5.Text)
                    AutoNum()
                    ProcessMode()

                Else
                    PackingHiddenRes = Format((Val(PackTbx1.Text) / Val(PackCmb.Text) * StandardsSalary), "0.")
                    WFGrid01.Rows.Add(WPTbx3.Text, WFCal.Text, TimeOfDay.ToString("T"), WPTbx1.Text, PackCmb.Text, PackTbx1.Text, PackingHiddenRes, WPTbx5.Text)
                    AutoNum()
                    ProcessMode()
                End If
            End If
        End If
    End Sub

    ' Mutu II GUI
    Private Sub MutuTbx2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MutuTbx2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InValid2.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If


        If e.KeyChar.ToString = ChrW(Keys.Enter) Then

            If MutuTbx1.Text = "" Or MutuTbx2.Text = "" Then
                MsgBox("Please Insert the Required Data")

            ElseIf WPTbx2.Text = "" Then
                MsgBox("Please Look for Personnel/ Orang Karyawan")

            ElseIf WPTbx3.Text = "" Then
                MsgBox("Kindly Select Your Desire Dept/Jabatan")

            ElseIf WPTbx5.Text = "" Then
                MsgBox("Please Input the Coupon Number")

            Else
                If HolMod = 2 And YearMod >= 365 Then
                    MutuHiddenRes = Format((Val(MutuTbx2.Text) / Val(MutuTbx1.Text) * StandardsSalary) * 2, "0.")
                    WFGrid01.Rows.Add(WPTbx3.Text, WFCal.Text, TimeOfDay.ToString("T"), WPTbx1.Text, MutuTbx1.Text, MutuTbx2.Text, MutuHiddenRes, WPTbx5.Text)
                    AutoNum()
                    ProcessMode()
                Else
                    MutuHiddenRes = Format((Val(MutuTbx2.Text) / Val(MutuTbx1.Text) * StandardsSalary), "0.")
                    WFGrid01.Rows.Add(WPTbx3.Text, WFCal.Text, TimeOfDay.ToString("T"), WPTbx1.Text, MutuTbx1.Text, MutuTbx2.Text, MutuHiddenRes, WPTbx5.Text)
                    AutoNum()
                    ProcessMode()

                End If
            End If
        End If
    End Sub

    Private Sub MutuTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MutuTbx1.KeyPress
        If Not InValid2.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If

        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            MutuTbx2.Focus()
        End If
    End Sub

    ' Wallet GUI
    Private Sub WallTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles WallTbx1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InValid2.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If


        If e.KeyChar.ToString = ChrW(Keys.Enter) Then

            If WallCmb.Text = "" Or WallTbx1.Text = "" Then
                MsgBox("Please Insert the Required Data")

            ElseIf WPTbx2.Text = "" Then
                MsgBox("Please Look for Personnel/ Orang Karyawan")

            ElseIf WPTbx3.Text = "" Then
                MsgBox("Kindly Select Your Desire Dept/Jabatan")

            ElseIf WPTbx5.Text = "" Then
                MsgBox("Please Input the Coupon Number")

            Else
                If HolMod = 2 And YearMod >= 365 Then
                    WalletHiddenRes = Format((Val(WallTbx1.Text) / Val(WallCmb.Text) * StandardsSalary) * 2, "0.")
                    WFGrid01.Rows.Add(WPTbx3.Text, WFCal.Text, TimeOfDay.ToString("T"), WPTbx1.Text, WallCmb.Text, WallTbx1.Text, WalletHiddenRes, WPTbx5.Text)
                    ProcessMode()
                    AutoNum()
                Else
                    WalletHiddenRes = Format((Val(WallTbx1.Text) / Val(WallCmb.Text) * StandardsSalary), "0.")
                    WFGrid01.Rows.Add(WPTbx3.Text, WFCal.Text, TimeOfDay.ToString("T"), WPTbx1.Text, WallCmb.Text, WallTbx1.Text, WalletHiddenRes, WPTbx5.Text)
                    ProcessMode()
                    AutoNum()
                End If
            End If
        End If
    End Sub
    Private Sub WallCmb_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles WallCmb.KeyPress
        If Not InValid2.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If

        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            WallTbx1.Focus()
        End If
    End Sub
    Private Sub WallCmb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WallCmb.SelectedIndexChanged
        WallTbx1.Focus()
    End Sub

    ' Sortasi GUI
    Private Sub SortMasTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles SortMasTbx1.KeyPress
        SortMasTbx1.Mask = "#,#"
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            e.Handled = True

            If SortCmb1.Text = "" Or SortCmb2.Text = "" Or SortMasTbx1.Text = "" Then
                MsgBox("Please Insert the Required Data")

            ElseIf WPTbx2.Text = "" Then
                MsgBox("Please Look for Personnel/ Orang Karyawan")

            ElseIf WPTbx3.Text = "" Then
                MsgBox("Kindly Select Your Desire Dept/Jabatan")

            ElseIf WPTbx5.Text = "" Then
                MsgBox("Please Input the Coupon Number")

            Else

                If HolMod = 2 And YearMod >= 365 Then

                    FixGram = Val(SortMasTbx1.Text.Replace(".", ",") / 10000)
                    SortTotalPieces = Format((Val(SortCmb1.Text) / Val(FixGram)) * Val(SortCmb2.Text), "###.00")
                    SortPiTbx.Text = SortTotalPieces
                    SortasiHiddenRes = Format((Val(SortPiTbx.Text / 20000)) * (StandardsSalary), "###.00")
                    WFGrid01.Rows.Add(WPTbx3.Text, WFCal.Text, TimeOfDay.ToString("T"), WPTbx1.Text, SortCmb1.Text, SortCmb2.Text, SortMasTbx1.Text, SortTotalPieces, SortasiHiddenRes, WPTbx5.Text)
                    AutoNum()
                    ProcessMode()
                    PiecesCounter()
                    SpecialSortasi()
                    WPTot2.Text = (Val(SortMTot) * 2).ToString("N0", CustomtoUS)

                Else

                    FixGram = Val(SortMasTbx1.Text.Replace(".", ",") / 10000)
                    SortTotalPieces = Format((Val(SortCmb1.Text) / Val(FixGram)) * Val(SortCmb2.Text), "###.00")
                    SortPiTbx.Text = SortTotalPieces
                    SortasiHiddenRes = Format((Val(SortPiTbx.Text / 20000)) * (StandardsSalary), "###.00")
                    WFGrid01.Rows.Add(WPTbx3.Text, WFCal.Text, TimeOfDay.ToString("T"), WPTbx1.Text, SortCmb1.Text, SortCmb2.Text, SortMasTbx1.Text, SortTotalPieces, SortasiHiddenRes, WPTbx5.Text)
                    AutoNum()
                    ProcessMode()
                    PiecesCounter()
                    SpecialSortasi()
                    WPTot2.Text = Val(SortMTot).ToString("N0", CustomtoUS)

                End If
            End If
        End If

    End Sub
    Private Sub SortPiTbx_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles SortPiTbx.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub
    Private Sub SortCmb1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles SortCmb1.KeyPress
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            SortCmb2.Focus()
            e.Handled = True
        End If

    End Sub
    Private Sub SortCmb1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SortCmb1.SelectedIndexChanged
        SortCmb2.Focus()
    End Sub

    ' Misc GUI
    Private Sub MiscTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MiscTbx1.KeyPress

        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InValid2.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If

        If e.KeyChar.ToString = ChrW(Keys.Enter) Then

            If MiscTbx1.Text = "" Then
                MsgBox("Please Insert the Required Data")

            ElseIf WPTbx2.Text = "" Then
                MsgBox("Please Look for Personnel/ Orang Karyawan")

            ElseIf WPTbx3.Text = "" Then
                MsgBox("Kindly Select Your Desire Dept/Jabatan")

            Else
                If HolMod = 2 And YearMod >= 365 Then
                    MiscHiddenRes = Format((Val(MiscTbx1.Text)) * 2, "0.")
                    WFGrid01.Rows.Add(WPTbx3.Text, WFCal.Text, TimeOfDay.ToString("T"), WPTbx1.Text, MiscHiddenRes)
                    AutoNum()
                    ProcessMode()
                Else
                    MiscHiddenRes = Format((Val(MiscTbx1.Text)), "0.")
                    WFGrid01.Rows.Add(WPTbx3.Text, WFCal.Text, TimeOfDay.ToString("T"), WPTbx1.Text, MiscHiddenRes)
                    AutoNum()
                    ProcessMode()
                End If
            End If
        End If
    End Sub

#Region "CheckBox Default"
    'Private Sub InceChkBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    If InceChkBox1.Checked = True Then

    '        SQL = ""
    '        SQL = SQL & "Select * From 08_Standard_Table "
    '        SQL = SQL & "Where Original = ('" & "IncentiveLock" & "') "
    '        OpenTbl(ADb, Atbl36, SQL)
    '        If Not Atbl36.RecordCount <> 0 Then
    '            Atbl36.AddNew()
    '        End If

    '        Atbl36("Original").Value = "IncentiveLock"
    '        Atbl36("Standard_Wage").Value = "True"


    '        Atbl36.Update()

    '    ElseIf InceChkBox1.Checked = False Then

    '        SQL = ""
    '        SQL = SQL & "Select * From 08_Standard_Table "
    '        SQL = SQL & "Where Original = ('" & "IncentiveLock" & "') "
    '        OpenTbl(ADb, Atbl36, SQL)

    '        If Not Atbl36.RecordCount <> 0 Then
    '            Atbl36.AddNew()
    '        End If

    '        Atbl36("Original").Value = "IncentiveLock"
    '        Atbl36("Standard_Wage").Value = "False"

    '        Atbl36.Update()

    '    End If
    'End Sub

    Private Sub DirChk1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DirChk1.CheckedChanged

        If DirChk1.Checked = True Then

            SQL = ""
            SQL = SQL & "Select * From 08_Standard_Table "
            SQL = SQL & "Where Original = ('" & "LockDirectTot" & "') "
            OpenTbl(ADb, Atbl36, SQL)
            If Not Atbl36.RecordCount <> 0 Then
                Atbl36.AddNew()
            End If

            Atbl36("Original").Value = "LockDirectTot"
            Atbl36("Standard_Wage").Value = "True"


            Atbl36.Update()

        ElseIf DirChk1.Checked = False Then

            SQL = ""
            SQL = SQL & "Select * From 08_Standard_Table "
            SQL = SQL & "Where Original = ('" & "LockDirectTot" & "') "
            OpenTbl(ADb, Atbl36, SQL)

            If Not Atbl36.RecordCount <> 0 Then
                Atbl36.AddNew()
            End If

            Atbl36("Original").Value = "LockDirectTot"
            Atbl36("Standard_Wage").Value = "False"

            Atbl36.Update()

        End If
    End Sub
    Private Sub DirChk2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DirChk2.CheckedChanged

        If DirChk2.Checked = True Then

            SQL = ""
            SQL = SQL & "Select * From 08_Standard_Table "
            SQL = SQL & "Where Original = ('" & "LockDirectSave" & "') "
            OpenTbl(ADb, Atbl36, SQL)
            If Not Atbl36.RecordCount <> 0 Then
                Atbl36.AddNew()
            End If

            Atbl36("Original").Value = "LockDirectSave"
            Atbl36("Standard_Wage").Value = "True"


            Atbl36.Update()

        ElseIf DirChk2.Checked = False Then

            SQL = ""
            SQL = SQL & "Select * From 08_Standard_Table "
            SQL = SQL & "Where Original = ('" & "LockDirectSave" & "') "
            OpenTbl(ADb, Atbl36, SQL)

            If Not Atbl36.RecordCount <> 0 Then
                Atbl36.AddNew()
            End If

            Atbl36("Original").Value = "LockDirectSave"
            Atbl36("Standard_Wage").Value = "False"

            Atbl36.Update()

        End If
    End Sub

#End Region


#Region "Calculation"

    Sub ReCalculation()

        If IndiLabel.Text = "Conveyour" Then

            If tr3 >= 46 Then
                WPTot2.Text = Val(ConMainResult).ToString("N0", CustomtoUS)

            ElseIf tr3 <= 45 Then

                If Val(ConCmb.Text) <= Val(ConPiecesRes) Then
                    WPTot2.Text = Val(ConMainResult).ToString("N0", CustomtoUS)

                ElseIf Val(ConCmb.Text) > Val(ConPiecesRes) Then
                    WPTot2.Text = Val(SubsidiSalary).ToString("N0", CustomtoUS)

                End If

            End If

        ElseIf IndiLabel.Text = "Mutu II" Then

            If tr3 >= 46 Then

                WPTot2.Text = Val(MutuMainResult).ToString("N0", CustomtoUS)

            ElseIf tr3 <= 45 Then

                If Val(MutuTbx1.Text) <= Val(MutPiecesRes) Then
                    WPTot2.Text = Val(MutuMainResult).ToString("N0", CustomtoUS)

                ElseIf Val(MutuTbx1.Text) > Val(MutPiecesRes) Then
                    WPTot2.Text = Val(SubsidiSalary).ToString("N0", CustomtoUS)

                End If

            End If

        ElseIf IndiLabel.Text = "Packing" Then

            If tr3 >= 46 Then
                WPTot2.Text = Val(PackMainResult).ToString("N0", CustomtoUS)

            ElseIf tr3 <= 45 Then

                If Val(PackCmb.Text) <= Val(PackPieceRes) Then
                    WPTot2.Text = Val(PackMainResult).ToString("N0", CustomtoUS)

                ElseIf Val(PackCmb.Text) > Val(PackPieceRes) Then
                    WPTot2.Text = Val(SubsidiSalary).ToString("N0", CustomtoUS)

                End If

            End If

        ElseIf IndiLabel.Text = "Wallet" Then

            If tr3 >= 46 Then
                WPTot2.Text = Val(WallMainResult).ToString("N0", CustomtoUS)

            ElseIf tr3 <= 45 Then

                If Val(WallCmb.Text) <= Val(WalPiecesRes) Then
                    WPTot2.Text = Val(WallMainResult).ToString("N0", CustomtoUS)

                ElseIf Val(WallCmb.Text) > Val(WalPiecesRes) Then
                    WPTot2.Text = Val(SubsidiSalary).ToString("N0", CustomtoUS)
                End If

            End If

        End If

    End Sub

    Sub ProcessMode()
        If DirChk1.Checked = True And IndiLabel.Text = "Conveyour" Then
            ConMainResult = ""
            ConPiecesRes = Nothing
            For i = 0 To WFGrid01.Rows.Count - 1

                GridVal6 = WFGrid01(6, i).Value
                ConMainResult = Val(ConMainResult) + Val(GridVal6)
                ConPiecesRes = Val(ConPiecesRes) + Val(WFGrid01(4, i).Value)


                'If tr3 >= 61 Then
                '    WPTot2.Text = Format(Va l(ConMainResult), "N0")

                'ElseIf tr3 <= 60 And Val(ConMainResult) > Val(SubsidiSalary) Then
                '    WPTot2.Text = Format(Val(ConMainResult), "N0")

                'ElseIf tr3 <= 60 And Val(ConMainResult) < Val(SubsidiSalary) Then
                '    WPTot2.Text = Format(Val(SubsidiSalary), "N0")

                'End If

            Next
            ReCalculation()

        ElseIf DirChk1.Checked = True And IndiLabel.Text = "Mutu II" Then

            MutuMainResult = ""
            MutPiecesRes = Nothing

            For i = 0 To WFGrid01.Rows.Count - 1

                GridVal6 = WFGrid01(6, i).Value
                MutuMainResult = Val(MutuMainResult) + Val(GridVal6)
                MutPiecesRes = Val(MutPiecesRes) + Val(WFGrid01(5, i).Value)


                'If tr3 >= 61 Then
                '    WPTot2.Text = Format(Val(MutuMainResult), "N0")

                'ElseIf tr3 <= 60 And Val(MutuMainResult) > Val(SubsidiSalary) Then
                '    WPTot2.Text = Format(Val(MutuMainResult), "N0")

                'ElseIf tr3 <= 60 And Val(MutuMainResult) < Val(SubsidiSalary) Then
                '    WPTot2.Text = Format(Val(SubsidiSalary), "N0")

                'End If
            Next
            ReCalculation()

        ElseIf DirChk1.Checked = True And IndiLabel.Text = "Wallet" Then
            WallMainResult = ""
            WalPiecesRes = Nothing

            For i = 0 To WFGrid01.Rows.Count - 1

                GridVal6 = WFGrid01(6, i).Value
                WallMainResult = Val(WallMainResult) + Val(GridVal6)
                WalPiecesRes = Val(WalPiecesRes) + Val(WFGrid01(5, i).Value)

                'If tr3 >= 61 Then
                '    WPTot2.Text = Format(Val(WallMainResult), "N0")

                'ElseIf tr3 <= 60 And Val(WallMainResult) > Val(SubsidiSalary) Then
                '    WPTot2.Text = Format(Val(WallMainResult), "N0")

                'ElseIf tr3 <= 60 And Val(WallMainResult) < Val(SubsidiSalary) Then
                '    WPTot2.Text = Format(Val(SubsidiSalary), "N0")

                'End If

            Next
            ReCalculation()


        ElseIf DirChk1.Checked = True And IndiLabel.Text = "Packing" Then
            PackMainResult = ""
            PackPieceRes = Nothing

            For i = 0 To WFGrid01.Rows.Count - 1

                GridVal6 = WFGrid01(6, i).Value
                PackMainResult = Val(PackMainResult) + Val(GridVal6)
                PackPieceRes = Val(PackPieceRes) + Val(WFGrid01(5, i).Value)

            Next
            ReCalculation()

            'PackPieceRes = Val(PackPieceRes) + Val(GridPiPack)

            'If tr3 >= 61 Then
            '    WPTot2.Text = Format(Val(PackMainResult), "N0")

            'ElseIf tr3 <= 60 And Val(PackCmb.Text) <= Val(PackPieceRes) Then
            '    WPTot2.Text = Format(Val(PackMainResult), "N0")

            'ElseIf tr3 <= 60 And Val(PackCmb.Text) > Val(PackPieceRes) Then
            '    WPTot2.Text = Format(Val(SubsidiSalary), "N0")

            'End If


        ElseIf DirChk1.Checked = True And IndiLabel.Text = "Sortasi" Then
            SortasiMainResult = ""

            For i = 0 To WFGrid01.Rows.Count - 1

                GridVal8 = WFGrid01(8, i).Value
                SortasiMainResult = Val(SortasiMainResult) + Val(GridVal8)

            Next

        ElseIf DirChk1.Checked = True And IndiLabel.Text = "Miscellaneous" Then
            MiscMainResult = ""

            For i = 0 To WFGrid01.Rows.Count - 1

                GridVal4 = WFGrid01(4, i).Value
                MiscMainResult = Val(MiscMainResult) + Val(GridVal4)
                WPTot2.Text = Format(Val(MiscMainResult), "N0")

            Next

        End If
    End Sub

    Sub PiecesCounter()
        If IndiLabel.Text = "Sortasi" Then
            SortPiecesTotal = ""

            For b = 0 To WFGrid01.Rows.Count - 1
                GridVal7 = WFGrid01(7, b).Value
                SortPiecesTotal = Val(SortPiecesTotal) + Val(GridVal7)
                WPTot1.Text = Format(Val(SortPiecesTotal), "N0")

            Next
        End If
    End Sub

    'Conveyour Save
    Sub SaveUpConv()

        For a = 0 To WFGrid01.Rows.Count - 1

            GridVal0 = WFGrid01(0, a).Value
            GridVal1 = WFGrid01(1, a).Value
            GridVal2 = WFGrid01(2, a).Value
            GridVal3 = WFGrid01(3, a).Value
            GridVal4 = WFGrid01(4, a).Value
            GridVal5 = WFGrid01(5, a).Value
            GridVal6 = WFGrid01(6, a).Value
            GridVal7 = WFGrid01(7, a).Value

            SQL = ""
            SQL = SQL & "Select * From 03_Conveyour_Table "
            SQL = SQL & "Where Process_ID = ('" & GridVal0 & "') "

            OpenTbl(ADb, Atb5, SQL)

            If Not Atb5.RecordCount <> 0 Then
                Atb5.AddNew()


                Atb5("Time").Value = GridVal2
                Atb5("Nik").Value = GridVal3
                Atb5("Target").Value = GridVal5
                Atb5("Pieces").Value = GridVal4
                Atb5("Salary").Value = GridVal6
                Atb5("Date").Value = GridVal1.ToString("yyyy-MM-dd")
                Atb5("Process_ID").Value = GridVal0
                Atb5.Update()

                WFGrid01(7, a).Value = "Has Been Saved"

            ElseIf Atb5.RecordCount > 0 Then

                WFGrid01(7, a).Value = "Data is Already Exist"

            End If

        Next

    End Sub

    ' Mutu II Save
    Sub SaveUpMutuII()
        For a = 0 To WFGrid01.Rows.Count - 1

            GridVal0 = WFGrid01(0, a).Value
            GridVal1 = WFGrid01(1, a).Value
            GridVal2 = WFGrid01(2, a).Value
            GridVal3 = WFGrid01(3, a).Value
            GridVal4 = WFGrid01(4, a).Value
            GridVal5 = WFGrid01(5, a).Value
            GridVal6 = WFGrid01(6, a).Value
            GridVal7 = WFGrid01(7, a).Value
            GridVal8 = WFGrid01(8, a).Value

            SQL = ""
            SQL = SQL & "Select * From 04_MutuII_Table "
            SQL = SQL & "Where Process_ID = ('" & GridVal0 & "') "

            OpenTbl(ADb, Atb5, SQL)

            If Not Atb5.RecordCount <> 0 Then
                Atb5.AddNew()

                Atb5("Time").Value = GridVal2
                Atb5("Nik").Value = GridVal3
                Atb5("Target").Value = GridVal4
                Atb5("Pieces").Value = GridVal5
                Atb5("Salary").Value = GridVal6
                Atb5("Date").Value = GridVal1.ToString("yyyy-MM-dd")
                Atb5("Process_ID").Value = GridVal0
                Atb5("Coupon").Value = GridVal7
                Atb5.Update()

                WFGrid01(8, a).Value = "Has Been Saved"

            ElseIf Atb5.RecordCount > 0 Then

                WFGrid01(8, a).Value = "Data is Already Exist"

            End If

        Next

    End Sub

    ' Packing Save 

    Sub SaveUpPack()

        For a = 0 To WFGrid01.Rows.Count - 1

            GridVal0 = WFGrid01(0, a).Value
            GridVal1 = WFGrid01(1, a).Value
            GridVal2 = WFGrid01(2, a).Value
            GridVal3 = WFGrid01(3, a).Value
            GridVal4 = WFGrid01(4, a).Value
            GridVal5 = WFGrid01(5, a).Value
            GridVal6 = WFGrid01(6, a).Value
            GridVal7 = WFGrid01(7, a).Value
            GridVal8 = WFGrid01(8, a).Value

            SQL = ""
            SQL = SQL & "Select * From 05_Packing_Table "
            SQL = SQL & "Where Process_ID = ('" & GridVal0 & "') "


            OpenTbl(ADb, Atb5, SQL)

            If Not Atb5.RecordCount <> 0 Then
                Atb5.AddNew()

                Atb5("Time").Value = GridVal2
                Atb5("Nik").Value = GridVal3
                Atb5("Target").Value = GridVal4
                Atb5("Salary").Value = GridVal6
                Atb5("Date").Value = GridVal1.ToString("yyyy-MM-dd")
                Atb5("Process_ID").Value = GridVal0
                Atb5("Carton").Value = GridVal5
                Atb5("Coupon").Value = GridVal7
                Atb5.Update()


                WFGrid01(8, a).Value = "Has Been Saved"

            ElseIf Atb5.RecordCount > 0 Then

                WFGrid01(8, a).Value = "Data is Already Exist"
            End If
        Next
    End Sub

    ' Wallet Save
    Sub SaveUpWall()

        For a = 0 To WFGrid01.Rows.Count - 1

            GridVal0 = WFGrid01(0, a).Value
            GridVal1 = WFGrid01(1, a).Value
            GridVal2 = WFGrid01(2, a).Value
            GridVal3 = WFGrid01(3, a).Value
            GridVal4 = WFGrid01(4, a).Value
            GridVal5 = WFGrid01(5, a).Value
            GridVal6 = WFGrid01(6, a).Value
            GridVal7 = WFGrid01(7, a).Value
            GridVal8 = WFGrid01(8, a).Value

            SQL = ""
            SQL = SQL & "Select * From 06_Wallet_Table "
            SQL = SQL & "Where Process_ID = ('" & GridVal0 & "') "

            OpenTbl(ADb, Atb5, SQL)

            If Not Atb5.RecordCount <> 0 Then
                Atb5.AddNew()

                Atb5("Time").Value = GridVal2
                Atb5("Nik").Value = GridVal3
                Atb5("Target").Value = GridVal4
                Atb5("Pieces").Value = GridVal5
                Atb5("Salary").Value = GridVal6
                Atb5("Date").Value = GridVal1.ToString("yyyy-MM-dd")
                Atb5("Process_ID").Value = GridVal0
                Atb5("Coupon").Value = GridVal7
                Atb5.Update()


                WFGrid01(8, a).Value = "Has Been Saved"

            ElseIf Atb5.RecordCount > 0 Then

                WFGrid01(8, a).Value = "Data is Already Exist"
            End If

        Next

    End Sub

    ' Misc Save
    Sub SaveUpMisc()
        For a = 0 To WFGrid01.Rows.Count - 1

            GridVal0 = WFGrid01(0, a).Value
            GridVal1 = WFGrid01(1, a).Value
            GridVal2 = WFGrid01(2, a).Value
            GridVal3 = WFGrid01(3, a).Value
            GridVal4 = WFGrid01(4, a).Value
            GridVal5 = WFGrid01(5, a).Value

            SQL = ""
            SQL = SQL & "Select * From 19_Miscellaneous_Table "
            SQL = SQL & "Where Process_ID = ('" & GridVal0 & "') "

            OpenTbl(ADb, Atb5, SQL)

            If Not Atb5.RecordCount <> 0 Then
                Atb5.AddNew()

                Atb5("Time").Value = GridVal2
                Atb5("Nik").Value = GridVal3
                Atb5("Salary").Value = GridVal4
                Atb5("Date").Value = GridVal1.ToString("yyyy-MM-dd")
                Atb5("Process_ID").Value = GridVal0

                Atb5.Update()

                WFGrid01(5, a).Value = "Has Been Saved"

            ElseIf Atb5.RecordCount > 0 Then

                WFGrid01(5, a).Value = "Data is Already Exist"

            End If

        Next

    End Sub

    ' Sort Save

    Sub SaveUpSort()

        For a = 0 To WFGrid01.Rows.Count - 1

            GridVal0 = WFGrid01(0, a).Value
            GridVal1 = WFGrid01(1, a).Value
            GridVal2 = WFGrid01(2, a).Value
            GridVal3 = WFGrid01(3, a).Value
            GridVal4 = WFGrid01(4, a).Value
            GridVal5 = WFGrid01(5, a).Value
            GridVal6 = WFGrid01(6, a).Value
            GridVal7 = WFGrid01(7, a).Value
            GridVal8 = WFGrid01(8, a).Value
            GridVal9 = WFGrid01(9, a).Value

            SQL = ""
            SQL = SQL & "Select * From 21_NewMiscellaneous_Table "
            SQL = SQL & "Where Process_ID = ('" & GridVal0 & "') "

            OpenTbl(ADb, Atb5, SQL)

            If Not Atb5.RecordCount <> 0 Then
                Atb5.AddNew()

                Atb5("Time").Value = GridVal2
                Atb5("Nik").Value = GridVal3
                Atb5("NoKg").Value = GridVal4
                Atb5("NoBag").Value = GridVal5
                Atb5("NoGr").Value = GridVal6
                Atb5("Pieces").Value = GridVal7
                Atb5("Coupon").Value = GridVal9
                Atb5("Salary").Value = GridVal8
                Atb5("Date").Value = GridVal1.ToString("yyyy-MM-dd")
                Atb5("Process_ID").Value = GridVal0
                Atb5.Update()

            ElseIf Atb5.RecordCount > 0 Then

                WFGrid01(10, a).Value = "Data is Already Exist"

            End If

        Next

    End Sub


#End Region

    Private Sub WPTot2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles WPTot2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub
    Private Sub WPTot1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles WPTot1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub
    Private Sub WPSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WPSave.Click
        If IndiLabel.Text = "Conveyour" Then
            SaveUpConv()
        ElseIf IndiLabel.Text = "Mutu II" Then
            SaveUpMutuII()
        ElseIf IndiLabel.Text = "Wallet" Then
            SaveUpWall()
        ElseIf IndiLabel.Text = "Packing" Then
            SaveUpPack()
        ElseIf IndiLabel.Text = "Miscellaneous" Then
            SaveUpMisc()
        ElseIf IndiLabel.Text = "Sortasi" Then
            SaveUpSort()
        ElseIf IndiLabel.Text = "Over All" Then

        End If

        If DirChk2.Checked = "True" Then
            RushDate = WFCal.Text
            If IndiLabel.Text = "Conveyour" Then
                ConveyourSalary()

            ElseIf IndiLabel.Text = "Mutu II" Then
                MutuIISalary()

            ElseIf IndiLabel.Text = "Wallet" Then
                WalletSalary()

            ElseIf IndiLabel.Text = "Packing" Then
                PackingSalary()

            ElseIf IndiLabel.Text = "Miscellaneous" Then
                MiscellaneousSalary()

            ElseIf IndiLabel.Text = "Sortasi" Then
                SortasiSalary()

            ElseIf IndiLabel.Text = "Over All" Then
                Try
                    OverAllSaver()
                Catch ex As Exception
                    MessageBox.Show(ex.Message, Me.Text)
                End Try



            End If
        End If

    End Sub

#Region "Per Dept Salary"
    Dim RushDate As Date
    Sub ConveyourSalary()


        SQL = ""
        SQL = SQL & "Select * From 13_Conveyour_Salary "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & RushDate.ToString("yyyy-MM-dd") & "') "

        OpenTbl(ADb, Atbl30, SQL)

        If Not Atbl30.RecordCount <> 0 Then
            Atbl30.AddNew()
        End If

        Atbl30("Date").Value = WFCal.Text
        Atbl30("Nik").Value = WPTbx1.Text
        Atbl30("Salary").Value = WPTot2.Text

        Atbl30.Update()
        Me.Refresh()

    End Sub
    Sub MutuIISalary()

        SQL = ""
        SQL = SQL & "Select * From 14_MutuII_Salary "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & RushDate.ToString("yyyy-MM-dd") & "') "
        OpenTbl(ADb, Atbl30, SQL)

        If Not Atbl30.RecordCount <> 0 Then
            Atbl30.AddNew()
        End If

        Atbl30("Date").Value = WFCal.Text
        Atbl30("Nik").Value = WPTbx1.Text
        Atbl30("Salary").Value = WPTot2.Text

        Atbl30.Update()
        Me.Refresh()

    End Sub
    Sub WalletSalary()

        SQL = ""
        SQL = SQL & "Select * From 15_Wallet_Salary "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & RushDate.ToString("yyyy-MM-dd") & "') "
        OpenTbl(ADb, Atbl30, SQL)

        If Not Atbl30.RecordCount <> 0 Then
            Atbl30.AddNew()
        End If

        Atbl30("Date").Value = WFCal.Text
        Atbl30("Nik").Value = WPTbx1.Text
        Atbl30("Salary").Value = WPTot2.Text

        Atbl30.Update()
        Me.Refresh()

    End Sub
    Sub PackingSalary()

        SQL = ""
        SQL = SQL & "Select * From 16_Packing_Salary "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & RushDate.ToString("yyyy-MM-dd") & "') "
        OpenTbl(ADb, Atbl30, SQL)

        If Not Atbl30.RecordCount <> 0 Then
            Atbl30.AddNew()
        End If

        Atbl30("Date").Value = WFCal.Text
        Atbl30("Nik").Value = WPTbx1.Text
        Atbl30("Salary").Value = WPTot2.Text

        Atbl30.Update()
        Me.Refresh()

    End Sub
    Sub MiscellaneousSalary()

        SQL = ""
        SQL = SQL & "Select * From 20_Miscellaneous_Salary "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & RushDate.ToString("yyyy-MM-dd") & "') "
        SQL = SQL & "And TypeCtrl = ('" & "Old" & "') "
        OpenTbl(ADb, Atbl30, SQL)

        If Not Atbl30.RecordCount <> 0 Then
            Atbl30.AddNew()
        End If



        Atbl30("Date").Value = WFCal.Text
        Atbl30("Nik").Value = WPTbx1.Text
        Atbl30("Salary").Value = WPTot2.Text
        Atbl30("TypeCtrl").Value = "Old"

        Atbl30.Update()

        Me.Refresh()

    End Sub
    Sub SortasiSalary()


        SQL = ""
        SQL = SQL & "Select * From 20_Miscellaneous_Salary "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & RushDate.ToString("yyyy-MM-dd") & "') "
        SQL = SQL & "And TypeCtrl = ('" & "New" & "') "
        OpenTbl(ADb, Atbl30, SQL)

        If Not Atbl30.RecordCount <> 0 Then
            Atbl30.AddNew()
        End If

        Atbl30("Date").Value = WFCal.Text
        Atbl30("Nik").Value = WPTbx1.Text
        Atbl30("Salary").Value = WPTot2.Text
        Atbl30("TypeCtrl").Value = "New"

        Atbl30.Update()


        Me.Refresh()

    End Sub

    Sub OverAllSaver()
        SQL = ""
        SQL = SQL & "Select * From SalarySync1_Table "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Periode = ('" & PeriodeCtrl & "') "
        SQL = SQL & "And PeriodeRange = ('" & PeriodeMonthCtrl & "') "
        OpenTbl(CBb, Ctbl1, SQL)

        If Not Ctbl1.RecordCount <> 0 Then
            Ctbl1.AddNew()
        End If
        StringCaller = WPTbx2.Text.Replace("?", "'")
        Ctbl1("Nik").Value = WPTbx1.Text
        Ctbl1("Name").Value = StringCaller
        Ctbl1("Periode").Value = PeriodeCtrl
        Ctbl1("PeriodeRange").Value = PeriodeMonthCtrl
        Ctbl1("Pay").Value = WorkSupTbx4.Text
        Ctbl1("AstekVal").Value = WorkSupTbx3.Text

        If WorkSupTbx8.Text = "1" Then

            Ctbl1("Salary1").Value = OverAllTot
            Ctbl1("Date1").Value = WFCal.Text

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "2" Then

            Ctbl1("Salary2").Value = OverAllTot
            Ctbl1("Date2").Value = WFCal.Text

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "3" Then

            Ctbl1("Salary3").Value = OverAllTot
            Ctbl1("Date3").Value = WFCal.Text
            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "4" Then

            Ctbl1("Salary4").Value = OverAllTot
            Ctbl1("Date4").Value = WFCal.Text

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "5" Then

            Ctbl1("Salary5").Value = OverAllTot
            Ctbl1("Date5").Value = WFCal.Text

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "6" Then

            Ctbl1("Salary6").Value = OverAllTot
            Ctbl1("Date6").Value = WFCal.Text

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "7" Then

            Ctbl1("Salary7").Value = OverAllTot
            Ctbl1("Date7").Value = WFCal.Text

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "8" Then

            Ctbl1("Salary8").Value = OverAllTot
            Ctbl1("Date8").Value = WFCal.Text

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "9" Then

            Ctbl1("Salary9").Value = OverAllTot
            Ctbl1("Date9").Value = WFCal.Text

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "10" Then

            Ctbl1("Salary10").Value = OverAllTot
            Ctbl1("Date10").Value = WFCal.Text

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "11" Then

            Ctbl1("Salary11").Value = OverAllTot
            Ctbl1("Date11").Value = WFCal.Text

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "12" Then

            Ctbl1("Salary12").Value = OverAllTot
            Ctbl1("Date12").Value = WFCal.Text

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "13" Then

            Ctbl1("Salary13").Value = OverAllTot
            Ctbl1("Date13").Value = WFCal.Text

            Ctbl1.Update()
            Me.Refresh()


        ElseIf WorkSupTbx8.Text = "14" Then

            Ctbl1("Salary14").Value = OverAllTot
            Ctbl1("Date14").Value = WFCal.Text

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "15" Then

            Ctbl1("Salary15").Value = OverAllTot
            Ctbl1("Date15").Value = WFCal.Text

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "16" Then

            Ctbl1("Salary16").Value = OverAllTot
            Ctbl1("Date16").Value = WFCal.Text

            Ctbl1.Update()
            Me.Refresh()

        End If


        'If InceChkBox1.Checked = True Then
        '    IncentivesControlSave()
        'End If

        MsgBox("Save!", vbInformation)

    End Sub

#End Region

#Region "Load Max Salary"

    Sub LoadAllSalary()

        SQL = ""
        SQL = SQL & "Select * from 13_Conveyour_Salary "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') "
        OpenTbl(ADb, Dbtb21, SQL)
        If Dbtb21.RecordCount > 0 Then

            ConCD = IIf(IsDBNull(Dbtb21("Salary").Value), "", Dbtb21("Salary").Value)
        Else
            ConCD = 0

        End If

        SQL = ""
        SQL = SQL & "Select * from 14_MutuII_Salary "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') "
        OpenTbl(ADb, Dbtb22, SQL)
        If Dbtb22.RecordCount > 0 Then

            MutuIICD = IIf(IsDBNull(Dbtb22("Salary").Value), "", Dbtb22("Salary").Value)

        Else
            MutuIICD = 0

        End If

        SQL = ""
        SQL = SQL & "Select * from 15_Wallet_Salary "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') "
        OpenTbl(ADb, Dbtb23, SQL)
        If Dbtb23.RecordCount > 0 Then
            WalletCD = IIf(IsDBNull(Dbtb23("Salary").Value), "", Dbtb23("Salary").Value)

        Else
            WalletCD = 0

        End If

        SQL = ""
        SQL = SQL & "Select * from 16_Packing_Salary "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') "
        OpenTbl(ADb, Dbtb24, SQL)
        If Dbtb24.RecordCount > 0 Then
            PackingCD = IIf(IsDBNull(Dbtb24("Salary").Value), "", Dbtb24("Salary").Value)

        Else
            PackingCD = 0

        End If

        SQL = ""
        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') "
        SQL = SQL & "And TypeCtrl = ('" & "Old" & "') "
        OpenTbl(ADb, Dbtb35, SQL)
        If Dbtb35.RecordCount > 0 Then
            MiscellaneousCD = IIf(IsDBNull(Dbtb35("Salary").Value), "", Dbtb35("Salary").Value)

        Else
            MiscellaneousCD = 0

        End If

        SQL = ""
        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') "
        SQL = SQL & "And TypeCtrl = ('" & "New" & "') "
        OpenTbl(ADb, Dbtb37, SQL)
        If Dbtb37.RecordCount > 0 Then
            NewMiscellaneousCD = IIf(IsDBNull(Dbtb37("Salary").Value), "", Dbtb37("Salary").Value)

        Else
            NewMiscellaneousCD = 0

        End If


        OverAllTot = ConCD + MutuIICD + WalletCD + PackingCD + NewMiscellaneousCD + MiscellaneousCD
        'WFGrid01.Invoke(DirectCast(Sub() WFGrid01.Rows.Add(WFCal.Text, ConCD, MutuIICD, WalletCD, PackingCD, NewMiscellaneousCD, MiscellaneousCD, Format(OverAllTot, "N0")), MethodInvoker))
        WFGrid01.Rows.Add(WFCal.Text, ConCD, MutuIICD, WalletCD, PackingCD, NewMiscellaneousCD, MiscellaneousCD, OverAllTot.ToString("N0", CustomtoUS))

    End Sub

    
    Sub LoadConveyour()
        DateGet = WFCal.Text

        SQL = ""
        SQL = SQL & "Select * from 03_conveyour_table "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') Limit 15 "
        OpenTbl(ADb, DbTbl4, SQL)

        If DbTbl4.RecordCount > 0 Then

            DbTbl4.MoveFirst()
            Do While Not DbTbl4.EOF

                NikCtrl = IIf(IsDBNull(DbTbl4("Nik").Value), "", DbTbl4("Nik").Value)
                ProID = IIf(IsDBNull(DbTbl4("Process_ID").Value), "", DbTbl4("Process_ID").Value)
                TimeCtrl = IIf(IsDBNull(DbTbl4("Time").Value), "", DbTbl4("Time").Value)
                DateCtrl = IIf(IsDBNull(DbTbl4("Date").Value), "", DbTbl4("Date").Value)
                PcsCtrl = IIf(IsDBNull(DbTbl4("Pieces").Value), "", DbTbl4("Pieces").Value)
                TarCtrl = IIf(IsDBNull(DbTbl4("Target").Value), "", DbTbl4("Target").Value)
                SalCtrl = IIf(IsDBNull(DbTbl4("Salary").Value), "", DbTbl4("Salary").Value)
                TimeCtrl = IIf(IsDBNull(DbTbl4("Time").Value), "", DbTbl4("Time").Value)

                'WFGrid01.Invoke(DirectCast(Sub() WFGrid01.Rows.Add(ProID, DateCtrl, TimeCtrl, NikCtrl, PcsCtrl, TarCtrl, SalCtrl), MethodInvoker))
                WFGrid01.Rows.Add(ProID, DateCtrl, TimeCtrl, NikCtrl, PcsCtrl, TarCtrl, SalCtrl)
                ProcessMode()

                DbTbl4.MoveNext()

            Loop

        End If

    End Sub

    Sub LoadMutuII()
        DateGet = WFCal.Text
        SQL = ""
        SQL = SQL & "Select * from 04_MutuII_Table "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') Limit 15 "
        OpenTbl(ADb, DbTbl7, SQL)
        If DbTbl7.RecordCount > 0 Then

            DbTbl7.MoveFirst()
            Do While Not DbTbl7.EOF

                NikCtrl = IIf(IsDBNull(DbTbl7("Nik").Value), "", DbTbl7("Nik").Value)
                ProID = IIf(IsDBNull(DbTbl7("Process_ID").Value), "", DbTbl7("Process_ID").Value)
                TimeCtrl = IIf(IsDBNull(DbTbl7("Time").Value), "", DbTbl7("Time").Value)
                DateCtrl = IIf(IsDBNull(DbTbl7("Date").Value), "", DbTbl7("Date").Value)
                PcsCtrl = IIf(IsDBNull(DbTbl7("Pieces").Value), "", DbTbl7("Pieces").Value)
                TarCtrl = IIf(IsDBNull(DbTbl7("Target").Value), "", DbTbl7("Target").Value)
                SalCtrl = IIf(IsDBNull(DbTbl7("Salary").Value), "", DbTbl7("Salary").Value)
                CouCtrl = IIf(IsDBNull(DbTbl7("Coupon").Value), "", DbTbl7("Coupon").Value)
                'WFGrid01.Invoke(DirectCast(Sub() WFGrid01.Rows.Add(ProID, DateCtrl, TimeCtrl, NikCtrl, TarCtrl, PcsCtrl, SalCtrl, CouCtrl), MethodInvoker))
                WFGrid01.Rows.Add(ProID, DateCtrl, TimeCtrl, NikCtrl, TarCtrl, PcsCtrl, SalCtrl, CouCtrl)
                ProcessMode()
                DbTbl7.MoveNext()


            Loop
        End If

    End Sub

    Sub LoadPacking()
        DateGet = WFCal.Text
        SQL = ""
        SQL = SQL & "Select * from 05_Packing_Table "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') Limit 15 "
        OpenTbl(ADb, DbTbl9, SQL)
        If DbTbl9.RecordCount > 0 Then

            DbTbl9.MoveFirst()
            Do While Not DbTbl9.EOF


                NikCtrl = IIf(IsDBNull(DbTbl9("Nik").Value), "", DbTbl9("Nik").Value)
                ProID = IIf(IsDBNull(DbTbl9("Process_ID").Value), "", DbTbl9("Process_ID").Value)
                TimeCtrl = IIf(IsDBNull(DbTbl9("Time").Value), "", DbTbl9("Time").Value)
                DateCtrl = IIf(IsDBNull(DbTbl9("Date").Value), "", DbTbl9("Date").Value)
                CartCtrl = IIf(IsDBNull(DbTbl9("Carton").Value), "", DbTbl9("Carton").Value)
                TarCtrl = IIf(IsDBNull(DbTbl9("Target").Value), "", DbTbl9("Target").Value)
                SalCtrl = IIf(IsDBNull(DbTbl9("Salary").Value), "", DbTbl9("Salary").Value)
                CouCtrl = IIf(IsDBNull(DbTbl9("Coupon").Value), "", DbTbl9("Coupon").Value)

                'WFGrid01.Invoke(DirectCast(Sub() WFGrid01.Rows.Add(ProID, DateCtrl, TimeCtrl, NikCtrl, TarCtrl, CartCtrl, SalCtrl, CouCtrl), MethodInvoker))
                WFGrid01.Rows.Add(ProID, DateCtrl, TimeCtrl, NikCtrl, TarCtrl, CartCtrl, SalCtrl, CouCtrl)
                ProcessMode()
                DbTbl9.MoveNext()

            Loop
        End If
    End Sub

    Sub LoadWallet()
        DateGet = WFCal.Text
        SQL = ""
        SQL = SQL & "Select * from 06_Wallet_Table "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') Limit 15 "
        OpenTbl(ADb, DbTbl8, SQL)
        If DbTbl8.RecordCount > 0 Then

            DbTbl8.MoveFirst()
            Do While Not DbTbl8.EOF



                NikCtrl = IIf(IsDBNull(DbTbl8("Nik").Value), "", DbTbl8("Nik").Value)
                ProID = IIf(IsDBNull(DbTbl8("Process_ID").Value), "", DbTbl8("Process_ID").Value)
                TimeCtrl = IIf(IsDBNull(DbTbl8("Time").Value), "", DbTbl8("Time").Value)
                DateCtrl = IIf(IsDBNull(DbTbl8("Date").Value), "", DbTbl8("Date").Value)
                PcsCtrl = IIf(IsDBNull(DbTbl8("Pieces").Value), "", DbTbl8("Pieces").Value)
                TarCtrl = IIf(IsDBNull(DbTbl8("Target").Value), "", DbTbl8("Target").Value)
                SalCtrl = IIf(IsDBNull(DbTbl8("Salary").Value), "", DbTbl8("Salary").Value)
                CouCtrl = IIf(IsDBNull(DbTbl8("Coupon").Value), "", DbTbl8("Coupon").Value)


                'WFGrid01.Invoke(DirectCast(Sub() WFGrid01.Rows.Add(ProID, DateCtrl, TimeCtrl, NikCtrl, TarCtrl, PcsCtrl, SalCtrl, CouCtrl), MethodInvoker))
                WFGrid01.Rows.Add(ProID, DateCtrl, TimeCtrl, NikCtrl, TarCtrl, PcsCtrl, SalCtrl, CouCtrl)
                ProcessMode()
                DbTbl8.MoveNext()


            Loop
        End If

    End Sub

    Sub LoadMisc()

        DateGet = WFCal.Text
        SQL = ""
        SQL = SQL & "Select * from 19_Miscellaneous_Table "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') Limit 15 "
        OpenTbl(ADb, Dbtb34, SQL)
        If Dbtb34.RecordCount > 0 Then
            Do While Not Dbtb34.EOF

                NikCtrl = IIf(IsDBNull(Dbtb34("Nik").Value), "", Dbtb34("Nik").Value)
                ProID = IIf(IsDBNull(Dbtb34("Process_ID").Value), "", Dbtb34("Process_ID").Value)
                TimeCtrl = IIf(IsDBNull(Dbtb34("Time").Value), "", Dbtb34("Time").Value)
                DateCtrl = IIf(IsDBNull(Dbtb34("Date").Value), "", Dbtb34("Date").Value)
                SalCtrl = IIf(IsDBNull(Dbtb34("Salary").Value), "", Dbtb34("Salary").Value)

                'WFGrid01.Invoke(DirectCast(Sub() WFGrid01.Rows.Add(ProID, DateCtrl, TimeCtrl, NikCtrl, SalCtrl, CouCtrl), MethodInvoker))
                WFGrid01.Rows.Add(ProID, DateCtrl, TimeCtrl, NikCtrl, SalCtrl, CouCtrl)
                ProcessMode()
                Dbtb34.MoveNext()

            Loop
        End If



    End Sub

    Sub LoadSortasi()

        DateGet = WFCal.Text
        SQL = ""
        SQL = SQL & "Select * from 21_NewMiscellaneous_Table "
        SQL = SQL & "Where Nik = ('" & WPTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') Limit 15 "
        OpenTbl(ADb, Dbtb37, SQL)
        If Dbtb37.RecordCount > 0 Then

            Do While Not Dbtb37.EOF

                NikCtrl = IIf(IsDBNull(Dbtb37("Nik").Value), "", Dbtb37("Nik").Value)
                ProID = IIf(IsDBNull(Dbtb37("Process_ID").Value), "", Dbtb37("Process_ID").Value)
                TimeCtrl = IIf(IsDBNull(Dbtb37("Time").Value), "", Dbtb37("Time").Value)
                DateCtrl = IIf(IsDBNull(Dbtb37("Date").Value), "", Dbtb37("Date").Value)
                PcsCtrl = IIf(IsDBNull(Dbtb37("Pieces").Value), "", Dbtb37("Pieces").Value)
                SalCtrl = IIf(IsDBNull(Dbtb37("Salary").Value), "", Dbtb37("Salary").Value)
                CouCtrl = IIf(IsDBNull(Dbtb37("Coupon").Value), "", Dbtb37("Coupon").Value)
                NoKgCtrl = IIf(IsDBNull(Dbtb37("NoKg").Value), "", Dbtb37("NoKg").Value)
                NoGrCtrl = IIf(IsDBNull(Dbtb37("NoGr").Value), "", Dbtb37("NoGr").Value)
                NoBagCtrl = IIf(IsDBNull(Dbtb37("NoBag").Value), "", Dbtb37("NoBag").Value)

                'WFGrid01.Invoke(DirectCast(Sub() WFGrid01.Rows.Add(ProID, DateCtrl, TimeCtrl, NikCtrl, NoKgCtrl, NoBagCtrl, NoGrCtrl, PcsCtrl, SalCtrl, CouCtrl), MethodInvoker))
                WFGrid01.Rows.Add(ProID, DateCtrl, TimeCtrl, NikCtrl, NoKgCtrl, NoBagCtrl, NoGrCtrl, PcsCtrl, SalCtrl, CouCtrl)
                Dbtb37.MoveNext()
                ProcessMode()

            Loop
        End If

    End Sub


    Sub ErrorSpecial()
        On Error GoTo Err
        SpecialSortasi()
        Exit Sub
Err:
    End Sub
#End Region

    Private Sub GlassButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WPClear.Click
        WFGrid01.Rows.Clear()
        WFGrid01.Columns.Clear()
        WFGridHeader()
        WPTot1.Text = ""
        WPTot2.Text = ""

        ConMainResult = Nothing
        MutuMainResult = Nothing
        WallMainResult = Nothing
        PackMainResult = Nothing
        SortasiMainResult = Nothing
        MiscMainResult = Nothing
        PackPieceRes = Nothing
        ConPiecesRes = Nothing
        MutPiecesRes = Nothing
        WalPiecesRes = Nothing

    End Sub


    Private Sub WpSData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WpSData.Click
        If IndiLabel.Text = "Conveyour" Then
            WFBtn1.PerformClick()

        ElseIf IndiLabel.Text = "Mutu II" Then
            WFBtn2.PerformClick()

        ElseIf IndiLabel.Text = "Packing" Then
            WFBtn4.PerformClick()

        ElseIf IndiLabel.Text = "Wallet" Then
            WFBtn3.PerformClick()

        ElseIf IndiLabel.Text = "Sortasi" Then
            WFBtn5.PerformClick()

        ElseIf IndiLabel.Text = "Miscellaneous" Then
            WFBtn6.PerformClick()
        End If
    End Sub



    Private Sub PackTbx1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PackTbx1.TextChanged

    End Sub


    Private Sub WFBgWorker_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles WFBgWorker.DoWork
        DateGet = WFCal.Text

        If IndiLabel.Text = "Conveyour" Then
            LoadConveyour()
        ElseIf IndiLabel.Text = "Mutu II" Then
            LoadMutuII()
        ElseIf IndiLabel.Text = "Wallet" Then
            LoadWallet()
        ElseIf IndiLabel.Text = "Packing" Then
            LoadPacking()
        ElseIf IndiLabel.Text = "Sortasi" Then
            LoadSortasi()
        ElseIf IndiLabel.Text = "Miscellaneous" Then
            LoadMisc()
        ElseIf IndiLabel.Text = "Over All" Then
            LoadAllSalary()

        End If
    End Sub


    Private Sub SortPiTbx_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SortPiTbx.TextChanged

    End Sub

    Private Sub SortMasTbx1_MaskInputRejected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MaskInputRejectedEventArgs) Handles SortMasTbx1.MaskInputRejected

    End Sub

    Private Sub WPTbx1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WPTbx1.TextChanged

    End Sub
End Class