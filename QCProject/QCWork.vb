Option Explicit On

Public Class WorkBlock

    Dim SizeMode As New System.Drawing.Size(728, 439)
    Dim SizeGrid As New System.Drawing.Size(534, 331)
    Dim LocMode As New System.Drawing.Point(308, 50)
    Dim LocGrid As New System.Drawing.Point(185, 24)
    Dim DateYearDate As New System.DateTime
    Dim date2 As New System.DateTime
    Dim ProcessNum As String
    Dim ProcessDig As String
    Dim m1 As String
    Dim y1 As String
    Dim m2 As String
    Dim tr3 As String
    Dim YearDate As String
    Dim AstekShow As String
    Dim AstekLook As String
    Dim TimeLimiter1 As Integer
    Dim TimeLimiter2 As Integer
    Dim PiecesNest As Double
    Dim TargetNest As Double
    Dim SalaryNest As Double
    Dim CtnNest As Integer
    Dim TimeNest As String
    Dim NoKgNest As Integer
    Dim NoGramNest As String
    Dim NoBagNest As Integer
    Dim GrControl As String
    Dim ConCD As Double
    Dim MutuIICD As Double
    Dim WalletCD As Double
    Dim PackingCD As Double
    Dim MiscellaneousCD As Double
    Dim NewMiscellaneousCD As Double
    Dim Incenload1 As String
    Dim Inceload2 As String
    Dim IncentiveCount As String
    Dim IncentiveLock As String
    Dim MiscHolTot As String
    Dim SorHolTot As String


    Private Sub WorkBlock_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        Me.Dispose(True)
    End Sub

    Private Sub WorkBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        OverAllStyle()
        LoadDB()
        LoadDB2()
        TimeLimiter1 = 0
        TimeLimiter2 = 0
        LoadHolidayMod()
        IncentivesControlRange()
        ErrorOverAll()
        IncentivesEnabler()
    End Sub


    Sub NestingtoZero()

        PiecesNest = 0
        TargetNest = 0
        SalaryNest = 0.0
        CtnNest = 0
        NoKgNest = 0
        NoGramNest = 0
        NoBagNest = 0
        TimeNest = ""

    End Sub

    Sub ConveyourStyle()

        WorkFrm1.AutoSize = True
        WorkFrm1.Visible = True
        WorkFrm1.Size = SizeMode
        WorkFrm1.Location = LocMode
        ConGrid.Location = LocGrid
        ConGrid.Size = SizeGrid


    End Sub

    Sub MutuIIStyle()

        WorkFrm2.AutoSize = True
        WorkFrm2.Visible = True
        WorkFrm2.Size = SizeMode
        WorkFrm2.Location = LocMode
        MutuGrid.Location = LocGrid
        MutuGrid.Size = SizeGrid
    End Sub

    Sub WalletStyle()
        WorkFrm3.AutoSize = True
        WorkFrm3.Visible = True
        WorkFrm3.Size = SizeMode
        WorkFrm3.Location = LocMode
        WalletGrid.Location = LocGrid
        WalletGrid.Size = SizeGrid
    End Sub

    Sub PackingStyle()

        WorkFrm4.AutoSize = True
        WorkFrm4.Visible = True
        WorkFrm4.Size = SizeMode
        WorkFrm4.Location = LocMode
        PackingGrid.Location = LocGrid
        PackingGrid.Size = SizeGrid
    End Sub

    Sub SortasiStyle()
        WorkFrm5.AutoSize = True
        WorkFrm5.Visible = True
        WorkFrm5.Size = SizeMode
        WorkFrm5.Location = LocMode
        SortGrid.Location = LocGrid
        SortGrid.Size = SizeGrid
    End Sub

    Sub MiscellaneousStyle()

        WorkFrm6.AutoSize = True
        WorkFrm6.Visible = True
        WorkFrm6.Size = SizeMode
        WorkFrm6.Location = LocMode
        MiscGrid.Location = LocGrid
        MiscGrid.Size = SizeGrid
    End Sub

    Sub OverAllStyle()
        WorkFrm7.AutoSize = True
        WorkFrm7.Visible = True
        WorkFrm7.Size = SizeMode
        WorkFrm7.Location = LocMode

    End Sub

    Private Sub WorkBtn3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WorkBtn3.Click
        WalletStyle()
        NestingtoZero()
        LoadWallet()
        WorkFrm1.Visible = False
        WorkFrm2.Visible = False
        WorkFrm3.Visible = True
        WorkFrm4.Visible = False
        WorkFrm5.Visible = False
        WorkFrm6.Visible = False
        WorkFrm7.Visible = False
        WorkTimer2.Enabled = True
        Me.Refresh()

    End Sub

    Private Sub WorkBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WorkBtn1.Click
        ConveyourStyle()
        NestingtoZero()
        LoadConveyour()
        WorkFrm1.Visible = True
        WorkFrm2.Visible = False
        WorkFrm3.Visible = False
        WorkFrm4.Visible = False
        WorkFrm5.Visible = False
        WorkFrm6.Visible = False
        WorkFrm7.Visible = False
        WorkTimer2.Enabled = True
        Me.Refresh()
    End Sub

    Private Sub WorkBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WorkBtn2.Click
        MutuIIStyle()
        NestingtoZero()
        LoadMutuII()
        WorkFrm1.Visible = False
        WorkFrm2.Visible = True
        WorkFrm3.Visible = False
        WorkFrm4.Visible = False
        WorkFrm5.Visible = False
        WorkFrm6.Visible = False
        WorkFrm7.Visible = False
        WorkTimer2.Enabled = True
        Me.Refresh()
    End Sub

    Private Sub WorkBtn4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WorkBtn4.Click
        PackingStyle()
        NestingtoZero()
        LoadPacking()
        IncentivesControlLoad()
        WorkFrm1.Visible = False
        WorkFrm2.Visible = False
        WorkFrm3.Visible = False
        WorkFrm4.Visible = True
        WorkFrm5.Visible = False
        WorkFrm6.Visible = False
        WorkFrm7.Visible = False
        WorkTimer2.Enabled = True
        Me.Refresh()

    End Sub

    Private Sub WorkBtn5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WorkBtn5.Click
        SortasiStyle()
        NestingtoZero()
        LoadSortasi()
        WorkFrm1.Visible = False
        WorkFrm2.Visible = False
        WorkFrm3.Visible = False
        WorkFrm4.Visible = False
        WorkFrm5.Visible = True
        WorkFrm6.Visible = False
        WorkFrm7.Visible = False
        WorkTimer2.Enabled = True
        Me.Refresh()

    End Sub

    Private Sub WorkBtn6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WorkBtn6.Click

        MiscellaneousStyle()
        NestingtoZero()
        LoadMisc()
        WorkFrm1.Visible = False
        WorkFrm2.Visible = False
        WorkFrm3.Visible = False
        WorkFrm4.Visible = False
        WorkFrm5.Visible = False
        WorkFrm6.Visible = True
        WorkFrm7.Visible = False
        Me.Refresh()

    End Sub

    Private Sub WorkBtn7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WorkBtn7.Click

        OverAllStyle()
        NestingtoZero()
        WorkFrm1.Visible = False
        WorkFrm2.Visible = False
        WorkFrm3.Visible = False
        WorkFrm4.Visible = False
        WorkFrm5.Visible = False
        WorkFrm6.Visible = False
        WorkFrm7.Visible = True
        WorkTimer2.Enabled = True
        Me.Refresh()

    End Sub

    Private Sub PanelBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    ' Action Mode
    Sub GenConCode()

        SQL = ""
        SQL = SQL & "Select * From 03_Conveyour_Table "
        SQL = SQL & "Order by Process_ID Desc "
        OpenTbl(ADb, DbTbl6, SQL)
        If DbTbl6.RecordCount <> 0 Then
            ProcessDig = DbTbl6("Process_ID").Value
            ProcessNum = Format(ProcessDig + 1, "0000000000")
        Else
            ProcessNum = "0000000001"
        End If
    End Sub
    Sub GenMutuIICode()

        SQL = ""
        SQL = SQL & "Select * From 04_MutuII_Table "
        SQL = SQL & "Order by Process_ID Desc "
        OpenTbl(ADb, DbTbl6, SQL)
        If DbTbl6.RecordCount <> 0 Then
            ProcessDig = DbTbl6("Process_ID").Value
            ProcessNum = Format(ProcessDig + 1, "0000000000")
        Else
            ProcessNum = "0000000001"
        End If
    End Sub
    Sub GenPackingCode()

        SQL = ""
        SQL = SQL & "Select * From 05_Packing_Table "
        SQL = SQL & "Order by Process_ID Desc "
        OpenTbl(ADb, DbTbl6, SQL)
        If DbTbl6.RecordCount <> 0 Then
            ProcessDig = DbTbl6("Process_ID").Value
            ProcessNum = Format(ProcessDig + 1, "0000000000")
        Else
            ProcessNum = "0000000001"
        End If
    End Sub
    Sub GenWalletCode()

        SQL = ""
        SQL = SQL & "Select * From 06_Wallet_Table "
        SQL = SQL & "Order by Process_ID Desc "
        OpenTbl(ADb, DbTbl6, SQL)
        If DbTbl6.RecordCount <> 0 Then
            ProcessDig = DbTbl6("Process_ID").Value
            ProcessNum = Format(ProcessDig + 1, "0000000000")
        Else
            ProcessNum = "0000000001"
        End If
    End Sub
    Sub GenMiscCode()

        SQL = ""
        SQL = SQL & "Select * From 19_Miscellaneous_Table "
        SQL = SQL & "Order by Process_ID Desc "
        OpenTbl(ADb, DbTbl6, SQL)
        If DbTbl6.RecordCount <> 0 Then
            ProcessDig = DbTbl6("Process_ID").Value
            ProcessNum = Format(ProcessDig + 1, "0000000000")
        Else
            ProcessNum = "0000000001"
        End If
    End Sub
    Sub GenNewMiscCode()

        SQL = ""
        SQL = SQL & "Select * From 21_NewMiscellaneous_Table "
        SQL = SQL & "Order by Process_ID Desc "
        OpenTbl(ADb, DbTbl6, SQL)
        If DbTbl6.RecordCount <> 0 Then
            ProcessDig = DbTbl6("Process_ID").Value
            ProcessNum = Format(ProcessDig + 1, "0000000000")
        Else
            ProcessNum = "0000000001"
        End If
    End Sub

    Sub LoadHolidayMod()

        SQL = ""
        SQL = SQL & "Select * from 17_Holiday_Table "
        SQL = SQL & "Where Date = ('" & WorkCalendar.SelectionStart & "') "
        OpenTbl(ADb, Dbtb29, SQL)

        If Dbtb29.RecordCount > 0 Then
            HolMod = Dbtb29("Salary_Mod").Value

        Else
            HolMod = "1"

        End If

    End Sub

    Private Sub WorkTimer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WorkTimer2.Tick

        LoadConveyour()
        LoadMutuII()
        LoadWallet()
        LoadPacking()
        LoadSortasi()
        LoadMisc()
        ErrorPeriode()
        WorkSupTbx8.Text = PeriodeDayCtrl
        WorkSupTbx6.Text = PeriodeCtrl
        WorkSupTbx7.Text = PeriodeMonthCtrl
        ErrorOverAll()
        WorkTimer2.Enabled = True
        PanelSaveLb1.ForeColor = Color.Black
        PanelSaveLb1.Text = "On Work: "

        PanelSaveLb2.ForeColor = Color.Green
        PanelSaveLb2.Text = " Yes "

        If TimeLimiter2 >= 30000 Then

            WorkTimer2.Enabled = False
            TimeLimiter2 = 0

        End If

    End Sub

    Sub PeriodeNest()

        PeriodeDayCtrl = ""
        PeriodeCtrl = ""
        PeriodeMonthCtrl = ""

    End Sub

    Sub YearHolMod()

        SQL = ""
        SQL = SQL & "Select * from 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        OpenTbl(ADb, Dbtb33, SQL)

        If Dbtb33.RecordCount > 0 Then

            YearDate = Dbtb33("DateStart").Value
            AstekLook = Dbtb33("Jamsostek").Value
            DateYearDate = WorkCalendar.SelectionStart

        End If

        YearMod = DateYearDate.Subtract(YearDate).Days
        WorkSupTbx2.Text = YearDate
        WorkSupTbx3.Text = AstekLook

    End Sub

    Sub MasaKerjaCtrl()

        tr3 = WorkSupTbx1.Text

        m1 = CInt(tr3) / 30

        If m1 <= 0 Then m1 = 0

        y1 = Int(CDbl(m1) / 12)

        m2 = Format((m1 - CDbl(y1) * 12), "#")

        PanelTbx4.Text = (y1) + " Tahun " + (m2) + " Bulan"

    End Sub

    Private Sub WorkTimer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WorkTimer1.Tick

        YearHolMod()
        LoadHolidayMod()
        WorkSupTbx4.Text = PayAsSetup
        WorkSupTbx1.Text = YearMod
        WorkSupTbx5.Text = HolMod

        TimeLimiter1 = TimeLimiter1 + WorkTimer1.Interval

        If TimeLimiter1 >= 4000 Then

            WorkTimer1.Enabled = False
            TimeLimiter1 = 0

        End If

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Sub EmpLookup()
        SQL = ""
        SQL = SQL & "Select * from 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "Order by Nik"
        OpenTbl(ADb, Atb3, SQL)

        If Atb3.RecordCount > 0 Then

            PanelTbx2.Text = Atb3("Name").Value
            WorkTimer1.Enabled = True
            PayAsSetup = Atb3("Pay").Value

        Else

            MsgBox("Employee Not Found", MsgBoxStyle.Information, "Codex ~ QC Build " & BuildCounter & " Warning!!")

        End If
    End Sub

#Region "GUI Control"

    Private Sub PanelTbx1_KeyPress1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PanelTbx1.KeyPress
        PanelTbx1.CharacterCasing = CharacterCasing.Upper
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            EmpLookup()
            PanelTbx1.Focus()
            e.Handled = True
            WorkTimer2.Enabled = True
            IncentivesControlLoad()

        End If
    End Sub

    Private Sub PanelTbx2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PanelTbx2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub PanelTbx3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PanelTbx3.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub


    Private Sub PanelTbx4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PanelTbx4.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub PanelBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PanelBtn1.Click

        If PanelTbx2.Text = "" Then
            MsgBox("Look For Personnel First", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

        Else

            If WorkFrm1.Visible = True Then
                GenConCode()
                PanelTbx3.Text = ProcessNum

            ElseIf WorkFrm2.Visible = True Then
                GenMutuIICode()
                PanelTbx3.Text = ProcessNum

            ElseIf WorkFrm3.Visible = True Then
                GenWalletCode()
                PanelTbx3.Text = ProcessNum


            ElseIf WorkFrm4.Visible = True Then
                GenPackingCode()
                PanelTbx3.Text = ProcessNum

            ElseIf WorkFrm5.Visible = True Then
                GenNewMiscCode()
                PanelTbx3.Text = ProcessNum

            ElseIf WorkFrm6.Visible = True Then
                GenMiscCode()
                PanelTbx3.Text = ProcessNum

            End If
            WorkTimer2.Enabled = True
        End If
    End Sub


    Private Sub PanelBtn4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PanelBtn4.Click
        If WorkSupFrame.Visible = False Then

            WorkSupFrame.Visible = True
        Else
            WorkSupFrame.Visible = False

        End If
    End Sub

    Private Sub WorkSupTbx1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WorkSupTbx1.TextChanged
        MasaKerjaCtrl()
    End Sub

    Private Sub InceChkBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InceChkBox1.CheckedChanged

        If InceChkBox1.Checked = True Then

            SQL = ""
            SQL = SQL & "Select * From 08_Standard_Table "
            SQL = SQL & "Where Original = ('" & "IncentiveLock" & "') "
            OpenTbl(ADb, Atbl36, SQL)
            If Not Atbl36.RecordCount <> 0 Then
                Atbl36.AddNew()
            End If

            Atbl36("Original").Value = "IncentiveLock"
            Atbl36("Standard_Wage").Value = "True"


            Atbl36.Update()

        ElseIf InceChkBox1.Checked = False Then

            SQL = ""
            SQL = SQL & "Select * From 08_Standard_Table "
            SQL = SQL & "Where Original = ('" & "IncentiveLock" & "') "
            OpenTbl(ADb, Atbl36, SQL)

            If Not Atbl36.RecordCount <> 0 Then
                Atbl36.AddNew()
            End If

            Atbl36("Original").Value = "IncentiveLock"
            Atbl36("Standard_Wage").Value = "False"

            Atbl36.Update()

        End If
    End Sub
#End Region

#Region "Conveyour"
    ' Conveyour Control Code

    Private Sub ConCmb_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ConCmb.KeyPress
        If e.KeyChar.ToString = Chr(Keys.Back) Then Exit Sub
        If Not InValid2.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If
    End Sub

    Private Sub ConTbx2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ConTbx2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub ConTbx3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ConTbx3.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub ConTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ConTbx1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InValid2.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If
    End Sub


    Private Sub ConTbx1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConTbx1.TextChanged
        If HolMod = 2 And YearMod >= 365 Then
            ConTbx2.Text = Format((Val(ConTbx1.Text) / Val(ConCmb.Text) * StandardsSalary) * 2, "0.")
        Else
            ConTbx2.Text = Format((Val(ConTbx1.Text) / Val(ConCmb.Text) * StandardsSalary), "0.")

        End If
    End Sub

    Private Sub ConCmb_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConCmb.SelectedIndexChanged
        If HolMod = 2 And YearMod >= 365 Then
            ConTbx2.Text = Format((Val(ConTbx1.Text) / Val(ConCmb.Text) * StandardsSalary) * 2, "0.")
        Else
            ConTbx2.Text = Format((Val(ConTbx1.Text) / Val(ConCmb.Text) * StandardsSalary), "0.")
        End If
        ConTbx1.Focus()
    End Sub

    Private Sub ConSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConSave.Click
        If PanelTbx3.Text = "" Then
            MsgBox("Click Add First", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        ElseIf ConTbx1.Text = "" Then
            MsgBox("Enter Required Pieces", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

        ElseIf ConTbx2.Text = "" Then
            MsgBox("Place Calculate First", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

        Else
            ConveyourSave()
            PanelTbx3.Text = ""
            ConTbx1.Text = ""

        End If
    End Sub

    Sub ConveyourSave()

        SQL = ""
        SQL = SQL & "Select * From 03_Conveyour_Table "
        SQL = SQL & "Where Process_ID = ('" & PanelTbx3.Text & "') "

        OpenTbl(ADb, Atb5, SQL)

        If Not Atb5.RecordCount <> 0 Then
            Atb5.AddNew()
        End If

        Atb5("Time").Value = TimeOfDay
        Atb5("Nik").Value = PanelTbx1.Text
        Atb5("Target").Value = ConCmb.Text
        Atb5("Pieces").Value = ConTbx1.Text
        Atb5("Salary").Value = ConTbx2.Text
        Atb5("Date").Value = WorkCalendar.SelectionStart
        Atb5("Process_ID").Value = PanelTbx3.Text
        Atb5.Update()

        PanelSaveLb1.ForeColor = Color.Green
        PanelSaveLb1.Text = "Saved: "
        PanelSaveLb2.ForeColor = Color.Green
        PanelSaveLb2.Text = "Conveyour Complete"

        Me.Refresh()

    End Sub

    Sub LoadConveyour()

        ConGrid.Rows.Clear()
        ConMainTot = 0

        SQL = ""
        SQL = SQL & "Select * from 03_Conveyour_Table "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & WorkCalendar.SelectionStart & "') "
        SQL = SQL & "Order by Nik"
        OpenTbl(ADb, DbTbl4, SQL)

        If DbTbl4.RecordCount > 0 Then

            DbTbl4.MoveFirst()
            Do While Not DbTbl4.EOF

                PiecesNest = DbTbl4("Pieces").Value
                TargetNest = DbTbl4("Target").Value
                SalaryNest = DbTbl4("Salary").Value
                TimeNest = DbTbl4("Time").Value


                ConGrid.Rows.Add(PiecesNest, TargetNest, SalaryNest, TimeNest)
                DbTbl4.MoveNext()
                ConMainTot = ConMainTot + SalaryNest

            Loop


        End If

        ConTbx3.Text = Format(ConMainTot, "###.00")
    End Sub

#End Region

#Region "Mutu II"

    ' Mutu II Control Code

    Private Sub MutuTbx3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MutuTbx3.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub MutuTbx4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MutuTbx4.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub MutuMask1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MutuMask1.KeyPress
        MutuMask1.Mask = "######-##"
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            MutuTbx1.Focus()
            e.Handled = True
        End If
    End Sub

    Private Sub MutuSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MutuSave.Click
        If PanelTbx3.Text = "" Then
            MsgBox("Click Add First", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        ElseIf MutuMask1.Text = "" Then
            MsgBox("Invalid Coupon Number", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        ElseIf MutuTbx2.Text = "" Then
            MsgBox("Enter Required Pieces", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        ElseIf MutuTbx3.Text = "" Then
            MsgBox("Please Calculate First", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        Else
            MutuIISave()
            PanelTbx3.Text = ""
            MutuTbx2.Text = ""
        End If
    End Sub

    Private Sub MutuTbx2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MutuTbx2.TextChanged
        If HolMod = 2 And YearMod >= 365 Then
            MutuTbx3.Text = Format((Val(MutuTbx2.Text) / Val(MutuTbx1.Text) * StandardsSalary) * 2, "0.")
        Else
            MutuTbx3.Text = Format((Val(MutuTbx2.Text) / Val(MutuTbx1.Text) * StandardsSalary), "0.")

        End If
    End Sub

    Private Sub MutuTbx1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MutuTbx1.TextChanged
        If HolMod = 2 And YearMod >= 365 Then
            MutuTbx3.Text = Format((Val(MutuTbx2.Text) / Val(MutuTbx1.Text) * StandardsSalary) * 2, "0.")
        Else
            MutuTbx3.Text = Format((Val(MutuTbx2.Text) / Val(MutuTbx1.Text) * StandardsSalary), "0.")

        End If
    End Sub

    Sub LoadMutuII()
        MutuGrid.Rows.Clear()
        MutuMainTot = 0

        SQL = ""
        SQL = SQL & "Select * from 04_MutuII_Table "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & WorkCalendar.SelectionStart & "') "
        SQL = SQL & "Order by Process_ID Desc "
        OpenTbl(ADb, DbTbl7, SQL)
        If DbTbl7.RecordCount > 0 Then

            DbTbl7.MoveFirst()
            Do While Not DbTbl7.EOF

                PiecesNest = DbTbl7("Pieces").Value
                TargetNest = DbTbl7("Target").Value
                SalaryNest = DbTbl7("Salary").Value
                TimeNest = DbTbl7("Time").Value


                MutuGrid.Rows.Add(PiecesNest, TargetNest, SalaryNest, TimeNest)
                DbTbl7.MoveNext()
                MutuMainTot = MutuMainTot + SalaryNest

            Loop
        End If
        MutuTbx4.Text = Format(MutuMainTot, "###.00")
    End Sub

    Sub MutuIISave()


        SQL = ""
        SQL = SQL & "Select * From 04_MutuII_Table "
        SQL = SQL & "Where Process_ID = ('" & PanelTbx3.Text & "') "
        OpenTbl(ADb, Atb6, SQL)

        If Not Atb6.RecordCount <> 0 Then
            Atb6.AddNew()
        End If

        Atb6("Time").Value = TimeOfDay
        Atb6("Nik").Value = PanelTbx1.Text
        Atb6("Target").Value = MutuTbx1.Text
        Atb6("Pieces").Value = MutuTbx2.Text
        Atb6("Salary").Value = MutuTbx3.Text
        Atb6("Coupon").Value = MutuMask1.Text
        Atb6("Date").Value = WorkCalendar.SelectionStart
        Atb6("Process_ID").Value = PanelTbx3.Text
        Atb6.Update()



        PanelSaveLb1.ForeColor = Color.Green
        PanelSaveLb1.Text = "Saved: "
        PanelSaveLb2.ForeColor = Color.Green
        PanelSaveLb2.Text = "MutuII Complete"

        Me.Refresh()
    End Sub

#End Region

#Region "Wallet"
    ' Wallet Control Code

    Private Sub WalletMask1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles WalletMask1.KeyPress
        WalletMask1.Mask = "######-##"
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            WalletCmb1.Focus()
            e.Handled = True
        End If
    End Sub

    Private Sub WalletCmb1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles WalletCmb1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InValid2.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If
    End Sub

    Private Sub WalletCmb1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WalletCmb1.SelectedIndexChanged
        If HolMod = 2 And YearMod >= 365 Then
            WalletTbx2.Text = Format((Val(WalletTbx1.Text) / Val(WalletCmb1.Text) * StandardsSalary) * 2, "0.")
        Else
            WalletTbx2.Text = Format((Val(WalletTbx1.Text) / Val(WalletCmb1.Text) * StandardsSalary), "0.")

        End If
        WalletTbx1.Focus()
    End Sub

    Private Sub WalletTbx2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles WalletTbx2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub WalletTbx3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles WalletTbx3.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Sub WalletSave()
        SQL = ""
        SQL = SQL & "Select * From 06_Wallet_Table "
        SQL = SQL & "Where Process_ID = ('" & PanelTbx3.Text & "') "

        OpenTbl(ADb, Atb7, SQL)

        If Not Atb7.RecordCount <> 0 Then
            Atb7.AddNew()
        End If

        Atb7("Time").Value = TimeOfDay
        Atb7("Nik").Value = PanelTbx1.Text
        Atb7("Target").Value = WalletCmb1.Text
        Atb7("Pieces").Value = WalletTbx1.Text
        Atb7("Salary").Value = WalletTbx2.Text
        Atb7("Coupon").Value = WalletMask1.Text
        Atb7("Date").Value = WorkCalendar.SelectionStart
        Atb7("Process_ID").Value = PanelTbx3.Text
        Atb7.Update()


        PanelSaveLb1.ForeColor = Color.Green
        PanelSaveLb1.Text = "Saved: "
        PanelSaveLb2.ForeColor = Color.Green
        PanelSaveLb2.Text = "Wallet Complete"

        Me.Refresh()
    End Sub

    Sub LoadWallet()
        WalletGrid.Rows.Clear()
        WalletMainTot = 0
        SQL = ""
        SQL = SQL & "Select * from 06_Wallet_Table "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & WorkCalendar.SelectionStart & "') "
        SQL = SQL & "Order by Process_ID Desc "
        OpenTbl(ADb, DbTbl8, SQL)
        If DbTbl8.RecordCount > 0 Then

            DbTbl8.MoveFirst()
            Do While Not DbTbl8.EOF

                PiecesNest = DbTbl8("Pieces").Value
                TargetNest = DbTbl8("Target").Value
                SalaryNest = DbTbl8("Salary").Value
                TimeNest = DbTbl8("Time").Value


                WalletGrid.Rows.Add(PiecesNest, TargetNest, SalaryNest, TimeNest)
                DbTbl8.MoveNext()
                WalletMainTot = WalletMainTot + SalaryNest

            Loop
        End If
        WalletTbx3.Text = Format(WalletMainTot, "###.00")
    End Sub

    Private Sub WalletTbx1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WalletTbx1.TextChanged
        If HolMod = 2 And YearMod >= 365 Then
            WalletTbx2.Text = Format((Val(WalletTbx1.Text) / Val(WalletCmb1.Text) * StandardsSalary) * 2, "0.")
        Else
            WalletTbx2.Text = Format((Val(WalletTbx1.Text) / Val(WalletCmb1.Text) * StandardsSalary), "0.")

        End If
    End Sub

    Private Sub WallSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WallSave.Click
        If PanelTbx3.Text = "" Then
            MsgBox("Click Add First", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        ElseIf WalletMask1.Text = "" Then
            MsgBox("Invalid Coupon Number", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        ElseIf WalletTbx1.Text = "" Then
            MsgBox("Enter Required Pieces", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        ElseIf WalletTbx2.Text = "" Then
            MsgBox("Please Calculate First", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        Else
            WalletSave()
            PanelTbx3.Text = ""
            WalletTbx2.Text = ""
        End If
    End Sub


#End Region

#Region "Packing"

    ' Packing Control Code

    Private Sub PackingCmb1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PackingCmb1.SelectedIndexChanged
        If HolMod = 2 And YearMod >= 365 Then
            PackingTbx2.Text = Format((Val(PackingTbx1.Text) / Val(PackingCmb1.Text) * StandardsSalary) * 2, "0.")
        Else
            PackingTbx2.Text = Format((Val(PackingTbx1.Text) / Val(PackingCmb1.Text) * StandardsSalary), "0.")
        End If
        PackingTbx1.Focus()
    End Sub

    Private Sub PackingCmb1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PackingCmb1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InValid2.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If
    End Sub

    Private Sub PackingMask1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PackingMask1.KeyPress
        PackingMask1.Mask = "######-##"
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            PackingCmb1.Focus()
            e.Handled = True
        End If
    End Sub

    Private Sub PackingTbx2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PackingTbx2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub PackingTbx3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PackingTbx3.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub PackingTbx1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PackingTbx1.TextChanged
        If HolMod = 2 And YearMod >= 365 Then
            PackingTbx2.Text = Format((Val(PackingTbx1.Text) / Val(PackingCmb1.Text) * StandardsSalary) * 2, "0.")
        Else
            PackingTbx2.Text = Format((Val(PackingTbx1.Text) / Val(PackingCmb1.Text) * StandardsSalary), "0.")

        End If
    End Sub

    Sub PackingSave()


        SQL = ""
        SQL = SQL & "Select * From 05_Packing_Table "
        SQL = SQL & "Where Process_ID = ('" & PanelTbx3.Text & "') "

        OpenTbl(ADb, Atb8, SQL)

        If Not Atb8.RecordCount <> 0 Then
            Atb8.AddNew()
        End If

        Atb8("Time").Value = TimeOfDay
        Atb8("Nik").Value = PanelTbx1.Text
        Atb8("Target").Value = PackingCmb1.Text
        Atb8("Carton").Value = PackingTbx1.Text
        Atb8("Salary").Value = PackingTbx2.Text
        Atb8("Coupon").Value = PackingMask1.Text
        Atb8("Date").Value = WorkCalendar.SelectionStart
        Atb8("Process_ID").Value = PanelTbx3.Text
        Atb8.Update()


        PanelSaveLb1.ForeColor = Color.Green
        PanelSaveLb1.Text = "Saved: "
        PanelSaveLb2.ForeColor = Color.Green
        PanelSaveLb2.Text = "Packing Complete"

        Me.Refresh()
    End Sub

    Sub LoadPacking()
        PackingGrid.Rows.Clear()
        PackingMainTot = 0
        SQL = ""
        SQL = SQL & "Select * from 05_Packing_Table "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & WorkCalendar.SelectionStart & "') "
        SQL = SQL & "Order by Process_ID Desc "
        OpenTbl(ADb, DbTbl9, SQL)
        If DbTbl9.RecordCount > 0 Then

            DbTbl9.MoveFirst()
            Do While Not DbTbl9.EOF

                CtnNest = DbTbl9("Carton").Value
                TargetNest = DbTbl9("Target").Value
                SalaryNest = DbTbl9("Salary").Value
                TimeNest = DbTbl9("Time").Value


                PackingGrid.Rows.Add(CtnNest, TargetNest, SalaryNest, TimeNest)
                DbTbl9.MoveNext()
                PackingMainTot = PackingMainTot + SalaryNest

            Loop
        End If
        PackingTbx3.Text = Format(PackingMainTot, "###.00")
    End Sub

    Private Sub PackSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PackSave.Click
        If PanelTbx3.Text = "" Then
            MsgBox("Click Add First", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        ElseIf PackingMask1.Text = "" Then
            MsgBox("Invalid Coupon Number", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        ElseIf PackingTbx1.Text = "" Then
            MsgBox("Enter Required Pieces", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        ElseIf PackingTbx2.Text = "" Then
            MsgBox("Please Calculate First", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        Else
            PackingSave()
            PanelTbx3.Text = ""
            PackingTbx1.Text = ""
        End If
    End Sub

#End Region

#Region "Sortasi"

    ' Sortasi Control Code

    Private Sub SortCmb1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles SortCmb1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InValid2.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If
    End Sub

    Private Sub SortCmb2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles SortCmb2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub SortMask1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles SortMask1.KeyPress
        SortMask1.Mask = "######-##"
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            SortMask2.Focus()
            e.Handled = True
        End If

    End Sub

    Private Sub SortMask2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles SortMask2.KeyPress
        SortMask2.Mask = "#,#"
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            SortCmb1.Focus()
            e.Handled = True
            GrControl = Val(SortMask2.Text / 10000)
            SortTbx1.Text = Format((Val(SortCmb1.Text) / Val(GrControl)) * Val(SortCmb2.Text), "###.00")
            SortTbx3.Text = Format((Val(SortTbx1.Text / 20000)) * (StandardsSalary), "###.00")
        End If
    End Sub

    Private Sub SortTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles SortTbx1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub


#End Region

#Region "GUI CONTROL 2"
    Private Sub SortTbx2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles SortTbx2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub SortTbx3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles SortTbx3.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub SortTbx4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles SortTbx4.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub SortSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SortSave.Click
        If PanelTbx3.Text = "" Then
            MsgBox("Click Add First", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        ElseIf SortMask1.Text = "" Then
            MsgBox("Invalid Coupon Number", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        ElseIf SortMask2.Text = "" Then
            MsgBox("Enter Required Gram", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        ElseIf SortTbx3.Text = "" Then
            MsgBox("Please Calculate First", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        ElseIf SortTbx1.Text = "" Then
            MsgBox("Please Calculate First", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        Else
            SortasiSave()
            PanelTbx3.Text = ""
            SortMask2.Text = ""

        End If

    End Sub

    Private Sub SortCmb1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SortCmb1.SelectedIndexChanged
        If SortMask2.Text = "" Then
            If WorkFrm5.Visible = True Then
                MsgBox("Please Input the Required Gram", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
                SortMask2.Focus()
            End If

        Else
            GrControl = Val(SortMask2.Text / 10000)
            SortTbx1.Text = Format((Val(SortCmb1.Text) / Val(GrControl)) * Val(SortCmb2.Text), "###.00")
            SortTbx3.Text = Format((Val(SortTbx1.Text / 20000)) * (StandardsSalary), "###.00")
            SortCmb2.Focus()
        End If
    End Sub

    Private Sub SortCmb2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SortCmb2.SelectedIndexChanged
        If SortMask2.Text = "" Then
            If WorkFrm5.Visible = True Then
                MsgBox("Please Input the Required Gram", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
                SortMask2.Focus()

            End If
        Else

            GrControl = Val(SortMask2.Text / 10000)
            SortTbx1.Text = Format((Val(SortCmb1.Text) / Val(GrControl)) * Val(SortCmb2.Text), "###.00")
            SortTbx3.Text = Format((Val(SortTbx1.Text / 20000)) * (StandardsSalary), "###.00")
            SortCmb2.Focus()

        End If
    End Sub

#End Region

    Sub SortasiSave()

        If HolMod = 2 And YearMod >= 365 Then

            SorHolTot = Format(Val(SortTbx3.Text) * 2, "N0")
        Else

            SorHolTot = Format(Val(SortTbx3.Text), "N0")
        End If

        SQL = ""
        SQL = SQL & "Select * From 21_NewMiscellaneous_Table "
        SQL = SQL & "Where Process_ID = ('" & PanelTbx3.Text & "') "

        OpenTbl(ADb, Atbl28, SQL)

        If Not Atbl28.RecordCount <> 0 Then
            Atbl28.AddNew()
        End If

        Atbl28("Time").Value = TimeOfDay
        Atbl28("Nik").Value = PanelTbx1.Text
        Atbl28("NoKg").Value = SortCmb1.Text
        Atbl28("NoBag").Value = SortCmb2.Text
        Atbl28("NoGr").Value = SortMask2.Text
        Atbl28("Pieces").Value = SortTbx1.Text
        Atbl28("Coupon").Value = SortMask1.Text
        Atbl28("Salary").Value = SorHolTot
        Atbl28("Date").Value = WorkCalendar.SelectionStart
        Atbl28("Process_ID").Value = PanelTbx3.Text
        Atbl28.Update()

        PanelSaveLb1.ForeColor = Color.Green
        PanelSaveLb1.Text = "Saved: "
        PanelSaveLb2.ForeColor = Color.Green
        PanelSaveLb2.Text = "Sortasi Complete"

        Me.Refresh()
        WorkTimer2.Enabled = True
    End Sub

    Sub LoadSortasi()

        SortGrid.Rows.Clear()
        SortMainTot = 0
        SortPiecesTot = 0

        SQL = ""
        SQL = SQL & "Select * from 21_NewMiscellaneous_Table "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & WorkCalendar.SelectionStart & "') "
        SQL = SQL & "Order by Process_ID Desc "

        OpenTbl(ADb, Dbtb37, SQL)
        If Dbtb37.RecordCount > 0 Then

            Do While Not Dbtb37.EOF

                NoBagNest = Dbtb37("NoBag").Value
                NoKgNest = Dbtb37("NoKg").Value
                NoGramNest = Dbtb37("NoGr").Value
                SalaryNest = Dbtb37("Salary").Value
                TimeNest = Dbtb37("Time").Value
                PiecesNest = Dbtb37("Pieces").Value

                SortGrid.Rows.Add(NoKgNest, NoGramNest, NoBagNest, PiecesNest, TimeNest, SalaryNest)
                Dbtb37.MoveNext()
                SortMainTot = SortMainTot + SalaryNest
                SortPiecesTot = SortPiecesTot + PiecesNest

            Loop
        End If

        SortTbx2.Text = Format(SortPiecesTot, "###.00")
        SpecialSortasi()

    End Sub

    Sub SpecialSortasi()

        If tr3 >= 61 And SortTbx2.Text >= 20001 Then
            SortSpecialTot = Val(((SortTbx2.Text - 20000) / 20000) * (1.2 * StandardsSalary) + StandardsSalary)
            SortTbx4.Text = Format(SortSpecialTot, "###.00")

        ElseIf tr3 <= 60 And SortTbx2.Text <= 15000 Then
            SortTbx4.Text = SubsidiSalary

        ElseIf tr3 <= 60 And SortTbx2.Text >= 20000 Then

            SortSpecialTot = Val(((SortTbx2.Text - 20000) / 20000) * (1.2 * StandardsSalary) + StandardsSalary)
            SortTbx4.Text = Format(SortSpecialTot, "###.00")

        ElseIf tr3 <= 60 And SortTbx2.Text >= 15001 Then
            SortTbx4.Text = Format(SortMainTot, "###.00")

        Else
            SortTbx4.Text = Format(SortMainTot, "###.00")
        End If

    End Sub

#Region "Misc"

    ' Miscellaneous Control Code

    Private Sub MiscSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MiscSave.Click
        If PanelTbx3.Text = "" Then
            MsgBox("Click Add First", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        ElseIf MiscMask1.Text = "" Then
            MsgBox("Enter Total Value ", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")
        Else
            MiscellaneousSave()
            MiscMask1.Text = ""
            PanelTbx3.Text = ""
        End If
    End Sub

    Private Sub MiscTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MiscTbx1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Sub MiscellaneousSave()

        If HolMod = 2 And YearMod >= 365 Then

            MiscHolTot = Format(Val(MiscMask1.Text) * 2, "N0")
        Else
            MiscHolTot = Format(Val(MiscMask1.Text), "N0")
        End If



        SQL = ""
        SQL = SQL & "Select * From 19_Miscellaneous_Table "
        SQL = SQL & "Where Process_ID = ('" & PanelTbx3.Text & "') "

        OpenTbl(ADb, Atbl24, SQL)

        If Not Atbl24.RecordCount <> 0 Then
            Atbl24.AddNew()
        End If

        Atbl24("Time").Value = TimeOfDay
        Atbl24("Nik").Value = PanelTbx1.Text
        Atbl24("Salary").Value = MiscHolTot
        Atbl24("Date").Value = WorkCalendar.SelectionStart
        Atbl24("Process_ID").Value = PanelTbx3.Text
        Atbl24.Update()


        PanelSaveLb1.ForeColor = Color.Green
        PanelSaveLb1.Text = "Saved: "
        PanelSaveLb2.ForeColor = Color.Green
        PanelSaveLb2.Text = "Miscellaneous Complete"

        Me.Refresh()

    End Sub

    Sub LoadMisc()

        MiscGrid.Rows.Clear()
        MiscMainTot = 0

        SQL = ""
        SQL = SQL & "Select * from 19_Miscellaneous_Table "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & WorkCalendar.SelectionStart & "') "
        SQL = SQL & "Order by Process_ID Desc "

        OpenTbl(ADb, Dbtb34, SQL)
        If Dbtb34.RecordCount > 0 Then
            Do While Not Dbtb34.EOF


                SalaryNest = Dbtb34("Salary").Value
                TimeNest = Dbtb34("Time").Value


                MiscGrid.Rows.Add(TimeNest, SalaryNest)
                Dbtb34.MoveNext()

                MiscMainTot = MiscMainTot + SalaryNest
            Loop
        End If
        MiscTbx1.Text = Format(MiscMainTot, "###.00")

    End Sub


#End Region

    Private Sub PanelBtn2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PanelBtn2.Click

        If WorkSupTbx8.Text = "" And WorkSupTbx6.Text = "" And WorkSupTbx7.Text = "" Then

            MsgBox("Please Look For (Periode) Data in Date SetUp", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

        End If

        If WorkFrm1.Visible = True Then
            If ConTbx3.Text < 1 Then
                MsgBox("Proceed the Calculation (Data Input on Converyour before you click me :) ")
            ElseIf PanelTbx2.Text = "" Then
                MsgBox("Kindly Look the Personnel to process his/her Work ")
            Else
                ConveyourSalary()
            End If

        ElseIf WorkFrm2.Visible = True Then
            If MutuTbx4.Text < 1 Then
                MsgBox("Proceed the Calculation (Data Input on MutuII before you click me :) ")
            ElseIf PanelTbx2.Text = "" Then
                MsgBox("Kindly Look the Personnel to process his/her Work ")
            Else
                MutuIISalary()
            End If

        ElseIf WorkFrm3.Visible = True Then
            If WalletTbx3.Text < 1 Then
                MsgBox("Proceed the Calculation (Data Input on Wallet before you click me :) ")
            ElseIf PanelTbx2.Text = "" Then
                MsgBox("Kindly Look the Personnel to process his/her Work ")
            Else
                WalletSalary()
            End If

        ElseIf WorkFrm4.Visible = True Then
            If PackingTbx3.Text < 1 Then
                MsgBox("Proceed the Calculation (Data Input on Packing before you click me :) ")
            ElseIf PanelTbx2.Text = "" Then
                MsgBox("Kindly Look the Personnel to process his/her Work ")
            Else
                PackingSalary()
            End If

        ElseIf WorkFrm5.Visible = True Then
            If SortTbx4.Text < 1 Then
                MsgBox("Proceed the Calculation (Data Input on Sortasi before you click me :) ")
            ElseIf PanelTbx2.Text = "" Then
                MsgBox("Kindly Look the Personnel to process his/her Work ")
            Else
                SortasiSalary()
            End If

        ElseIf WorkFrm6.Visible = True Then
            If MiscTbx1.Text < 1 Then
                MsgBox("Proceed the Calculation (Data Input on Miscellaneous before you click me :) ")
            ElseIf PanelTbx2.Text = "" Then
                MsgBox("Kindly Look the Personnel to process his/her Work ")
            Else
                MiscellaneousSalary()
            End If

        ElseIf WorkFrm7.Visible = True Then
            If OverAllTot < 1 And Not PanelTbx2.Text = "" Then
                MsgBox("Proceed the Calculation (Data Input on any deparment) before you click me :) ")
            ElseIf PanelTbx2.Text = "" Then
                MsgBox("Kindly Look the Personnel to process his/her Work ")
            Else
                OverAllSaver()
            End If

        End If

    End Sub

#Region "Incentives"

    Sub IncentivesControlSave()


        IncentiveCount = WorkSupTbx9.Text + 1

        SQL = ""
        SQL = SQL & "Select * From 12_Incentives_Ctrl "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And MonthPeriodeRange = ('" & WorkSupTbx10.Text & "') "
        OpenTbl(ADb, Atbl31, SQL)
        If Not Atbl31.RecordCount <> 0 Then
            Atbl31.AddNew()
        End If

        Atbl31("MonthPeriodeRange").Value = WorkSupTbx10.Text
        Atbl31("Nik").Value = PanelTbx1.Text
        Atbl31("Count").Value = IncentiveCount

        Atbl31.Update()


    End Sub

    Sub IncentivesControlLoad()
        'SQL = ""
        'SQL = SQL & "Select * From 12_Incentives_Ctrl "
        'SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        'SQL = SQL & "And MonthPeriodeRange = ('" & WorkSupTbx10.Text & "') "
        'OpenTbl(ADb, Atbl32, SQL)

        'If Atbl32.RecordCount <> 0 Then
        '    WorkSupTbx9.Text = Atbl32("Count").Value

        'Else
        '    WorkSupTbx9.Text = "0"


        'End If

    End Sub


    Sub IncentivesControlRange()
        'SQL = ""
        'SQL = SQL & "Select * From 22_Incentives_Setup "
        'SQL = SQL & "Where Actives = ('" & "Yes" & "') "
        'OpenTbl(ADb, Atbl33, SQL)

        'If Atbl33.RecordCount > 0 Then

        '    WorkSupTbx10.Text = Atbl33("MonthPeriodeRange").Value

        'End If
        'Me.Refresh()
    End Sub

    Sub IncentivesEnabler()

        'SQL = ""
        'SQL = SQL & "Select * From 08_Standard_Table "
        'SQL = SQL & "Where Original = ('" & "IncentiveLock" & "') "
        'OpenTbl(ADb, Atbl35, SQL)

        'If Atbl35.RecordCount > 0 Then

        '    IncentiveLock = Atbl35("Standard_Wage").Value

        'End If
        'Me.Refresh()


        'If IncentiveLock = "True" Then

        '    InceChkBox1.Checked = True

        'End If
    End Sub



#End Region

#Region "Division Salary"

    Sub ConveyourSalary()


        SQL = ""
        SQL = SQL & "Select * From 13_Conveyour_Salary "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & WorkCalendar.SelectionStart & "') "

        OpenTbl(ADb, Atbl30, SQL)

        If Not Atbl30.RecordCount <> 0 Then
            Atbl30.AddNew()
        End If

        Atbl30("Date").Value = WorkCalendar.SelectionStart
        Atbl30("Nik").Value = PanelTbx1.Text
        Atbl30("Salary").Value = ConTbx3.Text

        Atbl30.Update()
        PanelSaveLb1.ForeColor = Color.Green
        PanelSaveLb1.Text = "Saved: "
        PanelSaveLb2.ForeColor = Color.Green
        PanelSaveLb2.Text = "Conveyour Salary "
        Me.Refresh()

    End Sub
    Sub MutuIISalary()

        SQL = ""
        SQL = SQL & "Select * From 14_MutuII_Salary "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & WorkCalendar.SelectionStart & "') "
        OpenTbl(ADb, Atbl30, SQL)

        If Not Atbl30.RecordCount <> 0 Then
            Atbl30.AddNew()
        End If

        Atbl30("Date").Value = WorkCalendar.SelectionStart
        Atbl30("Nik").Value = PanelTbx1.Text
        Atbl30("Salary").Value = MutuTbx4.Text

        Atbl30.Update()
        PanelSaveLb1.ForeColor = Color.Green
        PanelSaveLb1.Text = "Saved: "
        PanelSaveLb2.ForeColor = Color.Green
        PanelSaveLb2.Text = "MutuII Salary "
        Me.Refresh()

    End Sub
    Sub WalletSalary()


        SQL = ""
        SQL = SQL & "Select * From 15_Wallet_Salary "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & WorkCalendar.SelectionStart & "') "
        OpenTbl(ADb, Atbl30, SQL)

        If Not Atbl30.RecordCount <> 0 Then
            Atbl30.AddNew()
        End If

        Atbl30("Date").Value = WorkCalendar.SelectionStart
        Atbl30("Nik").Value = PanelTbx1.Text
        Atbl30("Salary").Value = WalletTbx3.Text

        Atbl30.Update()
        PanelSaveLb1.ForeColor = Color.Green
        PanelSaveLb1.Text = "Saved: "
        PanelSaveLb2.ForeColor = Color.Green
        PanelSaveLb2.Text = "Wallet Salary "
        Me.Refresh()

    End Sub
    Sub PackingSalary()

        SQL = ""
        SQL = SQL & "Select * From 16_Packing_Salary "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & WorkCalendar.SelectionStart & "') "
        OpenTbl(ADb, Atbl30, SQL)

        If Not Atbl30.RecordCount <> 0 Then
            Atbl30.AddNew()
        End If

        Atbl30("Date").Value = WorkCalendar.SelectionStart
        Atbl30("Nik").Value = PanelTbx1.Text
        Atbl30("Salary").Value = PackingTbx3.Text

        Atbl30.Update()
        PanelSaveLb1.ForeColor = Color.Green
        PanelSaveLb1.Text = "Saved: "
        PanelSaveLb2.ForeColor = Color.Green
        PanelSaveLb2.Text = "Packing Salary "
        Me.Refresh()



    End Sub
    Sub MiscellaneousSalary()

        SQL = ""
        SQL = SQL & "Select * From 20_Miscellaneous_Salary "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & WorkCalendar.SelectionStart & "') "
        SQL = SQL & "And TypeCtrl = ('" & "Old" & "') "
        OpenTbl(ADb, Atbl30, SQL)

        If Not Atbl30.RecordCount <> 0 Then
            Atbl30.AddNew()
        End If



        Atbl30("Date").Value = WorkCalendar.SelectionStart
        Atbl30("Nik").Value = PanelTbx1.Text
        Atbl30("Salary").Value = MiscTbx1.Text
        Atbl30("TypeCtrl").Value = "Old"

        Atbl30.Update()
        PanelSaveLb1.ForeColor = Color.Green
        PanelSaveLb1.Text = "Saved: "
        PanelSaveLb2.ForeColor = Color.Green
        PanelSaveLb2.Text = "Miscellaneous Salary "
        Me.Refresh()

    End Sub
    Sub SortasiSalary()


        SQL = ""
        SQL = SQL & "Select * From 20_Miscellaneous_Salary "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & WorkCalendar.SelectionStart & "') "
        SQL = SQL & "And TypeCtrl = ('" & "New" & "') "
        OpenTbl(ADb, Atbl30, SQL)

        If Not Atbl30.RecordCount <> 0 Then
            Atbl30.AddNew()
        End If

        Atbl30("Date").Value = WorkCalendar.SelectionStart
        Atbl30("Nik").Value = PanelTbx1.Text
        Atbl30("Salary").Value = SortTbx4.Text
        Atbl30("TypeCtrl").Value = "New"

        Atbl30.Update()

        PanelSaveLb1.ForeColor = Color.Green
        PanelSaveLb1.Text = "Saved: "
        PanelSaveLb2.ForeColor = Color.Green
        PanelSaveLb2.Text = "Sortasi Salary "
        Me.Refresh()

    End Sub
    Sub LoadDayPeriodeCtrl()

        SQL = ""
        SQL = SQL & "Select * from Periode_CounterTable "
        SQL = SQL & "Where Date = ('" & WorkCalendar.SelectionStart & "') "
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
    End Sub
    Sub OverAllSaver()
        SQL = ""
        SQL = SQL & "Select * From SalarySync1_Table "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Periode = ('" & PeriodeCtrl & "') "
        SQL = SQL & "And PeriodeRange = ('" & PeriodeMonthCtrl & "') "
        OpenTbl(CBb, Ctbl1, SQL)


        If Not Ctbl1.RecordCount <> 0 Then
            Ctbl1.AddNew()
        End If

        Ctbl1("Nik").Value = PanelTbx1.Text
        Ctbl1("Name").Value = PanelTbx2.Text
        Ctbl1("Periode").Value = PeriodeCtrl
        Ctbl1("PeriodeRange").Value = PeriodeMonthCtrl
        Ctbl1("Pay").Value = WorkSupTbx4.Text
        Ctbl1("AstekVal").Value = WorkSupTbx3.Text


        If WorkSupTbx8.Text = "1" Then

            Ctbl1("Salary1").Value = OverAllTot
            Ctbl1("Date1").Value = WorkCalendar.SelectionStart

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "2" Then

            Ctbl1("Salary2").Value = OverAllTot
            Ctbl1("Date2").Value = WorkCalendar.SelectionStart

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "3" Then

            Ctbl1("Salary3").Value = OverAllTot
            Ctbl1("Date3").Value = WorkCalendar.SelectionStart
            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "4" Then

            Ctbl1("Salary4").Value = OverAllTot
            Ctbl1("Date4").Value = WorkCalendar.SelectionStart

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "5" Then

            Ctbl1("Salary5").Value = OverAllTot
            Ctbl1("Date5").Value = WorkCalendar.SelectionStart

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "6" Then

            Ctbl1("Salary6").Value = OverAllTot
            Ctbl1("Date6").Value = WorkCalendar.SelectionStart

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "7" Then

            Ctbl1("Salary7").Value = OverAllTot
            Ctbl1("Date7").Value = WorkCalendar.SelectionStart

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "8" Then

            Ctbl1("Salary8").Value = OverAllTot
            Ctbl1("Date8").Value = WorkCalendar.SelectionStart

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "9" Then

            Ctbl1("Salary9").Value = OverAllTot
            Ctbl1("Date9").Value = WorkCalendar.SelectionStart

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "10" Then

            Ctbl1("Salary10").Value = OverAllTot
            Ctbl1("Date10").Value = WorkCalendar.SelectionStart

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "11" Then

            Ctbl1("Salary11").Value = OverAllTot
            Ctbl1("Date11").Value = WorkCalendar.SelectionStart

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "12" Then

            Ctbl1("Salary12").Value = OverAllTot
            Ctbl1("Date12").Value = WorkCalendar.SelectionStart

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "13" Then

            Ctbl1("Salary13").Value = OverAllTot
            Ctbl1("Date13").Value = WorkCalendar.SelectionStart

            Ctbl1.Update()
            Me.Refresh()


        ElseIf WorkSupTbx8.Text = "14" Then

            Ctbl1("Salary14").Value = OverAllTot
            Ctbl1("Date14").Value = WorkCalendar.SelectionStart

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "15" Then

            Ctbl1("Salary15").Value = OverAllTot
            Ctbl1("Date15").Value = WorkCalendar.SelectionStart

            Ctbl1.Update()
            Me.Refresh()

        ElseIf WorkSupTbx8.Text = "16" Then

            Ctbl1("Salary16").Value = OverAllTot
            Ctbl1("Date16").Value = WorkCalendar.SelectionStart

            Ctbl1.Update()
            Me.Refresh()


        End If


        If InceChkBox1.Checked = True Then
            IncentivesControlSave()
        End If

        MsgBox("Save!", vbInformation)
    End Sub
    Private Sub WorkCalendar_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles WorkCalendar.DateChanged
        ErrorOverAll()
        LoadConveyour()
        LoadMutuII()
        LoadWallet()
        LoadPacking()
        LoadSortasi()
        LoadMisc()
        WorkTimer2.Enabled = True
    End Sub
    Sub OALNest()
        ConCD = 0
        MutuIICD = 0
        PackingCD = 0
        WalletCD = 0
        NewMiscellaneousCD = 0
        MiscellaneousCD = 0
        OverAllTot = 0

    End Sub
    Sub OverAllLoad()

        OALNest()
        OverGrid.Rows.Clear()


        SQL = ""
        SQL = SQL & "Select * from 13_Conveyour_Salary "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & WorkCalendar.SelectionStart & "') "
        OpenTbl(ADb, Dbtb21, SQL)
        If Dbtb21.RecordCount > 0 Then
            ConCD = Dbtb21("Salary").Value
        Else
            ConCD = 0

        End If

        SQL = ""
        SQL = SQL & "Select * from 14_MutuII_Salary "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & WorkCalendar.SelectionStart & "') "
        OpenTbl(ADb, Dbtb22, SQL)
        If Dbtb22.RecordCount > 0 Then
            MutuIICD = Dbtb22("Salary").Value

        Else
            MutuIICD = 0

        End If

        SQL = ""
        SQL = SQL & "Select * from 15_Wallet_Salary "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & WorkCalendar.SelectionStart & "') "
        OpenTbl(ADb, Dbtb23, SQL)
        If Dbtb23.RecordCount > 0 Then
            WalletCD = Dbtb23("Salary").Value

        Else
            WalletCD = 0

        End If


        SQL = ""
        SQL = SQL & "Select * from 16_Packing_Salary "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & WorkCalendar.SelectionStart & "') "
        OpenTbl(ADb, Dbtb24, SQL)
        If Dbtb24.RecordCount > 0 Then
            PackingCD = Dbtb24("Salary").Value

        Else
            PackingCD = 0

        End If

        SQL = ""
        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & WorkCalendar.SelectionStart & "') "
        SQL = SQL & "And TypeCtrl = ('" & "Old" & "') "
        OpenTbl(ADb, Dbtb35, SQL)
        If Dbtb35.RecordCount > 0 Then
            MiscellaneousCD = Dbtb35("Salary").Value

        Else
            MiscellaneousCD = 0

        End If

        SQL = ""
        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
        SQL = SQL & "Where Nik = ('" & PanelTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & WorkCalendar.SelectionStart & "') "
        SQL = SQL & "And TypeCtrl = ('" & "New" & "') "
        OpenTbl(ADb, Dbtb37, SQL)
        If Dbtb37.RecordCount > 0 Then
            NewMiscellaneousCD = Dbtb37("Salary").Value

        Else
            NewMiscellaneousCD = 0


        End If

        OverAllTot = ConCD + MutuIICD + WalletCD + PackingCD + NewMiscellaneousCD + MiscellaneousCD
        OverGrid.Rows.Add(Format(WorkCalendar.SelectionStart, "dd/MM/yyyy"), ConCD, MutuIICD, WalletCD, PackingCD, NewMiscellaneousCD, MiscellaneousCD, OverAllTot)

    End Sub
    Sub ErrorPeriode()
        On Error GoTo Err
        LoadDayPeriodeCtrl()
        Exit Sub
Err:
    End Sub

    Sub ErrorOverAll()
        On Error GoTo Err
        OverAllLoad()
        Exit Sub
Err:
    End Sub

#End Region

    Private Sub MiscMask1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MiscMask1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InValid2.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If
    End Sub


End Class

