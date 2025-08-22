Option Explicit On

Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel

Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.ComponentModel


Public Class ReportBlock
    Dim ExcelAP As Excel.Application
    Dim ExcelWB As Excel.Workbook
    Dim ExcelWS As Excel.Worksheet

    Dim ExcelName As String
    Dim CmbDater As String
    Dim CmbDater2 As String
    Dim CmbDater3 As String
    Dim GetSalaryTot As String
    Dim GetSalaryTotFormat As String
    Dim RecordCounting As Integer = 0
    Dim GetSalaryMaxTotal As String
    Dim GetInceRange As String
    Dim GetInceDay As String
    Dim GetInceCount As String
    Dim GetInceSum As String
    Dim GetInceSumAgain As String
    Dim GetInceDif As String
    Dim TotRoundoff As Int64
    Dim TotRound1 As Int64
    Dim TotRound2 As String
    Dim UpSlip1 As String
    Dim UpSlip2 As String
    Dim UpSlip3 As String
    Dim GetSal1Up As String
    Dim JabatanDat As String
    Dim CGaji As String
    Dim NoRekFinder As String
    Dim GajiCtrlAl As String = 1

    Dim CallIncentives As String = Nothing
    Dim EmpAddress As String

    ' For PPh21 variables
    Dim UpNik As String
    Dim UpName As String
    Dim KTP As String
    Dim NPWP As String
    Dim FGaji As String
    Dim LGaji As String
    Dim UpAstek As String
    Dim UpPay As String
    Dim UpIncentif As String


    Private Sub Report1Block_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LoadDB()
        LoadDB2()
        LoadDBPPh21()
 
    End Sub

    Sub ConIncentives()

        If CBRep2.Checked = True Then

            SQL = ""
            SQL = SQL & "Select * From Periode_CounterTable "
            SQL = SQL & "Where PeriodeRange = ('" & PerCmb1.Text & "') "
            SQL = SQL & "And Nik_Val = ('" & NewGSal(0) & "') "
            SQL = SQL & "And Name_Nik = ('" & NewGSal(1) & "') "
            OpenTbl(CBb, Ctbl60, SQL)

            If Ctbl60.RecordCount > 0 Then

                CallIncentives = IIf(IsDBNull(Ctbl24("IncenValue").Value), "", Ctbl24("IncenValue").Value)

            End If

        End If

    End Sub

#Region "Formmating Entity / Nest to 0"

    Sub GajiCtrlMove()

        If CBRep3.Checked = True Then

            GajiCtrlAl = 1 + (GajiCtrlSalary / 100)

        Else

            GajiCtrlAl = 1

        End If

    End Sub

    Sub GetSalaryNest()

        For i = 0 To 17
            NewGSal(i) = Nothing
        Next

        GetAstek = 0
        GetSalaryTot = 0
        GetSalaryTotFormat = 0
        GetSalaryMaxTotal = 0

        For i = 0 To 15

            NewTotSalRow(i) = Nothing
        Next
        TotAstek = 0

    End Sub


    Sub SalTbxNest()
        SalTbx1.Text = ""
        SalTbx2.Text = ""
        SalTbx3.Text = ""
        SalTbx4.Text = ""
        SalTbx5.Text = ""
        SalTbx6.Text = ""
        SalTbx7.Text = ""
        SalTbx8.Text = ""
        SalTbx9.Text = ""
        SalTbx10.Text = ""
        SalTbx11.Text = ""
        SalTbx12.Text = ""
        SalTbx13.Text = ""
        SalTbx14.Text = ""
        SalTbx15.Text = ""
        SalTbx16.Text = ""

        SalTbx21.Text = ""
        SalTbx18.Text = ""
        SalTbx19.Text = ""
        RecordCounting = 0

    End Sub
    Sub SalaryRowTotalMode()

        SalTbx1.Text = Format(Val(NewTotSalRow(0)), "#,#.")
        SalTbx2.Text = Format(Val(NewTotSalRow(1)), "#,#.")
        SalTbx3.Text = Format(Val(NewTotSalRow(2)), "#,#.")
        SalTbx4.Text = Format(Val(NewTotSalRow(3)), "#,#.")
        SalTbx5.Text = Format(Val(NewTotSalRow(4)), "#,#.")
        SalTbx6.Text = Format(Val(NewTotSalRow(5)), "#,#.")
        SalTbx7.Text = Format(Val(NewTotSalRow(6)), "#,#.")
        SalTbx8.Text = Format(Val(NewTotSalRow(7)), "#,#.")
        SalTbx9.Text = Format(Val(NewTotSalRow(8)), "#,#.")
        SalTbx10.Text = Format(Val(NewTotSalRow(9)), "#,#.")
        SalTbx11.Text = Format(Val(NewTotSalRow(10)), "#,#.")
        SalTbx12.Text = Format(Val(NewTotSalRow(11)), "#,#.")
        SalTbx13.Text = Format(Val(NewTotSalRow(12)), "#,#.")
        SalTbx14.Text = Format(Val(NewTotSalRow(13)), "#,#.")
        SalTbx15.Text = Format(Val(NewTotSalRow(14)), "#,#.")
        SalTbx16.Text = Format(Val(NewTotSalRow(15)), "#,#.")
        SalTbx18.Text = Format(Val(GetSalaryMaxTotal), "#,#.")
        SalTbx19.Text = Format(Val(TotAstek), "#,#.")


    End Sub

    Sub GridFormatNum()

        GajiCtrlMove()
        ConIncentives()

        ' Salary 1
        If NewGSal(2) = "" Or NewGSal(2) = "0" Then
            NewFormT(0) = ""
        Else
            NewFormT(0) = Val(NewGSal(2) * GajiCtrlAl).ToString("N0", CustomtoUS)
        End If

        ' Salary 2
        If NewGSal(3) = "" Or NewGSal(3) = "0" Then
            NewFormT(1) = ""
        Else
            NewFormT(1) = Val(NewGSal(3) * GajiCtrlAl).ToString("N0", CustomtoUS)
        End If

        'Salary 3
        If NewGSal(4) = "" Or NewGSal(4) = "0" Then
            NewFormT(2) = ""
        Else
            NewFormT(2) = Val(NewGSal(4) * GajiCtrlAl).ToString("N0", CustomtoUS)
        End If

        ' Salary 4
        If NewGSal(5) = "" Or NewGSal(5) = "0" Then
            NewFormT(3) = ""
        Else
            NewFormT(3) = Val(NewGSal(5) * GajiCtrlAl).ToString("N0", CustomtoUS)
        End If

        ' Salary 5
        If NewGSal(6) = "" Or NewGSal(6) = "0" Then
            NewFormT(4) = ""
        Else
            NewFormT(4) = Val(NewGSal(6) * GajiCtrlAl).ToString("N0", CustomtoUS)
        End If

        ' Salary 6
        If NewGSal(7) = "" Or NewGSal(7) = "0" Then
            NewFormT(5) = ""
        Else
            NewFormT(5) = Val(NewGSal(7) * GajiCtrlAl).ToString("N0", CustomtoUS)
        End If

        'Salary 7
        If NewGSal(8) = "" Or NewGSal(8) = "0" Then
            NewFormT(6) = ""
        Else
            NewFormT(6) = Val(NewGSal(8) * GajiCtrlAl).ToString("N0", CustomtoUS)
        End If

        ' Salary 8
        If NewGSal(9) = "" Or NewGSal(9) = "0" Then
            NewFormT(7) = ""
        Else
            NewFormT(7) = Val(NewGSal(9) * GajiCtrlAl).ToString("N0", CustomtoUS)
        End If

        'Salary 9
        If NewGSal(10) = "" Or NewGSal(10) = "0" Then
            NewFormT(8) = ""
        Else
            NewFormT(8) = Val(NewGSal(10) * GajiCtrlAl).ToString("N0", CustomtoUS)
        End If

        ' Salary 10
        If NewGSal(11) = "" Or NewGSal(11) = "0" Then
            NewFormT(9) = ""
        Else
            NewFormT(9) = Val(NewGSal(11) * GajiCtrlAl).ToString("N0", CustomtoUS)
        End If

        'Salary 11
        If NewGSal(12) = "" Or NewGSal(12) = "0" Then
            NewFormT(10) = ""
        Else
            NewFormT(10) = Val(NewGSal(12) * GajiCtrlAl).ToString("N0", CustomtoUS)
        End If

        ' Salary 12
        If NewGSal(13) = "" Or NewGSal(13) = "0" Then
            NewFormT(11) = ""
        Else
            NewFormT(11) = Val(NewGSal(13) * GajiCtrlAl).ToString("N0", CustomtoUS)
        End If

        ' Salary 13
        If NewGSal(14) = "" Or NewGSal(14) = "0" Then
            NewFormT(12) = ""
        Else
            NewFormT(12) = Val(NewGSal(14) * GajiCtrlAl).ToString("N0", CustomtoUS)
        End If

        ' Salary 14
        If NewGSal(15) = "" Or NewGSal(15) = "0" Then
            NewFormT(13) = ""
        Else
            NewFormT(13) = Val(NewGSal(15) * GajiCtrlAl).ToString("N0", CustomtoUS)
        End If

        ' Salary 15
        If NewGSal(16) = "" Or NewGSal(16) = "0" Then
            NewFormT(14) = ""
        Else
            NewFormT(14) = Val(NewGSal(16) * GajiCtrlAl).ToString("N0", CustomtoUS)
        End If

        ' Salary 16
        If NewGSal(17) = "" Or NewGSal(17) = "0" Then
            NewFormT(15) = ""
        Else
            NewFormT(15) = Val(NewGSal(17) * GajiCtrlAl).ToString("N0", CustomtoUS)
        End If

        'Astek Value

        If GetAstek2 = "" Or GetAstek2 = "0" Or GetAstek2 = "0.00" Then

            GetAstekFormat = ""

        Else

            GetAstekFormat = Format(Val(GetAstek2), "N0")

        End If

        RoundProcess()
        TotRound2 = Val(TotRound1).ToString("N0", CustomtoUS)

        NewGSal(1) = NewGSal(1).Replace("?", "'")
        GetSal1Up = NewGSal(0).ToUpper
        PerGrid01.Rows.Add(GetSal1Up, NewGSal(1), NoRekFinder, NewFormT(0), NewFormT(1), NewFormT(2), NewFormT(3), NewFormT(4), NewFormT(5), NewFormT(6), NewFormT(7), NewFormT(8), NewFormT(9), NewFormT(10), NewFormT(11), NewFormT(12), NewFormT(13), NewFormT(14), NewFormT(15), "", "", GetAstekFormat, CallIncentives, "", "", GetPotlain, GetSalaryTot, TotRound2)
    End Sub

    Sub LetsCount() ' Counting the Record in Grid

        RecordCounting += 1
        SalTbx21.Text = RecordCounting.ToString

    End Sub

#End Region

#Region "Control Setting Code"


    Private Sub PerCmb1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub PerCmb2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PerCmb2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub PerCmb3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PerCmb3.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub PerBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PerBtn1.Click

        If PerCmb1.Text = "" Or PerCmb2.Text = "" Then
            MsgBox("Please Select the Periode Range or Periode Entity for data request", MsgBoxStyle.Exclamation)
        Else

            Dim NewMDIChild As New LoadingBlock()
            LoadingBlock.MdiParent = MainMenu

            LoadHeadDater()
            LoadEmpSalSep()
            RecordCounting = 0
        End If

    End Sub


    Private Sub PerBtn4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PerBtn4.Click

        PerGrid01.Rows.Clear()
        SalTbxNest()
        GetSalaryNest()
        GetSalaryMaxTotal = "0"
        Me.Refresh()

    End Sub

    Private Sub PerBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PerBtn2.Click

        If PerCmb1.Text = "" Then

            MsgBox("Please Select the Periode")

        Else
            SaveFileLink()
            'ExportExcel()
            OnClickTheWorker()
        End If

    End Sub

    Private Sub PerBtn3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PerBtn3.Click
        UPandSlipFil()

        If PPhPanel1.Visible = True Then

            PPhPanel2.Visible = False
            PPhPanel1.Visible = False

        Else

            PPhPanel2.Visible = True
            PPhPanel1.Visible = True

        End If
    End Sub

#End Region

#Region "For PPH21 Up and Load"

    Sub UPandSlipFil()
        PerCmb4.Items.Clear()
        UpSlip1 = Now.ToString("yyyy")
        UpSlip2 = UpSlip1 - 1
        UpSlip3 = UpSlip2 + 1

        With PerCmb4

            .Items.Add("Dec " + UpSlip2 + " - " + "Jan " + Format(Now, "yyyy"))
            .Items.Add("Jan " + Format(Now, "yyyy") + " - " + "Feb " + Format(Now, "yyyy"))
            .Items.Add("Feb " + Format(Now, "yyyy") + " - " + "Mar " + Format(Now, "yyyy"))
            .Items.Add("Mar " + Format(Now, "yyyy") + " - " + "Apr " + Format(Now, "yyyy"))
            .Items.Add("Apr " + Format(Now, "yyyy") + " - " + "May " + Format(Now, "yyyy"))
            .Items.Add("May " + Format(Now, "yyyy") + " - " + "Jun " + Format(Now, "yyyy"))
            .Items.Add("Jun " + Format(Now, "yyyy") + " - " + "Jul " + Format(Now, "yyyy"))
            .Items.Add("Jul " + Format(Now, "yyyy") + " - " + "Aug " + Format(Now, "yyyy"))
            .Items.Add("Aug " + Format(Now, "yyyy") + " - " + "Sep " + Format(Now, "yyyy"))
            .Items.Add("Sep " + Format(Now, "yyyy") + " - " + "Oct " + Format(Now, "yyyy"))
            .Items.Add("Oct " + Format(Now, "yyyy") + " - " + "Nov " + Format(Now, "yyyy"))
            .Items.Add("Nov " + Format(Now, "yyyy") + " - " + "Dec " + Format(Now, "yyyy"))
            .Items.Add("Dec " + Format(Now, "yyyy") + " - " + "Jan " + UpSlip3)

        End With
    End Sub

    Sub DataNester()

        KTP = Nothing
        NPWP = Nothing
        UpPay = Nothing
        UpAstek = Nothing
        JabatanDat = Nothing
        EmpAddress = Nothing

    End Sub

    Sub UpPPhValue1()

        For a = 0 To PerGrid01.Rows.Count - 1

            UpNik = PerGrid01(0, a).Value
            UpName = PerGrid01(1, a).Value
            FGaji = PerGrid01(26, a).Value
            UpIncentif = PerGrid01(22, a).Value

            SQL = ""
            SQL = SQL & "Select * From Emp_PPHTable "
            SQL = SQL & "Where PeriodeGajian = ('" & PerCmb4.Text & "') "
            SQL = SQL & "And Nik = ('" & UpNik & "') "
            OpenTbl(PPhDB, PPhTb5, SQL)

            If Not PPhTb5.RecordCount <> 0 Then

                PPhTb5.AddNew()

            End If

            NomorLookerSolo()
            PPhTb5("PeriodeGajian").Value = PerCmb4.Text
            PPhTb5("Periode").Value = PerCmb2.Text
            PPhTb5("PeriodeRange").Value = PerCmb1.Text
            PPhTb5("Nik").Value = UpNik
            PPhTb5("Name").Value = UpName
            PPhTb5("KTP").Value = KTP
            PPhTb5("NPWP").Value = NPWP
            PPhTb5("MainSalary1").Value = FGaji
            PPhTb5("Astek").Value = UpAstek
            PPhTb5("Pay").Value = UpPay
            PPhTb5("Incentif").Value = UpIncentif
            PPhTb5("EmAdd").Value = EmpAddress

            PPhTb5.Update()

            DataNester()

        Next

        MsgBox("Data for PPH21 Gaji 1 is now Completed")
        PPhPanel2.Visible = False
        PPhPanel1.Visible = False

    End Sub

    Sub UpPPhValue2()

        For a = 0 To PerGrid01.Rows.Count - 1

            UpNik = PerGrid01(0, a).Value
            UpName = PerGrid01(1, a).Value
            LGaji = PerGrid01(26, a).Value
            UpIncentif = PerGrid01(22, a).Value

            SQL = ""
            SQL = SQL & "Select * From Emp_PPHTable "
            SQL = SQL & "Where PeriodeGajian = ('" & PerCmb4.Text & "') "
            SQL = SQL & "And Nik = ('" & UpNik & "') "
            OpenTbl(PPhDB, PPhTb5, SQL)

            If Not PPhTb5.RecordCount <> 0 Then

                PPhTb5.AddNew()

            End If

            NomorLookerSolo()
            PPhTb5("PeriodeGajian").Value = PerCmb4.Text
            PPhTb5("Periode").Value = PerCmb2.Text
            PPhTb5("PeriodeRange").Value = PerCmb1.Text
            PPhTb5("Nik").Value = UpNik
            PPhTb5("Name").Value = UpName
            PPhTb5("KTP").Value = KTP
            PPhTb5("NPWP").Value = NPWP
            PPhTb5("MainSalary2").Value = LGaji
            PPhTb5("Pay").Value = UpPay
            PPhTb5("Incentif").Value = UpIncentif
            PPhTb5("EmAdd").Value = EmpAddress
            PPhTb5.Update()
            DataNester()

        Next

        MsgBox("Data for PPH21 Gaji 2 is now Completed")
        PPhPanel2.Visible = False
        PPhPanel1.Visible = False

    End Sub

    Sub UpPPhValue3()

        For a = 0 To PerGrid01.Rows.Count - 1

            UpNik = PerGrid01(0, a).Value
            UpName = PerGrid01(1, a).Value
            LGaji = PerGrid01(26, a).Value
            UpIncentif = PerGrid01(22, a).Value

            SQL = ""
            SQL = SQL & "Select * From Emp_PPHTable "
            SQL = SQL & "Where PeriodeGajian = ('" & PerCmb4.Text & "') "
            SQL = SQL & "And Nik = ('" & UpNik & "') "
            OpenTbl(PPhDB, PPhTb5, SQL)

            If Not PPhTb5.RecordCount <> 0 Then

                PPhTb5.AddNew()

            End If

            NomorLookerSolo()
            PPhTb5("PeriodeGajian").Value = PerCmb4.Text
            PPhTb5("Periode").Value = PerCmb2.Text
            PPhTb5("PeriodeRange").Value = PerCmb1.Text
            PPhTb5("Nik").Value = UpNik
            PPhTb5("Name").Value = UpName
            PPhTb5("KTP").Value = KTP
            PPhTb5("NPWP").Value = NPWP
            PPhTb5("MainSalary3").Value = LGaji
            PPhTb5("Pay").Value = UpPay
            PPhTb5("Incentif").Value = UpIncentif
            PPhTb5("EmAdd").Value = EmpAddress
            PPhTb5.Update()
            DataNester()

        Next

        MsgBox("Data for PPH21 Gaji EX is now Completed")
        PPhPanel2.Visible = False
        PPhPanel1.Visible = False

    End Sub

    Sub NomorLookerSolo()

        SQL = ""
        SQL = SQL & "Select * From 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & UpNik & " ') "
        OpenTbl(ADb, Atb2, SQL)
        If Atb2.RecordCount > 0 Then

            KTP = IIf(IsDBNull(Atb2("NKTP").Value), "", Atb2("NKTP").Value)
            NPWP = IIf(IsDBNull(Atb2("NPWP").Value), "", Atb2("NPWP").Value)
            UpPay = IIf(IsDBNull(Atb2("Pay").Value), "", Atb2("Pay").Value)
            UpAstek = IIf(IsDBNull(Atb2("Jamsostek").Value), "", Atb2("Jamsostek").Value)
            JabatanDat = IIf(IsDBNull(Atb2("JabData").Value), "", Atb2("JabData").Value)
            EmpAddress = IIf(IsDBNull(Atb2("Alamat").Value), "", Atb2("Alamat").Value)

        End If

    End Sub
    'Sub AddressLooker()

    '    SQL = ""
    '    SQL = SQL & "Select * From Emp_Table001 "
    '    SQL = SQL & "Where Nik = ('" & UpNik & " ')"
    '    OpenTbl(PPhDB, PPhTb9, SQL)
    '    If PPhTb9.RecordCount > 0 Then
    '        EmpAddress = IIf(IsDBNull(PPhTb9("Alamat").Value), "", PPhTb9("Alamat").Value)
    '    End If


    'End Sub

#End Region

#Region "For Coupon"

    Sub CouponUp()

        For a = 0 To PerGrid01.Rows.Count - 1

            UpNik = PerGrid01(0, a).Value
            UpName = PerGrid01(1, a).Value
            CGaji = PerGrid01(27, a).Value

            SQL = ""
            SQL = SQL & "Select * From Emp_Coupon "
            SQL = SQL & "Where PeriodeRange = ('" & PerCmb1.Text & "') "
            SQL = SQL & "And Nik = ('" & UpNik & "') "
            OpenTbl(PPhDB, PPhTb5, SQL)

            If Not PPhTb5.RecordCount <> 0 Then

                PPhTb5.AddNew()

            End If

            NomorLookerSolo()
            PPhTb5("PeriodeRange").Value = PerCmb1.Text
            PPhTb5("Periode").Value = PerCmb2.Text
            PPhTb5("Nik").Value = UpNik
            PPhTb5("Name").Value = UpName
            PPhTb5("Gaji").Value = CGaji
            PPhTb5("Jab").Value = JabatanDat
            PPhTb5("Tanggal").Value = RepDatePick3.Text
            PPhTb5("PeriodeDate").Value = RepDatePick1.Text + " - " + RepDatePick2.Text

            PPhTb5.Update()

        Next

        MsgBox("Data for Coupon is now Completed")
        PPhPanel2.Visible = False
        PPhPanel1.Visible = False

    End Sub

    Sub CouponDelete()

        SQL = ""
        SQL = SQL & "Select * From Emp_Coupon "
        SQL = SQL & "Where Periode = ('" & "Periode I" & "') "
        SQL = SQL & "Or Periode= ('" & "Periode II" & "') "
        OpenTbl(PPhDB, PPhTb7, SQL)

        If PPhTb7.RecordCount > 0 Then
            PPhTb7.MoveFirst()
            Do While Not PPhTb7.EOF

                PPhTb7.Delete()
                PPhTb7.Update()
                PPhTb7.MoveNext()

            Loop

        End If

    End Sub

#End Region

#Region "Loading Data Code"

    Sub LoadHeadDater()

        SQL = ""
        SQL = SQL & "Select * From DateCounter2Table "
        SQL = SQL & "Where Periode = ('" & PerCmb2.Text & "') "
        SQL = SQL & "And PeriodeRange = ('" & PerCmb1.Text & "') "
        OpenTbl(CBb, Ctbl24, SQL)

        If Ctbl24.RecordCount > 0 Then

            NewSald(0) = IIf(IsDBNull(Ctbl24("Date1").Value), "", Ctbl24("Date1").Value)
            NewSald(1) = IIf(IsDBNull(Ctbl24("Date2").Value), "", Ctbl24("Date2").Value)
            NewSald(2) = IIf(IsDBNull(Ctbl24("Date3").Value), "", Ctbl24("Date3").Value)
            NewSald(3) = IIf(IsDBNull(Ctbl24("Date4").Value), "", Ctbl24("Date4").Value)
            NewSald(4) = IIf(IsDBNull(Ctbl24("Date5").Value), "", Ctbl24("Date5").Value)
            NewSald(5) = IIf(IsDBNull(Ctbl24("Date6").Value), "", Ctbl24("Date6").Value)
            NewSald(6) = IIf(IsDBNull(Ctbl24("Date7").Value), "", Ctbl24("Date7").Value)
            NewSald(7) = IIf(IsDBNull(Ctbl24("Date8").Value), "", Ctbl24("Date8").Value)
            NewSald(8) = IIf(IsDBNull(Ctbl24("Date9").Value), "", Ctbl24("Date9").Value)
            NewSald(9) = IIf(IsDBNull(Ctbl24("Date10").Value), "", Ctbl24("Date10").Value)
            NewSald(10) = IIf(IsDBNull(Ctbl24("Date11").Value), "", Ctbl24("Date11").Value)
            NewSald(11) = IIf(IsDBNull(Ctbl24("Date12").Value), "", Ctbl24("Date12").Value)
            NewSald(12) = IIf(IsDBNull(Ctbl24("Date13").Value), "", Ctbl24("Date13").Value)
            NewSald(13) = IIf(IsDBNull(Ctbl24("Date14").Value), "", Ctbl24("Date14").Value)
            NewSald(14) = IIf(IsDBNull(Ctbl24("Date15").Value), "", Ctbl24("Date15").Value)
            NewSald(15) = IIf(IsDBNull(Ctbl24("Date16").Value), "", Ctbl24("Date16").Value)

        Else

            For i = 0 To 15
                NewSald(i) = Nothing
            Next

            'Sald1 = ""
            'Sald2 = ""
            'Sald3 = ""
            'Sald4 = ""
            'Sald5 = ""
            'Sald6 = ""
            'Sald7 = ""
            'Sald8 = ""
            'Sald9 = ""
            'Sald10 = ""
            'Sald11 = ""
            'Sald12 = ""
            'Sald13 = ""
            'Sald14 = ""
            'Sald15 = ""
            'Sald16 = ""

        End If

        PerGrid01.Columns(3).HeaderText = NewSald(0)
        PerGrid01.Columns(4).HeaderText = NewSald(1)
        PerGrid01.Columns(5).HeaderText = NewSald(2)
        PerGrid01.Columns(6).HeaderText = NewSald(3)
        PerGrid01.Columns(7).HeaderText = NewSald(4)
        PerGrid01.Columns(8).HeaderText = NewSald(5)
        PerGrid01.Columns(9).HeaderText = NewSald(6)
        PerGrid01.Columns(10).HeaderText = NewSald(7)
        PerGrid01.Columns(11).HeaderText = NewSald(8)
        PerGrid01.Columns(12).HeaderText = NewSald(9)
        PerGrid01.Columns(13).HeaderText = NewSald(10)
        PerGrid01.Columns(14).HeaderText = NewSald(11)
        PerGrid01.Columns(15).HeaderText = NewSald(12)
        PerGrid01.Columns(16).HeaderText = NewSald(13)
        PerGrid01.Columns(17).HeaderText = NewSald(14)
        PerGrid01.Columns(18).HeaderText = NewSald(15)

        SalLbl1.Text = NewSald(0)
        SalLbl2.Text = NewSald(1)
        SalLbl3.Text = NewSald(2)
        SalLbl4.Text = NewSald(3)
        SalLbl5.Text = NewSald(4)
        SalLbl6.Text = NewSald(5)
        SalLbl7.Text = NewSald(6)
        SalLbl8.Text = NewSald(7)
        SalLbl9.Text = NewSald(8)
        SalLbl10.Text = NewSald(9)
        SalLbl11.Text = NewSald(10)
        SalLbl12.Text = NewSald(11)
        SalLbl13.Text = NewSald(12)
        SalLbl14.Text = NewSald(13)
        SalLbl15.Text = NewSald(14)
        SalLbl16.Text = NewSald(15)

        Me.Refresh()

    End Sub

    Sub LoadEmpSalSep() ' Seperator for Periode on factoring the required Entity
        If PerCmb2.Text = "Periode I" Then
            LoadEmpSalaryCtrl()

        ElseIf PerCmb2.Text = "Periode II" Then
            LoadEmpSalaryCtrl2()

        ElseIf PerCmb2.Text = "Periode II-Ex" Then
            LoadEmpSalaryCtrl()

        End If

    End Sub

    Sub LoadEmpSalaryCtrl()
        GetSalaryNest()

        If PerCmb3.Text = "All" Then

            PerGrid01.Rows.Clear()

            SQL = ""
            SQL = SQL & "Select * from SalarySync1_Table "
            SQL = SQL & "Where Periode = ('" & PerCmb2.Text & "') "
            SQL = SQL & "And PeriodeRange = ('" & PerCmb1.Text & "') "
            SQL = SQL & "Order by Nik "
            OpenTbl(CBb, Ctbl25, SQL)
            If Ctbl25.RecordCount > 0 Then
                Ctbl25.MoveFirst()
                Do While Not Ctbl25.EOF

                    NewGSal(0) = IIf(IsDBNull(Ctbl25("Nik").Value), "", Ctbl25("Nik").Value)
                    NewGSal(1) = IIf(IsDBNull(Ctbl25("Name").Value), "", Ctbl25("Name").Value)
                    NewGSal(2) = IIf(IsDBNull(Ctbl25("Salary1").Value), "", Ctbl25("Salary1").Value)
                    NewGSal(3) = IIf(IsDBNull(Ctbl25("Salary2").Value), "", Ctbl25("Salary2").Value)
                    NewGSal(4) = IIf(IsDBNull(Ctbl25("Salary3").Value), "", Ctbl25("Salary3").Value)
                    NewGSal(5) = IIf(IsDBNull(Ctbl25("Salary4").Value), "", Ctbl25("Salary4").Value)
                    NewGSal(6) = IIf(IsDBNull(Ctbl25("Salary5").Value), "", Ctbl25("Salary5").Value)
                    NewGSal(7) = IIf(IsDBNull(Ctbl25("Salary6").Value), "", Ctbl25("Salary6").Value)
                    NewGSal(8) = IIf(IsDBNull(Ctbl25("Salary7").Value), "", Ctbl25("Salary7").Value)
                    NewGSal(9) = IIf(IsDBNull(Ctbl25("Salary8").Value), "", Ctbl25("Salary8").Value)
                    NewGSal(10) = IIf(IsDBNull(Ctbl25("Salary9").Value), "", Ctbl25("Salary9").Value)
                    NewGSal(11) = IIf(IsDBNull(Ctbl25("Salary10").Value), "", Ctbl25("Salary10").Value)
                    NewGSal(12) = IIf(IsDBNull(Ctbl25("Salary11").Value), "", Ctbl25("Salary11").Value)
                    NewGSal(13) = IIf(IsDBNull(Ctbl25("Salary12").Value), "", Ctbl25("Salary12").Value)
                    NewGSal(14) = IIf(IsDBNull(Ctbl25("Salary13").Value), "", Ctbl25("Salary13").Value)
                    NewGSal(15) = IIf(IsDBNull(Ctbl25("Salary14").Value), "", Ctbl25("Salary14").Value)
                    NewGSal(16) = IIf(IsDBNull(Ctbl25("Salary15").Value), "", Ctbl25("Salary15").Value)
                    NewGSal(17) = IIf(IsDBNull(Ctbl25("Salary16").Value), "", Ctbl25("Salary16").Value)
                    NoRekFinder = IIf(IsDBNull(Ctbl25("PNoRek").Value), "", Ctbl25("PNoRek").Value)
                    GetPotlain = IIf(IsDBNull(Ctbl25("PotLain").Value), "", Ctbl25("PotLain").Value)

                    GetSalaryTot = Format(Val(NewGSal(2)) + Val(NewGSal(3)) + Val(NewGSal(4)) + Val(NewGSal(5)) + Val(NewGSal(6)) + Val(NewGSal(7)) + Val(NewGSal(8)) + Val(NewGSal(9)) + Val(NewGSal(10)) + Val(NewGSal(11)) + Val(NewGSal(12)) + Val(NewGSal(13)) + Val(NewGSal(14)) + Val(NewGSal(15)) + Val(NewGSal(16)) + Val(NewGSal(17)), "#.")
                    GetSalaryTotFormat = Format(Val(GetSalaryTot) - Val(GetAstek), "#.")
                    GetSalaryTot = Format(Val(GetSalaryTot) - Val(GetAstek), "N0")
                    GetSalaryMaxTotal = Val(GetSalaryMaxTotal) + Val(GetSalaryTotFormat)
                    GridFormatNum()
                    LetsCount()

                    NewTotSalRow(0) = Format(Val(NewTotSalRow(0)) + Val(NewGSal(2)), "#.")
                    NewTotSalRow(1) = Format(Val(NewTotSalRow(1)) + Val(NewGSal(3)), "#.")
                    NewTotSalRow(2) = Format(Val(NewTotSalRow(2)) + Val(NewGSal(4)), "#.")
                    NewTotSalRow(3) = Format(Val(NewTotSalRow(3)) + Val(NewGSal(5)), "#.")
                    NewTotSalRow(4) = Format(Val(NewTotSalRow(4)) + Val(NewGSal(6)), "#.")
                    NewTotSalRow(5) = Format(Val(NewTotSalRow(5)) + Val(NewGSal(7)), "#.")
                    NewTotSalRow(6) = Format(Val(NewTotSalRow(6)) + Val(NewGSal(8)), "#.")
                    NewTotSalRow(7) = Format(Val(NewTotSalRow(7)) + Val(NewGSal(9)), "#.")
                    NewTotSalRow(8) = Format(Val(NewTotSalRow(8)) + Val(NewGSal(10)), "#.")
                    NewTotSalRow(9) = Format(Val(NewTotSalRow(9)) + Val(NewGSal(11)), "#.")
                    NewTotSalRow(10) = Format(Val(NewTotSalRow(10)) + Val(NewGSal(12)), "#.")
                    NewTotSalRow(11) = Format(Val(NewTotSalRow(11)) + Val(NewGSal(13)), "#.")
                    NewTotSalRow(12) = Format(Val(NewTotSalRow(12)) + Val(NewGSal(14)), "#.")
                    NewTotSalRow(13) = Format(Val(NewTotSalRow(13)) + Val(NewGSal(15)), "#.")
                    NewTotSalRow(14) = Format(Val(NewTotSalRow(14)) + Val(NewGSal(16)), "#.")
                    NewTotSalRow(15) = Format(Val(NewTotSalRow(15)) + Val(NewGSal(17)), "#.")
                    NoRekFinder = IIf(IsDBNull(Ctbl25("PNoRek").Value), "", Ctbl25("PNoRek").Value)

                    SalaryRowTotalMode()

                    Ctbl25.MoveNext()

                Loop
                MsgBox("DONE", vbInformation, "Codex")
            End If

        ElseIf PerCmb3.Text = "CASH" Then

            PerGrid01.Rows.Clear()

            SQL = ""
            SQL = SQL & "Select * from SalarySync1_Table "
            SQL = SQL & "Where Periode = ('" & PerCmb2.Text & "') "
            SQL = SQL & "And PeriodeRange = ('" & PerCmb1.Text & "') "
            SQL = SQL & "And Pay = ('" & "CASH" & "') "
            SQL = SQL & "Order by Nik "
            OpenTbl(CBb, Ctbl25, SQL)
            If Ctbl25.RecordCount > 0 Then
                Ctbl25.MoveFirst()
                Do While Not Ctbl25.EOF

                    NewGSal(0) = IIf(IsDBNull(Ctbl25("Nik").Value), "", Ctbl25("Nik").Value)
                    NewGSal(1) = IIf(IsDBNull(Ctbl25("Name").Value), "", Ctbl25("Name").Value)
                    NewGSal(2) = IIf(IsDBNull(Ctbl25("Salary1").Value), "", Ctbl25("Salary1").Value)
                    NewGSal(3) = IIf(IsDBNull(Ctbl25("Salary2").Value), "", Ctbl25("Salary2").Value)
                    NewGSal(4) = IIf(IsDBNull(Ctbl25("Salary3").Value), "", Ctbl25("Salary3").Value)
                    NewGSal(5) = IIf(IsDBNull(Ctbl25("Salary4").Value), "", Ctbl25("Salary4").Value)
                    NewGSal(6) = IIf(IsDBNull(Ctbl25("Salary5").Value), "", Ctbl25("Salary5").Value)
                    NewGSal(7) = IIf(IsDBNull(Ctbl25("Salary6").Value), "", Ctbl25("Salary6").Value)
                    NewGSal(8) = IIf(IsDBNull(Ctbl25("Salary7").Value), "", Ctbl25("Salary7").Value)
                    NewGSal(9) = IIf(IsDBNull(Ctbl25("Salary8").Value), "", Ctbl25("Salary8").Value)
                    NewGSal(10) = IIf(IsDBNull(Ctbl25("Salary9").Value), "", Ctbl25("Salary9").Value)
                    NewGSal(11) = IIf(IsDBNull(Ctbl25("Salary10").Value), "", Ctbl25("Salary10").Value)
                    NewGSal(12) = IIf(IsDBNull(Ctbl25("Salary11").Value), "", Ctbl25("Salary11").Value)
                    NewGSal(13) = IIf(IsDBNull(Ctbl25("Salary12").Value), "", Ctbl25("Salary12").Value)
                    NewGSal(14) = IIf(IsDBNull(Ctbl25("Salary13").Value), "", Ctbl25("Salary13").Value)
                    NewGSal(15) = IIf(IsDBNull(Ctbl25("Salary14").Value), "", Ctbl25("Salary14").Value)
                    NewGSal(16) = IIf(IsDBNull(Ctbl25("Salary15").Value), "", Ctbl25("Salary15").Value)
                    NewGSal(17) = IIf(IsDBNull(Ctbl25("Salary16").Value), "", Ctbl25("Salary16").Value)
                    GetPotlain = IIf(IsDBNull(Ctbl25("PotLain").Value), "", Ctbl25("PotLain").Value)

                    GetSalaryTot = Format(Val(NewGSal(2)) + Val(NewGSal(3)) + Val(NewGSal(4)) + Val(NewGSal(5)) + Val(NewGSal(6)) + Val(NewGSal(7)) + Val(NewGSal(8)) + Val(NewGSal(9)) + Val(NewGSal(10)) + Val(NewGSal(11)) + Val(NewGSal(12)) + Val(NewGSal(13)) + Val(NewGSal(14)) + Val(NewGSal(15)) + Val(NewGSal(16)) + Val(NewGSal(17)), "#.")
                    GetSalaryTotFormat = Format(Val(GetSalaryTot) - Val(GetAstek), "#.")
                    GetSalaryTot = Format(Val(GetSalaryTot) - Val(GetAstek), "N0")
                    GetSalaryMaxTotal = Val(GetSalaryMaxTotal) + Val(GetSalaryTotFormat)
                    GridFormatNum()
                    LetsCount()

                    NewTotSalRow(0) = Format(Val(NewTotSalRow(0)) + Val(NewGSal(2)), "#.")
                    NewTotSalRow(1) = Format(Val(NewTotSalRow(1)) + Val(NewGSal(3)), "#.")
                    NewTotSalRow(2) = Format(Val(NewTotSalRow(2)) + Val(NewGSal(4)), "#.")
                    NewTotSalRow(3) = Format(Val(NewTotSalRow(3)) + Val(NewGSal(5)), "#.")
                    NewTotSalRow(4) = Format(Val(NewTotSalRow(4)) + Val(NewGSal(6)), "#.")
                    NewTotSalRow(5) = Format(Val(NewTotSalRow(5)) + Val(NewGSal(7)), "#.")
                    NewTotSalRow(6) = Format(Val(NewTotSalRow(6)) + Val(NewGSal(8)), "#.")
                    NewTotSalRow(7) = Format(Val(NewTotSalRow(7)) + Val(NewGSal(9)), "#.")
                    NewTotSalRow(8) = Format(Val(NewTotSalRow(8)) + Val(NewGSal(10)), "#.")
                    NewTotSalRow(9) = Format(Val(NewTotSalRow(9)) + Val(NewGSal(11)), "#.")
                    NewTotSalRow(10) = Format(Val(NewTotSalRow(10)) + Val(NewGSal(12)), "#.")
                    NewTotSalRow(11) = Format(Val(NewTotSalRow(11)) + Val(NewGSal(13)), "#.")
                    NewTotSalRow(12) = Format(Val(NewTotSalRow(12)) + Val(NewGSal(14)), "#.")
                    NewTotSalRow(13) = Format(Val(NewTotSalRow(13)) + Val(NewGSal(15)), "#.")
                    NewTotSalRow(14) = Format(Val(NewTotSalRow(14)) + Val(NewGSal(16)), "#.")
                    NewTotSalRow(15) = Format(Val(NewTotSalRow(15)) + Val(NewGSal(17)), "#.")
                    NoRekFinder = IIf(IsDBNull(Ctbl25("PNoRek").Value), "", Ctbl25("PNoRek").Value)

                    SalaryRowTotalMode()

                    Ctbl25.MoveNext()

                Loop
                MsgBox("DONE", vbInformation, "Codex")

            End If

        ElseIf PerCmb3.Text = "BTN" Then

            PerGrid01.Rows.Clear()

            SQL = ""
            SQL = SQL & "Select * from SalarySync1_Table "
            SQL = SQL & "Where Periode = ('" & PerCmb2.Text & "') "
            SQL = SQL & "And PeriodeRange = ('" & PerCmb1.Text & "') "
            SQL = SQL & "And Pay = ('" & "BTN" & "') "
            SQL = SQL & "Order by Nik "
            OpenTbl(CBb, Ctbl25, SQL)
            If Ctbl25.RecordCount > 0 Then
                Ctbl25.MoveFirst()
                Do While Not Ctbl25.EOF

                    NewGSal(0) = IIf(IsDBNull(Ctbl25("Nik").Value), "", Ctbl25("Nik").Value)
                    NewGSal(1) = IIf(IsDBNull(Ctbl25("Name").Value), "", Ctbl25("Name").Value)
                    NewGSal(2) = IIf(IsDBNull(Ctbl25("Salary1").Value), "", Ctbl25("Salary1").Value)
                    NewGSal(3) = IIf(IsDBNull(Ctbl25("Salary2").Value), "", Ctbl25("Salary2").Value)
                    NewGSal(4) = IIf(IsDBNull(Ctbl25("Salary3").Value), "", Ctbl25("Salary3").Value)
                    NewGSal(5) = IIf(IsDBNull(Ctbl25("Salary4").Value), "", Ctbl25("Salary4").Value)
                    NewGSal(6) = IIf(IsDBNull(Ctbl25("Salary5").Value), "", Ctbl25("Salary5").Value)
                    NewGSal(7) = IIf(IsDBNull(Ctbl25("Salary6").Value), "", Ctbl25("Salary6").Value)
                    NewGSal(8) = IIf(IsDBNull(Ctbl25("Salary7").Value), "", Ctbl25("Salary7").Value)
                    NewGSal(9) = IIf(IsDBNull(Ctbl25("Salary8").Value), "", Ctbl25("Salary8").Value)
                    NewGSal(10) = IIf(IsDBNull(Ctbl25("Salary9").Value), "", Ctbl25("Salary9").Value)
                    NewGSal(11) = IIf(IsDBNull(Ctbl25("Salary10").Value), "", Ctbl25("Salary10").Value)
                    NewGSal(12) = IIf(IsDBNull(Ctbl25("Salary11").Value), "", Ctbl25("Salary11").Value)
                    NewGSal(13) = IIf(IsDBNull(Ctbl25("Salary12").Value), "", Ctbl25("Salary12").Value)
                    NewGSal(14) = IIf(IsDBNull(Ctbl25("Salary13").Value), "", Ctbl25("Salary13").Value)
                    NewGSal(15) = IIf(IsDBNull(Ctbl25("Salary14").Value), "", Ctbl25("Salary14").Value)
                    NewGSal(16) = IIf(IsDBNull(Ctbl25("Salary15").Value), "", Ctbl25("Salary15").Value)
                    NewGSal(17) = IIf(IsDBNull(Ctbl25("Salary16").Value), "", Ctbl25("Salary16").Value)
                    NoRekFinder = IIf(IsDBNull(Ctbl25("PNoRek").Value), "", Ctbl25("PNoRek").Value)
                    GetPotlain = IIf(IsDBNull(Ctbl25("PotLain").Value), "", Ctbl25("PotLain").Value)

                    GetSalaryTot = Format(Val(NewGSal(2)) + Val(NewGSal(3)) + Val(NewGSal(4)) + Val(NewGSal(5)) + Val(NewGSal(6)) + Val(NewGSal(7)) + Val(NewGSal(8)) + Val(NewGSal(9)) + Val(NewGSal(10)) + Val(NewGSal(11)) + Val(NewGSal(12)) + Val(NewGSal(13)) + Val(NewGSal(14)) + Val(NewGSal(15)) + Val(NewGSal(16)) + Val(NewGSal(17)), "#.")
                    GetSalaryTotFormat = Format(Val(GetSalaryTot) - Val(GetAstek), "#.")
                    GetSalaryTot = Format(Val(GetSalaryTot) - Val(GetAstek), "N0")
                    GetSalaryMaxTotal = Val(GetSalaryMaxTotal) + Val(GetSalaryTotFormat)
                    GridFormatNum()
                    LetsCount()

                    NewTotSalRow(0) = Format(Val(NewTotSalRow(0)) + Val(NewGSal(2)), "#.")
                    NewTotSalRow(1) = Format(Val(NewTotSalRow(1)) + Val(NewGSal(3)), "#.")
                    NewTotSalRow(2) = Format(Val(NewTotSalRow(2)) + Val(NewGSal(4)), "#.")
                    NewTotSalRow(3) = Format(Val(NewTotSalRow(3)) + Val(NewGSal(5)), "#.")
                    NewTotSalRow(4) = Format(Val(NewTotSalRow(4)) + Val(NewGSal(6)), "#.")
                    NewTotSalRow(5) = Format(Val(NewTotSalRow(5)) + Val(NewGSal(7)), "#.")
                    NewTotSalRow(6) = Format(Val(NewTotSalRow(6)) + Val(NewGSal(8)), "#.")
                    NewTotSalRow(7) = Format(Val(NewTotSalRow(7)) + Val(NewGSal(9)), "#.")
                    NewTotSalRow(8) = Format(Val(NewTotSalRow(8)) + Val(NewGSal(10)), "#.")
                    NewTotSalRow(9) = Format(Val(NewTotSalRow(9)) + Val(NewGSal(11)), "#.")
                    NewTotSalRow(10) = Format(Val(NewTotSalRow(10)) + Val(NewGSal(12)), "#.")
                    NewTotSalRow(11) = Format(Val(NewTotSalRow(11)) + Val(NewGSal(13)), "#.")
                    NewTotSalRow(12) = Format(Val(NewTotSalRow(12)) + Val(NewGSal(14)), "#.")
                    NewTotSalRow(13) = Format(Val(NewTotSalRow(13)) + Val(NewGSal(15)), "#.")
                    NewTotSalRow(14) = Format(Val(NewTotSalRow(14)) + Val(NewGSal(16)), "#.")
                    NewTotSalRow(15) = Format(Val(NewTotSalRow(15)) + Val(NewGSal(17)), "#.")

                    SalaryRowTotalMode()

                    Ctbl25.MoveNext()

                Loop

                MsgBox("DONE", vbInformation, "Codex")

            End If

        End If

    End Sub
    Sub LoadEmpSalaryCtrl2()
        GetSalaryNest()

        If PerCmb3.Text = "All" Then

            PerGrid01.Rows.Clear()

            SQL = ""
            SQL = SQL & "Select * from SalarySync1_Table "
            SQL = SQL & "Where Periode = ('" & PerCmb2.Text & "') "
            SQL = SQL & "And PeriodeRange = ('" & PerCmb1.Text & "') "
            SQL = SQL & "Order by Nik "
            OpenTbl(CBb, Ctbl25, SQL)
            If Ctbl25.RecordCount > 0 Then
                Ctbl25.MoveFirst()
                Do While Not Ctbl25.EOF

                    NewGSal(0) = IIf(IsDBNull(Ctbl25("Nik").Value), "", Ctbl25("Nik").Value)
                    NewGSal(1) = IIf(IsDBNull(Ctbl25("Name").Value), "", Ctbl25("Name").Value)
                    NewGSal(2) = IIf(IsDBNull(Ctbl25("Salary1").Value), "", Ctbl25("Salary1").Value)
                    NewGSal(3) = IIf(IsDBNull(Ctbl25("Salary2").Value), "", Ctbl25("Salary2").Value)
                    NewGSal(4) = IIf(IsDBNull(Ctbl25("Salary3").Value), "", Ctbl25("Salary3").Value)
                    NewGSal(5) = IIf(IsDBNull(Ctbl25("Salary4").Value), "", Ctbl25("Salary4").Value)
                    NewGSal(6) = IIf(IsDBNull(Ctbl25("Salary5").Value), "", Ctbl25("Salary5").Value)
                    NewGSal(7) = IIf(IsDBNull(Ctbl25("Salary6").Value), "", Ctbl25("Salary6").Value)
                    NewGSal(8) = IIf(IsDBNull(Ctbl25("Salary7").Value), "", Ctbl25("Salary7").Value)
                    NewGSal(9) = IIf(IsDBNull(Ctbl25("Salary8").Value), "", Ctbl25("Salary8").Value)
                    NewGSal(10) = IIf(IsDBNull(Ctbl25("Salary9").Value), "", Ctbl25("Salary9").Value)
                    NewGSal(11) = IIf(IsDBNull(Ctbl25("Salary10").Value), "", Ctbl25("Salary10").Value)
                    NewGSal(12) = IIf(IsDBNull(Ctbl25("Salary11").Value), "", Ctbl25("Salary11").Value)
                    NewGSal(13) = IIf(IsDBNull(Ctbl25("Salary12").Value), "", Ctbl25("Salary12").Value)
                    NewGSal(14) = IIf(IsDBNull(Ctbl25("Salary13").Value), "", Ctbl25("Salary13").Value)
                    NewGSal(15) = IIf(IsDBNull(Ctbl25("Salary14").Value), "", Ctbl25("Salary14").Value)
                    NewGSal(16) = IIf(IsDBNull(Ctbl25("Salary15").Value), "", Ctbl25("Salary15").Value)
                    NewGSal(17) = IIf(IsDBNull(Ctbl25("Salary16").Value), "", Ctbl25("Salary16").Value)
                    GetPotlain = IIf(IsDBNull(Ctbl25("PotLain").Value), "", Ctbl25("PotLain").Value)

                    AstekLoad()

                    GetSalaryTot = Format(Val(NewGSal(2)) + Val(NewGSal(3)) + Val(NewGSal(4)) + Val(NewGSal(5)) + Val(NewGSal(6)) + Val(NewGSal(7)) + Val(NewGSal(8)) + Val(NewGSal(9)) + Val(NewGSal(10)) + Val(NewGSal(11)) + Val(NewGSal(12)) + Val(NewGSal(13)) + Val(NewGSal(14)) + Val(NewGSal(15)) + Val(NewGSal(16)) + Val(NewGSal(17)), "#.")
                    GetSalaryTotFormat = Format(Val(GetSalaryTot) - Val(GetAstek), "#.")
                    GetSalaryTot = Format(Val(GetSalaryTot) - Val(GetAstek), "N0")
                    GetSalaryMaxTotal = Val(GetSalaryMaxTotal) + Val(GetSalaryTotFormat)
                    GridFormatNum()
                    LetsCount()

                    NewTotSalRow(0) = Format(Val(NewTotSalRow(0)) + Val(NewGSal(2)), "#.")
                    NewTotSalRow(1) = Format(Val(NewTotSalRow(1)) + Val(NewGSal(3)), "#.")
                    NewTotSalRow(2) = Format(Val(NewTotSalRow(2)) + Val(NewGSal(4)), "#.")
                    NewTotSalRow(3) = Format(Val(NewTotSalRow(3)) + Val(NewGSal(5)), "#.")
                    NewTotSalRow(4) = Format(Val(NewTotSalRow(4)) + Val(NewGSal(6)), "#.")
                    NewTotSalRow(5) = Format(Val(NewTotSalRow(5)) + Val(NewGSal(7)), "#.")
                    NewTotSalRow(6) = Format(Val(NewTotSalRow(6)) + Val(NewGSal(8)), "#.")
                    NewTotSalRow(7) = Format(Val(NewTotSalRow(7)) + Val(NewGSal(9)), "#.")
                    NewTotSalRow(8) = Format(Val(NewTotSalRow(8)) + Val(NewGSal(10)), "#.")
                    NewTotSalRow(9) = Format(Val(NewTotSalRow(9)) + Val(NewGSal(11)), "#.")
                    NewTotSalRow(10) = Format(Val(NewTotSalRow(10)) + Val(NewGSal(12)), "#.")
                    NewTotSalRow(11) = Format(Val(NewTotSalRow(11)) + Val(NewGSal(13)), "#.")
                    NewTotSalRow(12) = Format(Val(NewTotSalRow(12)) + Val(NewGSal(14)), "#.")
                    NewTotSalRow(13) = Format(Val(NewTotSalRow(13)) + Val(NewGSal(15)), "#.")
                    NewTotSalRow(14) = Format(Val(NewTotSalRow(14)) + Val(NewGSal(16)), "#.")
                    NewTotSalRow(15) = Format(Val(NewTotSalRow(15)) + Val(NewGSal(17)), "#.")
                    TotAstek = Val(TotAstek).ToString("N0", CustomtoUS) + Val(GetAstek).ToString("N0", CustomtoUS)
                    SalaryRowTotalMode()

                    Ctbl25.MoveNext()

                Loop

                MsgBox("DONE", vbInformation, "Codex")

            End If

        ElseIf PerCmb3.Text = "CASH" Then

            PerGrid01.Rows.Clear()

            SQL = ""
            SQL = SQL & "Select * from SalarySync1_Table "
            SQL = SQL & "Where Periode = ('" & PerCmb2.Text & "') "
            SQL = SQL & "And PeriodeRange = ('" & PerCmb1.Text & "') "
            SQL = SQL & "And Pay = ('" & "CASH" & "') "
            SQL = SQL & "Order by Nik "
            OpenTbl(CBb, Ctbl25, SQL)
            If Ctbl25.RecordCount > 0 Then
                Ctbl25.MoveFirst()
                Do While Not Ctbl25.EOF

                    NewGSal(0) = IIf(IsDBNull(Ctbl25("Nik").Value), "", Ctbl25("Nik").Value)
                    NewGSal(1) = IIf(IsDBNull(Ctbl25("Name").Value), "", Ctbl25("Name").Value)
                    NewGSal(2) = IIf(IsDBNull(Ctbl25("Salary1").Value), "", Ctbl25("Salary1").Value)
                    NewGSal(3) = IIf(IsDBNull(Ctbl25("Salary2").Value), "", Ctbl25("Salary2").Value)
                    NewGSal(4) = IIf(IsDBNull(Ctbl25("Salary3").Value), "", Ctbl25("Salary3").Value)
                    NewGSal(5) = IIf(IsDBNull(Ctbl25("Salary4").Value), "", Ctbl25("Salary4").Value)
                    NewGSal(6) = IIf(IsDBNull(Ctbl25("Salary5").Value), "", Ctbl25("Salary5").Value)
                    NewGSal(7) = IIf(IsDBNull(Ctbl25("Salary6").Value), "", Ctbl25("Salary6").Value)
                    NewGSal(8) = IIf(IsDBNull(Ctbl25("Salary7").Value), "", Ctbl25("Salary7").Value)
                    NewGSal(9) = IIf(IsDBNull(Ctbl25("Salary8").Value), "", Ctbl25("Salary8").Value)
                    NewGSal(10) = IIf(IsDBNull(Ctbl25("Salary9").Value), "", Ctbl25("Salary9").Value)
                    NewGSal(11) = IIf(IsDBNull(Ctbl25("Salary10").Value), "", Ctbl25("Salary10").Value)
                    NewGSal(12) = IIf(IsDBNull(Ctbl25("Salary11").Value), "", Ctbl25("Salary11").Value)
                    NewGSal(13) = IIf(IsDBNull(Ctbl25("Salary12").Value), "", Ctbl25("Salary12").Value)
                    NewGSal(14) = IIf(IsDBNull(Ctbl25("Salary13").Value), "", Ctbl25("Salary13").Value)
                    NewGSal(15) = IIf(IsDBNull(Ctbl25("Salary14").Value), "", Ctbl25("Salary14").Value)
                    NewGSal(16) = IIf(IsDBNull(Ctbl25("Salary15").Value), "", Ctbl25("Salary15").Value)
                    NewGSal(17) = IIf(IsDBNull(Ctbl25("Salary16").Value), "", Ctbl25("Salary16").Value)
                    'GetAstek = IIf(IsDBNull(Ctbl25("AstekVal").Value), "", Ctbl25("Astekval").Value)
                    GetPotlain = IIf(IsDBNull(Ctbl25("PotLain").Value), "", Ctbl25("PotLain").Value)

                    AstekLoad()

                    GetSalaryTot = Format(Val(NewGSal(2)) + Val(NewGSal(3)) + Val(NewGSal(4)) + Val(NewGSal(5)) + Val(NewGSal(6)) + Val(NewGSal(7)) + Val(NewGSal(8)) + Val(NewGSal(9)) + Val(NewGSal(10)) + Val(NewGSal(11)) + Val(NewGSal(12)) + Val(NewGSal(13)) + Val(NewGSal(14)) + Val(NewGSal(15)) + Val(NewGSal(16)) + Val(NewGSal(17)), "#.")
                    GetSalaryTotFormat = Format(Val(GetSalaryTot) - Val(GetAstek), "#.")
                    GetSalaryTot = Format(Val(GetSalaryTot) - Val(GetAstek), "N0")
                    GetSalaryMaxTotal = Val(GetSalaryMaxTotal) + Val(GetSalaryTotFormat)
                    GridFormatNum()
                    LetsCount()

                    NewTotSalRow(0) = Format(Val(NewTotSalRow(0)) + Val(NewGSal(2)), "#.")
                    NewTotSalRow(1) = Format(Val(NewTotSalRow(1)) + Val(NewGSal(3)), "#.")
                    NewTotSalRow(2) = Format(Val(NewTotSalRow(2)) + Val(NewGSal(4)), "#.")
                    NewTotSalRow(3) = Format(Val(NewTotSalRow(3)) + Val(NewGSal(5)), "#.")
                    NewTotSalRow(4) = Format(Val(NewTotSalRow(4)) + Val(NewGSal(6)), "#.")
                    NewTotSalRow(5) = Format(Val(NewTotSalRow(5)) + Val(NewGSal(7)), "#.")
                    NewTotSalRow(6) = Format(Val(NewTotSalRow(6)) + Val(NewGSal(8)), "#.")
                    NewTotSalRow(7) = Format(Val(NewTotSalRow(7)) + Val(NewGSal(9)), "#.")
                    NewTotSalRow(8) = Format(Val(NewTotSalRow(8)) + Val(NewGSal(10)), "#.")
                    NewTotSalRow(9) = Format(Val(NewTotSalRow(9)) + Val(NewGSal(11)), "#.")
                    NewTotSalRow(10) = Format(Val(NewTotSalRow(10)) + Val(NewGSal(12)), "#.")
                    NewTotSalRow(11) = Format(Val(NewTotSalRow(11)) + Val(NewGSal(13)), "#.")
                    NewTotSalRow(12) = Format(Val(NewTotSalRow(12)) + Val(NewGSal(14)), "#.")
                    NewTotSalRow(13) = Format(Val(NewTotSalRow(13)) + Val(NewGSal(15)), "#.")
                    NewTotSalRow(14) = Format(Val(NewTotSalRow(14)) + Val(NewGSal(16)), "#.")
                    NewTotSalRow(15) = Format(Val(NewTotSalRow(15)) + Val(NewGSal(17)), "#.")

                    TotAstek = Format(Val(TotAstek) + Val(GetAstek), "#.")

                    SalaryRowTotalMode()

                    Ctbl25.MoveNext()

                Loop
                MsgBox("DONE", vbInformation, "Codex")

            End If

        ElseIf PerCmb3.Text = "BTN" Then

            PerGrid01.Rows.Clear()

            SQL = ""
            SQL = SQL & "Select * from SalarySync1_Table "
            SQL = SQL & "Where Periode = ('" & PerCmb2.Text & "') "
            SQL = SQL & "And PeriodeRange = ('" & PerCmb1.Text & "') "
            SQL = SQL & "And Pay = ('" & "BTN" & "') "
            SQL = SQL & "Order by Nik "
            OpenTbl(CBb, Ctbl25, SQL)
            If Ctbl25.RecordCount > 0 Then
                Ctbl25.MoveFirst()
                Do While Not Ctbl25.EOF

                    NewGSal(0) = IIf(IsDBNull(Ctbl25("Nik").Value), "", Ctbl25("Nik").Value)
                    NewGSal(1) = IIf(IsDBNull(Ctbl25("Name").Value), "", Ctbl25("Name").Value)
                    NewGSal(2) = IIf(IsDBNull(Ctbl25("Salary1").Value), "", Ctbl25("Salary1").Value)
                    NewGSal(3) = IIf(IsDBNull(Ctbl25("Salary2").Value), "", Ctbl25("Salary2").Value)
                    NewGSal(4) = IIf(IsDBNull(Ctbl25("Salary3").Value), "", Ctbl25("Salary3").Value)
                    NewGSal(5) = IIf(IsDBNull(Ctbl25("Salary4").Value), "", Ctbl25("Salary4").Value)
                    NewGSal(6) = IIf(IsDBNull(Ctbl25("Salary5").Value), "", Ctbl25("Salary5").Value)
                    NewGSal(7) = IIf(IsDBNull(Ctbl25("Salary6").Value), "", Ctbl25("Salary6").Value)
                    NewGSal(8) = IIf(IsDBNull(Ctbl25("Salary7").Value), "", Ctbl25("Salary7").Value)
                    NewGSal(9) = IIf(IsDBNull(Ctbl25("Salary8").Value), "", Ctbl25("Salary8").Value)
                    NewGSal(10) = IIf(IsDBNull(Ctbl25("Salary9").Value), "", Ctbl25("Salary9").Value)
                    NewGSal(11) = IIf(IsDBNull(Ctbl25("Salary10").Value), "", Ctbl25("Salary10").Value)
                    NewGSal(12) = IIf(IsDBNull(Ctbl25("Salary11").Value), "", Ctbl25("Salary11").Value)
                    NewGSal(13) = IIf(IsDBNull(Ctbl25("Salary12").Value), "", Ctbl25("Salary12").Value)
                    NewGSal(14) = IIf(IsDBNull(Ctbl25("Salary13").Value), "", Ctbl25("Salary13").Value)
                    NewGSal(15) = IIf(IsDBNull(Ctbl25("Salary14").Value), "", Ctbl25("Salary14").Value)
                    NewGSal(16) = IIf(IsDBNull(Ctbl25("Salary15").Value), "", Ctbl25("Salary15").Value)
                    NewGSal(17) = IIf(IsDBNull(Ctbl25("Salary16").Value), "", Ctbl25("Salary16").Value)
                    'GetAstek = IIf(IsDBNull(Ctbl25("AstekVal").Value), "", Ctbl25("Astekval").Value)
                    GetPotlain = IIf(IsDBNull(Ctbl25("PotLain").Value), "", Ctbl25("PotLain").Value)

                    AstekLoad()

                    GetSalaryTot = Format(Val(NewGSal(2)) + Val(NewGSal(3)) + Val(NewGSal(4)) + Val(NewGSal(5)) + Val(NewGSal(6)) + Val(NewGSal(7)) + Val(NewGSal(8)) + Val(NewGSal(9)) + Val(NewGSal(10)) + Val(NewGSal(11)) + Val(NewGSal(12)) + Val(NewGSal(13)) + Val(NewGSal(14)) + Val(NewGSal(15)) + Val(NewGSal(16)) + Val(NewGSal(17)), "#.")
                    GetSalaryTotFormat = Format(Val(GetSalaryTot) - Val(GetAstek), "#.")
                    GetSalaryTot = Format(Val(GetSalaryTot) - Val(GetAstek), "N0")
                    GetSalaryMaxTotal = Val(GetSalaryMaxTotal) + Val(GetSalaryTotFormat)
                    GridFormatNum()
                    LetsCount()

                    NewTotSalRow(0) = Format(Val(NewTotSalRow(0)) + Val(NewGSal(2)), "#.")
                    NewTotSalRow(1) = Format(Val(NewTotSalRow(1)) + Val(NewGSal(3)), "#.")
                    NewTotSalRow(2) = Format(Val(NewTotSalRow(2)) + Val(NewGSal(4)), "#.")
                    NewTotSalRow(3) = Format(Val(NewTotSalRow(3)) + Val(NewGSal(5)), "#.")
                    NewTotSalRow(4) = Format(Val(NewTotSalRow(4)) + Val(NewGSal(6)), "#.")
                    NewTotSalRow(5) = Format(Val(NewTotSalRow(5)) + Val(NewGSal(7)), "#.")
                    NewTotSalRow(6) = Format(Val(NewTotSalRow(6)) + Val(NewGSal(8)), "#.")
                    NewTotSalRow(7) = Format(Val(NewTotSalRow(7)) + Val(NewGSal(9)), "#.")
                    NewTotSalRow(8) = Format(Val(NewTotSalRow(8)) + Val(NewGSal(10)), "#.")
                    NewTotSalRow(9) = Format(Val(NewTotSalRow(9)) + Val(NewGSal(11)), "#.")
                    NewTotSalRow(10) = Format(Val(NewTotSalRow(10)) + Val(NewGSal(12)), "#.")
                    NewTotSalRow(11) = Format(Val(NewTotSalRow(11)) + Val(NewGSal(13)), "#.")
                    NewTotSalRow(12) = Format(Val(NewTotSalRow(12)) + Val(NewGSal(14)), "#.")
                    NewTotSalRow(13) = Format(Val(NewTotSalRow(13)) + Val(NewGSal(15)), "#.")
                    NewTotSalRow(14) = Format(Val(NewTotSalRow(14)) + Val(NewGSal(16)), "#.")
                    NewTotSalRow(15) = Format(Val(NewTotSalRow(15)) + Val(NewGSal(17)), "#.")

                    TotAstek = Format(Val(TotAstek) + Val(GetAstek2), "#.")

                    SalaryRowTotalMode()

                    Ctbl25.MoveNext()

                Loop
                MsgBox("DONE", vbInformation, "Codex")

            End If

        End If

    End Sub

#End Region

#Region "Astek Looker"

    Sub AstekLoad()
        SQL = ""
        SQL = SQL & "Select * From 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & NewGSal(0) & "') "
        OpenTbl(ADb, Atbl37, SQL)

        If Not Atbl37.RecordCount = Nothing Then

            GetAstek2 = IIf(IsDBNull(Atbl37("Jamsostek").Value), "", Atbl37("Jamsostek").Value)

        End If

        If GetAstek2 = "" Or GetAstek2 = "0" Then
            GetAstek = ""

        Else
            GetAstek = GetAstek2

        End If



    End Sub
#End Region

#Region "Excel Codes"

    Sub GenExcel()

        ExcelName = "Sortasi Report" & "_" & PerCmb1.Text & "_" & Format(Now, "dd.MM.yyyy Hmmss")

        KillExcel()
        StartExcel()
        CreateWorkSheet()
        PopWorkSheet()
        SaveWorkSheet()
        CloseWorkSheet()
        OpenMe()

        'If Dir("C:\Program Files\Microsoft Office\Office12\excel.exe", vbDirectory) <> "" Then
        '    Shell("C:\Program Files\Microsoft Office\Office12\Excel " & Application.StartupPath & "\Report Excel\" & ExcelName & ".xls", vbMaximizedFocus)

        'ElseIf Dir("C:\Program Files\Microsoft Office\OFFICE11\excel.exe", vbDirectory) <> "" Then
        '    Shell("C:\C:\Program Files\Microsoft Office\OFFICE11\Excel " & Application.StartupPath & "\Report Excel\" & ExcelName & ".xls", vbMaximizedFocus)

        'ElseIf Dir("C:\Program Files\Microsoft Office\Office10\excel.exe", vbDirectory) <> "" Then
        '    Shell("C:\Program Files\Microsoft Office\Office11\Excel " & Application.StartupPath & "\Report Excel\" & ExcelName & ".xls", vbMaximizedFocus)

        'ElseIf Dir("C:\Program Files\Microsoft Office\Office\excel.exe", vbDirectory) <> "" Then
        '    Shell("C:\Program Files\Microsoft Office\Office11\Excel " & Application.StartupPath & "\Report Excel\" & ExcelName & ".xls", vbMaximizedFocus)

        'Else
        '    MsgBox("Microsoft Excel has not been found.", vbOKOnly + 64, "")

        'End If

    End Sub

    Sub KillExcel()

        If Dir(Application.StartupPath & "\Reports Excel\" & ExcelName & ".xls") <> "" Then
            Kill(Application.StartupPath & "\Reports Excel\" & ExcelName & ".xls")
        End If

    End Sub

    Sub StartExcel()

        On Error GoTo Err
        ExcelAP = GetObject("Excel.Application")
        Exit Sub
Err:
        ExcelAP = CreateObject("Excel.Application")

    End Sub

    Sub CreateWorkSheet()

        ExcelWB = ExcelAP.Workbooks.Add
        ExcelWS = ExcelWB.Worksheets(1)

    End Sub

    Sub PopWorkSheet()


        ExcelWS.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperLegal
        ExcelWS.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape
        ExcelWS.PageSetup.PrintTitleRows = "A7"
        ExcelWS.PageSetup.Zoom = 85

        With ExcelAP.Range("A1:AA1")

            .Merge()
            .Cells.Value = "PT. UNIVERSAL GLOVES"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2

        End With

        With ExcelAP.Range("A2:AA2")

            .Merge()
            .Cells.Value = "JL. Pertahanan No. 17 Patumbak 20361 Deli Serdang  - Indonesia"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2

        End With

        With ExcelAP.Range("A3:AA3")

            .Merge()
            .Cells.Value = "DAFTAR GAJI BORONGAN PER PERIODE"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2

        End With

        With ExcelAP.Range("A4:AA4")

            .Merge()
            .Cells.Value = "BAGIAN : SORTASI"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2

        End With

        With ExcelAP.Range("A5:AA5")

            .Merge()
            .Font.Bold = True
            .Cells.Value = PerCmb1.Text
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .Font.Name = "Calibri"
            .Font.Size = 10

        End With

        '-----------------------------------------------------------------------------------------------

        With ExcelAP.Range("A7:A7")

            .Cells.Value = "NIK"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("B7:B7")

            .Cells.Value = "NAMA"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 20

        End With

        With ExcelAP.Range("C7:C7")

            .Cells.Value = "No. Rek"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("D7:D7")

            .Cells.Value = NewSald(0)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("E7:E7")

            .Cells.Value = NewSald(1)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("F7:F7")

            .Cells.Value = NewSald(2)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("G7:G7")

            .Cells.Value = NewSald(3)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("H7:H7")

            .Cells.Value = NewSald(4)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("I7:I7")

            .Cells.Value = NewSald(5)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("J7:J7")


            .Cells.Value = NewSald(6)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With


        With ExcelAP.Range("K7:K7")


            .Cells.Value = NewSald(7)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With


        With ExcelAP.Range("L7:L7")

            .Cells.Value = NewSald(8)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("M7:M7")

            .Cells.Value = NewSald(9)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("N7:N7")

            .Cells.Value = NewSald(10)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("O7:O7")

            .Cells.Value = NewSald(11)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("P7:P7")

            .Cells.Value = NewSald(12)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("Q7:Q7")

            .Cells.Value = NewSald(13)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("R7:R7")

            .Cells.Value = NewSald(14)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("S7:S7")

            .Cells.Value = NewSald(15)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("T7:T7")

            .Cells.Value = ""
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("U7:U7")

            .Cells.Value = ""
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("V7:V7")

            .Cells.Value = "ASTEK"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("W7:W7")


            .Cells.Value = "TJ. LAIN"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("X7:X7")


            .Cells.Value = "TJ. PPH21"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With


        With ExcelAP.Range("Y7:Y7")


            .Cells.Value = "PPH21"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("Z7:Z7")


            .Cells.Value = "POT LAIN"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("AA7:AA7")


            .Cells.Value = "TOTAL"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With


        With ExcelAP.Range("AB7:AB7")


            .Cells.Value = "GAJI"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("AC7:AC7")

            .Cells.Value = "PARAF"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("AD7:AD7")

            .Cells.Value = "Pay"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With


        For i = 0 To PerGrid01.Rows.Count - 1
            For j = 0 To PerGrid01.ColumnCount - 1
                ExcelWS.Cells(i + 8, j + 1) = PerGrid01(j, i).Value
                ExcelWS.Cells(i + 8, j + 1).Borders.LineStyle = 1
            Next
        Next

        i = i + 2

        ExcelWS.Cells(i + 9, 1) = SalLbl1.Text
        ExcelWS.Cells(i + 9, 2) = SalTbx1.Text
        ExcelWS.Cells(i + 10, 1) = SalLbl2.Text
        ExcelWS.Cells(i + 10, 2) = SalTbx2.Text
        ExcelWS.Cells(i + 11, 1) = SalLbl3.Text
        ExcelWS.Cells(i + 11, 2) = SalTbx3.Text
        ExcelWS.Cells(i + 12, 1) = SalLbl4.Text
        ExcelWS.Cells(i + 12, 2) = SalTbx4.Text
        ExcelWS.Cells(i + 13, 1) = SalLbl5.Text
        ExcelWS.Cells(i + 13, 2) = SalTbx5.Text
        ExcelWS.Cells(i + 14, 1) = SalLbl6.Text
        ExcelWS.Cells(i + 14, 2) = SalTbx6.Text
        ExcelWS.Cells(i + 15, 1) = SalLbl7.Text
        ExcelWS.Cells(i + 15, 2) = SalTbx7.Text
        ExcelWS.Cells(i + 16, 1) = SalLbl8.Text
        ExcelWS.Cells(i + 16, 2) = SalTbx8.Text
        ExcelWS.Cells(i + 17, 1) = SalLbl9.Text
        ExcelWS.Cells(i + 17, 2) = SalTbx9.Text
        ExcelWS.Cells(i + 18, 1) = SalLbl10.Text
        ExcelWS.Cells(i + 18, 2) = SalTbx10.Text
        ExcelWS.Cells(i + 19, 1) = SalLbl11.Text
        ExcelWS.Cells(i + 19, 2) = SalTbx11.Text
        ExcelWS.Cells(i + 20, 1) = SalLbl12.Text
        ExcelWS.Cells(i + 20, 2) = SalTbx12.Text
        ExcelWS.Cells(i + 21, 1) = SalLbl13.Text
        ExcelWS.Cells(i + 21, 2) = SalTbx13.Text
        ExcelWS.Cells(i + 22, 1) = SalLbl14.Text
        ExcelWS.Cells(i + 22, 2) = SalTbx14.Text
        ExcelWS.Cells(i + 23, 1) = SalLbl15.Text
        ExcelWS.Cells(i + 23, 2) = SalTbx15.Text
        ExcelWS.Cells(i + 24, 1) = SalLbl16.Text
        ExcelWS.Cells(i + 24, 2) = SalTbx16.Text
        ExcelWS.Cells(i + 26, 1) = "Incentives: "
        ExcelWS.Cells(i + 26, 2) = SalTbx17.Text
        ExcelWS.Cells(i + 27, 1) = "Total: "
        ExcelWS.Cells(i + 27, 2) = SalTbx18.Text
        ExcelWS.Cells(i + 28, 1) = "Astek: "
        ExcelWS.Cells(i + 28, 2) = SalTbx19.Text
        ExcelWS.Cells(i + 29, 1) = "Rounded Total: "
        ExcelWS.Cells(i + 29, 2) = SalTbx20.Text




    End Sub

    Sub SaveWorkSheet()
        On Error GoTo Err
        SaveWorkBook()
Err:
    End Sub

    Sub SaveWorkBook()
        ExcelWB.SaveAs(Application.StartupPath & "\Report Excel\" & ExcelName & ".xls")
    End Sub

    Sub CloseWorkSheet()
        ExcelAP.Workbooks.Close()
        ExcelAP.Quit()
    End Sub

    Sub OpenMe()

        Dim oXLApp As Object, oXLWorkbook As Object

        oXLApp = CreateObject("Excel.Application")


        oXLWorkbook = oXLApp.Workbooks.Open(FileName:=Application.StartupPath & "\Report Excel\" & ExcelName & ".xls")

        oXLApp.Visible = True


    End Sub

#End Region

#Region "NEW Excel CALL"
    Dim SaveName As String
    Sub SaveFileLink()

        Dim SaveFileName As New SaveFileDialog
        SaveFileName.Filter = "Excel File (*.xlsx)|*.xlsx"
        SaveFileName.FilterIndex = 1
        If SaveFileName.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            SaveName = SaveFileName.FileName
        End If

    End Sub

    Sub ExportExcel()

        Try
            'SaveFileLink()
            Dim NewFile As New FileInfo(SaveName)
            If NewFile.Exists Then
                NewFile.Delete()
            End If

            Using ExcelModPkg = New ExcelPackage(NewFile)

                ' Create Work Sheet

                Dim ExcelNewWSH As ExcelWorksheet = ExcelModPkg.Workbook.Worksheets.Add("GAJIAN")

                ExcelNewWSH.PrinterSettings.PaperSize = ePaperSize.Legal
                ExcelNewWSH.PrinterSettings.Orientation = eOrientation.Landscape

                With ExcelNewWSH.Cells("A1:AA1")

                    .Merge = True
                    .Value = "PT. UNIVERSAL GLOVES"
                    .Style.Font.Bold = True
                    .Style.Font.Name = "Calibri"
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

                End With

                With ExcelNewWSH.Cells("A2:AA2")

                    .Merge = True
                    .Value = "JL. Pertahanan No. 17 Patumbak 20361 Deli Serdang  - Indonesia"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

                End With

                With ExcelNewWSH.Cells("A3:AA3")

                    .Merge = True
                    .Value = "DAFTAR GAJI BORONGAN PER PERIODE"
                    .Style.Font.Name = "Calibri"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

                End With

                With ExcelNewWSH.Cells("A4:AA4")
                    .Merge = True
                    .Value = "BAGIAN : SORTASI"
                    .Style.Font.Name = "Calibri"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

                End With

                With ExcelNewWSH.Cells("A6")

                    .Value = "NIK"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With

                ' Block

                With ExcelNewWSH.Cells("B6")

                    .Value = "Name"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()

                End With
                With ExcelNewWSH.Cells("C6")
                    .Value = "No. Rek"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("D6")

                    .Value = NewSald(0)
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("E6")

                    .Value = NewSald(1)
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("F6")

                    .Value = NewSald(2)
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("G6")

                    .Value = NewSald(3)
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("H6")

                    .Value = NewSald(4)
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("I6")

                    .Value = NewSald(5)
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("J6")

                    .Value = NewSald(6)
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("K6")

                    .Value = NewSald(7)
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("L6")

                    .Value = NewSald(8)
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("M6")

                    .Value = NewSald(9)
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("N6")

                    .Value = NewSald(10)
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("O6")

                    .Value = NewSald(11)
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("P6")

                    .Value = NewSald(12)
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("Q6")

                    .Value = NewSald(13)
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("R6")

                    .Value = NewSald(14)
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("S6")

                    .Value = NewSald(15)
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("T6")

                    .Value = ""
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With

                With ExcelNewWSH.Cells("U6")

                    .Value = ""
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("V6")

                    .Value = "ASTEK"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("W6")

                    .Value = "TJ. LAIN"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("X6")

                    .Value = "TJ. PPH21"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("Y6")

                    .Value = "PPH21"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With

                With ExcelNewWSH.Cells("Z6")

                    .Value = "POT. LAIN"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With

                With ExcelNewWSH.Cells("AA6")

                    .Value = "TOTAL"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With

                With ExcelNewWSH.Cells("AB6")

                    .Value = "GAJI"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With

                With ExcelNewWSH.Cells("AC6")

                    .Value = "PARAF"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With

                With ExcelNewWSH.Cells("AD6")

                    .Value = "Pay"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With

                ' For GAJI

                For i = 0 To PerGrid01.Rows.Count - 1

                    For j = 0 To PerGrid01.ColumnCount - 1
                        ExcelNewWSH.Cells(i + 7, j + 1).Value = PerGrid01(j, i).Value
                        ExcelNewWSH.Cells(i + 7, j + 1).Style.Border.BorderAround(ExcelBorderStyle.Thin)
                        ExcelNewWSH.Cells.Style.Font.Size = 8
                    Next

                    ExcelNewWSH.Cells.AutoFitColumns()
                Next


                i = i + 2

                ExcelNewWSH.Cells(i + 9, 1).Value = SalLbl1.Text
                ExcelNewWSH.Cells(i + 9, 2).Value = SalTbx1.Text
                ExcelNewWSH.Cells(i + 10, 1).Value = SalLbl2.Text
                ExcelNewWSH.Cells(i + 10, 2).Value = SalTbx2.Text
                ExcelNewWSH.Cells(i + 11, 1).Value = SalLbl3.Text
                ExcelNewWSH.Cells(i + 11, 2).Value = SalTbx3.Text
                ExcelNewWSH.Cells(i + 12, 1).Value = SalLbl4.Text
                ExcelNewWSH.Cells(i + 12, 2).Value = SalTbx4.Text
                ExcelNewWSH.Cells(i + 13, 1).Value = SalLbl5.Text
                ExcelNewWSH.Cells(i + 13, 2).Value = SalTbx5.Text
                ExcelNewWSH.Cells(i + 14, 1).Value = SalLbl6.Text
                ExcelNewWSH.Cells(i + 14, 2).Value = SalTbx6.Text
                ExcelNewWSH.Cells(i + 15, 1).Value = SalLbl7.Text
                ExcelNewWSH.Cells(i + 15, 2).Value = SalTbx7.Text
                ExcelNewWSH.Cells(i + 16, 1).Value = SalLbl8.Text
                ExcelNewWSH.Cells(i + 16, 2).Value = SalTbx8.Text
                ExcelNewWSH.Cells(i + 17, 1).Value = SalLbl9.Text
                ExcelNewWSH.Cells(i + 17, 2).Value = SalTbx9.Text
                ExcelNewWSH.Cells(i + 18, 1).Value = SalLbl10.Text
                ExcelNewWSH.Cells(i + 18, 2).Value = SalTbx10.Text
                ExcelNewWSH.Cells(i + 19, 1).Value = SalLbl11.Text
                ExcelNewWSH.Cells(i + 19, 2).Value = SalTbx11.Text
                ExcelNewWSH.Cells(i + 20, 1).Value = SalLbl12.Text
                ExcelNewWSH.Cells(i + 20, 2).Value = SalTbx12.Text
                ExcelNewWSH.Cells(i + 21, 1).Value = SalLbl13.Text
                ExcelNewWSH.Cells(i + 21, 2).Value = SalTbx13.Text
                ExcelNewWSH.Cells(i + 22, 1).Value = SalLbl14.Text
                ExcelNewWSH.Cells(i + 22, 2).Value = SalTbx14.Text
                ExcelNewWSH.Cells(i + 23, 1).Value = SalLbl15.Text
                ExcelNewWSH.Cells(i + 23, 2).Value = SalTbx15.Text
                ExcelNewWSH.Cells(i + 24, 1).Value = SalLbl16.Text
                ExcelNewWSH.Cells(i + 24, 2).Value = SalTbx16.Text
                ExcelNewWSH.Cells(i + 26, 1).Value = "Incentives: "
                ExcelNewWSH.Cells(i + 26, 2).Value = SalTbx17.Text
                ExcelNewWSH.Cells(i + 27, 1).Value = "Total: "
                ExcelNewWSH.Cells(i + 27, 2).Value = SalTbx18.Text
                ExcelNewWSH.Cells(i + 28, 1).Value = "Astek: "
                ExcelNewWSH.Cells(i + 28, 2).Value = SalTbx19.Text
                ExcelNewWSH.Cells(i + 29, 1).Value = "Rounded Total: "
                ExcelNewWSH.Cells(i + 29, 2).Value = SalTbx20.Text


                ExcelModPkg.Save()
                Dim LookMe As New ProcessStartInfo(SaveName)
                Process.Start(LookMe)

            End Using

        Catch ex As Exception

        End Try

    End Sub
#End Region

#Region "BGM on MODE"
    Private BGWorkMode() As BackgroundWorker
    Private i = 0
    Sub OnClickTheWorker()

        i += 1
        ReDim BGWorkMode(i)
        BGWorkMode(i) = New BackgroundWorker
        BGWorkMode(i).WorkerReportsProgress = True
        BGWorkMode(i).WorkerSupportsCancellation = True
        AddHandler BGWorkMode(i).DoWork, AddressOf WorkerDoWork
        AddHandler BGWorkMode(i).ProgressChanged, AddressOf WorkerProgressChanged
        AddHandler BGWorkMode(i).RunWorkerCompleted, AddressOf WorkerCompleted
        BGWorkMode(i).RunWorkerAsync()

    End Sub
    Private Sub WorkerDoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs)
        ExportExcel()
    End Sub

    Private Sub WorkerProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs)

    End Sub

    Private Sub WorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)

    End Sub

#End Region

#Region "Round Off Mode"

    Sub RoundProcess()

        TotRoundoff = GetSalaryTot
        TotRound1 = CustomRound(TotRoundoff)
    End Sub


#End Region

    Private Sub PerBtn25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PerBtn25.Click

        If PerCmb4.Text = "" Or PerCmb5.Text = "" Then
            MsgBox("Please Select the Required Item for PPh21/Slip")
        Else
            If PerCmb5.Text = "1" Then
                UpPPhValue1()
            ElseIf PerCmb5.Text = "2" Then
                UpPPhValue2()
            ElseIf PerCmb5.Text = "Ex" Then
                UpPPhValue3()
            End If
        End If

    End Sub

    Private Sub PerCmb4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PerCmb4.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub PerCmb5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PerCmb5.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub PerBtn6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PerBtn6.Click

        If PerCmb3.Text = "CASH" Then

            CouponDelete()
            CouponUp()

        Else

            MsgBox("This function is for NON-BTN/CASH only")

        End If

    End Sub

 
End Class