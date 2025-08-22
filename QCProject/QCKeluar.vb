Option Explicit On

Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Globalization
Imports System.Threading



Public Class KeluarBlock

    Dim ExcelAP As Excel.Application
    Dim ExcelWB As Excel.Workbook
    Dim ExcelWS As Excel.Worksheet

    Dim ExcelName As String

    Dim SubGaji As String
    Dim TotalGaji As String
    Dim TotRoundoff As Integer
    Dim SubGaji2 As String
    Dim KelomVal As String

    Dim Date1a As String = Nothing
    Dim Date2a As String = Nothing
    Dim Date3a As String = Nothing
    Dim Date4a As String = Nothing
    Dim Date5a As String = Nothing
    Dim Date6a As String = Nothing
    Dim Date7a As String = Nothing
    Dim Date8a As String = Nothing
    Dim Date9a As String = Nothing
    Dim Date10a As String = Nothing
    Dim Date11a As String = Nothing
    Dim Date12a As String = Nothing
    Dim Date13a As String = Nothing
    Dim Date14a As String = Nothing
    Dim Date15a As String = Nothing
    Dim Date16a As String = Nothing

    Dim NewDate1 As New DateTime
    Dim NewDate2 As New DateTime
    Dim NewDate3 As New DateTime
    Dim NewDate4 As New DateTime
    Dim NewDate5 As New DateTime
    Dim NewDate6 As New DateTime
    Dim NewDate7 As New DateTime
    Dim NewDate8 As New DateTime
    Dim NewDate9 As New DateTime
    Dim NewDate10 As New DateTime
    Dim NewDate11 As New DateTime
    Dim NewDate12 As New DateTime
    Dim NewDate13 As New DateTime
    Dim NewDate14 As New DateTime
    Dim NewDate15 As New DateTime
    Dim NewDate16 As New DateTime

    Dim DelAstek As String = ""
    Dim DelMasuk As String = ""


    Dim Salr1 As String = Nothing
    Dim Salr2 As String = Nothing
    Dim Salr3 As String = Nothing
    Dim Salr4 As String = Nothing
    Dim Salr5 As String = Nothing
    Dim Salr6 As String = Nothing
    Dim Salr7 As String = Nothing
    Dim Salr8 As String = Nothing
    Dim Salr9 As String = Nothing
    Dim Salr10 As String = Nothing
    Dim Salr11 As String = Nothing
    Dim Salr12 As String = Nothing
    Dim Salr13 As String = Nothing
    Dim Salr14 As String = Nothing
    Dim Salr15 As String = Nothing
    Dim Salr16 As String = Nothing
    Dim AstSlr As String = Nothing


    Dim Dt1 As String
    Dim Dt2 As String
    Dim Dt3 As String
    Dim Dt4 As String
    Dim Dt5 As String
    Dim Dt6 As String
    Dim Dt7 As String
    Dim Dt8 As String
    Dim Dt9 As String
    Dim Dt10 As String
    Dim Dt11 As String
    Dim Dt12 As String
    Dim Dt13 As String
    Dim Dt14 As String
    Dim Dt15 As String
    Dim Dt16 As String
    Dim SSDate As String
    Dim SEDate As String


    Dim NikKel As String
    Dim CmbDater As String
    Dim CmbDater1 As String
    Dim CmbDater2 As String
    Dim CmbDater3 As String
    Dim AddressOne As String
    Dim AddressTwo As String
    Dim AddressThree As String
    Dim AddressFour As String
    Dim Signing As String

    Private Sub KeluarBlock_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        MainMenu.Refresh()
    End Sub

    Private Sub QCKeluar_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        AddresssLoader()
        KDTPick01.Text = Today
        KDTPick01.Format = DateTimePickerFormat.Custom
        KDTPick01.CustomFormat = "dd MMM yyyy"

        KDTPick02.Text = Today
        KDTPick02.Format = DateTimePickerFormat.Custom
        KDTPick02.CustomFormat = "dd MMM yyyy"


        LoadDB()
        LoadDB2()
        LoadDB3()

    End Sub


    Sub DeleteKeluarPerson()
        SQL = ""
        SQL = SQL & "Select * From Keluar_Table "
        OpenTbl(CBb, Ctbl53, SQL)
        If Ctbl53.RecordCount > 0 Then
            Ctbl53.Delete()
        End If

    End Sub



    Sub SaveKeluarPerson()

        SQL = ""
        SQL = SQL & "Select * From Keluar_Table "
        SQL = SQL & "Where Nik_Num = ('" & DelTbx1.Text & "') "
        SQL = SQL & "And Name = ('" & DelTbx2.Text & "') "
        OpenTbl(CBb, Ctbl53, SQL)

        If Not Ctbl53.RecordCount <> 0 Then
            Ctbl53.AddNew()
        End If

        Ctbl53("Nik_Num").Value = NikKel
        Ctbl53("Name").Value = DelTbx2.Text
        Ctbl53("From_Date").Value = KDTPick01.Text
        Ctbl53("To_Date").Value = KDTPick02.Text
        Ctbl53("TglMsk").Value = DelMasuk
        Ctbl53("KeloF").Value = KelomVal

        'Date Blank

        Ctbl53("Date01").Value = IIf(KeluarLb1.Text = Nothing, DBNull.Value, KeluarLb1.Text)
        Ctbl53("Date02").Value = IIf(KeluarLb2.Text = Nothing, DBNull.Value, KeluarLb2.Text)
        Ctbl53("Date03").Value = IIf(KeluarLb3.Text = Nothing, DBNull.Value, KeluarLb3.Text)
        Ctbl53("Date04").Value = IIf(KeluarLb4.Text = Nothing, DBNull.Value, KeluarLb4.Text)
        Ctbl53("Date05").Value = IIf(KeluarLb5.Text = Nothing, DBNull.Value, KeluarLb5.Text)
        Ctbl53("Date06").Value = IIf(KeluarLb6.Text = Nothing, DBNull.Value, KeluarLb6.Text)
        Ctbl53("Date07").Value = IIf(KeluarLb7.Text = Nothing, DBNull.Value, KeluarLb7.Text)
        Ctbl53("Date08").Value = IIf(KeluarLb8.Text = Nothing, DBNull.Value, KeluarLb8.Text)
        Ctbl53("Date09").Value = IIf(KeluarLb9.Text = Nothing, DBNull.Value, KeluarLb9.Text)
        Ctbl53("Date10").Value = IIf(KeluarLb10.Text = Nothing, DBNull.Value, KeluarLb10.Text)
        Ctbl53("Date11").Value = IIf(KeluarLb11.Text = Nothing, DBNull.Value, KeluarLb11.Text)
        Ctbl53("Date12").Value = IIf(KeluarLb12.Text = Nothing, DBNull.Value, KeluarLb12.Text)
        Ctbl53("Date13").Value = IIf(KeluarLb13.Text = Nothing, DBNull.Value, KeluarLb13.Text)
        Ctbl53("Date14").Value = IIf(KeluarLb14.Text = Nothing, DBNull.Value, KeluarLb14.Text)
        Ctbl53("Date15").Value = IIf(KeluarLb15.Text = Nothing, DBNull.Value, KeluarLb15.Text)
        Ctbl53("Date16").Value = IIf(KeluarLb16.Text = Nothing, DBNull.Value, KeluarLb16.Text)
    
        ' Salary Blank
        Ctbl53("Sal01").Value = IIf(KeluarTbx1.Text = Nothing, Nothing, KeluarTbx1.Text)
        Ctbl53("Sal02").Value = IIf(KeluarTbx2.Text = Nothing, Nothing, KeluarTbx2.Text)
        Ctbl53("Sal03").Value = IIf(KeluarTbx3.Text = Nothing, Nothing, KeluarTbx3.Text)
        Ctbl53("Sal04").Value = IIf(KeluarTbx4.Text = Nothing, Nothing, KeluarTbx4.Text)
        Ctbl53("Sal05").Value = IIf(KeluarTbx5.Text = Nothing, Nothing, KeluarTbx5.Text)
        Ctbl53("Sal06").Value = IIf(KeluarTbx6.Text = Nothing, Nothing, KeluarTbx6.Text)
        Ctbl53("Sal07").Value = IIf(KeluarTbx7.Text = Nothing, Nothing, KeluarTbx7.Text)
        Ctbl53("Sal08").Value = IIf(KeluarTbx8.Text = Nothing, Nothing, KeluarTbx8.Text)
        Ctbl53("Sal09").Value = IIf(KeluarTbx9.Text = Nothing, Nothing, KeluarTbx9.Text)
        Ctbl53("Sal10").Value = IIf(KeluarTbx10.Text = Nothing, Nothing, KeluarTbx10.Text)
        Ctbl53("Sal11").Value = IIf(KeluarTbx11.Text = Nothing, Nothing, KeluarTbx11.Text)
        Ctbl53("Sal12").Value = IIf(KeluarTbx12.Text = Nothing, Nothing, KeluarTbx12.Text)
        Ctbl53("Sal13").Value = IIf(KeluarTbx13.Text = Nothing, Nothing, KeluarTbx13.Text)
        Ctbl53("Sal14").Value = IIf(KeluarTbx14.Text = Nothing, Nothing, KeluarTbx14.Text)
        Ctbl53("Sal15").Value = IIf(KeluarTbx15.Text = Nothing, Nothing, KeluarTbx15.Text)
        Ctbl53("Sal16").Value = IIf(KeluarTbx16.Text = Nothing, Nothing, KeluarTbx16.Text)
        Ctbl53("Total").Value = IIf(KeluarSub.Text = Nothing, Nothing, KeluarSub.Text)

        Ctbl53("Tunjangan").Value = IIf(KeluarTun.Text = Nothing, Nothing, KeluarTun.Text)
        Ctbl53("Potongan").Value = IIf(KeluarPot.Text = Nothing, Nothing, KeluarPot.Text)
        Ctbl53("Astek").Value = IIf(AstSlr = Nothing, Nothing, AstSlr)
        Ctbl53("GajiDet").Value = IIf(KeluarTot.Text = Nothing, Nothing, KeluarTot.Text)

        Ctbl53("Sign1").Value = User
        Ctbl53("Sign2").Value = Signing

        Ctbl53.Update()

        MessageBox.Show("You may now click the [New Report] Button", "Success", MessageBoxButtons.OK)
        Me.Refresh()


    End Sub



    Sub GenerateDate()

        SQL = ""
        SQL = SQL & "Select * From DateCounter2Table "
        SQL = SQL & "Where Periode =  ('" & KelCmb2.Text & "') "
        SQL = SQL & "And PeriodeRange =  ('" & KelCmb1.Text & "')"
        OpenTbl(CBb, Ctbl22, SQL)

        If Ctbl22.RecordCount <> 0 Then
            Ctbl22.MoveLast()

            Date1a = If(IsDBNull(Ctbl22("Date1").Value), "", Format(Ctbl22("Date1").Value))
            Date2a = If(IsDBNull(Ctbl22("Date2").Value), "", Format(Ctbl22("Date2").Value))
            Date3a = If(IsDBNull(Ctbl22("Date3").Value), "", Format(Ctbl22("Date3").Value))
            Date4a = If(IsDBNull(Ctbl22("Date4").Value), "", Format(Ctbl22("Date4").Value))
            Date5a = If(IsDBNull(Ctbl22("Date5").Value), "", Format(Ctbl22("Date5").Value))
            Date6a = If(IsDBNull(Ctbl22("Date6").Value), "", Format(Ctbl22("Date6").Value))
            Date7a = If(IsDBNull(Ctbl22("Date7").Value), "", Format(Ctbl22("Date7").Value))
            Date8a = If(IsDBNull(Ctbl22("Date8").Value), "", Format(Ctbl22("Date8").Value))
            Date9a = If(IsDBNull(Ctbl22("Date9").Value), "", Format(Ctbl22("Date9").Value))
            Date10a = If(IsDBNull(Ctbl22("Date10").Value), "", Format(Ctbl22("Date10").Value))
            Date11a = If(IsDBNull(Ctbl22("Date11").Value), "", Format(Ctbl22("Date11").Value))
            Date12a = If(IsDBNull(Ctbl22("Date12").Value), "", Format(Ctbl22("Date12").Value))
            Date13a = If(IsDBNull(Ctbl22("Date13").Value), "", Format(Ctbl22("Date13").Value))
            Date14a = If(IsDBNull(Ctbl22("Date14").Value), "", Format(Ctbl22("Date14").Value))
            Date15a = If(IsDBNull(Ctbl22("Date15").Value), "", Format(Ctbl22("Date15").Value))
            Date16a = If(IsDBNull(Ctbl22("Date16").Value), "", Format(Ctbl22("Date16").Value))


            If Not Date1a = "" Then
                NewDate1 = Date1a
                KeluarLb1.Text = NewDate1.ToString("dd MMM yyyy")
            Else
                KeluarLb1.Text = ""

            End If

            If Not Date2a = "" Then
                NewDate2 = Date2a
                KeluarLb2.Text = NewDate2.ToString("dd MMM yyyy")
            Else
                KeluarLb2.Text = ""

            End If

            If Not Date3a = Nothing Then
                NewDate3 = Date3a
                KeluarLb3.Text = NewDate3.ToString("dd MMM yyyy")
            Else
                KeluarLb3.Text = ""

            End If

            If Not Date4a = "" Then
                NewDate4 = Date4a
                KeluarLb4.Text = NewDate4.ToString("dd MMM yyyy")
            Else
                KeluarLb4.Text = ""

            End If

            If Not Date5a = "" Then
                NewDate5 = Date5a
                KeluarLb5.Text = NewDate5.ToString("dd MMM yyyy")
            Else
                KeluarLb5.Text = ""

            End If

            If Not Date6a = "" Then
                NewDate6 = Date6a
                KeluarLb6.Text = NewDate6.ToString("dd MMM yyyy")
            Else
                KeluarLb6.Text = ""

            End If

            If Not Date7a = "" Then
                NewDate7 = Date7a
                KeluarLb7.Text = NewDate7.ToString("dd MMM yyyy")
            Else
                KeluarLb7.Text = ""

            End If

            If Not Date8a = "" Then
                NewDate8 = Date8a
                KeluarLb8.Text = NewDate8.ToString("dd MMM yyyy")
            Else
                KeluarLb8.Text = ""

            End If

            If Not Date9a = "" Then
                NewDate9 = Date9a
                KeluarLb9.Text = NewDate9.ToString("dd MMM yyyy")
            Else
                KeluarLb9.Text = ""

            End If

            If Not Date10a = "" Then
                NewDate10 = Date10a
                KeluarLb10.Text = NewDate10.ToString("dd MMM yyyy")
            Else
                KeluarLb10.Text = ""

            End If

            If Not Date11a = "" Then
                NewDate11 = Date11a
                KeluarLb11.Text = NewDate11.ToString("dd MMM yyyy")
            Else
                KeluarLb11.Text = ""

            End If

            If Not Date12a = "" Then
                NewDate12 = Date12a
                KeluarLb12.Text = NewDate12.ToString("dd MMM yyyy")
            Else
                KeluarLb12.Text = ""

            End If

            If Not Date13a = "" Then
                NewDate13 = Date13a
                KeluarLb13.Text = NewDate13.ToString("dd MMM yyyy")
            Else
                KeluarLb13.Text = ""

            End If

            If Not Date14a = "" Then
                NewDate14 = Date14a
                KeluarLb14.Text = NewDate14.ToString("dd MMM yyyy")
            Else
                KeluarLb14.Text = ""

            End If

            If Not Date15a = "" Then
                NewDate15 = Date15a
                KeluarLb15.Text = NewDate15.ToString("dd MMM yyyy")
            Else
                KeluarLb15.Text = ""

            End If

            If Not Date16a = "" Then
                NewDate16 = Date16a
                KeluarLb16.Text = NewDate16.ToString("dd MMM yyyy")
            Else
                KeluarLb16.Text = ""

            End If

        End If

        Me.Refresh()

    End Sub

    Sub EmpLook()

        SQL = ""
        SQL = SQL & "Select * from SalarySync1_Table "
        SQL = SQL & "Where Nik =  ('" & DelTbx1.Text & "')"
        SQL = SQL & "And PeriodeRange = ('" & KelCmb1.Text & "')"
        SQL = SQL & "And Periode = ('" & KelCmb2.Text & "')"
        SQL = SQL & "Order by Nik"
        OpenTbl(CBb, Ctbl45, SQL)

        If Ctbl45.RecordCount <> 0 Then
            Ctbl45.MoveLast()

            NikKel = Ctbl45("Nik").Value
            DelTbx2.Text = Ctbl45("Name").Value
            Salr1 = If(IsDBNull(Ctbl45("Salary1").Value), "", Ctbl45("Salary1").Value)
            Salr2 = If(IsDBNull(Ctbl45("Salary2").Value), "", Ctbl45("Salary2").Value)
            Salr3 = If(IsDBNull(Ctbl45("Salary3").Value), "", Ctbl45("Salary3").Value)
            Salr4 = If(IsDBNull(Ctbl45("Salary4").Value), "", Ctbl45("Salary4").Value)
            Salr5 = If(IsDBNull(Ctbl45("Salary5").Value), "", Ctbl45("Salary5").Value)
            Salr6 = If(IsDBNull(Ctbl45("Salary6").Value), "", Ctbl45("Salary6").Value)
            Salr7 = If(IsDBNull(Ctbl45("Salary7").Value), "", Ctbl45("Salary7").Value)
            Salr8 = If(IsDBNull(Ctbl45("Salary8").Value), "", Ctbl45("Salary8").Value)
            Salr9 = If(IsDBNull(Ctbl45("Salary9").Value), "", Ctbl45("Salary9").Value)
            Salr10 = If(IsDBNull(Ctbl45("Salary10").Value), "", Ctbl45("Salary10").Value)
            Salr11 = If(IsDBNull(Ctbl45("Salary11").Value), "", Ctbl45("Salary11").Value)
            Salr12 = If(IsDBNull(Ctbl45("Salary12").Value), "", Ctbl45("Salary12").Value)
            Salr13 = If(IsDBNull(Ctbl45("Salary13").Value), "", Ctbl45("Salary13").Value)
            Salr14 = If(IsDBNull(Ctbl45("Salary14").Value), "", Ctbl45("Salary14").Value)
            Salr15 = If(IsDBNull(Ctbl45("Salary15").Value), "", Ctbl45("Salary15").Value)
            Salr16 = If(IsDBNull(Ctbl45("Salary16").Value), "", Ctbl45("Salary16").Value)

            If KelCmb2.Text = "Periode II" Then

                AstSlr = IIf(IsDBNull(Ctbl45("AstekVal").Value), "", Ctbl45("AstekVal").Value)

            End If

            If Not AstSlr = "" Or AstSlr = "0" Then

                KelAstek.Text = Format(Val(AstSlr), "N1")

            End If

            If Not Salr1 = "" Then
                KeluarTbx1.Text = Format(Val(Salr1), "N1")

            End If

            If Not Salr2 = "" Then
                KeluarTbx2.Text = Format(Val(Salr2), "N1")

            End If

            If Not Salr3 = "" Then
                KeluarTbx3.Text = Format(Val(Salr3), "N1")

            End If

            If Not Salr4 = "" Then
                KeluarTbx4.Text = Format(Val(Salr4), "N1")

            End If

            If Not Salr5 = "" Then
                KeluarTbx5.Text = Format(Val(Salr5), "N1")

            End If

            If Not Salr6 = "" Then
                KeluarTbx6.Text = Format(Val(Salr6), "N1")

            End If

            If Not Salr7 = "" Then
                KeluarTbx7.Text = Format(Val(Salr7), "N1")

            End If

            If Not Salr8 = "" Then
                KeluarTbx8.Text = Format(Val(Salr8), "N1")

            End If

            If Not Salr9 = "" Then
                KeluarTbx9.Text = Format(Val(Salr9), "N1")

            End If

            If Not Salr10 = "" Then
                KeluarTbx10.Text = Format(Val(Salr10), "N1")

            End If

            If Not Salr11 = "" Then
                KeluarTbx11.Text = Format(Val(Salr11), "N1")

            End If

            If Not Salr12 = "" Then
                KeluarTbx12.Text = Format(Val(Salr12), "N1")

            End If

            If Not Salr13 = "" Then
                KeluarTbx13.Text = Format(Val(Salr13), "N1")

            End If

            If Not Salr14 = "" Then
                KeluarTbx14.Text = Format(Val(Salr14), "N1")

            End If

            If Not Salr15 = "" Then
                KeluarTbx15.Text = Format(Val(Salr15), "N1")

            End If

            If Not Salr16 = "" Then
                KeluarTbx16.Text = Format(Val(Salr16), "N1")

            End If

            SubGaji = Val(Salr1) + Val(Salr2) + Val(Salr3) + Val(Salr4) + Val(Salr5) + Val(Salr6) + Val(Salr7) + Val(Salr8) + Val(Salr9) + Val(Salr10) + Val(Salr11) + Val(Salr12) + Val(Salr13) + Val(Salr14) + Val(Salr15) + Val(Salr16)

            KeluarSub.Text = Format(Val(SubGaji), "N1")

            SubGaji2 = (Val(SubGaji)) - (Val(AstSlr))

            RoundProcess()

            KeluarTot.Text = Format(Val(TotalGaji), "N1")

        End If

    End Sub

    Sub RoundProcess()

        TotRoundoff = SubGaji2
        TotalGaji = CustomRound(TotRoundoff)

    End Sub

    Sub KelDelete()
        SQL = ""
        SQL = SQL & "Select * From SalarySync1_Table "
        SQL = SQL & "Where Nik = ('" & DelTbx1.Text & "') "
        SQL = SQL & "And PeriodeRange = ('" & KelCmb1.Text & "')"
        SQL = SQL & "And Periode = ('" & KelCmb2.Text & "')"
        OpenTbl(CBb, Ctbl49, SQL)
        If Ctbl49.RecordCount <> 0 Then

            Ctbl49.Delete()
            MsgBox("Data has been Deleted", vbInformation)

        End If

    End Sub


    Private Sub DelTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DelTbx1.KeyPress

        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            GenerateDate()
            EmpLook()
            DelEmpLooker()
        End If

    End Sub

    Private Sub KeluarBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeluarBtn1.Click
        Unactive()
        GenerateDate()
        DeleteKeluarPerson()
        SaveKeluarPerson()

    End Sub

    Sub DelEmpLooker()

        SQL = ""
        SQL = SQL & "Select * from 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & DelTbx1.Text & "')"
        SQL = SQL & "Order by Nik"
        OpenTbl(ADb, Atbl29, SQL)

        If Atbl29.RecordCount <> 0 Then
            Atbl29.MoveLast()

            DelMasuk = IIf(IsDBNull(Atbl29("DateStart").Value), "", Format(Atbl29("DateStart").Value, "dd MMM yy"))
            DelAstek = IIf(IsDBNull(Atbl29("Jamsostek").Value), "", Atbl29("Jamsostek").Value)
            KelomVal = If(IsDBNull(Atbl29("Dept").Value), "", Format(Atbl29("Dept").Value))

        End If

    End Sub

    Sub Unactive()
        SQL = ""
        SQL = SQL & "Select * from 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & DelTbx1.Text & "')"
        SQL = SQL & "Order by Nik"
        OpenTbl(ADb, Atbl40, SQL)

        If Not Atbl40.RecordCount < 0 Then

            Atbl40("Active").Value = "No"
            Atbl40.Update()

        Else
            MsgBox("Person not Found")

        End If

    End Sub

    Sub AddresssLoader()

        SQL = ""
        SQL = SQL & "Select * From 08_Standard_Table "
        SQL = SQL & "Where Original = ('" & "SignBy" & "') "
        OpenTbl(ADb, Atbl22, SQL)
        If Atbl22.RecordCount <> 0 Then
            Atbl22.MoveLast()
            Signing = IIf(IsDBNull(Atbl22("Standard_Wage").Value), "", Atbl22("Standard_Wage").Value)

        End If
    End Sub

#Region "Excel Codes"

    Sub GenExcel()

        ExcelName = "Keluar Report" & "_" & KelCmb1.Text & "_" & Format(Now, "dd.MM.yyyy Hmmss")

        KillExcel()
        StartExcel()
        CreateWorkSheet()
        PopWorkSheet()
        SaveWorkSheet()
        CloseWorkSheet()
        OpenMe()

        If Dir("C:\Program Files\Microsoft Office\Office12\excel.exe", vbDirectory) <> "" Then
            Shell("C:\Program Files\Microsoft Office\Office12\Excel " & Application.StartupPath & "\Report Excel\" & ExcelName & ".xls", vbMaximizedFocus)

        ElseIf Dir("C:\Program Files\Microsoft Office\OFFICE11\excel.exe", vbDirectory) <> "" Then
            Shell("C:\C:\Program Files\Microsoft Office\OFFICE11\Excel " & Application.StartupPath & "\Report Excel\" & ExcelName & ".xls", vbMaximizedFocus)

        ElseIf Dir("C:\Program Files\Microsoft Office\Office10\excel.exe", vbDirectory) <> "" Then
            Shell("C:\Program Files\Microsoft Office\Office11\Excel " & Application.StartupPath & "\Report Excel\" & ExcelName & ".xls", vbMaximizedFocus)

        ElseIf Dir("C:\Program Files\Microsoft Office\Office\excel.exe", vbDirectory) <> "" Then
            Shell("C:\Program Files\Microsoft Office\Office11\Excel " & Application.StartupPath & "\Report Excel\" & ExcelName & ".xls", vbMaximizedFocus)

        Else
            MsgBox("Microsoft Excel has not been found.", vbOKOnly + 64, "")
        End If

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

        Dim Row As Integer

        ExcelWS.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperLegal
        ExcelWS.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape
        ExcelWS.PageSetup.PrintTitleRows = "A7"
        ExcelWS.PageSetup.Zoom = 85

        With ExcelAP.Range("A1:J1")

            .Merge()
            .Cells.Value = "PT. UNIVERSAL GLOVES"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2

        End With

        With ExcelAP.Range("A2:J2")

            .Merge()
            .Cells.Value = "JL. Pertahanan No. 17 Patumbak 20361 Deli Serdang  - Indonesia"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2

        End With

        With ExcelAP.Range("A3:J3")

            .Merge()
            .Cells.Value = "DAFTAR GAJI BORONGAN PER PERIODE"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2

        End With

        With ExcelAP.Range("A4:J4")

            .Merge()
            .Cells.Value = "BAGIAN : SORTASI"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2

        End With

        With ExcelAP.Range("A5:J5")

            .Merge()
            .Font.Bold = True
            .Cells.Value = KelCmb1.Text
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .Font.Name = "Calibri"
            .Font.Size = 10

        End With

        '-----------------------------------------------------------------------------------------------

        Row = Row + 5

        ExcelWS.Cells(Row + 2, 1) = "Nik"
        ExcelWS.Cells(Row + 2, 2) = ":  " + DelTbx1.Text
        ExcelWS.Cells(Row + 3, 1) = "Nama Karyawan"
        ExcelWS.Cells(Row + 3, 2) = ":  " + DelTbx2.Text
        ExcelWS.Cells(Row + 4, 1) = "Alamat"
        ExcelWS.Cells(Row + 4, 2) = ":  " + ""
        ExcelWS.Cells(Row + 5, 1) = "Kelompok"
        ExcelWS.Cells(Row + 5, 2) = ":  " + ""
        ExcelWS.Cells(Row + 6, 1) = "Tgl Msk Kerja"
        ExcelWS.Cells(Row + 6, 2) = ":  " + DelMasuk
        ExcelWS.Cells(Row + 8, 1) = KeluarLb1.Text
        ExcelWS.Cells(Row + 8, 2) = ":      " + KeluarTbx1.Text
        ExcelWS.Cells(Row + 8, 6) = KeluarLb2.Text
        ExcelWS.Cells(Row + 8, 7) = "      :     " + KeluarTbx2.Text
        ExcelWS.Cells(Row + 9, 1) = KeluarLb3.Text
        ExcelWS.Cells(Row + 9, 2) = ":      " + KeluarTbx3.Text
        ExcelWS.Cells(Row + 9, 6) = KeluarLb4.Text
        ExcelWS.Cells(Row + 9, 7) = "      :     " + KeluarTbx4.Text
        ExcelWS.Cells(Row + 10, 1) = KeluarLb5.Text
        ExcelWS.Cells(Row + 10, 2) = ":     " + KeluarTbx5.Text
        ExcelWS.Cells(Row + 10, 6) = KeluarLb6.Text
        ExcelWS.Cells(Row + 10, 7) = "      :     " + KeluarTbx6.Text
        ExcelWS.Cells(Row + 11, 1) = KeluarLb7.Text
        ExcelWS.Cells(Row + 11, 2) = ":     " + KeluarTbx7.Text
        ExcelWS.Cells(Row + 11, 6) = KeluarLb8.Text
        ExcelWS.Cells(Row + 11, 7) = "      :     " + KeluarTbx8.Text
        ExcelWS.Cells(Row + 12, 1) = KeluarLb9.Text
        ExcelWS.Cells(Row + 12, 2) = ":     " + KeluarTbx9.Text
        ExcelWS.Cells(Row + 12, 6) = KeluarLb10.Text
        ExcelWS.Cells(Row + 12, 7) = "      :     " + KeluarTbx10.Text
        ExcelWS.Cells(Row + 13, 1) = KeluarLb11.Text
        ExcelWS.Cells(Row + 13, 2) = ":     " + KeluarTbx11.Text
        ExcelWS.Cells(Row + 13, 6) = KeluarLb12.Text
        ExcelWS.Cells(Row + 13, 7) = "      :     " + KeluarTbx12.Text
        ExcelWS.Cells(Row + 14, 1) = KeluarLb13.Text
        ExcelWS.Cells(Row + 14, 2) = ":     " + KeluarTbx13.Text
        ExcelWS.Cells(Row + 14, 6) = KeluarLb14.Text
        ExcelWS.Cells(Row + 14, 7) = "      :     " + KeluarTbx14.Text
        ExcelWS.Cells(Row + 15, 1) = KeluarLb15.Text
        ExcelWS.Cells(Row + 15, 2) = ":     " + KeluarTbx15.Text
        ExcelWS.Cells(Row + 15, 6) = KeluarLb16.Text
        ExcelWS.Cells(Row + 15, 7) = "      :     " + KeluarTbx16.Text


        ExcelWS.Cells(Row + 17, 1) = "Total"
        ExcelWS.Cells(Row + 17, 3) = KeluarSub.Text
        ExcelWS.Cells(Row + 18, 1) = "Tunjangan"
        ExcelWS.Cells(Row + 18, 3) = KeluarTun.Text
        ExcelWS.Cells(Row + 19, 1) = "Potongan"
        ExcelWS.Cells(Row + 19, 3) = KeluarPot.Text
        ExcelWS.Cells(Row + 20, 1) = "Astek"
        ExcelWS.Cells(Row + 20, 3) = KelAstek.Text
        ExcelWS.Cells(Row + 21, 1) = "Gaji Diterima"
        ExcelWS.Cells(Row + 21, 3) = KeluarTot.Text
        ExcelWS.Cells(Row + 24, 2) = "Patumbak  " + Format(Now, "dddd, dd MMM yyyy")
        ExcelWS.Cells(Row + 25, 7) = "DIKETAHUI OLEH"
        ExcelWS.Cells(Row + 25, 3) = "DIBUAT OLEH"
        ExcelWS.Cells(Row + 27, 3) = "LINA"
        ExcelWS.Cells(Row + 27, 7) = "BUNGA MARI TARIGAN"


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

    Private Sub KelCmb1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True

    End Sub

    Private Sub KelCmb2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles KelCmb2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub KeluarBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeluarBtn2.Click
        GenExcel()
    End Sub

    Private Sub GlassButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' Generate Crystal Report
        SDelMasuk = DelMasuk
        SSigning = Signing

        Dim NewMDIChild As New KeluarReportView()
        KeluarReportView.MdiParent = MainMenu
        KeluarReportView.Show()
        Me.Refresh()

    End Sub

    Private Sub KelBtn4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If DelTbx1.Text = "" Or KelCmb1.Text = "" Or KelCmb2.Text = "" Then
            MsgBox("Please Complete your Data First")
        Else
            KelDelete()
        End If

    End Sub

    Private Sub KeluarBtn4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeluarBtn4.Click

        KeluarNewKerja02.MdiParent = MainMenu
        KeluarNewKerja02.Show()

    End Sub

    Private Sub KeluarPot_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles KeluarPot.KeyPress
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            SubGaji2 = SubGaji2 - (Val(KeluarPot.Text))
            KeluarSub.Text = Format(Val(SubGaji2), "N1")
            RoundProcess()
            KeluarTot.Text = Format(Val(TotalGaji), "N1")

        End If
    End Sub

    Private Sub KeluarBtn3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeluarBtn3.Click

        SQL = ""
        SQL = SQL & "Select * from SalarySync1_Table "
        SQL = SQL & "Where Nik = ('" & DelTbx1.Text & "') "
        SQL = SQL & "And Name = ('" & DelTbx2.Text & "') "
        SQL = SQL & "And Periode = ('" & KelCmb1.Text & "') "
        SQL = SQL & "And PeriodeRange = ('" & KelCmb2.Text & "') "
        OpenTbl(CBb, Ctbl52, SQL)

        If Ctbl52.RecordCount > 0 Then
            Ctbl52.Delete()
        End If

    End Sub

  
    Private Sub DelTbx1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DelTbx1.TextChanged

    End Sub
End Class