Option Explicit On

Imports System
Imports System.Reflection
Imports System.Threading
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel


Public Class Report2Block

    Dim ExcelName As String
    Dim ExcelAP As Excel.Application
    Dim ExcelWB As Excel.Workbook
    Dim ExcelWS As Excel.Worksheet
    Dim CmbDater As String
    Dim CmbDater2 As String
    Dim CmbDater3 As String
    Dim GetSalaryTot As String
    Dim NikLooker As String
    Dim NameLooker As String
    Dim AstekLookMode As String = 0
    Dim CtrlShow As String = 0
    Dim RecordCounting As String = 0
    Dim TableLooker As String
    Dim MiscLook As String
    Dim GetSalaryTotFormat As String
    Dim Chooser As String
    Dim Bracket As String


    Dim SalaryLooker1 As String
    Dim SalaryLooker2 As String
    Dim SalaryLooker3 As String
    Dim SalaryLooker4 As String
    Dim SalaryLooker5 As String
    Dim SalaryLooker6 As String
    Dim SalaryLooker7 As String
    Dim SalaryLooker8 As String
    Dim SalaryLooker9 As String
    Dim SalaryLooker10 As String
    Dim SalaryLooker11 As String
    Dim SalaryLooker12 As String
    Dim SalaryLooker13 As String
    Dim SalaryLooker14 As String
    Dim SalaryLooker15 As String
    Dim SalaryLooker16 As String
    Dim GetInceRange As String
    Dim GetInceCount As String
    Dim GetInceDay As String
    Dim GetInceDif As String
    Dim GetInceSum As String
    Dim GetInceSumAgain As String
    Dim TotRoundoff As Integer
    Dim TotRound1 As Integer
    Dim TotRound2 As String



    Dim i As Integer
    Dim j As Integer

    Private Sub QCReport2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LoadDB()
        LoadDB2()
        LoadDB3()

    End Sub

#Region "Formmating Entity / Nest to 0"

    Sub GetSalaryNest()

        For i = 0 To 17
            NewGSal(i) = Nothing


        Next
        GetAstek = 0

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
        SalTbx19.Text = ""
        SalTbx21.Text = ""
    End Sub

    Sub SalaryLookerNest()

        SalaryLooker1 = ""
        SalaryLooker2 = ""
        SalaryLooker3 = ""
        SalaryLooker4 = ""
        SalaryLooker5 = ""
        SalaryLooker6 = ""
        SalaryLooker7 = ""
        SalaryLooker8 = ""
        SalaryLooker9 = ""
        SalaryLooker10 = ""
        SalaryLooker11 = ""
        SalaryLooker12 = ""
        SalaryLooker13 = ""
        SalaryLooker14 = ""
        SalaryLooker15 = ""
        SalaryLooker16 = ""


    End Sub
    Sub SalaryRowTotalMode()

        SalTbx1.Invoke(DirectCast(Sub() SalTbx1.Text = Format(Val(NewTotSalRow(0)), "#,#."), MethodInvoker))
        SalTbx2.Invoke(DirectCast(Sub() SalTbx2.Text = Format(Val(NewTotSalRow(1)), "#,#."), MethodInvoker))
        SalTbx3.Invoke(DirectCast(Sub() SalTbx3.Text = Format(Val(NewTotSalRow(2)), "#,#."), MethodInvoker))
        SalTbx4.Invoke(DirectCast(Sub() SalTbx4.Text = Format(Val(NewTotSalRow(3)), "#,#."), MethodInvoker))
        SalTbx5.Invoke(DirectCast(Sub() SalTbx5.Text = Format(Val(NewTotSalRow(4)), "#,#."), MethodInvoker))
        SalTbx6.Invoke(DirectCast(Sub() SalTbx6.Text = Format(Val(NewTotSalRow(5)), "#,#."), MethodInvoker))
        SalTbx7.Invoke(DirectCast(Sub() SalTbx7.Text = Format(Val(NewTotSalRow(6)), "#,#."), MethodInvoker))
        SalTbx8.Invoke(DirectCast(Sub() SalTbx8.Text = Format(Val(NewTotSalRow(7)), "#,#."), MethodInvoker))
        SalTbx9.Invoke(DirectCast(Sub() SalTbx9.Text = Format(Val(NewTotSalRow(8)), "#,#."), MethodInvoker))
        SalTbx10.Invoke(DirectCast(Sub() SalTbx10.Text = Format(Val(NewTotSalRow(9)), "#,#."), MethodInvoker))
        SalTbx11.Invoke(DirectCast(Sub() SalTbx11.Text = Format(Val(NewTotSalRow(10)), "#,#."), MethodInvoker))
        SalTbx12.Invoke(DirectCast(Sub() SalTbx12.Text = Format(Val(NewTotSalRow(11)), "#,#."), MethodInvoker))
        SalTbx13.Invoke(DirectCast(Sub() SalTbx13.Text = Format(Val(NewTotSalRow(12)), "#,#."), MethodInvoker))
        SalTbx14.Invoke(DirectCast(Sub() SalTbx14.Text = Format(Val(NewTotSalRow(13)), "#,#."), MethodInvoker))
        SalTbx15.Invoke(DirectCast(Sub() SalTbx15.Text = Format(Val(NewTotSalRow(14)), "#,#."), MethodInvoker))
        SalTbx16.Invoke(DirectCast(Sub() SalTbx16.Text = Format(Val(NewTotSalRow(15)), "#,#."), MethodInvoker))
        SalTbx19.Invoke(DirectCast(Sub() SalTbx19.Text = Format(Val(TotAstek), "#,#."), MethodInvoker))
    End Sub

    Sub LetsCount() ' Counting the Record in Grid

        RecordCounting = RecordCounting + 1
        SalTbx21.Invoke(DirectCast(Sub() SalTbx21.Text = RecordCounting, MethodInvoker))

    End Sub

    Sub FieldGridFormatNum()

        ' Salary 1
        If SalaryLooker1 = "" Or SalaryLooker1 = "0" Then
            NewFormT(0) = ""
        Else
            NewFormT(0) = SalaryLooker1
        End If

        ' Salary 2
        If SalaryLooker2 = "" Or SalaryLooker2 = "0" Then
            NewFormT(1) = ""
        Else
            NewFormT(1) = SalaryLooker2
        End If

        'Salary 3
        If SalaryLooker3 = "" Or SalaryLooker3 = "0" Then
            NewFormT(2) = ""
        Else
            NewFormT(2) = SalaryLooker3
        End If

        ' Salary 4
        If SalaryLooker4 = "" Or SalaryLooker4 = "0" Then
            NewFormT(3) = ""

        Else
            NewFormT(3) = SalaryLooker4

        End If


        ' Salary 5
        If SalaryLooker5 = "" Or SalaryLooker5 = "0" Then
            NewFormT(4) = ""
        Else
            NewFormT(4) = SalaryLooker5
        End If

        ' Salary 6
        If SalaryLooker6 = "" Or SalaryLooker6 = "0" Then
            NewFormT(5) = ""
        Else
            NewFormT(5) = SalaryLooker6
        End If

        'Salary 7
        If SalaryLooker7 = "" Or SalaryLooker7 = "0" Then
            NewFormT(6) = ""
        Else
            NewFormT(6) = SalaryLooker7
        End If

        ' Salary 8
        If SalaryLooker8 = "" Or SalaryLooker8 = "0" Then
            NewFormT(7) = ""
        Else
            NewFormT(7) = SalaryLooker8
        End If

        'Salary 9
        If SalaryLooker9 = "" Or SalaryLooker9 = "0" Then
            NewFormT(8) = ""
        Else
            NewFormT(8) = NewGSal(10)
        End If

        ' Salary 10
        If SalaryLooker10 = "" Or SalaryLooker10 = "0" Then
            NewFormT(9) = ""
        Else
            NewFormT(9) = SalaryLooker10
        End If

        'Salary 11
        If SalaryLooker11 = "" Or SalaryLooker11 = "0" Then
            NewFormT(10) = ""
        Else
            NewFormT(10) = SalaryLooker11
        End If

        ' Salary 12
        If SalaryLooker12 = "" Or SalaryLooker12 = "0" Then
            NewFormT(11) = ""
        Else
            NewFormT(11) = SalaryLooker12
        End If

        ' Salary 13
        If SalaryLooker13 = "" Or SalaryLooker13 = "0" Then
            NewFormT(12) = ""
        Else
            NewFormT(12) = SalaryLooker13
        End If

        ' Salary 14
        If SalaryLooker14 = "" Or SalaryLooker14 = "0" Then
            NewFormT(13) = ""
        Else
            NewFormT(13) = SalaryLooker14
        End If

        ' Salary 15
        If SalaryLooker15 = "" Or SalaryLooker15 = "0" Then
            NewFormT(14) = ""
        Else
            NewFormT(14) = SalaryLooker15
        End If

        ' Salary 16
        If SalaryLooker16 = "" Or SalaryLooker16 = "0" Then
            NewFormT(15) = ""
        Else
            NewFormT(15) = SalaryLooker16
        End If

        RoundProcess()
        TotRound2 = Format(TotRound1, "N0")

    End Sub

    Sub TotalTbxFiller() ' For Total per Row on Grid

        NewTotSalRow(0) = Format(Val(NewTotSalRow(0)) + Val(SalaryLooker1), "#.")
        NewTotSalRow(1) = Format(Val(NewTotSalRow(1)) + Val(SalaryLooker2), "#.")
        NewTotSalRow(2) = Format(Val(NewTotSalRow(2)) + Val(SalaryLooker3), "#.")
        NewTotSalRow(3) = Format(Val(NewTotSalRow(3)) + Val(SalaryLooker4), "#.")
        NewTotSalRow(4) = Format(Val(NewTotSalRow(4)) + Val(SalaryLooker5), "#.")
        NewTotSalRow(5) = Format(Val(NewTotSalRow(5)) + Val(SalaryLooker6), "#.")
        NewTotSalRow(6) = Format(Val(NewTotSalRow(6)) + Val(SalaryLooker7), "#.")
        NewTotSalRow(7) = Format(Val(NewTotSalRow(7)) + Val(SalaryLooker8), "#.")
        NewTotSalRow(8) = Format(Val(NewTotSalRow(8)) + Val(SalaryLooker9), "#.")
        NewTotSalRow(9) = Format(Val(NewTotSalRow(9)) + Val(SalaryLooker10), "#.")
        NewTotSalRow(10) = Format(Val(NewTotSalRow(10)) + Val(SalaryLooker11), "#.")
        NewTotSalRow(11) = Format(Val(NewTotSalRow(11)) + Val(SalaryLooker12), "#.")
        NewTotSalRow(12) = Format(Val(NewTotSalRow(12)) + Val(SalaryLooker13), "#.")
        NewTotSalRow(13) = Format(Val(NewTotSalRow(13)) + Val(SalaryLooker14), "#.")
        NewTotSalRow(14) = Format(Val(NewTotSalRow(14)) + Val(SalaryLooker15), "#.")
        NewTotSalRow(15) = Format(Val(NewTotSalRow(15)) + Val(SalaryLooker16), "#.")


        SalaryRowTotalMode()

    End Sub

    Sub AstekLoadOld()



        If PerFieldCmb2.Text = "Periode II" Then

            GetSalaryTotFormat = Format(Val(GetSalaryTot) - Val(AstekLookMode), "#.")
            GetSalaryTot = Format(Val(GetSalaryTot) - Val(AstekLookMode), "N0")

            If AstekLookMode = "" Or AstekLookMode = "0" Then
                AstekLookMode = ""
            End If



        ElseIf PerFieldCmb2.Text = "Periode I" Then
            AstekLookMode = ""
        End If



    End Sub

#End Region


#Region "Astek Looker"

    Sub AstekLoad()
        SQL = ""
        SQL = SQL & "Select * From 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & NewGSal(0) & "') "
        OpenTbl(ADb, Atbl37, SQL)

        If Atbl37.RecordCount <> 0 Then
            GetAstek2 = Atbl37("Jamsostek").Value
        End If

        If GetAstek2 = "" Or GetAstek2 = "0" Then
            GetAstek = ""

        Else
            GetAstek = GetAstek2

        End If



    End Sub

#End Region

    Sub ClockCodeNew()

        If Chooser = "Conveyour" Then
            PerDataLoader()

        ElseIf Chooser = "Mutu II" Then
            PerDataLoader2()

        ElseIf Chooser = "Packing" Then
            PerDataLoader3()
        ElseIf Chooser = "Wallet" Then

            PerDataLoader4()

        ElseIf Chooser = "Sortasi" Then

            PerDataLoader5()
        ElseIf Chooser = "Miscellaneous" Then

            PerDataLoader5()


        End If
    End Sub

#Region "Incentives Pattern"

    Sub IncentivesControlLoad() ' 2nd process in Incentive Section
        SQL = ""
        SQL = SQL & "Select * From 12_Incentives_Ctrl "
        SQL = SQL & "Where Nik = ('" & NewGSal(0) & "') "
        SQL = SQL & "And MonthPeriodeRange = ('" & GetInceRange & "') "
        OpenTbl(ADb, Atbl32, SQL)

        If Atbl32.RecordCount <> 0 Then
            GetInceCount = Atbl32("Count").Value


        ElseIf Not Atbl32.RecordCount <> 0 Then
            GetInceCount = "0"


        End If
    End Sub




    Sub IncentivesControlRange() '  must be first search on Incetives Section
        SQL = ""
        SQL = SQL & "Select * From 22_Incentives_Setup "
        SQL = SQL & "Where Actives = ('" & "Yes" & "') "
        OpenTbl(ADb, Atbl33, SQL)

        If Atbl33.RecordCount > 0 Then

            GetInceRange = Atbl33("MonthPeriodeRange").Value
            GetInceDay = Atbl33("Day").Value

        End If

    End Sub


    Sub IncentiveCalculation() ' 3rd process Control


        If CBRep2.Checked = True Then

            GetInceDif = GetInceDay - GetInceCount

            If GetInceDif = 4 Then

                GetInceSum = "57,600"
                GetInceSumAgain = "57600"

            ElseIf GetInceDif = 3 Then

                GetInceSum = "115,200"
                GetInceSumAgain = "115200"

            ElseIf GetInceDif = 2 Then

                GetInceSum = "172,800"
                GetInceSumAgain = "172800"

            ElseIf GetInceDif <= 1 Then

                GetInceSum = "230,000"
                GetInceSumAgain = "230000"

            ElseIf GetInceDif >= 5 Then

                GetInceSum = "0"
                GetInceSumAgain = "230000"

            End If

        ElseIf CBRep2.Checked = False Then

            GetInceSum = ""
            GetInceSumAgain = ""

        End If


    End Sub

    Sub IncentiveRun()

        IncentivesControlRange()
        IncentivesControlLoad()
        IncentiveCalculation()

    End Sub


#End Region

#Region "Loading Field Data"

    Sub LoadHeadDater()

        SQL = ""
        SQL = SQL & "Select * From DateCounter2Table "
        SQL = SQL & "Where Periode = ('" & PerFieldCmb2.Text & "') "
        SQL = SQL & "And PeriodeRange = ('" & PerFieldCmb1.Text & "') "
        OpenTbl(CBb, Ctbl24, SQL)


        If Ctbl24.RecordCount > 0 Then

            NewSald(0) = Ctbl24("Date1").Value
            NewSald(1) = Ctbl24("Date2").Value
            NewSald(2) = Ctbl24("Date3").Value
            NewSald(3) = Ctbl24("Date4").Value
            NewSald(4) = Ctbl24("Date5").Value
            NewSald(5) = Ctbl24("Date6").Value
            NewSald(6) = Ctbl24("Date7").Value
            NewSald(7) = Ctbl24("Date8").Value
            NewSald(8) = Ctbl24("Date9").Value
            NewSald(9) = Ctbl24("Date10").Value
            NewSald(10) = Ctbl24("Date11").Value
            NewSald(11) = Ctbl24("Date12").Value
            NewSald(12) = Ctbl24("Date13").Value
            NewSald(13) = Ctbl24("Date14").Value
            NewSald(14) = Ctbl24("Date15").Value
            NewSald(15) = Ctbl24("Date16").Value

        Else

            For i = 0 To 15
                NewSald(i) = Nothing
            Next


        End If

        PerFieldGrid01.Columns(2).HeaderText = NewSald(0)
        PerFieldGrid01.Columns(3).HeaderText = NewSald(1)
        PerFieldGrid01.Columns(4).HeaderText = NewSald(2)
        PerFieldGrid01.Columns(5).HeaderText = NewSald(3)
        PerFieldGrid01.Columns(6).HeaderText = NewSald(4)
        PerFieldGrid01.Columns(7).HeaderText = NewSald(5)
        PerFieldGrid01.Columns(8).HeaderText = NewSald(6)
        PerFieldGrid01.Columns(9).HeaderText = NewSald(7)
        PerFieldGrid01.Columns(10).HeaderText = NewSald(8)
        PerFieldGrid01.Columns(11).HeaderText = NewSald(9)
        PerFieldGrid01.Columns(12).HeaderText = NewSald(10)
        PerFieldGrid01.Columns(13).HeaderText = NewSald(11)
        PerFieldGrid01.Columns(14).HeaderText = NewSald(12)
        PerFieldGrid01.Columns(15).HeaderText = NewSald(13)
        PerFieldGrid01.Columns(16).HeaderText = NewSald(14)
        PerFieldGrid01.Columns(17).HeaderText = NewSald(15)


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

    End Sub

    Sub PerDataLoader() ' Conveyour

        NikLooker = ""
        NameLooker = ""
        SalaryLooker1 = ""
        RecordCounting = "0"

        If Bracket = "All" Then

            SQL = ""
            SQL = SQL & "Select * from 02_Name_Table "
            SQL = SQL & "Where Active = ('" & "Yes" & "') "
            SQL = SQL & "Order by Nik "

            OpenTbl(ADb, Dtbb40, SQL)
            If Dtbb40.RecordCount <> 0 Then

                Dtbb40.MoveFirst()
                Do While Not Dtbb40.EOF

                    NikLooker = IIf(IsDBNull(Dtbb40("Nik").Value), "", Dtbb40("Nik").Value)
                    NameLooker = IIf(IsDBNull(Dtbb40("Name").Value), "", Dtbb40("Name").Value)
                    AstekLookMode = IIf(IsDBNull(Dtbb40("Jamsostek").Value), "", Dtbb40("Jamsostek").Value)

                    If AstekLookMode = "" Or AstekLookMode = "0" Then
                        AstekLookMode = "0"

                    End If

                    SalaryLookerNest()
                    CtrlShow = "0"

                    If Not NewSald(0) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(0) & " ') "
                        OpenTbl(ADb, Dtbb39a, SQL)
                        If Dtbb39a.RecordCount > 0 Then
                            Dtbb39a.MoveFirst()

                            SalaryLooker1 = IIf(IsDBNull(Dtbb39a("Salary").Value), "", Dtbb39a("Salary").Value)
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(1) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(1) & "') "
                        OpenTbl(ADb, Dtbb39b, SQL)
                        If Dtbb39b.RecordCount > 0 Then
                            Dtbb39b.MoveFirst()

                            SalaryLooker2 = IIf(IsDBNull(Dtbb39b("Salary").Value), "", Dtbb39b("Salary").Value)
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(2) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(2) & "') "
                        OpenTbl(ADb, Dtbb39c, SQL)
                        If Dtbb39c.RecordCount > 0 Then
                            Dtbb39c.MoveFirst()

                            SalaryLooker3 = IIf(IsDBNull(Dtbb39c("Salary").Value), "", Dtbb39c("Salary").Value)
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(3) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(3) & "') "
                        OpenTbl(ADb, Dtbb39d, SQL)
                        If Dtbb39d.RecordCount > 0 Then
                            Dtbb39d.MoveFirst()

                            SalaryLooker4 = IIf(IsDBNull(Dtbb39d("Salary").Value), "", Dtbb39d("Salary").Value)
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(4) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(4) & "') "
                        OpenTbl(ADb, Dtbb39f, SQL)
                        If Dtbb39f.RecordCount > 0 Then
                            Dtbb39f.MoveFirst()

                            SalaryLooker5 = IIf(IsDBNull(Dtbb39f("Salary").Value), "", Dtbb39f("Salary").Value)
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(5) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(5) & "') "
                        OpenTbl(ADb, Dtbb39g, SQL)
                        If Dtbb39g.RecordCount > 0 Then
                            Dtbb39g.MoveFirst()

                            SalaryLooker6 = IIf(IsDBNull(Dtbb39g("Salary").Value), "", Dtbb39g("Salary").Value)
                            CtrlShow = CtrlShow + 1


                        End If
                    End If

                    If Not NewSald(6) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(6) & "') "
                        OpenTbl(ADb, Dtbb39h, SQL)
                        If Dtbb39h.RecordCount > 0 Then
                            Dtbb39h.MoveFirst()

                            SalaryLooker7 = IIf(IsDBNull(Dtbb39h("Salary").Value), "", Dtbb39h("Salary").Value)
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(7) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(7) & "') "
                        OpenTbl(ADb, Dtbb39i, SQL)
                        If Dtbb39i.RecordCount > 0 Then
                            Dtbb39i.MoveFirst()

                            SalaryLooker8 = IIf(IsDBNull(Dtbb39i("Salary").Value), "", Dtbb39i("Salary").Value)
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(8) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(8) & "') "
                        OpenTbl(ADb, Dtbb39j, SQL)
                        If Dtbb39j.RecordCount > 0 Then
                            Dtbb39j.MoveFirst()
                            SalaryLooker9 = IIf(IsDBNull(Dtbb39j("Salary").Value), "", Dtbb39j("Salary").Value)
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(9) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(9) & "') "
                        OpenTbl(ADb, Dtbb39k, SQL)
                        If Dtbb39k.RecordCount > 0 Then
                            Dtbb39k.MoveFirst()

                            SalaryLooker10 = IIf(IsDBNull(Dtbb39k("Salary").Value), "", Dtbb39k("Salary").Value)
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(10) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(10) & "') "
                        OpenTbl(ADb, Dtbb39l, SQL)
                        If Dtbb39l.RecordCount > 0 Then
                            Dtbb39l.MoveFirst()

                            SalaryLooker11 = IIf(IsDBNull(Dtbb39l("Salary").Value), "", Dtbb39l("Salary").Value)
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(11) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(11) & "') "
                        OpenTbl(ADb, Dtbb39m, SQL)
                        If Dtbb39m.RecordCount > 0 Then
                            Dtbb39m.MoveFirst()
                            SalaryLooker12 = IIf(IsDBNull(Dtbb39m("Salary").Value), "", Dtbb39m("Salary").Value)
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(12) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(12) & "') "
                        OpenTbl(ADb, Dtbb39n, SQL)
                        If Dtbb39n.RecordCount > 0 Then
                            Dtbb39n.MoveFirst()

                            SalaryLooker13 = IIf(IsDBNull(Dtbb39n("Salary").Value), "", Dtbb39n("Salary").Value)
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(13) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(13) & "') "
                        OpenTbl(ADb, Dtbb39o, SQL)
                        If Dtbb39o.RecordCount > 0 Then
                            Dtbb39o.MoveFirst()

                            SalaryLooker14 = IIf(IsDBNull(Dtbb39o("Salary").Value), "", Dtbb39o("Salary").Value)
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(14) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(14) & "') "
                        OpenTbl(ADb, Dtbb39p, SQL)
                        If Dtbb39p.RecordCount > 0 Then
                            Dtbb39p.MoveFirst()

                            SalaryLooker15 = IIf(IsDBNull(Dtbb39p("Salary").Value), "", Dtbb39p("Salary").Value)
                            CtrlShow = CtrlShow + 1

                        End If

                    End If

                    If Not NewSald(15) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(15) & "') "
                        OpenTbl(ADb, Dtbb39q, SQL)
                        If Dtbb39q.RecordCount > 0 Then
                            Dtbb39q.MoveFirst()

                            SalaryLooker16 = IIf(IsDBNull(Dtbb39q("Salary").Value), "", Dtbb39q("Salary").Value)
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If CtrlShow >= 1 Then

                        GetSalaryTot = Val(SalaryLooker1) + Val(SalaryLooker2) + Val(SalaryLooker3) + Val(SalaryLooker4) + Val(SalaryLooker5) + Val(SalaryLooker6) + Val(SalaryLooker7) + Val(SalaryLooker8) + Val(SalaryLooker9) + Val(SalaryLooker10) + Val(SalaryLooker11) + Val(SalaryLooker12) + Val(SalaryLooker13) + Val(SalaryLooker14) + Val(SalaryLooker15) + Val(SalaryLooker16)
                        FieldGridFormatNum()
                        PerFieldGrid01.Invoke(DirectCast(Sub() PerFieldGrid01.Rows.Add(NikLooker, NameLooker, NewFormT(0), NewFormT(1), NewFormT(2), NewFormT(3), NewFormT(4), NewFormT(5), NewFormT(6), NewFormT(7), NewFormT(8), NewFormT(9), NewFormT(10), NewFormT(11), NewFormT(12), NewFormT(13), NewFormT(14), NewFormT(15), "", "", AstekLookMode, "", "", "", "", Format(Val(GetSalaryTot), "N0"), TotRound2), MethodInvoker))
                        LetsCount()

                        TotalTbxFiller()

                    End If

                    Dtbb40.MoveNext()

                Loop

                MsgBox("Done")

            End If 'Termination of Dtbb40

        Else

            SQL = ""
            SQL = SQL & "Select * from 02_Name_Table "
            SQL = SQL & "Where Active = ('" & "Yes" & "') "
            SQL = SQL & "And Pay = ('" & Bracket & " ') "
            SQL = SQL & "Order by Nik "

            OpenTbl(ADb, Dtbb40, SQL)
            If Dtbb40.RecordCount <> 0 Then

                Dtbb40.MoveFirst()
                Do While Not Dtbb40.EOF

                    NikLooker = IIf(IsDBNull(Dtbb40("Nik").Value), "", Dtbb40("Nik").Value)
                    NameLooker = IIf(IsDBNull(Dtbb40("Name").Value), "", Dtbb40("Name").Value)
                    AstekLookMode = IIf(IsDBNull(Dtbb40("Jamsostek").Value), "", Dtbb40("Jamsostek").Value)

                    If AstekLookMode = "" Or AstekLookMode = "0" Then
                        AstekLookMode = "0"

                    End If

                    SalaryLookerNest()
                    CtrlShow = "0"

                    If Not NewSald(0) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(0) & " ') "
                        OpenTbl(ADb, Dtbb39a, SQL)
                        If Dtbb39a.RecordCount > 0 Then
                            Dtbb39a.MoveFirst()

                            SalaryLooker1 = Dtbb39a("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(1) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(1) & "') "
                        OpenTbl(ADb, Dtbb39b, SQL)
                        If Dtbb39b.RecordCount > 0 Then
                            Dtbb39b.MoveFirst()

                            SalaryLooker2 = Dtbb39b("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(2) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(2) & "') "
                        OpenTbl(ADb, Dtbb39c, SQL)
                        If Dtbb39c.RecordCount > 0 Then
                            Dtbb39c.MoveFirst()

                            SalaryLooker3 = Dtbb39c("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(3) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(3) & "') "
                        OpenTbl(ADb, Dtbb39d, SQL)
                        If Dtbb39d.RecordCount > 0 Then
                            Dtbb39d.MoveFirst()

                            SalaryLooker4 = Dtbb39d("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(4) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(4) & "') "
                        OpenTbl(ADb, Dtbb39f, SQL)
                        If Dtbb39f.RecordCount > 0 Then
                            Dtbb39f.MoveFirst()

                            SalaryLooker5 = Dtbb39f("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(5) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(5) & "') "
                        OpenTbl(ADb, Dtbb39g, SQL)
                        If Dtbb39g.RecordCount > 0 Then
                            Dtbb39g.MoveFirst()
                            SalaryLooker6 = Dtbb39g("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(6) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(6) & "') "
                        OpenTbl(ADb, Dtbb39h, SQL)
                        If Dtbb39h.RecordCount > 0 Then
                            Dtbb39h.MoveFirst()
                            SalaryLooker7 = Dtbb39h("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(7) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(7) & "') "
                        OpenTbl(ADb, Dtbb39i, SQL)
                        If Dtbb39i.RecordCount > 0 Then
                            Dtbb39i.MoveFirst()

                            SalaryLooker8 = Dtbb39i("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(8) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(8) & "') "
                        OpenTbl(ADb, Dtbb39j, SQL)
                        If Dtbb39j.RecordCount > 0 Then
                            Dtbb39j.MoveFirst()
                            SalaryLooker9 = Dtbb39j("Salary").Value
                            CtrlShow = CtrlShow + 1
                        End If
                    End If

                    If Not NewSald(9) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(9) & "') "
                        OpenTbl(ADb, Dtbb39k, SQL)
                        If Dtbb39k.RecordCount > 0 Then
                            Dtbb39k.MoveFirst()

                            SalaryLooker10 = Dtbb39k("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(10) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(10) & "') "
                        OpenTbl(ADb, Dtbb39l, SQL)
                        If Dtbb39l.RecordCount > 0 Then
                            Dtbb39l.MoveFirst()

                            SalaryLooker11 = Dtbb39l("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(11) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(11) & "') "
                        OpenTbl(ADb, Dtbb39m, SQL)
                        If Dtbb39m.RecordCount > 0 Then
                            Dtbb39m.MoveFirst()
                            SalaryLooker12 = Dtbb39m("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(12) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(12) & "') "
                        OpenTbl(ADb, Dtbb39n, SQL)
                        If Dtbb39n.RecordCount > 0 Then
                            Dtbb39n.MoveFirst()

                            SalaryLooker13 = Dtbb39n("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(13) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(13) & "') "
                        OpenTbl(ADb, Dtbb39o, SQL)
                        If Dtbb39o.RecordCount > 0 Then
                            Dtbb39o.MoveFirst()

                            SalaryLooker14 = Dtbb39o("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(14) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(14) & "') "
                        OpenTbl(ADb, Dtbb39p, SQL)
                        If Dtbb39p.RecordCount > 0 Then
                            Dtbb39p.MoveFirst()

                            SalaryLooker15 = Dtbb39p("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If

                    End If

                    If Not NewSald(15) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 13_Conveyour_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(15) & "') "
                        OpenTbl(ADb, Dtbb39q, SQL)
                        If Dtbb39q.RecordCount > 0 Then
                            Dtbb39q.MoveFirst()

                            SalaryLooker16 = Dtbb39q("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If CtrlShow >= 1 Then

                        GetSalaryTot = Val(SalaryLooker1) + Val(SalaryLooker2) + Val(SalaryLooker3) + Val(SalaryLooker4) + Val(SalaryLooker5) + Val(SalaryLooker6) + Val(SalaryLooker7) + Val(SalaryLooker8) + Val(SalaryLooker9) + Val(SalaryLooker10) + Val(SalaryLooker11) + Val(SalaryLooker12) + Val(SalaryLooker13) + Val(SalaryLooker14) + Val(SalaryLooker15) + Val(SalaryLooker16)
                        FieldGridFormatNum()
                        PerFieldGrid01.Invoke(DirectCast(Sub() PerFieldGrid01.Rows.Add(NikLooker, NameLooker, NewFormT(0), NewFormT(1), NewFormT(2), NewFormT(3), NewFormT(4), NewFormT(5), NewFormT(6), NewFormT(7), NewFormT(8), NewFormT(9), NewFormT(10), NewFormT(11), NewFormT(12), NewFormT(13), NewFormT(14), NewFormT(15), "", "", AstekLookMode, "", "", "", "", Format(Val(GetSalaryTot), "N0"), TotRound2), MethodInvoker))
                        LetsCount()

                        TotalTbxFiller()

                    End If

                    Dtbb40.MoveNext()
                Loop

                MsgBox("Done")

            End If 'Termination of Dtbb40

        End If

    End Sub

    Sub PerDataLoader2()

        NikLooker = ""
        NameLooker = ""
        SalaryLooker1 = ""
        RecordCounting = "0"

        If Bracket = "All" Then

            SQL = ""
            SQL = SQL & "Select * from 02_Name_Table "
            SQL = SQL & "Where Active = ('" & "Yes" & "') "
            SQL = SQL & "Order by Nik "
            OpenTbl(ADb, Dtbb40, SQL)
            If Dtbb40.RecordCount <> 0 Then


                Dtbb40.MoveFirst()
                Do While Not Dtbb40.EOF

                    NikLooker = IIf(IsDBNull(Dtbb40("Nik").Value), "", Dtbb40("Nik").Value)
                    NameLooker = IIf(IsDBNull(Dtbb40("Name").Value), "", Dtbb40("Name").Value)
                    AstekLookMode = IIf(IsDBNull(Dtbb40("Jamsostek").Value), "", Dtbb40("Jamsostek").Value)

                    If AstekLookMode = "" Or AstekLookMode = "0" Then
                        AstekLookMode = "0"

                    End If

                    SalaryLookerNest()
                    CtrlShow = "0"

                    If Not NewSald(0) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(0) & " ') "
                        OpenTbl(ADb, Dtbb39a, SQL)
                        If Dtbb39a.RecordCount > 0 Then
                            Dtbb39a.MoveFirst()

                            SalaryLooker1 = Dtbb39a("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If


                    If Not NewSald(1) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(1) & "') "
                        OpenTbl(ADb, Dtbb39b, SQL)
                        If Dtbb39b.RecordCount > 0 Then
                            Dtbb39b.MoveFirst()

                            SalaryLooker2 = Dtbb39b("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If


                    If Not NewSald(2) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(2) & "') "
                        OpenTbl(ADb, Dtbb39c, SQL)
                        If Dtbb39c.RecordCount > 0 Then
                            Dtbb39c.MoveFirst()

                            SalaryLooker3 = Dtbb39c("Salary").Value
                            CtrlShow = CtrlShow + 1
                        End If
                    End If

                    If Not NewSald(3) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(3) & "') "
                        OpenTbl(ADb, Dtbb39d, SQL)
                        If Dtbb39d.RecordCount > 0 Then
                            Dtbb39d.MoveFirst()

                            SalaryLooker4 = Dtbb39d("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If


                    If Not NewSald(4) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(4) & "') "
                        OpenTbl(ADb, Dtbb39f, SQL)
                        If Dtbb39f.RecordCount > 0 Then
                            Dtbb39f.MoveFirst()

                            SalaryLooker5 = Dtbb39f("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(5) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(5) & "') "
                        OpenTbl(ADb, Dtbb39g, SQL)
                        If Dtbb39g.RecordCount > 0 Then
                            Dtbb39g.MoveFirst()
                            SalaryLooker6 = Dtbb39g("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(6) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(6) & "') "
                        OpenTbl(ADb, Dtbb39h, SQL)
                        If Dtbb39h.RecordCount > 0 Then
                            Dtbb39h.MoveFirst()
                            SalaryLooker7 = Dtbb39h("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(7) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(7) & "') "
                        OpenTbl(ADb, Dtbb39i, SQL)
                        If Dtbb39i.RecordCount > 0 Then
                            Dtbb39i.MoveFirst()

                            SalaryLooker8 = Dtbb39i("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(8) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(8) & "') "
                        OpenTbl(ADb, Dtbb39j, SQL)
                        If Dtbb39j.RecordCount > 0 Then
                            Dtbb39j.MoveFirst()
                            SalaryLooker9 = Dtbb39j("Salary").Value
                            CtrlShow = CtrlShow + 1
                        End If
                    End If

                    If Not NewSald(9) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(9) & "') "
                        OpenTbl(ADb, Dtbb39k, SQL)
                        If Dtbb39k.RecordCount > 0 Then
                            Dtbb39k.MoveFirst()

                            SalaryLooker10 = Dtbb39k("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(10) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(10) & "') "
                        OpenTbl(ADb, Dtbb39l, SQL)
                        If Dtbb39l.RecordCount > 0 Then
                            Dtbb39l.MoveFirst()

                            SalaryLooker11 = Dtbb39l("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(11) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(11) & "') "
                        OpenTbl(ADb, Dtbb39m, SQL)
                        If Dtbb39m.RecordCount > 0 Then
                            Dtbb39m.MoveFirst()
                            SalaryLooker12 = Dtbb39m("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(12) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(12) & "') "
                        OpenTbl(ADb, Dtbb39n, SQL)
                        If Dtbb39n.RecordCount > 0 Then
                            Dtbb39n.MoveFirst()

                            SalaryLooker13 = Dtbb39n("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(13) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(13) & "') "
                        OpenTbl(ADb, Dtbb39o, SQL)
                        If Dtbb39o.RecordCount > 0 Then
                            Dtbb39o.MoveFirst()

                            SalaryLooker14 = Dtbb39o("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(14) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(14) & "') "
                        OpenTbl(ADb, Dtbb39p, SQL)
                        If Dtbb39p.RecordCount > 0 Then
                            Dtbb39p.MoveFirst()

                            SalaryLooker15 = Dtbb39p("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If

                    End If

                    If Not NewSald(15) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(15) & "') "
                        OpenTbl(ADb, Dtbb39q, SQL)
                        If Dtbb39q.RecordCount > 0 Then
                            Dtbb39q.MoveFirst()

                            SalaryLooker16 = Dtbb39q("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If CtrlShow >= 1 Then

                        GetSalaryTot = Val(SalaryLooker1) + Val(SalaryLooker2) + Val(SalaryLooker3) + Val(SalaryLooker4) + Val(SalaryLooker5) + Val(SalaryLooker6) + Val(SalaryLooker7) + Val(SalaryLooker8) + Val(SalaryLooker9) + Val(SalaryLooker10) + Val(SalaryLooker11) + Val(SalaryLooker12) + Val(SalaryLooker13) + Val(SalaryLooker14) + Val(SalaryLooker15) + Val(SalaryLooker16)
                        FieldGridFormatNum()
                        PerFieldGrid01.Invoke(DirectCast(Sub() PerFieldGrid01.Rows.Add(NikLooker, NameLooker, NewFormT(0), NewFormT(1), NewFormT(2), NewFormT(3), NewFormT(4), NewFormT(5), NewFormT(6), NewFormT(7), NewFormT(8), NewFormT(9), NewFormT(10), NewFormT(11), NewFormT(12), NewFormT(13), NewFormT(14), NewFormT(15), "", "", AstekLookMode, "", "", "", "", Format(Val(GetSalaryTot), "N0"), TotRound2), MethodInvoker))
                        LetsCount()

                        TotalTbxFiller()

                    End If

                    Dtbb40.MoveNext()
                Loop

                MsgBox("Done")

            End If 'Termination of Dtbb40

        Else

            SQL = ""
            SQL = SQL & "Select * from 02_Name_Table "
            SQL = SQL & "Where Active = ('" & "Yes" & "') "
            SQL = SQL & "And Pay = ('" & Bracket & " ') "
            SQL = SQL & "Order by Nik "
            OpenTbl(ADb, Dtbb40, SQL)
            If Dtbb40.RecordCount <> 0 Then


                Dtbb40.MoveFirst()
                Do While Not Dtbb40.EOF

                    NikLooker = IIf(IsDBNull(Dtbb40("Nik").Value), "", Dtbb40("Nik").Value)
                    NameLooker = IIf(IsDBNull(Dtbb40("Name").Value), "", Dtbb40("Name").Value)
                    AstekLookMode = IIf(IsDBNull(Dtbb40("Jamsostek").Value), "", Dtbb40("Jamsostek").Value)

                    If AstekLookMode = "" Or AstekLookMode = "0" Then
                        AstekLookMode = "0"

                    End If

                    SalaryLookerNest()
                    CtrlShow = "0"

                    If Not NewSald(0) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(0) & " ') "
                        OpenTbl(ADb, Dtbb39a, SQL)
                        If Dtbb39a.RecordCount > 0 Then
                            Dtbb39a.MoveFirst()

                            SalaryLooker1 = Dtbb39a("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If


                    If Not NewSald(1) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(1) & "') "
                        OpenTbl(ADb, Dtbb39b, SQL)
                        If Dtbb39b.RecordCount > 0 Then
                            Dtbb39b.MoveFirst()

                            SalaryLooker2 = Dtbb39b("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If


                    If Not NewSald(2) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(2) & "') "
                        OpenTbl(ADb, Dtbb39c, SQL)
                        If Dtbb39c.RecordCount > 0 Then
                            Dtbb39c.MoveFirst()

                            SalaryLooker3 = Dtbb39c("Salary").Value
                            CtrlShow = CtrlShow + 1
                        End If
                    End If


                    If Not NewSald(3) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(3) & "') "
                        OpenTbl(ADb, Dtbb39d, SQL)
                        If Dtbb39d.RecordCount > 0 Then
                            Dtbb39d.MoveFirst()

                            SalaryLooker4 = Dtbb39d("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If


                    If Not NewSald(4) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(4) & "') "
                        OpenTbl(ADb, Dtbb39f, SQL)
                        If Dtbb39f.RecordCount > 0 Then
                            Dtbb39f.MoveFirst()

                            SalaryLooker5 = Dtbb39f("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If


                    If Not NewSald(5) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(5) & "') "
                        OpenTbl(ADb, Dtbb39g, SQL)
                        If Dtbb39g.RecordCount > 0 Then
                            Dtbb39g.MoveFirst()
                            SalaryLooker6 = Dtbb39g("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If


                    If Not NewSald(6) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(6) & "') "
                        OpenTbl(ADb, Dtbb39h, SQL)
                        If Dtbb39h.RecordCount > 0 Then
                            Dtbb39h.MoveFirst()
                            SalaryLooker7 = Dtbb39h("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(7) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(7) & "') "
                        OpenTbl(ADb, Dtbb39i, SQL)
                        If Dtbb39i.RecordCount > 0 Then
                            Dtbb39i.MoveFirst()

                            SalaryLooker8 = Dtbb39i("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(8) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(8) & "') "
                        OpenTbl(ADb, Dtbb39j, SQL)
                        If Dtbb39j.RecordCount > 0 Then
                            Dtbb39j.MoveFirst()
                            SalaryLooker9 = Dtbb39j("Salary").Value
                            CtrlShow = CtrlShow + 1
                        End If
                    End If

                    If Not NewSald(9) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(9) & "') "
                        OpenTbl(ADb, Dtbb39k, SQL)
                        If Dtbb39k.RecordCount > 0 Then
                            Dtbb39k.MoveFirst()

                            SalaryLooker10 = Dtbb39k("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(10) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(10) & "') "
                        OpenTbl(ADb, Dtbb39l, SQL)
                        If Dtbb39l.RecordCount > 0 Then
                            Dtbb39l.MoveFirst()

                            SalaryLooker11 = Dtbb39l("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(11) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(11) & "') "
                        OpenTbl(ADb, Dtbb39m, SQL)
                        If Dtbb39m.RecordCount > 0 Then
                            Dtbb39m.MoveFirst()
                            SalaryLooker12 = Dtbb39m("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(12) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(12) & "') "
                        OpenTbl(ADb, Dtbb39n, SQL)
                        If Dtbb39n.RecordCount > 0 Then
                            Dtbb39n.MoveFirst()

                            SalaryLooker13 = Dtbb39n("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(13) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(13) & "') "
                        OpenTbl(ADb, Dtbb39o, SQL)
                        If Dtbb39o.RecordCount > 0 Then
                            Dtbb39o.MoveFirst()

                            SalaryLooker14 = Dtbb39o("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(14) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(14) & "') "
                        OpenTbl(ADb, Dtbb39p, SQL)
                        If Dtbb39p.RecordCount > 0 Then
                            Dtbb39p.MoveFirst()

                            SalaryLooker15 = Dtbb39p("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If

                    End If

                    If Not NewSald(15) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 14_MutuII_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(15) & "') "
                        OpenTbl(ADb, Dtbb39q, SQL)
                        If Dtbb39q.RecordCount > 0 Then
                            Dtbb39q.MoveFirst()

                            SalaryLooker16 = Dtbb39q("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If CtrlShow >= 1 Then

                        GetSalaryTot = Val(SalaryLooker1) + Val(SalaryLooker2) + Val(SalaryLooker3) + Val(SalaryLooker4) + Val(SalaryLooker5) + Val(SalaryLooker6) + Val(SalaryLooker7) + Val(SalaryLooker8) + Val(SalaryLooker9) + Val(SalaryLooker10) + Val(SalaryLooker11) + Val(SalaryLooker12) + Val(SalaryLooker13) + Val(SalaryLooker14) + Val(SalaryLooker15) + Val(SalaryLooker16)
                        FieldGridFormatNum()
                        PerFieldGrid01.Invoke(DirectCast(Sub() PerFieldGrid01.Rows.Add(NikLooker, NameLooker, NewFormT(0), NewFormT(1), NewFormT(2), NewFormT(3), NewFormT(4), NewFormT(5), NewFormT(6), NewFormT(7), NewFormT(8), NewFormT(9), NewFormT(10), NewFormT(11), NewFormT(12), NewFormT(13), NewFormT(14), NewFormT(15), "", "", AstekLookMode, "", "", "", "", Format(Val(GetSalaryTot), "N0"), TotRound2), MethodInvoker))
                        LetsCount()

                        TotalTbxFiller()

                    End If

                    Dtbb40.MoveNext()
                Loop

                MsgBox("Done")

            End If 'Termination of Dtbb40

        End If


    End Sub

    Sub PerDataLoader3()
        NikLooker = ""
        NameLooker = ""
        SalaryLooker1 = ""
        RecordCounting = "0"


        If Bracket = "All" Then

            SQL = ""
            SQL = SQL & "Select * from 02_Name_Table "
            SQL = SQL & "Where Active = ('" & "Yes" & "') "
            SQL = SQL & "Order by Nik "
            OpenTbl(ADb, Dtbb40, SQL)
            If Dtbb40.RecordCount <> 0 Then


                Dtbb40.MoveFirst()
                Do While Not Dtbb40.EOF

                    NikLooker = IIf(IsDBNull(Dtbb40("Nik").Value), "", Dtbb40("Nik").Value)
                    NameLooker = IIf(IsDBNull(Dtbb40("Name").Value), "", Dtbb40("Name").Value)
                    AstekLookMode = IIf(IsDBNull(Dtbb40("Jamsostek").Value), "", Dtbb40("Jamsostek").Value)

                    If AstekLookMode = "" Or AstekLookMode = "0" Then
                        AstekLookMode = "0"

                    End If

                    SalaryLookerNest()
                    CtrlShow = "0"

                    If Not NewSald(0) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(0) & " ') "
                        OpenTbl(ADb, Dtbb39a, SQL)
                        If Dtbb39a.RecordCount > 0 Then
                            Dtbb39a.MoveFirst()

                            SalaryLooker1 = Dtbb39a("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(1) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(1) & "') "
                        OpenTbl(ADb, Dtbb39b, SQL)
                        If Dtbb39b.RecordCount > 0 Then
                            Dtbb39b.MoveFirst()

                            SalaryLooker2 = Dtbb39b("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(2) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(2) & "') "
                        OpenTbl(ADb, Dtbb39c, SQL)
                        If Dtbb39c.RecordCount > 0 Then
                            Dtbb39c.MoveFirst()

                            SalaryLooker3 = Dtbb39c("Salary").Value
                            CtrlShow = CtrlShow + 1
                        End If
                    End If

                    If Not NewSald(3) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(3) & "') "
                        OpenTbl(ADb, Dtbb39d, SQL)
                        If Dtbb39d.RecordCount > 0 Then
                            Dtbb39d.MoveFirst()

                            SalaryLooker4 = Dtbb39d("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(4) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(4) & "') "
                        OpenTbl(ADb, Dtbb39f, SQL)
                        If Dtbb39f.RecordCount > 0 Then
                            Dtbb39f.MoveFirst()

                            SalaryLooker5 = Dtbb39f("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(5) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(5) & "') "
                        OpenTbl(ADb, Dtbb39g, SQL)
                        If Dtbb39g.RecordCount > 0 Then
                            Dtbb39g.MoveFirst()
                            SalaryLooker6 = Dtbb39g("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(6) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(6) & "') "
                        OpenTbl(ADb, Dtbb39h, SQL)
                        If Dtbb39h.RecordCount > 0 Then
                            Dtbb39h.MoveFirst()
                            SalaryLooker7 = Dtbb39h("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(7) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(7) & "') "
                        OpenTbl(ADb, Dtbb39i, SQL)
                        If Dtbb39i.RecordCount > 0 Then
                            Dtbb39i.MoveFirst()

                            SalaryLooker8 = Dtbb39i("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(8) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(8) & "') "
                        OpenTbl(ADb, Dtbb39j, SQL)
                        If Dtbb39j.RecordCount > 0 Then
                            Dtbb39j.MoveFirst()
                            SalaryLooker9 = Dtbb39j("Salary").Value
                            CtrlShow = CtrlShow + 1
                        End If
                    End If

                    If Not NewSald(9) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(9) & "') "
                        OpenTbl(ADb, Dtbb39k, SQL)
                        If Dtbb39k.RecordCount > 0 Then
                            Dtbb39k.MoveFirst()

                            SalaryLooker10 = Dtbb39k("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(10) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(10) & "') "
                        OpenTbl(ADb, Dtbb39l, SQL)
                        If Dtbb39l.RecordCount > 0 Then
                            Dtbb39l.MoveFirst()

                            SalaryLooker11 = Dtbb39l("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(11) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(11) & "') "
                        OpenTbl(ADb, Dtbb39m, SQL)
                        If Dtbb39m.RecordCount > 0 Then
                            Dtbb39m.MoveFirst()
                            SalaryLooker12 = Dtbb39m("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(12) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(12) & "') "
                        OpenTbl(ADb, Dtbb39n, SQL)
                        If Dtbb39n.RecordCount > 0 Then
                            Dtbb39n.MoveFirst()

                            SalaryLooker13 = Dtbb39n("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(13) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(13) & "') "
                        OpenTbl(ADb, Dtbb39o, SQL)
                        If Dtbb39o.RecordCount > 0 Then
                            Dtbb39o.MoveFirst()

                            SalaryLooker14 = Dtbb39o("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(14) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(14) & "') "
                        OpenTbl(ADb, Dtbb39p, SQL)
                        If Dtbb39p.RecordCount > 0 Then
                            Dtbb39p.MoveFirst()

                            SalaryLooker15 = Dtbb39p("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If

                    End If

                    If Not NewSald(15) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(15) & "') "
                        OpenTbl(ADb, Dtbb39q, SQL)
                        If Dtbb39q.RecordCount > 0 Then
                            Dtbb39q.MoveFirst()

                            SalaryLooker16 = Dtbb39q("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If CtrlShow >= 1 Then

                        GetSalaryTot = Val(SalaryLooker1) + Val(SalaryLooker2) + Val(SalaryLooker3) + Val(SalaryLooker4) + Val(SalaryLooker5) + Val(SalaryLooker6) + Val(SalaryLooker7) + Val(SalaryLooker8) + Val(SalaryLooker9) + Val(SalaryLooker10) + Val(SalaryLooker11) + Val(SalaryLooker12) + Val(SalaryLooker13) + Val(SalaryLooker14) + Val(SalaryLooker15) + Val(SalaryLooker16)
                        FieldGridFormatNum()
                        PerFieldGrid01.Invoke(DirectCast(Sub() PerFieldGrid01.Rows.Add(NikLooker, NameLooker, NewFormT(0), NewFormT(1), NewFormT(2), NewFormT(3), NewFormT(4), NewFormT(5), NewFormT(6), NewFormT(7), NewFormT(8), NewFormT(9), NewFormT(10), NewFormT(11), NewFormT(12), NewFormT(13), NewFormT(14), NewFormT(15), "", "", AstekLookMode, "", "", "", "", Format(Val(GetSalaryTot), "N0"), TotRound2), MethodInvoker))
                        LetsCount()

                        TotalTbxFiller()

                    End If

                    Dtbb40.MoveNext()
                Loop

                MsgBox("Done")

            End If 'Termination of Dtbb40

        Else

            SQL = ""
            SQL = SQL & "Select * from 02_Name_Table "
            SQL = SQL & "Where Active = ('" & "Yes" & "') "
            SQL = SQL & "And Pay = ('" & Bracket & " ') "
            SQL = SQL & "Order by Nik "
            OpenTbl(ADb, Dtbb40, SQL)
            If Dtbb40.RecordCount <> 0 Then


                Dtbb40.MoveFirst()
                Do While Not Dtbb40.EOF

                    NikLooker = IIf(IsDBNull(Dtbb40("Nik").Value), "", Dtbb40("Nik").Value)
                    NameLooker = IIf(IsDBNull(Dtbb40("Name").Value), "", Dtbb40("Name").Value)
                    AstekLookMode = IIf(IsDBNull(Dtbb40("Jamsostek").Value), "", Dtbb40("Jamsostek").Value)

                    If AstekLookMode = "" Or AstekLookMode = "0" Then
                        AstekLookMode = "0"

                    End If

                    SalaryLookerNest()
                    CtrlShow = "0"

                    If Not NewSald(0) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(0) & " ') "
                        OpenTbl(ADb, Dtbb39a, SQL)
                        If Dtbb39a.RecordCount > 0 Then
                            Dtbb39a.MoveFirst()

                            SalaryLooker1 = Dtbb39a("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(1) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(1) & "') "
                        OpenTbl(ADb, Dtbb39b, SQL)
                        If Dtbb39b.RecordCount > 0 Then
                            Dtbb39b.MoveFirst()

                            SalaryLooker2 = Dtbb39b("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(2) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(2) & "') "
                        OpenTbl(ADb, Dtbb39c, SQL)
                        If Dtbb39c.RecordCount > 0 Then
                            Dtbb39c.MoveFirst()

                            SalaryLooker3 = Dtbb39c("Salary").Value
                            CtrlShow = CtrlShow + 1
                        End If
                    End If

                    If Not NewSald(3) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(3) & "') "
                        OpenTbl(ADb, Dtbb39d, SQL)
                        If Dtbb39d.RecordCount > 0 Then
                            Dtbb39d.MoveFirst()

                            SalaryLooker4 = Dtbb39d("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(4) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(4) & "') "
                        OpenTbl(ADb, Dtbb39f, SQL)
                        If Dtbb39f.RecordCount > 0 Then
                            Dtbb39f.MoveFirst()

                            SalaryLooker5 = Dtbb39f("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(5) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(5) & "') "
                        OpenTbl(ADb, Dtbb39g, SQL)
                        If Dtbb39g.RecordCount > 0 Then
                            Dtbb39g.MoveFirst()
                            SalaryLooker6 = Dtbb39g("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(6) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(6) & "') "
                        OpenTbl(ADb, Dtbb39h, SQL)
                        If Dtbb39h.RecordCount > 0 Then
                            Dtbb39h.MoveFirst()
                            SalaryLooker7 = Dtbb39h("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(7) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(7) & "') "
                        OpenTbl(ADb, Dtbb39i, SQL)
                        If Dtbb39i.RecordCount > 0 Then
                            Dtbb39i.MoveFirst()

                            SalaryLooker8 = Dtbb39i("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(8) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(8) & "') "
                        OpenTbl(ADb, Dtbb39j, SQL)
                        If Dtbb39j.RecordCount > 0 Then
                            Dtbb39j.MoveFirst()
                            SalaryLooker9 = Dtbb39j("Salary").Value
                            CtrlShow = CtrlShow + 1
                        End If
                    End If

                    If Not NewSald(9) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(9) & "') "
                        OpenTbl(ADb, Dtbb39k, SQL)
                        If Dtbb39k.RecordCount > 0 Then
                            Dtbb39k.MoveFirst()

                            SalaryLooker10 = Dtbb39k("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(10) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(10) & "') "
                        OpenTbl(ADb, Dtbb39l, SQL)
                        If Dtbb39l.RecordCount > 0 Then
                            Dtbb39l.MoveFirst()

                            SalaryLooker11 = Dtbb39l("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(11) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(11) & "') "
                        OpenTbl(ADb, Dtbb39m, SQL)
                        If Dtbb39m.RecordCount > 0 Then
                            Dtbb39m.MoveFirst()
                            SalaryLooker12 = Dtbb39m("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(12) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(12) & "') "
                        OpenTbl(ADb, Dtbb39n, SQL)
                        If Dtbb39n.RecordCount > 0 Then
                            Dtbb39n.MoveFirst()

                            SalaryLooker13 = Dtbb39n("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(13) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(13) & "') "
                        OpenTbl(ADb, Dtbb39o, SQL)
                        If Dtbb39o.RecordCount > 0 Then
                            Dtbb39o.MoveFirst()

                            SalaryLooker14 = Dtbb39o("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(14) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(14) & "') "
                        OpenTbl(ADb, Dtbb39p, SQL)
                        If Dtbb39p.RecordCount > 0 Then
                            Dtbb39p.MoveFirst()

                            SalaryLooker15 = Dtbb39p("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If

                    End If

                    If Not NewSald(15) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 16_Packing_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(15) & "') "
                        OpenTbl(ADb, Dtbb39q, SQL)
                        If Dtbb39q.RecordCount > 0 Then
                            Dtbb39q.MoveFirst()

                            SalaryLooker16 = Dtbb39q("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If CtrlShow >= 1 Then

                        GetSalaryTot = Val(SalaryLooker1) + Val(SalaryLooker2) + Val(SalaryLooker3) + Val(SalaryLooker4) + Val(SalaryLooker5) + Val(SalaryLooker6) + Val(SalaryLooker7) + Val(SalaryLooker8) + Val(SalaryLooker9) + Val(SalaryLooker10) + Val(SalaryLooker11) + Val(SalaryLooker12) + Val(SalaryLooker13) + Val(SalaryLooker14) + Val(SalaryLooker15) + Val(SalaryLooker16)
                        FieldGridFormatNum()
                        PerFieldGrid01.Invoke(DirectCast(Sub() PerFieldGrid01.Rows.Add(NikLooker, NameLooker, NewFormT(0), NewFormT(1), NewFormT(2), NewFormT(3), NewFormT(4), NewFormT(5), NewFormT(6), NewFormT(7), NewFormT(8), NewFormT(9), NewFormT(10), NewFormT(11), NewFormT(12), NewFormT(13), NewFormT(14), NewFormT(15), "", "", AstekLookMode, "", "", "", "", Format(Val(GetSalaryTot), "N0"), TotRound2), MethodInvoker))
                        LetsCount()

                        TotalTbxFiller()

                    End If

                    Dtbb40.MoveNext()

                Loop

                MsgBox("Done")

            End If 'Termination of Dtbb40

        End If


    End Sub

    Sub PerDataLoader4()

        NikLooker = ""
        NameLooker = ""
        SalaryLooker1 = ""
        RecordCounting = "0"


        If Bracket = "All" Then

            SQL = ""
            SQL = SQL & "Select * from 02_Name_Table "
            SQL = SQL & "Where Active = ('" & "Yes" & "') "
            SQL = SQL & "Order by Nik "
            OpenTbl(ADb, Dtbb40, SQL)
            If Dtbb40.RecordCount <> 0 Then


                Dtbb40.MoveFirst()
                Do While Not Dtbb40.EOF

                    NikLooker = IIf(IsDBNull(Dtbb40("Nik").Value), "", Dtbb40("Nik").Value)
                    NameLooker = IIf(IsDBNull(Dtbb40("Name").Value), "", Dtbb40("Name").Value)
                    AstekLookMode = IIf(IsDBNull(Dtbb40("Jamsostek").Value), "", Dtbb40("Jamsostek").Value)

                    SalaryLookerNest()
                    CtrlShow = "0"

                    If Not NewSald(0) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(0) & " ') "
                        OpenTbl(ADb, Dtbb39a, SQL)
                        If Dtbb39a.RecordCount > 0 Then
                            Dtbb39a.MoveFirst()

                            SalaryLooker1 = Dtbb39a("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If


                    If Not NewSald(1) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(1) & "') "
                        OpenTbl(ADb, Dtbb39b, SQL)
                        If Dtbb39b.RecordCount > 0 Then
                            Dtbb39b.MoveFirst()

                            SalaryLooker2 = Dtbb39b("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If


                    If Not NewSald(2) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(2) & "') "
                        OpenTbl(ADb, Dtbb39c, SQL)
                        If Dtbb39c.RecordCount > 0 Then
                            Dtbb39c.MoveFirst()

                            SalaryLooker3 = Dtbb39c("Salary").Value
                            CtrlShow = CtrlShow + 1
                        End If
                    End If


                    If Not NewSald(3) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(3) & "') "
                        OpenTbl(ADb, Dtbb39d, SQL)
                        If Dtbb39d.RecordCount > 0 Then
                            Dtbb39d.MoveFirst()

                            SalaryLooker4 = Dtbb39d("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If


                    If Not NewSald(4) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(4) & "') "
                        OpenTbl(ADb, Dtbb39f, SQL)
                        If Dtbb39f.RecordCount > 0 Then
                            Dtbb39f.MoveFirst()

                            SalaryLooker5 = Dtbb39f("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If


                    If Not NewSald(5) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(5) & "') "
                        OpenTbl(ADb, Dtbb39g, SQL)
                        If Dtbb39g.RecordCount > 0 Then
                            Dtbb39g.MoveFirst()
                            SalaryLooker6 = Dtbb39g("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If


                    If Not NewSald(6) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(6) & "') "
                        OpenTbl(ADb, Dtbb39h, SQL)
                        If Dtbb39h.RecordCount > 0 Then
                            Dtbb39h.MoveFirst()
                            SalaryLooker7 = Dtbb39h("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If


                    If Not NewSald(7) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(7) & "') "
                        OpenTbl(ADb, Dtbb39i, SQL)
                        If Dtbb39i.RecordCount > 0 Then
                            Dtbb39i.MoveFirst()

                            SalaryLooker8 = Dtbb39i("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(8) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(8) & "') "
                        OpenTbl(ADb, Dtbb39j, SQL)
                        If Dtbb39j.RecordCount > 0 Then
                            Dtbb39j.MoveFirst()
                            SalaryLooker9 = Dtbb39j("Salary").Value
                            CtrlShow = CtrlShow + 1
                        End If
                    End If

                    If Not NewSald(9) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(9) & "') "
                        OpenTbl(ADb, Dtbb39k, SQL)
                        If Dtbb39k.RecordCount > 0 Then
                            Dtbb39k.MoveFirst()

                            SalaryLooker10 = Dtbb39k("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(10) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(10) & "') "
                        OpenTbl(ADb, Dtbb39l, SQL)
                        If Dtbb39l.RecordCount > 0 Then
                            Dtbb39l.MoveFirst()

                            SalaryLooker11 = Dtbb39l("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(11) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(11) & "') "
                        OpenTbl(ADb, Dtbb39m, SQL)
                        If Dtbb39m.RecordCount > 0 Then
                            Dtbb39m.MoveFirst()
                            SalaryLooker12 = Dtbb39m("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(12) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(12) & "') "
                        OpenTbl(ADb, Dtbb39n, SQL)
                        If Dtbb39n.RecordCount > 0 Then
                            Dtbb39n.MoveFirst()

                            SalaryLooker13 = Dtbb39n("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(13) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(13) & "') "
                        OpenTbl(ADb, Dtbb39o, SQL)
                        If Dtbb39o.RecordCount > 0 Then
                            Dtbb39o.MoveFirst()

                            SalaryLooker14 = Dtbb39o("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(14) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(14) & "') "
                        OpenTbl(ADb, Dtbb39p, SQL)
                        If Dtbb39p.RecordCount > 0 Then
                            Dtbb39p.MoveFirst()

                            SalaryLooker15 = Dtbb39p("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If

                    End If

                    If Not NewSald(15) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(15) & "') "
                        OpenTbl(ADb, Dtbb39q, SQL)
                        If Dtbb39q.RecordCount > 0 Then
                            Dtbb39q.MoveFirst()

                            SalaryLooker16 = Dtbb39q("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If CtrlShow >= 1 Then

                        GetSalaryTot = Val(SalaryLooker1) + Val(SalaryLooker2) + Val(SalaryLooker3) + Val(SalaryLooker4) + Val(SalaryLooker5) + Val(SalaryLooker6) + Val(SalaryLooker7) + Val(SalaryLooker8) + Val(SalaryLooker9) + Val(SalaryLooker10) + Val(SalaryLooker11) + Val(SalaryLooker12) + Val(SalaryLooker13) + Val(SalaryLooker14) + Val(SalaryLooker15) + Val(SalaryLooker16)
                        FieldGridFormatNum()
                        PerFieldGrid01.Invoke(DirectCast(Sub() PerFieldGrid01.Rows.Add(NikLooker, NameLooker, NewFormT(0), NewFormT(1), NewFormT(2), NewFormT(3), NewFormT(4), NewFormT(5), NewFormT(6), NewFormT(7), NewFormT(8), NewFormT(9), NewFormT(10), NewFormT(11), NewFormT(12), NewFormT(13), NewFormT(14), NewFormT(15), "", "", AstekLookMode, "", "", "", "", Format(Val(GetSalaryTot), "N0"), TotRound2), MethodInvoker))
                        LetsCount()

                        TotalTbxFiller()

                    End If

                    Dtbb40.MoveNext()
                Loop

                MsgBox("Done")

            End If 'Termination of Dtbb40

        Else

            SQL = ""
            SQL = SQL & "Select * from 02_Name_Table "
            SQL = SQL & "Where Active = ('" & "Yes" & "') "
            SQL = SQL & "And Pay = ('" & Bracket & " ') "
            SQL = SQL & "Order by Nik "
            OpenTbl(ADb, Dtbb40, SQL)
            If Dtbb40.RecordCount <> 0 Then

                Dtbb40.MoveFirst()
                Do While Not Dtbb40.EOF

                    NikLooker = IIf(IsDBNull(Dtbb40("Nik").Value), "", Dtbb40("Nik").Value)
                    NameLooker = IIf(IsDBNull(Dtbb40("Name").Value), "", Dtbb40("Name").Value)
                    AstekLookMode = IIf(IsDBNull(Dtbb40("Jamsostek").Value), "", Dtbb40("Jamsostek").Value)

                    SalaryLookerNest()
                    CtrlShow = "0"

                    If Not NewSald(0) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(0) & " ') "
                        OpenTbl(ADb, Dtbb39a, SQL)
                        If Dtbb39a.RecordCount > 0 Then
                            Dtbb39a.MoveFirst()

                            SalaryLooker1 = Dtbb39a("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If


                    If Not NewSald(1) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(1) & "') "
                        OpenTbl(ADb, Dtbb39b, SQL)
                        If Dtbb39b.RecordCount > 0 Then
                            Dtbb39b.MoveFirst()

                            SalaryLooker2 = Dtbb39b("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(2) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(2) & "') "
                        OpenTbl(ADb, Dtbb39c, SQL)
                        If Dtbb39c.RecordCount > 0 Then
                            Dtbb39c.MoveFirst()

                            SalaryLooker3 = Dtbb39c("Salary").Value
                            CtrlShow = CtrlShow + 1
                        End If
                    End If

                    If Not NewSald(3) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(3) & "') "
                        OpenTbl(ADb, Dtbb39d, SQL)
                        If Dtbb39d.RecordCount > 0 Then
                            Dtbb39d.MoveFirst()

                            SalaryLooker4 = Dtbb39d("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(4) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(4) & "') "
                        OpenTbl(ADb, Dtbb39f, SQL)
                        If Dtbb39f.RecordCount > 0 Then
                            Dtbb39f.MoveFirst()

                            SalaryLooker5 = Dtbb39f("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(5) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(5) & "') "
                        OpenTbl(ADb, Dtbb39g, SQL)
                        If Dtbb39g.RecordCount > 0 Then
                            Dtbb39g.MoveFirst()
                            SalaryLooker6 = Dtbb39g("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(6) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(6) & "') "
                        OpenTbl(ADb, Dtbb39h, SQL)
                        If Dtbb39h.RecordCount > 0 Then
                            Dtbb39h.MoveFirst()
                            SalaryLooker7 = Dtbb39h("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(7) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = cdate('" & NewSald(7) & "') "
                        OpenTbl(ADb, Dtbb39i, SQL)
                        If Dtbb39i.RecordCount > 0 Then
                            Dtbb39i.MoveFirst()

                            SalaryLooker8 = Dtbb39i("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(8) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = cdate('" & NewSald(8) & "') "
                        OpenTbl(ADb, Dtbb39j, SQL)
                        If Dtbb39j.RecordCount > 0 Then
                            Dtbb39j.MoveFirst()
                            SalaryLooker9 = Dtbb39j("Salary").Value
                            CtrlShow = CtrlShow + 1
                        End If
                    End If

                    If Not NewSald(9) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = cdate('" & NewSald(9) & "') "
                        OpenTbl(ADb, Dtbb39k, SQL)
                        If Dtbb39k.RecordCount > 0 Then
                            Dtbb39k.MoveFirst()

                            SalaryLooker10 = Dtbb39k("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(10) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = cdate('" & NewSald(10) & "') "
                        OpenTbl(ADb, Dtbb39l, SQL)
                        If Dtbb39l.RecordCount > 0 Then
                            Dtbb39l.MoveFirst()

                            SalaryLooker11 = Dtbb39l("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(11) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = cdate('" & NewSald(11) & "') "
                        OpenTbl(ADb, Dtbb39m, SQL)
                        If Dtbb39m.RecordCount > 0 Then
                            Dtbb39m.MoveFirst()
                            SalaryLooker12 = Dtbb39m("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(12) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = cdate('" & NewSald(12) & "') "
                        OpenTbl(ADb, Dtbb39n, SQL)
                        If Dtbb39n.RecordCount > 0 Then
                            Dtbb39n.MoveFirst()

                            SalaryLooker13 = Dtbb39n("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(13) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = cdate('" & NewSald(13) & "') "
                        OpenTbl(ADb, Dtbb39o, SQL)
                        If Dtbb39o.RecordCount > 0 Then
                            Dtbb39o.MoveFirst()

                            SalaryLooker14 = Dtbb39o("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(14) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = cdate('" & NewSald(14) & "') "
                        OpenTbl(ADb, Dtbb39p, SQL)
                        If Dtbb39p.RecordCount > 0 Then
                            Dtbb39p.MoveFirst()

                            SalaryLooker15 = Dtbb39p("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If

                    End If

                    If Not NewSald(15) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 15_Wallet_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = cdate('" & NewSald(15) & "') "
                        OpenTbl(ADb, Dtbb39q, SQL)
                        If Dtbb39q.RecordCount > 0 Then
                            Dtbb39q.MoveFirst()

                            SalaryLooker16 = Dtbb39q("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If CtrlShow >= 1 Then

                        GetSalaryTot = Val(SalaryLooker1) + Val(SalaryLooker2) + Val(SalaryLooker3) + Val(SalaryLooker4) + Val(SalaryLooker5) + Val(SalaryLooker6) + Val(SalaryLooker7) + Val(SalaryLooker8) + Val(SalaryLooker9) + Val(SalaryLooker10) + Val(SalaryLooker11) + Val(SalaryLooker12) + Val(SalaryLooker13) + Val(SalaryLooker14) + Val(SalaryLooker15) + Val(SalaryLooker16)
                        FieldGridFormatNum()
                        PerFieldGrid01.Invoke(DirectCast(Sub() PerFieldGrid01.Rows.Add(NikLooker, NameLooker, NewFormT(0), NewFormT(1), NewFormT(2), NewFormT(3), NewFormT(4), NewFormT(5), NewFormT(6), NewFormT(7), NewFormT(8), NewFormT(9), NewFormT(10), NewFormT(11), NewFormT(12), NewFormT(13), NewFormT(14), NewFormT(15), "", "", AstekLookMode, "", "", "", "", Format(Val(GetSalaryTot), "N0"), TotRound2), MethodInvoker))
                        LetsCount()

                        TotalTbxFiller()

                    End If

                    Dtbb40.MoveNext()
                Loop

                MsgBox("Done")

            End If 'Termination of Dtbb40


        End If


    End Sub

    Sub PerDataLoader5()

        NikLooker = ""
        NameLooker = ""
        SalaryLooker1 = ""
        RecordCounting = "0"


        If PerFieldCmb4.Text = "Sortasi" Then
            MiscLook = "New"

        ElseIf PerFieldCmb4.Text = "Miscellaneous" Then
            MiscLook = "Old"
        End If

        If Bracket = "All" Then

            SQL = ""
            SQL = SQL & "Select * from 02_Name_Table "
            SQL = SQL & "Where Active = ('" & "Yes" & "') "
            SQL = SQL & "Order by Nik "
            OpenTbl(ADb, Dtbb40, SQL)
            If Dtbb40.RecordCount <> 0 Then

                Dtbb40.MoveFirst()
                Do While Not Dtbb40.EOF

                    NikLooker = IIf(IsDBNull(Dtbb40("Nik").Value), "", Dtbb40("Nik").Value)
                    NameLooker = IIf(IsDBNull(Dtbb40("Name").Value), "", Dtbb40("Name").Value)
                    AstekLookMode = IIf(IsDBNull(Dtbb40("Jamsostek").Value), "", Dtbb40("Jamsostek").Value)

                    SalaryLookerNest()
                    CtrlShow = "0"

                    If Not NewSald(0) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(0) & " ') "
                        SQL = SQL & "And TypeCtrl = ('" & MiscLook & " ') "
                        OpenTbl(ADb, Dtbb39a, SQL)
                        If Dtbb39a.RecordCount > 0 Then
                            Dtbb39a.MoveFirst()

                            SalaryLooker1 = Dtbb39a("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(1) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(1) & "') "
                        SQL = SQL & "And TypeCtrl = ('" & MiscLook & " ') "
                        OpenTbl(ADb, Dtbb39b, SQL)
                        If Dtbb39b.RecordCount > 0 Then
                            Dtbb39b.MoveFirst()

                            SalaryLooker2 = Dtbb39b("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(2) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(2) & "') "
                        SQL = SQL & "And TypeCtrl = ('" & MiscLook & " ') "
                        OpenTbl(ADb, Dtbb39c, SQL)
                        If Dtbb39c.RecordCount > 0 Then
                            Dtbb39c.MoveFirst()

                            SalaryLooker3 = Dtbb39c("Salary").Value
                            CtrlShow = CtrlShow + 1
                        End If
                    End If

                    If Not NewSald(3) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(3) & "') "
                        SQL = SQL & "And TypeCtrl = ('" & MiscLook & " ') "
                        OpenTbl(ADb, Dtbb39d, SQL)
                        If Dtbb39d.RecordCount > 0 Then
                            Dtbb39d.MoveFirst()

                            SalaryLooker4 = Dtbb39d("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(4) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(4) & "') "
                        SQL = SQL & "And TypeCtrl = ('" & MiscLook & " ') "
                        OpenTbl(ADb, Dtbb39f, SQL)
                        If Dtbb39f.RecordCount > 0 Then

                            Dtbb39f.MoveFirst()

                            SalaryLooker5 = Dtbb39f("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(5) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(5) & "') "
                        SQL = SQL & "And TypeCtrl = ('" & MiscLook & " ') "
                        OpenTbl(ADb, Dtbb39g, SQL)
                        If Dtbb39g.RecordCount > 0 Then
                            Dtbb39g.MoveFirst()
                            SalaryLooker6 = Dtbb39g("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(6) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(6) & "') "
                        SQL = SQL & "And TypeCtrl = ('" & MiscLook & " ') "
                        OpenTbl(ADb, Dtbb39h, SQL)
                        If Dtbb39h.RecordCount > 0 Then
                            Dtbb39h.MoveFirst()
                            SalaryLooker7 = Dtbb39h("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(7) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(7) & "') "
                        SQL = SQL & "And TypeCtrl = ('" & MiscLook & " ') "
                        OpenTbl(ADb, Dtbb39i, SQL)
                        If Dtbb39i.RecordCount > 0 Then
                            Dtbb39i.MoveFirst()

                            SalaryLooker8 = Dtbb39i("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(8) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(8) & "') "
                        SQL = SQL & "And TypeCtrl = ('" & MiscLook & " ') "
                        OpenTbl(ADb, Dtbb39j, SQL)
                        If Dtbb39j.RecordCount > 0 Then
                            Dtbb39j.MoveFirst()
                            SalaryLooker9 = Dtbb39j("Salary").Value
                            CtrlShow = CtrlShow + 1
                        End If
                    End If

                    If Not NewSald(9) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(9) & "') "
                        SQL = SQL & "And TypeCtrl = ('" & MiscLook & " ') "
                        OpenTbl(ADb, Dtbb39k, SQL)
                        If Dtbb39k.RecordCount > 0 Then
                            Dtbb39k.MoveFirst()

                            SalaryLooker10 = Dtbb39k("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(10) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(10) & "') "
                        SQL = SQL & "And TypeCtrl = ('" & MiscLook & " ') "
                        OpenTbl(ADb, Dtbb39l, SQL)
                        If Dtbb39l.RecordCount > 0 Then
                            Dtbb39l.MoveFirst()

                            SalaryLooker11 = Dtbb39l("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(11) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(11) & "') "
                        SQL = SQL & "And TypeCtrl = ('" & MiscLook & " ') "
                        OpenTbl(ADb, Dtbb39m, SQL)
                        If Dtbb39m.RecordCount > 0 Then
                            Dtbb39m.MoveFirst()
                            SalaryLooker12 = Dtbb39m("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(12) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(12) & "') "
                        SQL = SQL & "And TypeCtrl = ('" & MiscLook & " ') "
                        OpenTbl(ADb, Dtbb39n, SQL)
                        If Dtbb39n.RecordCount > 0 Then
                            Dtbb39n.MoveFirst()

                            SalaryLooker13 = Dtbb39n("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(13) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(13) & "') "
                        SQL = SQL & "And TypeCtrl = ('" & MiscLook & " ') "
                        OpenTbl(ADb, Dtbb39o, SQL)
                        If Dtbb39o.RecordCount > 0 Then
                            Dtbb39o.MoveFirst()

                            SalaryLooker14 = Dtbb39o("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(14) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(14) & "') "
                        SQL = SQL & "And TypeCtrl = ('" & MiscLook & " ') "
                        OpenTbl(ADb, Dtbb39p, SQL)
                        If Dtbb39p.RecordCount > 0 Then
                            Dtbb39p.MoveFirst()

                            SalaryLooker15 = Dtbb39p("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If

                    End If

                    If Not NewSald(15) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(15) & "') "
                        SQL = SQL & "And TypeCtrl = ('" & MiscLook & " ') "
                        OpenTbl(ADb, Dtbb39q, SQL)
                        If Dtbb39q.RecordCount > 0 Then
                            Dtbb39q.MoveFirst()

                            SalaryLooker16 = Dtbb39q("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If CtrlShow >= 1 Then

                        GetSalaryTot = Val(SalaryLooker1) + Val(SalaryLooker2) + Val(SalaryLooker3) + Val(SalaryLooker4) + Val(SalaryLooker5) + Val(SalaryLooker6) + Val(SalaryLooker7) + Val(SalaryLooker8) + Val(SalaryLooker9) + Val(SalaryLooker10) + Val(SalaryLooker11) + Val(SalaryLooker12) + Val(SalaryLooker13) + Val(SalaryLooker14) + Val(SalaryLooker15) + Val(SalaryLooker16)
                        AstekLoad()
                        FieldGridFormatNum()
                        PerFieldGrid01.Invoke(DirectCast(Sub() PerFieldGrid01.Rows.Add(NikLooker, NameLooker, NewFormT(0), NewFormT(1), NewFormT(2), NewFormT(3), NewFormT(4), NewFormT(5), NewFormT(6), NewFormT(7), NewFormT(8), NewFormT(9), NewFormT(10), NewFormT(11), NewFormT(12), NewFormT(13), NewFormT(14), NewFormT(15), "", "", AstekLookMode, "", "", "", "", Format(Val(GetSalaryTot), "N0"), TotRound2), MethodInvoker))
                        LetsCount()

                        TotalTbxFiller()

                    End If

                    Dtbb40.MoveNext()
                Loop

                MsgBox("Done")

            End If 'Termination of Dtbb40

        Else

            SQL = ""
            SQL = SQL & "Select * from 02_Name_Table "
            SQL = SQL & "Where Active = ('" & "Yes" & "') "
            SQL = SQL & "And Pay = ('" & Bracket & " ') "
            SQL = SQL & "Order by Nik "
            OpenTbl(ADb, Dtbb40, SQL)
            If Dtbb40.RecordCount <> 0 Then

                Dtbb40.MoveFirst()
                Do While Not Dtbb40.EOF

                    NikLooker = IIf(IsDBNull(Dtbb40("Nik").Value), "", Dtbb40("Nik").Value)
                    NameLooker = IIf(IsDBNull(Dtbb40("Name").Value), "", Dtbb40("Name").Value)
                    AstekLookMode = IIf(IsDBNull(Dtbb40("Jamsostek").Value), "", Dtbb40("Jamsostek").Value)

                    SalaryLookerNest()
                    CtrlShow = "0"

                    If Not NewSald(0) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(0) & " ') "
                        OpenTbl(ADb, Dtbb39a, SQL)
                        If Dtbb39a.RecordCount > 0 Then
                            Dtbb39a.MoveFirst()

                            SalaryLooker1 = Dtbb39a("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(1) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(1) & "') "
                        OpenTbl(ADb, Dtbb39b, SQL)
                        If Dtbb39b.RecordCount > 0 Then
                            Dtbb39b.MoveFirst()

                            SalaryLooker2 = Dtbb39b("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If

                    End If

                    If Not NewSald(2) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(2) & "') "
                        OpenTbl(ADb, Dtbb39c, SQL)
                        If Dtbb39c.RecordCount > 0 Then
                            Dtbb39c.MoveFirst()

                            SalaryLooker3 = Dtbb39c("Salary").Value
                            CtrlShow = CtrlShow + 1
                        End If
                    End If

                    If Not NewSald(3) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(3) & "') "
                        OpenTbl(ADb, Dtbb39d, SQL)
                        If Dtbb39d.RecordCount > 0 Then
                            Dtbb39d.MoveFirst()

                            SalaryLooker4 = Dtbb39d("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(4) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(4) & "') "
                        OpenTbl(ADb, Dtbb39f, SQL)
                        If Dtbb39f.RecordCount > 0 Then
                            Dtbb39f.MoveFirst()

                            SalaryLooker5 = Dtbb39f("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(5) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(5) & "') "
                        OpenTbl(ADb, Dtbb39g, SQL)
                        If Dtbb39g.RecordCount > 0 Then
                            Dtbb39g.MoveFirst()

                            SalaryLooker6 = Dtbb39g("Salary").Value

                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(6) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = ('" & NewSald(6) & "') "
                        OpenTbl(ADb, Dtbb39h, SQL)
                        If Dtbb39h.RecordCount > 0 Then
                            Dtbb39h.MoveFirst()

                            SalaryLooker7 = Dtbb39h("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(7) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = cdate('" & NewSald(7) & "') "
                        OpenTbl(ADb, Dtbb39i, SQL)
                        If Dtbb39i.RecordCount > 0 Then
                            Dtbb39i.MoveFirst()

                            SalaryLooker8 = Dtbb39i("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(8) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = cdate('" & NewSald(8) & "') "
                        OpenTbl(ADb, Dtbb39j, SQL)
                        If Dtbb39j.RecordCount > 0 Then
                            Dtbb39j.MoveFirst()

                            SalaryLooker9 = Dtbb39j("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(9) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = cdate('" & NewSald(9) & "') "
                        OpenTbl(ADb, Dtbb39k, SQL)
                        If Dtbb39k.RecordCount > 0 Then
                            Dtbb39k.MoveFirst()

                            SalaryLooker10 = Dtbb39k("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(10) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = cdate('" & NewSald(10) & "') "
                        OpenTbl(ADb, Dtbb39l, SQL)
                        If Dtbb39l.RecordCount > 0 Then
                            Dtbb39l.MoveFirst()

                            SalaryLooker11 = Dtbb39l("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(11) = "" Then

                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = cdate('" & NewSald(11) & "') "
                        OpenTbl(ADb, Dtbb39m, SQL)
                        If Dtbb39m.RecordCount > 0 Then
                            Dtbb39m.MoveFirst()
                            SalaryLooker12 = Dtbb39m("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(12) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = cdate('" & NewSald(12) & "') "
                        OpenTbl(ADb, Dtbb39n, SQL)
                        If Dtbb39n.RecordCount > 0 Then
                            Dtbb39n.MoveFirst()

                            SalaryLooker13 = Dtbb39n("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(13) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = cdate('" & NewSald(13) & "') "
                        OpenTbl(ADb, Dtbb39o, SQL)
                        If Dtbb39o.RecordCount > 0 Then
                            Dtbb39o.MoveFirst()

                            SalaryLooker14 = Dtbb39o("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If Not NewSald(14) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = cdate('" & NewSald(14) & "') "
                        OpenTbl(ADb, Dtbb39p, SQL)
                        If Dtbb39p.RecordCount > 0 Then
                            Dtbb39p.MoveFirst()

                            SalaryLooker15 = Dtbb39p("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If

                    End If

                    If Not NewSald(15) = "" Then
                        SQL = ""
                        SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                        SQL = SQL & "Where Nik = ('" & NikLooker & "') "
                        SQL = SQL & "And Date = cdate('" & NewSald(15) & "') "
                        OpenTbl(ADb, Dtbb39q, SQL)
                        If Dtbb39q.RecordCount > 0 Then
                            Dtbb39q.MoveFirst()

                            SalaryLooker16 = Dtbb39q("Salary").Value
                            CtrlShow = CtrlShow + 1

                        End If
                    End If

                    If CtrlShow >= 1 Then

                        GetSalaryTot = Val(SalaryLooker1) + Val(SalaryLooker2) + Val(SalaryLooker3) + Val(SalaryLooker4) + Val(SalaryLooker5) + Val(SalaryLooker6) + Val(SalaryLooker7) + Val(SalaryLooker8) + Val(SalaryLooker9) + Val(SalaryLooker10) + Val(SalaryLooker11) + Val(SalaryLooker12) + Val(SalaryLooker13) + Val(SalaryLooker14) + Val(SalaryLooker15) + Val(SalaryLooker16)
                        AstekLoad()
                        FieldGridFormatNum()
                        PerFieldGrid01.Invoke(DirectCast(Sub() PerFieldGrid01.Rows.Add(NikLooker, NameLooker, NewFormT(0), NewFormT(1), NewFormT(2), NewFormT(3), NewFormT(4), NewFormT(5), NewFormT(6), NewFormT(7), NewFormT(8), NewFormT(9), NewFormT(10), NewFormT(11), NewFormT(12), NewFormT(13), NewFormT(14), NewFormT(15), "", "", AstekLookMode, "", "", "", "", Format(Val(GetSalaryTot), "N0"), TotRound2), MethodInvoker))

                        LetsCount()

                        TotalTbxFiller()

                    End If

                    Dtbb40.MoveNext()

                Loop

                MsgBox("Done")

            End If 'Termination of Dtbb40

        End If

    End Sub


#End Region

#Region "GUI Control"

    Private Sub PerFieldBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PerFieldBtn1.Click
        LoadHeadDater()
        Dim NewMDIChild As New LoadingBlock()
        MainMenu.SSTab1.Visible = True
        MainMenu.SSTab2.Visible = True
        LoadingBlock.MdiParent = MainMenu
        LoadingBlock.Show()
        Chooser = PerFieldCmb4.Text
        Bracket = PerFieldCmb3.Text

        FRbgw.RunWorkerAsync()


    End Sub


    Private Sub PerFieldBtn3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PerFieldBtn3.Click
        PerFieldGrid01.Rows.Clear()
        SalTbxNest()
    End Sub

    Private Sub PerFieldCmb4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PerFieldCmb4.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub PerFieldCmb3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PerFieldCmb3.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub PerFieldCmb2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PerFieldCmb2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub PerFieldCmb1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub PerFieldBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PerFieldBtn2.Click
        If PerFieldCmb1.Text = "" Then
            MsgBox("Please Select the Periode")

        Else
            GenExcel()

        End If
    End Sub
#End Region

#Region "Excel Codes"

    Sub GenExcel()

        ExcelName = Bracket & "_" & "Report" & "_" & PerFieldCmb1.Text & "_" & Format(Now, "dd.MM.yyyy Hmmss")

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
        '    MsgBox("Done", vbOKOnly + 64, "")
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
            .Cells.Value = "BAGIAN : " + PerFieldCmb4.Text
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2

        End With

        With ExcelAP.Range("A5:AA5")

            .Merge()
            .Font.Bold = True
            .Cells.Value = PerFieldCmb1.Text
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


            .Cells.Value = NewSald(0)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("D7:D7")


            .Cells.Value = NewSald(1)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("E7:E7")


            .Cells.Value = NewSald(2)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With


        With ExcelAP.Range("F7:F7")


            .Cells.Value = NewSald(3)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With


        With ExcelAP.Range("G7:G7")


            .Cells.Value = NewSald(4)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("H7:H7")


            .Cells.Value = NewSald(5)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("I7:I7")


            .Cells.Value = NewSald(6)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With


        With ExcelAP.Range("J7:J7")


            .Cells.Value = NewSald(7)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With


        With ExcelAP.Range("K7:K7")


            .Cells.Value = NewSald(8)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With


        With ExcelAP.Range("L7:L7")


            .Cells.Value = NewSald(0)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("M7:M7")


            .Cells.Value = NewSald(10)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("N7:N7")


            .Cells.Value = NewSald(11)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("O7:O7")


            .Cells.Value = NewSald(12)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("P7:P7")


            .Cells.Value = NewSald(13)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With


        With ExcelAP.Range("Q7:Q7")


            .Cells.Value = NewSald(14)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("R7:R7")


            .Cells.Value = NewSald(15)
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With


        With ExcelAP.Range("S7:S7")


            .Cells.Value = ""
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


            .Cells.Value = "ASTEK"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("V7:V7")


            .Cells.Value = "TJ. LAIN"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("W7:W7")


            .Cells.Value = "TJ. PPH21"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With


        With ExcelAP.Range("X7:X7")


            .Cells.Value = "PPH21"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("Y7:Y7")


            .Cells.Value = "POT LAIN"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("Z7:Z7")


            .Cells.Value = "TOTAL"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With


        With ExcelAP.Range("AA7:AA7")


            .Cells.Value = "GAJI"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("AB7:AB7")

            .Cells.Value = "PARAF"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("AC7:AC7")

            .Cells.Value = "Pay"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With


        For Me.i = 0 To PerFieldGrid01.Rows.Count - 1
            For Me.j = 0 To PerFieldGrid01.ColumnCount - 1
                ExcelWS.Cells(i + 8, j + 1) = PerFieldGrid01(j, i).Value
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

#Region "Round Off Mode"

    Sub RoundProcess()

        TotRoundoff = GetSalaryTot
        TotRound1 = CustomRound(TotRoundoff)

    End Sub


#End Region

  
    Private Sub FRbgw_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles FRbgw.DoWork
     
        ClockCodeNew()
    End Sub

    Private Sub FRbgw_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles FRbgw.RunWorkerCompleted
        LoadingBlock.Close()
    End Sub

End Class