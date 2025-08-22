Option Explicit On
Public Class DateBlock

    Dim NextDay As Date
    Dim DateNum As String
    Dim DateDig As String
    Dim DateLook As Integer
    Dim CmbDater As String
    Dim CmbDater2 As String
    Dim CmbDater3 As String

    Private Sub DateBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LoadDB()
        LoadDB2()
        LoadDateSetup()
        DateCmb2.Text = IIf(DateCmb2.Text = Nothing, "Periode I", DateCmb2.Text)

    End Sub

    'Sub DateSet()
    '    CmbDater = Format(Now, "yyyy")
    '    CmbDater2 = CmbDater - 1
    '    CmbDater3 = CmbDater + 1

    '    With DateCmb1

    '        .Items.Add("Dec " + CmbDater2)
    '        .Items.Add("Jan " + Format(Now, "yyyy"))
    '        .Items.Add("Feb " + Format(Now, "yyyy"))
    '        .Items.Add("Mar " + Format(Now, "yyyy"))
    '        .Items.Add("Apr " + Format(Now, "yyyy"))
    '        .Items.Add("May " + Format(Now, "yyyy"))
    '        .Items.Add("Jun " + Format(Now, "yyyy"))
    '        .Items.Add("Jul " + Format(Now, "yyyy"))
    '        .Items.Add("Aug " + Format(Now, "yyyy"))
    '        .Items.Add("Sep " + Format(Now, "yyyy"))
    '        .Items.Add("Oct " + Format(Now, "yyyy"))
    '        .Items.Add("Nov " + Format(Now, "yyyy"))
    '        .Items.Add("Dec " + Format(Now, "yyyy"))
    '        .Items.Add("Jan " + CmbDater3)

    '    End With
    'End Sub

    Sub NewEraser()
        DateCmb1.Text = ""
        DateCmb2.Text = ""
        DateTbx1.Text = ""
        DateTbx2.Text = ""
        DateTbx3.Text = ""
        DateTbx4.Text = ""
        DateTbx5.Text = ""
        DateTbx6.Text = ""
        DateTbx7.Text = ""
        DateTbx8.Text = ""
        DateTbx9.Text = ""
        DateTbx10.Text = ""
        DateTbx11.Text = ""
        DateTbx12.Text = ""
        DateTbx13.Text = ""
        DateTbx14.Text = ""
        DateTbx15.Text = ""
        DateTbx16.Text = ""

    End Sub

    Private Sub DateCmb1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub
    ' Auto 1 Day Code for Lazy People
    ' Date Saving Code
    Sub DateDataSave()
        DateAutoCode()

        SQL = ""
        SQL = SQL & "Select * From DateCounter2Table "
        SQL = SQL & "Where PeriodeRange =  ('" & DateCmb1.Text & "') "
        SQL = SQL & "And Periode =  ('" & DateCmb2.Text & "') "
        OpenTbl(CBb, Ctbl21, SQL)

        If Not Ctbl21.RecordCount <> 0 Then
            Ctbl21.AddNew()
        End If

        Ctbl21("IDDate").Value = DateNum
        Ctbl21("Periode").Value = DateCmb2.Text
        Ctbl21("PeriodeRange").Value = DateCmb1.Text
        Ctbl21("PeriodeValid").Value = "Yes"
        Ctbl21("Date1").Value = DateTbx1.Text
        Ctbl21("Date2").Value = DateTbx2.Text
        Ctbl21("Date3").Value = DateTbx3.Text
        Ctbl21("Date4").Value = DateTbx4.Text
        Ctbl21("Date5").Value = DateTbx5.Text
        Ctbl21("Date6").Value = DateTbx6.Text
        Ctbl21("Date7").Value = DateTbx7.Text
        Ctbl21("Date8").Value = DateTbx8.Text
        Ctbl21("Date9").Value = DateTbx9.Text
        Ctbl21("Date10").Value = DateTbx10.Text
        Ctbl21("Date11").Value = DateTbx11.Text
        Ctbl21("Date12").Value = DateTbx12.Text
        Ctbl21("Date13").Value = DateTbx13.Text
        Ctbl21("Date14").Value = DateTbx14.Text
        Ctbl21("Date15").Value = DateTbx15.Text
        Ctbl21("Date16").Value = DateTbx16.Text

        Ctbl21.Update()

        Me.Refresh()
        Me.Dispose()

    End Sub

    Sub DateDataSave2()

        SQL = ""
        SQL = SQL & "Select * From DateCounter2Table "
        SQL = SQL & "Where PeriodeValid =  ('" & "Yes" & "') "
        OpenTbl(CBb, Ctbl23, SQL)

        If Ctbl23.RecordCount > 0 Then
            Ctbl23.Delete()
        End If

    End Sub

    Sub NewDateSave()

        DateAutoCode()

        SQL = ""
        SQL = SQL & "Select * From DateCounter2Table "
        SQL = SQL & "Where PeriodeRange =  ('" & DateCmb1.Text & "') "
        SQL = SQL & "And Periode=  ('" & DateCmb2.Text & "') "
        OpenTbl(CBb, Ctbl23, SQL)

        If Not Ctbl23.RecordCount <> 0 Then
            Ctbl23.AddNew()
        End If

        If Not DateCmb1.Text = "" Then

            Ctbl23("IDDate").Value = DateNum
            Ctbl23("PeriodeValid").Value = "No"
            Ctbl23("Periode").Value = DateCmb2.Text
            Ctbl23("PeriodeRange").Value = DateCmb1.Text
            Ctbl23("Date1").Value = DateTbx1.Text
            Ctbl23("Date2").Value = DateTbx2.Text
            Ctbl23("Date3").Value = DateTbx3.Text
            Ctbl23("Date4").Value = DateTbx4.Text
            Ctbl23("Date5").Value = DateTbx5.Text
            Ctbl23("Date6").Value = DateTbx6.Text
            Ctbl23("Date7").Value = DateTbx7.Text
            Ctbl23("Date8").Value = DateTbx8.Text
            Ctbl23("Date9").Value = DateTbx9.Text
            Ctbl23("Date10").Value = DateTbx10.Text
            Ctbl23("Date11").Value = DateTbx11.Text
            Ctbl23("Date12").Value = DateTbx12.Text
            Ctbl23("Date13").Value = DateTbx13.Text
            Ctbl23("Date14").Value = DateTbx14.Text
            Ctbl23("Date15").Value = DateTbx15.Text
            Ctbl23("Date16").Value = DateTbx16.Text
            Ctbl23.Update()
        End If
        Me.Refresh()

    End Sub

    Sub DateDataSave3()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Date =  ('" & DateTbx1.Text & "') "
        OpenTbl(CBb, Ctbl26, SQL)

        If Not Ctbl26.RecordCount <> 0 Then
            Ctbl26.AddNew()
        End If

        If Not DateTbx1.Text = "" Then

            Ctbl26("Periode").Value = DateCmb2.Text
            Ctbl26("PeriodeRange").Value = DateCmb1.Text
            Ctbl26("Date").Value = IIf(DateTbx1.Text = Nothing, DBNull.Value, DateTbx1.Text)
            Ctbl26("Counter").Value = "1"
            Ctbl26.Update()

        End If
        Me.Refresh()
    End Sub

    Sub DateDataSave4()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Date =  ('" & DateTbx2.Text & "') "
        OpenTbl(CBb, Ctbl27, SQL)

        If Not Ctbl27.RecordCount <> 0 Then
            Ctbl27.AddNew()
        End If

        If Not DateTbx2.Text = "" Then

            Ctbl27("Periode").Value = DateCmb2.Text
            Ctbl27("PeriodeRange").Value = DateCmb1.Text
            Ctbl27("Date").Value = IIf(DateTbx2.Text = Nothing, DBNull.Value, DateTbx2.Text)
            Ctbl27("Counter").Value = "2"
            Ctbl27.Update()

        End If
        Me.Refresh()
    End Sub

    Sub DateDataSave5()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Date =  ('" & DateTbx3.Text & "') "
        OpenTbl(CBb, Ctbl28, SQL)

        If Not Ctbl28.RecordCount <> 0 Then
            Ctbl28.AddNew()
        End If

        If Not DateTbx3.Text = "" Then

            Ctbl28("Periode").Value = DateCmb2.Text
            Ctbl28("PeriodeRange").Value = DateCmb1.Text
            Ctbl28("Date").Value = IIf(DateTbx3.Text = Nothing, DBNull.Value, DateTbx3.Text)
            Ctbl28("Counter").Value = "3"
            Ctbl28.Update()
        End If
        Me.Refresh()

    End Sub

    Sub DateDataSave6()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Date =  ('" & DateTbx4.Text & "') "
        OpenTbl(CBb, Ctbl29, SQL)

        If Not Ctbl29.RecordCount <> 0 Then
            Ctbl29.AddNew()
        End If

        If Not DateTbx4.Text = "" Then

            Ctbl29("Periode").Value = DateCmb2.Text
            Ctbl29("PeriodeRange").Value = DateCmb1.Text
            Ctbl29("Date").Value = IIf(DateTbx4.Text = Nothing, DBNull.Value, DateTbx4.Text)
            Ctbl29("Counter").Value = "4"
            Ctbl29.Update()
        End If
        Me.Refresh()

    End Sub
    Sub DateDataSave7()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Date =  ('" & DateTbx5.Text & "') "
        OpenTbl(CBb, Ctbl30, SQL)

        If Not Ctbl30.RecordCount <> 0 Then
            Ctbl30.AddNew()
        End If

        If Not DateTbx5.Text = "" Then

            Ctbl30("Periode").Value = DateCmb2.Text
            Ctbl30("PeriodeRange").Value = DateCmb1.Text
            Ctbl30("Date").Value = IIf(DateTbx5.Text = Nothing, DBNull.Value, DateTbx5.Text)
            Ctbl30("Counter").Value = "5"
            Ctbl30.Update()

        End If

        Me.Refresh()
    End Sub

    Sub DateDataSave8()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Date =  ('" & DateTbx6.Text & "') "
        OpenTbl(CBb, Ctbl31, SQL)

        If Not Ctbl31.RecordCount <> 0 Then
            Ctbl31.AddNew()
        End If

        If Not DateTbx6.Text = "" Then

            Ctbl31("Periode").Value = DateCmb2.Text
            Ctbl31("PeriodeRange").Value = DateCmb1.Text
            Ctbl31("Date").Value = IIf(DateTbx6.Text = Nothing, DBNull.Value, DateTbx6.Text)
            Ctbl31("Counter").Value = "6"
            Ctbl31.Update()
        End If
        Me.Refresh()

    End Sub
    Sub DateDataSave9()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Date =  ('" & DateTbx7.Text & "') "
        OpenTbl(CBb, Ctbl32, SQL)

        If Not Ctbl32.RecordCount <> 0 Then
            Ctbl32.AddNew()
        End If

        If Not DateTbx7.Text = "" Then

            Ctbl32("Periode").Value = DateCmb2.Text
            Ctbl32("PeriodeRange").Value = DateCmb1.Text
            Ctbl32("Date").Value = IIf(DateTbx7.Text = Nothing, DBNull.Value, DateTbx7.Text)
            Ctbl32("Counter").Value = "7"
            Ctbl32.Update()

        End If
        Me.Refresh()

    End Sub

    Sub DateDataSave10()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Date =  ('" & DateTbx8.Text & "') "
        OpenTbl(CBb, Ctbl33, SQL)

        If Not Ctbl33.RecordCount <> 0 Then
            Ctbl33.AddNew()
        End If

        If Not DateTbx8.Text = "" Then

            Ctbl33("Periode").Value = DateCmb2.Text
            Ctbl33("PeriodeRange").Value = DateCmb1.Text
            Ctbl33("Date").Value = IIf(DateTbx8.Text = Nothing, DBNull.Value, DateTbx8.Text)
            Ctbl33("Counter").Value = "8"
            Ctbl33.Update()

        End If
        Me.Refresh()

    End Sub

    Sub DateDataSave11()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Date =  ('" & DateTbx9.Text & "') "
        OpenTbl(CBb, Ctbl34, SQL)

        If Not Ctbl34.RecordCount <> 0 Then
            Ctbl34.AddNew()
        End If

        If Not DateTbx9.Text = "" Then

            Ctbl34("Periode").Value = DateCmb2.Text
            Ctbl34("PeriodeRange").Value = DateCmb1.Text
            Ctbl34("Date").Value = IIf(DateTbx9.Text = Nothing, DBNull.Value, DateTbx9.Text)
            Ctbl34("Counter").Value = "9"
            Ctbl34.Update()

        End If
        Me.Refresh()
    End Sub

    Sub DateDataSave12()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Date =  ('" & DateTbx10.Text & "') "
        OpenTbl(CBb, Ctbl35, SQL)

        If Not Ctbl35.RecordCount <> 0 Then
            Ctbl35.AddNew()
        End If

        If Not DateTbx10.Text = "" Then

            Ctbl35("Periode").Value = DateCmb2.Text
            Ctbl35("PeriodeRange").Value = DateCmb1.Text
            Ctbl35("Date").Value = IIf(DateTbx10.Text = Nothing, DBNull.Value, DateTbx10.Text)
            Ctbl35("Counter").Value = "10"
            Ctbl35.Update()
        End If
        Me.Refresh()

    End Sub

    Sub DateDataSave13()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Date =  ('" & DateTbx11.Text & "') "
        OpenTbl(CBb, Ctbl36, SQL)

        If Not Ctbl36.RecordCount <> 0 Then
            Ctbl36.AddNew()
        End If

        If Not DateTbx11.Text = "" Then

            Ctbl36("Periode").Value = DateCmb2.Text
            Ctbl36("PeriodeRange").Value = DateCmb1.Text
            Ctbl36("Date").Value = IIf(DateTbx11.Text = Nothing, DBNull.Value, DateTbx11.Text)
            Ctbl36("Counter").Value = "11"
            Ctbl36.Update()

        End If
        Me.Refresh()

    End Sub

    Sub DateDataSave14()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Date =  ('" & DateTbx12.Text & "') "
        OpenTbl(CBb, Ctbl37, SQL)

        If Not Ctbl37.RecordCount <> 0 Then
            Ctbl37.AddNew()
        End If

        If Not DateTbx12.Text = "" Then

            Ctbl37("Periode").Value = DateCmb2.Text
            Ctbl37("PeriodeRange").Value = DateCmb1.Text
            Ctbl37("Date").Value = IIf(DateTbx12.Text = Nothing, DBNull.Value, DateTbx12.Text)
            Ctbl37("Counter").Value = "12"
            Ctbl37.Update()
        End If
        Me.Refresh()

    End Sub

    Sub DateDataSave15()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Date =  ('" & DateTbx13.Text & "') "
        OpenTbl(CBb, Ctbl38, SQL)

        If Not Ctbl38.RecordCount <> 0 Then
            Ctbl38.AddNew()
        End If

        If Not DateTbx13.Text = "" Then

            Ctbl38("Periode").Value = DateCmb2.Text
            Ctbl38("PeriodeRange").Value = DateCmb1.Text
            Ctbl38("Date").Value = IIf(DateTbx13.Text = Nothing, DBNull.Value, DateTbx13.Text)
            Ctbl38("Counter").Value = "13"
            Ctbl38.Update()
        End If
        Me.Refresh()

    End Sub
    Sub DateDataSave16()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Date =  ('" & DateTbx14.Text & "') "
        OpenTbl(CBb, Ctbl39, SQL)

        If Not Ctbl39.RecordCount <> 0 Then
            Ctbl39.AddNew()
        End If

        If Not DateTbx14.Text = "" Then

            Ctbl39("Periode").Value = DateCmb2.Text
            Ctbl39("PeriodeRange").Value = DateCmb1.Text
            Ctbl39("Date").Value = IIf(DateTbx14.Text = Nothing, DBNull.Value, DateTbx14.Text)
            Ctbl39("Counter").Value = "14"
            Ctbl39.Update()
        End If
        Me.Refresh()

    End Sub

    Sub DateDataSave17()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Date =  ('" & DateTbx15.Text & "') "
        OpenTbl(CBb, Ctbl40, SQL)

        If Not Ctbl40.RecordCount <> 0 Then
            Ctbl40.AddNew()
        End If

        If Not DateTbx15.Text = "" Then

            Ctbl40("Periode").Value = DateCmb2.Text
            Ctbl40("PeriodeRange").Value = DateCmb1.Text
            Ctbl40("Date").Value = IIf(DateTbx15.Text = Nothing, DBNull.Value, DateTbx15.Text)
            Ctbl40("Counter").Value = "15"
            Ctbl40.Update()
        End If
        Me.Refresh()

    End Sub

    Sub DateDataSave18()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Date =  ('" & DateTbx16.Text & "') "
        OpenTbl(CBb, Ctbl41, SQL)

        If Not Ctbl41.RecordCount <> 0 Then
            Ctbl41.AddNew()
        End If

        If Not DateTbx16.Text = "" Then

            Ctbl41("Periode").Value = DateCmb2.Text
            Ctbl41("PeriodeRange").Value = DateCmb1.Text
            Ctbl41("Date").Value = IIf(DateTbx16.Text = Nothing, DBNull.Value, DateTbx16.Text)
            Ctbl41("Counter").Value = "16"
            Ctbl41.Update()
        End If
        Me.Refresh()

    End Sub

    Sub LoadDateSetup()

        SQL = ""
        SQL = SQL & "Select * From DateCounter2Table "
        SQL = SQL & "Where PeriodeValid =  ('" & "Yes" & "') "
        OpenTbl(CBb, Ctbl22, SQL)

        If Ctbl22.RecordCount <> 0 Then
            Ctbl22.MoveLast()

            DateLook = IIf(IsDBNull(Ctbl22("IDDate").Value), "", Ctbl22("IDDate").Value)
            DateCmb1.Text = IIf(IsDBNull(Ctbl22("PeriodeRange").Value), "", Ctbl22("PeriodeRange").Value)
            DateCmb2.Text = IIf(IsDBNull(Ctbl22("Periode").Value), "", Ctbl22("Periode").Value)
            DateTbx1.Text = IIf(IsDBNull(Ctbl22("Date1").Value), "", Ctbl22("Date1").Value)
            DateTbx2.Text = IIf(IsDBNull(Ctbl22("Date2").Value), "", Ctbl22("Date2").Value)
            DateTbx3.Text = IIf(IsDBNull(Ctbl22("Date3").Value), "", Ctbl22("Date3").Value)
            DateTbx4.Text = IIf(IsDBNull(Ctbl22("Date4").Value), "", Ctbl22("Date4").Value)
            DateTbx5.Text = IIf(IsDBNull(Ctbl22("Date5").Value), "", Ctbl22("Date5").Value)
            DateTbx6.Text = IIf(IsDBNull(Ctbl22("Date6").Value), "", Ctbl22("Date6").Value)
            DateTbx7.Text = IIf(IsDBNull(Ctbl22("Date7").Value), "", Ctbl22("Date7").Value)
            DateTbx8.Text = IIf(IsDBNull(Ctbl22("Date8").Value), "", Ctbl22("Date8").Value)
            DateTbx9.Text = IIf(IsDBNull(Ctbl22("Date9").Value), "", Ctbl22("Date9").Value)
            DateTbx10.Text = IIf(IsDBNull(Ctbl22("Date10").Value), "", Ctbl22("Date10").Value)
            DateTbx11.Text = IIf(IsDBNull(Ctbl22("Date11").Value), "", Ctbl22("Date11").Value)
            DateTbx12.Text = IIf(IsDBNull(Ctbl22("Date12").Value), "", Ctbl22("Date12").Value)
            DateTbx13.Text = IIf(IsDBNull(Ctbl22("Date13").Value), "", Ctbl22("Date13").Value)
            DateTbx14.Text = IIf(IsDBNull(Ctbl22("Date14").Value), "", Ctbl22("Date14").Value)
            DateTbx15.Text = IIf(IsDBNull(Ctbl22("Date15").Value), "", Ctbl22("Date15").Value)
            DateTbx16.Text = IIf(IsDBNull(Ctbl22("Date16").Value), "", Ctbl22("Date16").Value)

        End If

        Me.Refresh()

    End Sub

    Sub DateAutoCode()

        SQL = ""
        SQL = SQL & "Select * From DateCounter2Table "
        SQL = SQL & "Order by IDDate Desc"
        OpenTbl(CBb, Ctbl43, SQL)
        If Ctbl43.RecordCount <> 0 Then
            DateDig = Ctbl43("IDDate").Value
            DateNum = Format(DateDig + 1, "00000000")
        Else
            DateNum = "00000001"
        End If

    End Sub

    Private Sub DateTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTbx1.KeyPress

        DateTbx1.Mask = IIf(DateTbx1.Text = Nothing, Nothing, "##/##/####")


        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            DateTbx2.Focus()
            If Not DateTbx1.Text = Nothing Then
                NextDay = DateTbx1.Text
                DateTbx2.Text = DateAdd(DateInterval.Day, 1, NextDay)
            End If
            e.Handled = True
        End If


    End Sub

    Private Sub DateTbx2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTbx2.KeyPress

        DateTbx2.Mask = IIf(DateTbx2.Text = Nothing, Nothing, "##/##/####")


        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            DateTbx3.Focus()
            If Not DateTbx2.Text = Nothing Then
                NextDay = DateTbx1.Text
                DateTbx3.Text = DateAdd(DateInterval.Day, 2, NextDay)
            End If
            e.Handled = True
        End If

    End Sub

    Private Sub DateTbx3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTbx3.KeyPress

        DateTbx3.Mask = IIf(DateTbx3.Text = Nothing, Nothing, "##/##/####")

        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            DateTbx4.Focus()
            If Not DateTbx3.Text = Nothing Then
                NextDay = DateTbx1.Text
                DateTbx4.Text = DateAdd(DateInterval.Day, 3, NextDay)
            End If
            e.Handled = True
        End If
    End Sub

    Private Sub DateTbx4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTbx4.KeyPress

        DateTbx4.Mask = IIf(DateTbx4.Text = Nothing, Nothing, "##/##/####")

        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            DateTbx5.Focus()
            If Not DateTbx4.Text = Nothing Then
                NextDay = DateTbx1.Text
                DateTbx5.Text = DateAdd(DateInterval.Day, 4, NextDay)
            End If
            e.Handled = True
        End If
    End Sub

    Private Sub DateTbx5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTbx5.KeyPress

        DateTbx5.Mask = IIf(DateTbx5.Text = Nothing, Nothing, "##/##/####")

        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            DateTbx6.Focus()
            If Not DateTbx5.Text = Nothing Then
                NextDay = DateTbx1.Text
                DateTbx6.Text = DateAdd(DateInterval.Day, 5, NextDay)
            End If
            e.Handled = True
        End If

    End Sub

    Private Sub DateTbx6_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTbx6.KeyPress

        DateTbx6.Mask = IIf(DateTbx6.Text = Nothing, Nothing, "##/##/####")


        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            DateTbx7.Focus()
            If Not DateTbx6.Text = Nothing Then
                NextDay = DateTbx1.Text
                DateTbx7.Text = DateAdd(DateInterval.Day, 6, NextDay)
            End If
            e.Handled = True
        End If

    End Sub

    Private Sub DateTbx7_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTbx7.KeyPress

        DateTbx7.Mask = IIf(DateTbx7.Text = Nothing, Nothing, "##/##/####")


        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            DateTbx8.Focus()
            If Not DateTbx7.Text = Nothing Then
                NextDay = DateTbx1.Text
                DateTbx8.Text = DateAdd(DateInterval.Day, 7, NextDay)
            End If
            e.Handled = True
        End If

    End Sub

    Private Sub DateTbx8_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTbx8.KeyPress

        DateTbx8.Mask = IIf(DateTbx8.Text = Nothing, Nothing, "##/##/####")


        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            DateTbx9.Focus()
            If Not DateTbx8.Text = Nothing Then
                NextDay = DateTbx1.Text
                DateTbx9.Text = DateAdd(DateInterval.Day, 8, NextDay)
            End If
            e.Handled = True
        End If

    End Sub

    Private Sub DateTbx9_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTbx9.KeyPress

        DateTbx9.Mask = IIf(DateTbx9.Text = Nothing, Nothing, "##/##/####")


        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            DateTbx10.Focus()
            If Not DateTbx9.Text = Nothing Then
                NextDay = DateTbx1.Text
                DateTbx10.Text = DateAdd(DateInterval.Day, 9, NextDay)
            End If

            e.Handled = True
        End If
    End Sub

    Private Sub DateTbx10_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTbx10.KeyPress

        DateTbx10.Mask = IIf(DateTbx10.Text = Nothing, Nothing, "##/##/####")


        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            DateTbx11.Focus()
            If Not DateTbx10.Text = Nothing Then
                NextDay = DateTbx1.Text
                DateTbx11.Text = DateAdd(DateInterval.Day, 10, NextDay)
            End If

            e.Handled = True
        End If
    End Sub

    Private Sub DateTbx11_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTbx11.KeyPress

        DateTbx11.Mask = IIf(DateTbx11.Text = Nothing, Nothing, "##/##/####")


        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            DateTbx12.Focus()
            If Not DateTbx11.Text = Nothing Then
                NextDay = DateTbx1.Text
                DateTbx12.Text = DateAdd(DateInterval.Day, 11, NextDay)
            End If
            e.Handled = True
        End If
    End Sub

    Private Sub DateTbx12_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTbx12.KeyPress

        DateTbx12.Mask = IIf(DateTbx12.Text = Nothing, Nothing, "##/##/####")


        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            DateTbx13.Focus()
            If Not DateTbx12.Text = Nothing Then
                NextDay = DateTbx1.Text
                DateTbx13.Text = DateAdd(DateInterval.Day, 12, NextDay)
            End If

            e.Handled = True
        End If
    End Sub

    Private Sub DateTbx13_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTbx13.KeyPress
        DateTbx13.Mask = IIf(DateTbx13.Text = Nothing, Nothing, "##/##/####")


        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            DateTbx14.Focus()
            If Not DateTbx13.Text = Nothing Then
                NextDay = DateTbx1.Text
                DateTbx14.Text = DateAdd(DateInterval.Day, 13, NextDay)
            End If

            e.Handled = True
        End If
    End Sub

    Private Sub DateTbx14_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTbx14.KeyPress

        DateTbx14.Mask = IIf(DateTbx14.Text = Nothing, Nothing, "##/##/####")


        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            DateTbx15.Focus()
            If Not DateTbx14.Text = Nothing Then
                NextDay = DateTbx1.Text
                DateTbx15.Text = DateAdd(DateInterval.Day, 14, NextDay)
            End If
            e.Handled = True
        End If
    End Sub

    Private Sub DateTbx15_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTbx15.KeyPress
        DateTbx15.Mask = IIf(DateTbx15.Text = Nothing, Nothing, "##/##/####")

        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            DateTbx16.Focus()
            If Not DateTbx15.Text = Nothing Then
                NextDay = DateTbx1.Text
                DateTbx16.Text = DateAdd(DateInterval.Day, 15, NextDay)
            End If
            e.Handled = True
        End If
    End Sub

    Private Sub DateTbx16_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateTbx16.KeyPress
        DateTbx16.Mask = IIf(DateTbx16.Text = Nothing, Nothing, "##/##/####")
    End Sub

    Private Sub DateBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateBtn1.Click
        DateDataSave2()
        NewDateSave()
        DateCmb2.Enabled = True
        DateCmb1.Enabled = True
        DateTbx1.Enabled = True
        DateTbx2.Enabled = True
        DateTbx3.Enabled = True
        DateTbx4.Enabled = True
        DateTbx5.Enabled = True
        DateTbx6.Enabled = True
        DateTbx7.Enabled = True
        DateTbx8.Enabled = True
        DateTbx9.Enabled = True
        DateTbx10.Enabled = True
        DateTbx11.Enabled = True
        DateTbx12.Enabled = True
        DateTbx13.Enabled = True
        DateTbx14.Enabled = True
        DateTbx15.Enabled = True
        DateTbx16.Enabled = True

    End Sub

    Private Sub DateBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateBtn2.Click

        DateDataSave3()
        DateDataSave4()
        DateDataSave5()
        DateDataSave6()
        DateDataSave7()
        DateDataSave8()
        DateDataSave9()
        DateDataSave10()
        DateDataSave11()
        DateDataSave12()
        DateDataSave13()
        DateDataSave14()
        DateDataSave15()
        DateDataSave16()
        DateDataSave17()
        DateDataSave18()
        DateDataSave()
    End Sub

    Private Sub DateCmb2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DateCmb2.KeyPress
        e.Handled = True

    End Sub

End Class