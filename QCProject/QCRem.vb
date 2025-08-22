Public Class RemoverBlock

    Dim CmbDater As String
    Dim CmbDater2 As String
    Dim CmbDater3 As String
    Dim NameEmp As String
    Dim NikCtrl As String
    Dim ProID As String
    Dim TimeCtrl As String
    Dim DateCtrl As String
    Dim PcsCtrl As String
    Dim TarCtrl As String
    Dim SalCtrl As String
    Dim CouCtrl As String
    Dim CartCtrl As String
    Dim NoKgCtrl As String
    Dim NoGrCtrl As String
    Dim NoBagCtrl As String
    Dim ConCD As String
    Dim MutuIICD As String
    Dim PackingCD As String
    Dim NewMiscellaneousCD As String
    Dim MiscellaneousCD As String
    Dim WalletCD As String
    Dim DepValue As String
    Dim DateSal As String

    Dim Gname As String
    Dim Gnik As String
    Dim GAskVal As String
    Dim GPay As String
    Dim GSal1 As String
    Dim GSal2 As String
    Dim GSal3 As String
    Dim GSal4 As String
    Dim GSal5 As String
    Dim GSal6 As String
    Dim GSal7 As String
    Dim GSal8 As String
    Dim GSal9 As String
    Dim GSal10 As String
    Dim GSal11 As String
    Dim GSal12 As String
    Dim GSal13 As String
    Dim GSal14 As String
    Dim GSal15 As String
    Dim GSal16 As String
    Dim GSPotLain As String
    Dim GPer1 As String
    Dim Gper2 As String

    Dim RemGidVal0 As String
    Dim RemGidVal1 As String
    Dim RemGidVal2 As String
    Dim RemGidVal3 As String
    Dim RemGidVal4 As String
    Dim RemGidVal5 As String
    Dim RemGidVal6 As String
    Dim RemGidVal7 As String
    Dim RemGidVal8 As String
    Dim RemGidVal9 As String
    Dim RemGidVal10 As String
    Dim RemGidVal11 As String
    Dim RemGidVal12 As String
    Dim RemGidVal13 As String
    Dim RemGidVal14 As String
    Dim RemGidVal15 As String
    Dim RemGidVal16 As String
    Dim RemGidVal17 As String
    Dim RemGidVal18 As String
    Dim RemGidVal19 As String
    Dim RemGidVal20 As String



    Private Sub RemTbx2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles RemTbx2.KeyPress

        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True

    End Sub

    Private Sub RemCmb1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles RemCmb1.KeyPress

        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True

    End Sub

    Private Sub RmChk1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RmChk1.CheckedChanged

        RemCmb1.Enabled = True
        RemCmb2.Enabled = False
        RemCmb3.Enabled = False
        RmBtn2.Enabled = True
        RemDP.Enabled = True

    End Sub

    Private Sub RemoverBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LoadDB()
        LoadDB2()
        LoadDB3()
        RmChk1.Checked = True
        'FillDateCmb()

    End Sub

    Private Sub RmChk2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RmChk2.CheckedChanged

        RemCmb1.Enabled = True
        RemCmb2.Enabled = False
        RemCmb3.Enabled = False
        RmBtn2.Enabled = False
        RemDP.Enabled = True

    End Sub

#Region "Process for Data"

    Sub LookData()

        SQL = ""
        SQL = SQL & "Select * From 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & RemTbx1.Text & "') "
        OpenTbl(ADb, Atb4, SQL)

        If Atb4.RecordCount > 0 Then

            RemTbx2.Text = Atb4("Name").Value

        Else

            MsgBox("Employee Not Found", MsgBoxStyle.Information, "Codex ~ QC Build " & BuildCounter & " Warning!!")
            RemTbx1.Clear()
            RemTbx2.Clear()

        End If

    End Sub
    Dim DateGet As Date
    Sub LookValue()
        DateGet = RemDP.Text
        ' For Bagian

        If RmChk1.Checked = True Then

            If RemCmb1.Text = "Conveyour" Then

                SQL = ""
                SQL = SQL & "Select * from 03_Conveyour_Table "
                SQL = SQL & "Where Nik = ('" & RemTbx1.Text & "') "
                SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') "
                OpenTbl(ADb, DbTbl4, SQL)

                If DbTbl4.RecordCount > 0 Then

                    DbTbl4.MoveFirst()

                    Do While Not DbTbl4.EOF

                        ProID = DbTbl4("Process_ID").Value
                        DateCtrl = DbTbl4("Date").Value
                        TimeCtrl = DbTbl4("Time").Value
                        NikCtrl = DbTbl4("Nik").Value
                        PcsCtrl = DbTbl4("Pieces").Value
                        TarCtrl = DbTbl4("Target").Value
                        SalCtrl = DbTbl4("Salary").Value
                        TimeCtrl = DbTbl4("Time").Value

                        UpNRGrid.Rows.Add(ProID, DateCtrl, TimeCtrl, NikCtrl, PcsCtrl, TarCtrl, SalCtrl)

                        DbTbl4.MoveNext()

                    Loop

                End If

            ElseIf RemCmb1.Text = "Mutu II" Then

                SQL = ""
                SQL = SQL & "Select * from 04_MutuII_Table "
                SQL = SQL & "Where Nik = ('" & RemTbx1.Text & "') "
                SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') "
                OpenTbl(ADb, DbTbl7, SQL)
                If DbTbl7.RecordCount > 0 Then

                    DbTbl7.MoveFirst()
                    Do While Not DbTbl7.EOF

                        NikCtrl = DbTbl7("Nik").Value
                        ProID = DbTbl7("Process_ID").Value
                        TimeCtrl = DbTbl7("Time").Value
                        DateCtrl = DbTbl7("Date").Value
                        PcsCtrl = DbTbl7("Pieces").Value
                        TarCtrl = DbTbl7("Target").Value
                        SalCtrl = DbTbl7("Salary").Value
                        CouCtrl = DbTbl7("Coupon").Value

                        UpNRGrid.Rows.Add(ProID, DateCtrl, TimeCtrl, NikCtrl, TarCtrl, PcsCtrl, CouCtrl, SalCtrl)
                        DbTbl7.MoveNext()

                    Loop
                End If

            ElseIf RemCmb1.Text = "Wallet" Then

                SQL = ""
                SQL = SQL & "Select * from 06_Wallet_Table "
                SQL = SQL & "Where Nik = ('" & RemTbx1.Text & "') "
                SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') "
                OpenTbl(ADb, DbTbl8, SQL)
                If DbTbl8.RecordCount > 0 Then

                    DbTbl8.MoveFirst()
                    Do While Not DbTbl8.EOF

                        NikCtrl = DbTbl8("Nik").Value
                        ProID = DbTbl8("Process_ID").Value
                        TimeCtrl = DbTbl8("Time").Value
                        DateCtrl = DbTbl8("Date").Value
                        PcsCtrl = DbTbl8("Pieces").Value
                        TarCtrl = DbTbl8("Target").Value
                        SalCtrl = DbTbl8("Salary").Value
                        CouCtrl = DbTbl8("Coupon").Value

                        UpNRGrid.Rows.Add(ProID, DateCtrl, TimeCtrl, NikCtrl, TarCtrl, PcsCtrl, CouCtrl, SalCtrl)
                        DbTbl8.MoveNext()

                    Loop
                End If

            ElseIf RemCmb1.Text = "Packing" Then

                SQL = ""
                SQL = SQL & "Select * from 05_Packing_Table "
                SQL = SQL & "Where Nik = ('" & RemTbx1.Text & "') "
                SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') "
                OpenTbl(ADb, DbTbl9, SQL)
                If DbTbl9.RecordCount > 0 Then

                    DbTbl9.MoveFirst()
                    Do While Not DbTbl9.EOF

                        NikCtrl = DbTbl9("Nik").Value
                        ProID = DbTbl9("Process_ID").Value
                        TimeCtrl = DbTbl9("Time").Value
                        DateCtrl = DbTbl9("Date").Value
                        CartCtrl = DbTbl9("Carton").Value
                        TarCtrl = DbTbl9("Target").Value
                        SalCtrl = DbTbl9("Salary").Value
                        CouCtrl = DbTbl9("Coupon").Value

                        UpNRGrid.Rows.Add(ProID, DateCtrl, TimeCtrl, NikCtrl, CartCtrl, TarCtrl, CouCtrl, SalCtrl)
                        DbTbl9.MoveNext()

                    Loop
                End If

            ElseIf RemCmb1.Text = "Sortasi" Then

                SQL = ""
                SQL = SQL & "Select * from 21_NewMiscellaneous_Table "
                SQL = SQL & "Where Nik = ('" & RemTbx1.Text & "') "
                SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') "
                SQL = SQL & "Order by Process_ID Desc "

                OpenTbl(ADb, Dbtb37, SQL)
                If Dbtb37.RecordCount > 0 Then

                    Do While Not Dbtb37.EOF

                        NikCtrl = Dbtb37("Nik").Value
                        ProID = Dbtb37("Process_ID").Value
                        TimeCtrl = Dbtb37("Time").Value
                        DateCtrl = Dbtb37("Date").Value
                        PcsCtrl = Dbtb37("Pieces").Value
                        SalCtrl = Dbtb37("Salary").Value
                        CouCtrl = Dbtb37("Coupon").Value
                        NoKgCtrl = Dbtb37("NoKg").Value
                        NoGrCtrl = Dbtb37("NoGr").Value
                        NoBagCtrl = Dbtb37("NoBag").Value

                        UpNRGrid.Rows.Add(ProID, DateCtrl, TimeCtrl, NikCtrl, PcsCtrl, NoKgCtrl, NoBagCtrl, NoGrCtrl, CouCtrl, SalCtrl)
                        Dbtb37.MoveNext()

                    Loop

                End If

            ElseIf RemCmb1.Text = "Miscellaneous" Then

                SQL = ""
                SQL = SQL & "Select * from 19_Miscellaneous_Table "
                SQL = SQL & "Where Nik = ('" & RemTbx1.Text & "') "
                SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') "

                OpenTbl(ADb, Dbtb34, SQL)
                If Dbtb34.RecordCount > 0 Then
                    Do While Not Dbtb34.EOF

                        NikCtrl = Dbtb34("Nik").Value
                        ProID = Dbtb34("Process_ID").Value
                        TimeCtrl = Dbtb34("Time").Value
                        DateCtrl = Dbtb34("Date").Value
                        SalCtrl = Dbtb34("Salary").Value

                        UpNRGrid.Rows.Add(ProID, DateCtrl, TimeCtrl, NikCtrl, SalCtrl, CouCtrl)
                        Dbtb34.MoveNext()

                    Loop

                End If

            End If

            ' Per Salary

        ElseIf RmChk2.Checked = True Then

            If RemCmb1.Text = "Conveyour" Then

                SQL = ""
                SQL = SQL & "Select * from 13_Conveyour_Salary "
                SQL = SQL & "Where Nik = ('" & RemTbx1.Text & "') "
                SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') "
                OpenTbl(ADb, Dbtb21, SQL)
                If Dbtb21.RecordCount > 0 Then

                    DateSal = Dbtb21("Date").Value

                    UpNRGrid.Rows.Add(DateSal, Dbtb21("Nik").Value, Dbtb21("Salary").Value)

                End If

            ElseIf RemCmb1.Text = "Mutu II" Then

                SQL = ""
                SQL = SQL & "Select * from 14_MutuII_Salary "
                SQL = SQL & "Where Nik = ('" & RemTbx1.Text & "') "
                SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') "
                OpenTbl(ADb, Dbtb22, SQL)
                If Dbtb22.RecordCount > 0 Then

                    DateSal = Dbtb22("Date").Value

                    UpNRGrid.Rows.Add(DateSal, Dbtb22("Nik").Value, Dbtb22("Salary").Value)

                End If

            ElseIf RemCmb1.Text = "Wallet" Then

                SQL = ""
                SQL = SQL & "Select * from 15_Wallet_Salary "
                SQL = SQL & "Where Nik = ('" & RemTbx1.Text & "') "
                SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') "
                OpenTbl(ADb, Dbtb23, SQL)
                If Dbtb23.RecordCount > 0 Then

                    DateSal = Dbtb23("Date").Value

                    UpNRGrid.Rows.Add(DateSal, Dbtb23("Nik").Value, Dbtb23("Salary").Value)

                End If

            ElseIf RemCmb1.Text = "Packing" Then

                SQL = ""
                SQL = SQL & "Select * from 16_Packing_Salary "
                SQL = SQL & "Where Nik = ('" & RemTbx1.Text & "') "
                SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') "
                OpenTbl(ADb, Dbtb24, SQL)
                If Dbtb24.RecordCount > 0 Then

                    DateSal = Dbtb24("Date").Value

                    UpNRGrid.Rows.Add(DateSal, Dbtb24("Nik").Value, Dbtb24("Salary").Value)

                End If

            ElseIf RemCmb1.Text = "Miscellaneous" Then

                SQL = ""
                SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                SQL = SQL & "Where Nik = ('" & RemTbx1.Text & "') "
                SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') "
                SQL = SQL & "And TypeCtrl = ('" & "Old" & "') "
                OpenTbl(ADb, Dbtb35, SQL)
                If Dbtb35.RecordCount > 0 Then

                    DateSal = Dbtb35("Date").Value

                    UpNRGrid.Rows.Add(DateSal, Dbtb35("Nik").Value, Dbtb35("Salary").Value)

                End If

            ElseIf RemCmb1.Text = "Sortasi" Then

                SQL = ""
                SQL = SQL & "Select * from 20_Miscellaneous_Salary "
                SQL = SQL & "Where Nik = ('" & RemTbx1.Text & "') "
                SQL = SQL & "And Date = ('" & DateGet.ToString("yyyy-MM-dd") & "') "
                SQL = SQL & "And TypeCtrl = ('" & "New" & "') "
                OpenTbl(ADb, Dbtb37, SQL)
                If Dbtb37.RecordCount > 0 Then

                    DateSal = Dbtb37("Date").Value

                    UpNRGrid.Rows.Add(DateSal, Dbtb37("Nik").Value, Dbtb37("Salary").Value)

                End If

            End If

        ElseIf RmChk3.Checked = True Then

            SQL = ""
            SQL = SQL & "Select * from SalarySync1_Table "
            SQL = SQL & "Where Nik = ('" & RemTbx1.Text & "') "
            SQL = SQL & "And Periode = ('" & RemCmb2.Text & "') "
            SQL = SQL & "And PeriodeRange = ('" & RemCmb3.Text & "') "
            SQL = SQL & "Order by Nik "
            OpenTbl(CBb, Ctbl25, SQL)
            If Ctbl25.RecordCount > 0 Then
                Ctbl25.MoveFirst()
                Do While Not Ctbl25.EOF

                    Gnik = IIf(IsDBNull(Ctbl25("Nik").Value), "", Ctbl25("Nik").Value)
                    Gname = IIf(IsDBNull(Ctbl25("Name").Value), "", Ctbl25("Name").Value)
                    GPay = IIf(IsDBNull(Ctbl25("Pay").Value), "", Ctbl25("Pay").Value)
                    GAskVal = IIf(IsDBNull(Ctbl25("AstekVal").Value), "", Ctbl25("AstekVal").Value)
                    GSal1 = IIf(IsDBNull(Ctbl25("Salary1").Value), "", Ctbl25("Salary1").Value)
                    GSal2 = IIf(IsDBNull(Ctbl25("Salary2").Value), "", Ctbl25("Salary2").Value)
                    GSal3 = IIf(IsDBNull(Ctbl25("Salary3").Value), "", Ctbl25("Salary3").Value)
                    GSal4 = IIf(IsDBNull(Ctbl25("Salary4").Value), "", Ctbl25("Salary4").Value)
                    GSal5 = IIf(IsDBNull(Ctbl25("Salary5").Value), "", Ctbl25("Salary5").Value)
                    GSal6 = IIf(IsDBNull(Ctbl25("Salary6").Value), "", Ctbl25("Salary6").Value)
                    GSal7 = IIf(IsDBNull(Ctbl25("Salary7").Value), "", Ctbl25("Salary7").Value)
                    GSal8 = IIf(IsDBNull(Ctbl25("Salary8").Value), "", Ctbl25("Salary8").Value)
                    GSal9 = IIf(IsDBNull(Ctbl25("Salary9").Value), "", Ctbl25("Salary9").Value)
                    GSal10 = IIf(IsDBNull(Ctbl25("Salary10").Value), "", Ctbl25("Salary10").Value)
                    GSal11 = IIf(IsDBNull(Ctbl25("Salary11").Value), "", Ctbl25("Salary11").Value)
                    GSal12 = IIf(IsDBNull(Ctbl25("Salary12").Value), "", Ctbl25("Salary12").Value)
                    GSal13 = IIf(IsDBNull(Ctbl25("Salary13").Value), "", Ctbl25("Salary13").Value)
                    GSal14 = IIf(IsDBNull(Ctbl25("Salary14").Value), "", Ctbl25("Salary14").Value)
                    GSal15 = IIf(IsDBNull(Ctbl25("Salary15").Value), "", Ctbl25("Salary15").Value)
                    GSal16 = IIf(IsDBNull(Ctbl25("Salary16").Value), "", Ctbl25("Salary16").Value)
                    GSPotLain = (IIf(IsDBNull(Ctbl25("PotLain").Value), "", Ctbl25("PotLain").Value))

                    UpNRGrid.Rows.Add(Gnik, Gname, GSal1, GSal2, GSal3, GSal4, GSal5, GSal6, GSal7, GSal8, GSal9, GSal10, GSal11, GSal12, GSal13, GSal14, GSal15, GSal16, GPay, GAskVal, GSPotLain)

                    Ctbl25.MoveNext()

                Loop

            End If

        End If

    End Sub
    Dim DateGrab As Date
    Sub LookSave()

        If RmChk1.Checked = True Then

            If RemCmb1.Text = "Conveyour" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    RemGidVal3 = UpNRGrid(3, a).Value
                    RemGidVal4 = UpNRGrid(4, a).Value
                    RemGidVal5 = UpNRGrid(5, a).Value
                    RemGidVal6 = UpNRGrid(6, a).Value
                    RemGidVal7 = UpNRGrid(7, a).Value
                    DateGrab = RemGidVal1
                    SQL = ""
                    SQL = SQL & "Select * From 03_Conveyour_Table "
                    SQL = SQL & "Where Process_ID = ('" & RemGidVal0 & "') "
                    SQL = SQL & "And Nik = ('" & RemGidVal3 & "') "

                    OpenTbl(ADb, Atb5, SQL)

                    If Atb5.RecordCount > 0 Then

                        Atb5("Process_ID").Value = RemGidVal0
                        Atb5("Date").Value = DateGrab.ToString("yyyy-MM-dd")
                        Atb5("Time").Value = RemGidVal2
                        Atb5("Nik").Value = RemGidVal3
                        Atb5("Pieces").Value = RemGidVal4
                        Atb5("Target").Value = RemGidVal5
                        Atb5("Salary").Value = RemGidVal6

                        Atb5.Update()

                        UpNRGrid(7, a).Value = "Has Been Saved"

                    End If

                Next

            ElseIf RemCmb1.Text = "Mutu II" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    RemGidVal3 = UpNRGrid(3, a).Value
                    RemGidVal4 = UpNRGrid(4, a).Value
                    RemGidVal5 = UpNRGrid(5, a).Value
                    RemGidVal6 = UpNRGrid(6, a).Value
                    RemGidVal7 = UpNRGrid(7, a).Value
                    DateGrab = RemGidVal1
                    SQL = ""
                    SQL = SQL & "Select * From 04_MutuII_Table "
                    SQL = SQL & "Where Process_ID = ('" & RemGidVal0 & "') "
                    SQL = SQL & "And Nik = ('" & RemGidVal3 & "') "

                    OpenTbl(ADb, Atb5, SQL)

                    If Atb5.RecordCount > 0 Then

                        Atb5("Process_ID").Value = RemGidVal0
                        Atb5("Date").Value = DateGrab.ToString("yyyy-MM-dd")
                        Atb5("Time").Value = RemGidVal2
                        Atb5("Nik").Value = RemGidVal3
                        Atb5("Target").Value = RemGidVal4
                        Atb5("Pieces").Value = RemGidVal5
                        Atb5("Coupon").Value = RemGidVal6
                        Atb5("Salary").Value = RemGidVal7

                        Atb5.Update()

                        UpNRGrid(8, a).Value = "Has Been Saved"

                    End If

                Next

            ElseIf RemCmb1.Text = "Wallet" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    RemGidVal3 = UpNRGrid(3, a).Value
                    RemGidVal4 = UpNRGrid(4, a).Value
                    RemGidVal5 = UpNRGrid(5, a).Value
                    RemGidVal6 = UpNRGrid(6, a).Value
                    RemGidVal7 = UpNRGrid(7, a).Value
                    DateGrab = RemGidVal1
                    SQL = ""
                    SQL = SQL & "Select * From 06_Wallet_Table "
                    SQL = SQL & "Where Process_ID = ('" & RemGidVal0 & "') "
                    SQL = SQL & "And Nik = ('" & RemGidVal3 & "') "

                    OpenTbl(ADb, Atb5, SQL)

                    If Atb5.RecordCount > 0 Then

                        Atb5("Process_ID").Value = RemGidVal0
                        Atb5("Date").Value = DateGrab.ToString("yyyy-MM-dd")
                        Atb5("Time").Value = RemGidVal2
                        Atb5("Nik").Value = RemGidVal3
                        Atb5("Target").Value = RemGidVal4
                        Atb5("Pieces").Value = RemGidVal5
                        Atb5("Coupon").Value = RemGidVal6
                        Atb5("Salary").Value = RemGidVal7

                        Atb5.Update()

                        UpNRGrid(8, a).Value = "Has Been Saved"

                    End If

                Next

            ElseIf RemCmb1.Text = "Packing" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    RemGidVal3 = UpNRGrid(3, a).Value
                    RemGidVal4 = UpNRGrid(4, a).Value
                    RemGidVal5 = UpNRGrid(5, a).Value
                    RemGidVal6 = UpNRGrid(6, a).Value
                    RemGidVal7 = UpNRGrid(7, a).Value
                    DateGrab = RemGidVal1
                    SQL = ""
                    SQL = SQL & "Select * From 05_Packing_Table "
                    SQL = SQL & "Where Process_ID = ('" & RemGidVal0 & "') "
                    SQL = SQL & "And Nik = ('" & RemGidVal3 & "') "

                    OpenTbl(ADb, Atb5, SQL)

                    If Atb5.RecordCount > 0 Then

                        Atb5("Process_ID").Value = RemGidVal0
                        Atb5("Date").Value = DateGrab.ToString("yyyy-MM-dd")
                        Atb5("Time").Value = RemGidVal2
                        Atb5("Nik").Value = RemGidVal3
                        Atb5("Carton").Value = RemGidVal4
                        Atb5("Target").Value = RemGidVal5
                        Atb5("Coupon").Value = RemGidVal6
                        Atb5("Salary").Value = RemGidVal7

                        Atb5.Update()

                        UpNRGrid(8, a).Value = "Has Been Saved"

                    End If

                Next

            ElseIf RemCmb1.Text = "Sortasi" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    RemGidVal3 = UpNRGrid(3, a).Value
                    RemGidVal4 = UpNRGrid(4, a).Value
                    RemGidVal5 = UpNRGrid(5, a).Value
                    RemGidVal6 = UpNRGrid(6, a).Value
                    RemGidVal7 = UpNRGrid(7, a).Value
                    RemGidVal8 = UpNRGrid(8, a).Value
                    RemGidVal9 = UpNRGrid(9, a).Value
                    DateGrab = RemGidVal1
                    SQL = ""
                    SQL = SQL & "Select * From 21_NewMiscellaneous_Table "
                    SQL = SQL & "Where Process_ID = ('" & RemGidVal0 & "') "
                    SQL = SQL & "And Nik = ('" & RemGidVal3 & "') "

                    OpenTbl(ADb, Atb5, SQL)

                    If Atb5.RecordCount > 0 Then

                        Atb5("Process_ID").Value = RemGidVal0
                        Atb5("Date").Value = DateGrab.ToString("yyyy-MM-dd")

                        Atb5("Time").Value = RemGidVal2
                        Atb5("Nik").Value = RemGidVal3
                        Atb5("Pieces").Value = RemGidVal4
                        Atb5("NoKg").Value = RemGidVal5
                        Atb5("NoBag").Value = RemGidVal6
                        Atb5("NoGr").Value = RemGidVal7
                        Atb5("Coupon").Value = RemGidVal8
                        Atb5("Salary").Value = RemGidVal9

                        Atb5.Update()

                        UpNRGrid(10, a).Value = "Has Been Saved"

                    End If

                Next

            ElseIf RemCmb1.Text = "Miscellaneous" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    RemGidVal3 = UpNRGrid(3, a).Value
                    RemGidVal4 = UpNRGrid(4, a).Value
                    RemGidVal5 = UpNRGrid(5, a).Value
                    DateGrab = RemGidVal1
                    SQL = ""
                    SQL = SQL & "Select * From 19_Miscellaneous_Table "
                    SQL = SQL & "Where Process_ID = ('" & RemGidVal0 & "') "
                    SQL = SQL & "And Nik = ('" & RemGidVal3 & "') "

                    OpenTbl(ADb, Atb5, SQL)

                    If Atb5.RecordCount > 0 Then

                        Atb5("Process_ID").Value = RemGidVal0
                        Atb5("Date").Value = DateGrab.ToString("yyyy-MM-dd")
                        Atb5("Time").Value = RemGidVal2
                        Atb5("Nik").Value = RemGidVal3
                        Atb5("Salary").Value = RemGidVal4

                        Atb5.Update()

                        UpNRGrid(5, a).Value = "Has Been Saved"

                    End If

                Next

            End If

        ElseIf RmChk2.Checked = True Then

            If RemCmb1.Text = "Conveyour" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    DateGrab = RemGidVal0
                    SQL = ""
                    SQL = SQL & "Select * From 13_Conveyour_Salary "
                    SQL = SQL & "Where Nik = ('" & RemGidVal1 & "') "
                    SQL = SQL & "And Date = ('" & DateGrab.ToString("yyyy-MM-dd") & "') "

                    OpenTbl(ADb, Atbl30, SQL)

                    If Not Atbl30.RecordCount <> 0 Then
                        Atbl30.AddNew()
                    End If

                    Atbl30("Date").Value = DateGrab.ToString("yyyy-MM-dd")
                    Atbl30("Nik").Value = RemGidVal1
                    Atbl30("Salary").Value = RemGidVal2

                    Atbl30.Update()
                    Me.Refresh()

                Next

            ElseIf RemCmb1.Text = "MutuII" Then


            ElseIf RemCmb1.Text = "Wallet" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    DateGrab = RemGidVal0
                    SQL = ""
                    SQL = SQL & "Select * From 15_Wallet_Salary "
                    SQL = SQL & "Where Nik = ('" & RemGidVal1 & "') "
                    SQL = SQL & "And Date = ('" & DateGrab.ToString("yyyy-MM-dd") & "') "
                    OpenTbl(ADb, Atbl30, SQL)

                    If Not Atbl30.RecordCount <> 0 Then
                        Atbl30.AddNew()
                    End If

                    Atbl30("Date").Value = DateGrab.ToString("yyyy-MM-dd")
                    Atbl30("Nik").Value = RemGidVal1
                    Atbl30("Salary").Value = RemGidVal2

                    Atbl30.Update()
                    Me.Refresh()

                Next

            ElseIf RemCmb1.Text = "Packing" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    DateGrab = RemGidVal0
                    SQL = ""
                    SQL = SQL & "Select * From 16_Packing_Salary "
                    SQL = SQL & "Where Nik = ('" & RemGidVal1 & "') "
                    SQL = SQL & "And Date = ('" & DateGrab.ToString("yyyy-MM-dd") & "') "
                    OpenTbl(ADb, Atbl30, SQL)

                    If Not Atbl30.RecordCount <> 0 Then
                        Atbl30.AddNew()
                    End If

                    Atbl30("Date").Value = RemGidVal0
                    Atbl30("Nik").Value = RemGidVal1
                    Atbl30("Salary").Value = RemGidVal2

                    Atbl30.Update()
                    Me.Refresh()

                Next

            ElseIf RemCmb1.Text = "Sortasi" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    DateGrab = RemGidVal0
                    SQL = ""
                    SQL = SQL & "Select * From 20_Miscellaneous_Salary "
                    SQL = SQL & "Where Nik = ('" & RemGidVal1 & "') "
                    SQL = SQL & "And Date = ('" & DateGrab.ToString("yyyy-MM-dd") & "') "
                    SQL = SQL & "And TypeCtrl = ('" & "New" & "') "
                    OpenTbl(ADb, Atbl30, SQL)

                    If Not Atbl30.RecordCount <> 0 Then
                        Atbl30.AddNew()
                    End If

                    Atbl30("Date").Value = RemGidVal0
                    Atbl30("Nik").Value = RemGidVal1
                    Atbl30("Salary").Value = RemGidVal2
                    Atbl30("TypeCtrl").Value = "New"

                    Atbl30.Update()
                    Me.Refresh()

                Next

            ElseIf RemCmb1.Text = "Miscellaneous" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    DateGrab = RemGidVal0
                    SQL = ""
                    SQL = SQL & "Select * From 20_Miscellaneous_Salary "
                    SQL = SQL & "Where Nik = ('" & RemGidVal1 & "') "
                    SQL = SQL & "And Date = ('" & DateGrab.ToString("yyyy-MM-dd") & "') "
                    SQL = SQL & "And TypeCtrl = ('" & "Old" & "') "
                    OpenTbl(ADb, Atbl30, SQL)

                    If Not Atbl30.RecordCount <> 0 Then
                        Atbl30.AddNew()
                    End If

                    Atbl30("Date").Value = RemGidVal0
                    Atbl30("Nik").Value = RemGidVal1
                    Atbl30("Salary").Value = RemGidVal2
                    Atbl30("TypeCtrl").Value = "Old"

                    Atbl30.Update()

                    Me.Refresh()

                Next

            End If

        ElseIf RmChk3.Checked = True Then

            For a = 0 To UpNRGrid.Rows.Count - 1

                RemGidVal0 = UpNRGrid(0, a).Value
                RemGidVal1 = UpNRGrid(1, a).Value
                RemGidVal2 = UpNRGrid(2, a).Value
                RemGidVal3 = UpNRGrid(3, a).Value
                RemGidVal4 = UpNRGrid(4, a).Value
                RemGidVal5 = UpNRGrid(5, a).Value
                RemGidVal6 = UpNRGrid(6, a).Value
                RemGidVal7 = UpNRGrid(7, a).Value
                RemGidVal8 = UpNRGrid(8, a).Value
                RemGidVal9 = UpNRGrid(9, a).Value
                RemGidVal10 = UpNRGrid(10, a).Value
                RemGidVal11 = UpNRGrid(11, a).Value
                RemGidVal12 = UpNRGrid(12, a).Value
                RemGidVal13 = UpNRGrid(13, a).Value
                RemGidVal14 = UpNRGrid(14, a).Value
                RemGidVal15 = UpNRGrid(15, a).Value
                RemGidVal16 = UpNRGrid(16, a).Value
                RemGidVal17 = UpNRGrid(17, a).Value
                RemGidVal18 = UpNRGrid(18, a).Value
                RemGidVal19 = UpNRGrid(19, a).Value
                RemGidVal20 = UpNRGrid(20, a).Value

                SQL = ""
                SQL = SQL & "Select * from SalarySync1_Table "
                SQL = SQL & "Where Nik = ('" & RemTbx1.Text & "') "
                SQL = SQL & "And Periode = ('" & RemCmb2.Text & "') "
                SQL = SQL & "And PeriodeRange = ('" & RemCmb3.Text & "') "
                SQL = SQL & "Order by Nik "
                OpenTbl(CBb, Ctbl52, SQL)

                If Not Ctbl52.RecordCount <> 0 Then
                    Ctbl52.AddNew()
                End If

                Ctbl52("Nik").Value = RemGidVal0
                Ctbl52("Name").Value = RemGidVal1
                Ctbl52("Salary1").Value = RemGidVal2
                Ctbl52("Salary2").Value = RemGidVal3
                Ctbl52("Salary3").Value = RemGidVal4
                Ctbl52("Salary4").Value = RemGidVal5
                Ctbl52("Salary5").Value = RemGidVal6
                Ctbl52("Salary6").Value = RemGidVal7
                Ctbl52("Salary7").Value = RemGidVal8
                Ctbl52("Salary8").Value = RemGidVal9
                Ctbl52("Salary9").Value = RemGidVal10
                Ctbl52("Salary10").Value = RemGidVal11
                Ctbl52("Salary11").Value = RemGidVal12
                Ctbl52("Salary12").Value = RemGidVal13
                Ctbl52("Salary13").Value = RemGidVal14
                Ctbl52("Salary14").Value = RemGidVal15
                Ctbl52("Salary15").Value = RemGidVal16
                Ctbl52("Salary16").Value = RemGidVal17
                Ctbl52("Pay").Value = RemGidVal18
                Ctbl52("AstekVal").Value = RemGidVal19
                Ctbl52("PotLain").Value = RemGidVal20

                Ctbl52.Update()

                UpNRGrid(21, a).Value = "Has Been Saved"
            Next

        End If

    End Sub

    Sub LookDelete()

        If RmChk1.Checked = True Then

            If RemCmb1.Text = "Conveyour" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    RemGidVal3 = UpNRGrid(3, a).Value
                    RemGidVal4 = UpNRGrid(4, a).Value
                    RemGidVal5 = UpNRGrid(5, a).Value
                    RemGidVal6 = UpNRGrid(6, a).Value
                    RemGidVal7 = UpNRGrid(7, a).Value

                    SQL = ""
                    SQL = SQL & "Select * From 03_Conveyour_Table "
                    SQL = SQL & "Where Process_ID = ('" & RemGidVal0 & "') "
                    SQL = SQL & "And Nik = ('" & RemGidVal3 & "') "

                    OpenTbl(ADb, Atb5, SQL)

                    If Atb5.RecordCount > 0 Then

                        Atb5.Delete()

                        UpNRGrid(7, a).Value = "Has Been Deleted"

                    End If

                Next

            ElseIf RemCmb1.Text = "Mutu II" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    RemGidVal3 = UpNRGrid(3, a).Value
                    RemGidVal4 = UpNRGrid(4, a).Value
                    RemGidVal5 = UpNRGrid(5, a).Value
                    RemGidVal6 = UpNRGrid(6, a).Value
                    RemGidVal7 = UpNRGrid(7, a).Value

                    SQL = ""
                    SQL = SQL & "Select * From 04_MutuII_Table "
                    SQL = SQL & "Where Process_ID = ('" & RemGidVal0 & "') "
                    SQL = SQL & "And Nik = ('" & RemGidVal3 & "') "

                    OpenTbl(ADb, Atb5, SQL)

                    If Atb5.RecordCount > 0 Then

                        Atb5.Delete()

                        UpNRGrid(8, a).Value = "Has Been Deleted"

                    End If

                Next

            ElseIf RemCmb1.Text = "Wallet" Then


                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    RemGidVal3 = UpNRGrid(3, a).Value
                    RemGidVal4 = UpNRGrid(4, a).Value
                    RemGidVal5 = UpNRGrid(5, a).Value
                    RemGidVal6 = UpNRGrid(6, a).Value
                    RemGidVal7 = UpNRGrid(7, a).Value

                    SQL = ""
                    SQL = SQL & "Select * From 06_Wallet_Table "
                    SQL = SQL & "Where Process_ID = ('" & RemGidVal0 & "') "
                    SQL = SQL & "And Nik = ('" & RemGidVal3 & "') "

                    OpenTbl(ADb, Atb5, SQL)

                    If Atb5.RecordCount > 0 Then

                        Atb5.Delete()

                        UpNRGrid(8, a).Value = "Has Been Deleted"

                    End If

                Next

            ElseIf RemCmb1.Text = "Packing" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    RemGidVal3 = UpNRGrid(3, a).Value
                    RemGidVal4 = UpNRGrid(4, a).Value
                    RemGidVal5 = UpNRGrid(5, a).Value
                    RemGidVal6 = UpNRGrid(6, a).Value
                    RemGidVal7 = UpNRGrid(7, a).Value

                    SQL = ""
                    SQL = SQL & "Select * From 05_Packing_Table "
                    SQL = SQL & "Where Process_ID = ('" & RemGidVal0 & "') "
                    SQL = SQL & "And Nik = ('" & RemGidVal3 & "') "

                    OpenTbl(ADb, Atb5, SQL)

                    If Atb5.RecordCount > 0 Then

                        Atb5.Delete()

                        UpNRGrid(8, a).Value = "Has Been Deleted"

                    End If
                Next

            ElseIf RemCmb1.Text = "Sortasi" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    RemGidVal3 = UpNRGrid(3, a).Value
                    RemGidVal4 = UpNRGrid(4, a).Value
                    RemGidVal5 = UpNRGrid(5, a).Value
                    RemGidVal6 = UpNRGrid(6, a).Value
                    RemGidVal7 = UpNRGrid(7, a).Value
                    RemGidVal8 = UpNRGrid(8, a).Value
                    RemGidVal9 = UpNRGrid(9, a).Value

                    SQL = ""
                    SQL = SQL & "Select * From 21_NewMiscellaneous_Table "
                    SQL = SQL & "Where Process_ID = ('" & RemGidVal0 & "') "
                    SQL = SQL & "And Nik = ('" & RemGidVal3 & "') "

                    OpenTbl(ADb, Atb5, SQL)

                    If Atb5.RecordCount > 0 Then

                        Atb5.Delete()

                        UpNRGrid(10, a).Value = "Has Been Deleted"

                    End If

                Next

            ElseIf RemCmb1.Text = "Miscellaneous" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    RemGidVal3 = UpNRGrid(3, a).Value
                    RemGidVal4 = UpNRGrid(4, a).Value
                    RemGidVal5 = UpNRGrid(5, a).Value

                    SQL = ""
                    SQL = SQL & "Select * From 19_Miscellaneous_Table "
                    SQL = SQL & "Where Process_ID = ('" & RemGidVal0 & "') "
                    SQL = SQL & "And Nik = ('" & RemGidVal3 & "') "

                    OpenTbl(ADb, Atb5, SQL)

                    If Atb5.RecordCount > 0 Then

                        Atb5.Delete()

                        UpNRGrid(5, a).Value = "Has Been Deleted"

                    End If

                Next

            End If

        ElseIf RmChk2.Checked = True Then

            If RemCmb1.Text = "Conveyour" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    DateGrab = RemGidVal0
                    SQL = ""
                    SQL = SQL & "Select * From 13_Conveyour_Salary "
                    SQL = SQL & "Where Nik = ('" & RemGidVal1 & "') "
                    SQL = SQL & "And Date = ('" & DateGrab.ToString("yyyy-MM-dd") & "') "

                    OpenTbl(ADb, Atbl30, SQL)

                    If Atbl30.RecordCount > 0 Then

                        Atbl30.Delete()
                        UpNRGrid(3, a).Value = "Has Been Deleted"

                    End If
                Next

            ElseIf RemCmb1.Text = "Mutu II" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    DateGrab = RemGidVal0
                    SQL = ""
                    SQL = SQL & "Select * From 14_MutuII_Salary "
                    SQL = SQL & "Where Nik = ('" & RemGidVal1 & "') "
                    SQL = SQL & "And Date = ('" & DateGrab.ToString("yyyy-MM-dd") & "') "
                    OpenTbl(ADb, Atbl30, SQL)

                    If Atbl30.RecordCount > 0 Then
                        Atbl30.Delete()
                        UpNRGrid(3, a).Value = "Has Been Deleted"
                    End If

                Next

            ElseIf RemCmb1.Text = "Wallet" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    DateGrab = RemGidVal0
                    SQL = ""
                    SQL = SQL & "Select * From 15_Wallet_Salary "
                    SQL = SQL & "Where Nik = ('" & RemGidVal1 & "') "
                    SQL = SQL & "And Date = ('" & DateGrab.ToString("yyyy-MM-dd") & "') "
                    OpenTbl(ADb, Atbl30, SQL)

                    If Atbl30.RecordCount > 0 Then
                        Atbl30.Delete()
                        UpNRGrid(3, a).Value = "Has Been Deleted"
                    End If

                Next

            ElseIf RemCmb1.Text = "Packing" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    DateGrab = RemGidVal0
                    SQL = ""
                    SQL = SQL & "Select * From 16_Packing_Salary "
                    SQL = SQL & "Where Nik = ('" & RemGidVal1 & "') "
                    SQL = SQL & "And Date = ('" & DateGrab.ToString("yyyy-MM-dd") & "') "
                    OpenTbl(ADb, Atbl30, SQL)

                    If Atbl30.RecordCount > 0 Then
                        Atbl30.Delete()
                        UpNRGrid(3, a).Value = "Has Been Deleted"
                    End If

                Next

            ElseIf RemCmb1.Text = "Sortasi" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    DateGrab = RemGidVal0
                    SQL = ""
                    SQL = SQL & "Select * From 20_Miscellaneous_Salary "
                    SQL = SQL & "Where Nik = ('" & RemGidVal1 & "') "
                    SQL = SQL & "And Date = ('" & DateGrab.ToString("yyyy-MM-dd") & "') "
                    SQL = SQL & "And TypeCtrl = ('" & "New" & "') "

                    OpenTbl(ADb, Atbl30, SQL)

                    If Atbl30.RecordCount > 0 Then

                        Atbl30.Delete()
                        UpNRGrid(3, a).Value = "Has Been Deleted"

                    End If
                Next

            ElseIf RemCmb1.Text = "Miscellaneous" Then

                For a = 0 To UpNRGrid.Rows.Count - 1

                    RemGidVal0 = UpNRGrid(0, a).Value
                    RemGidVal1 = UpNRGrid(1, a).Value
                    RemGidVal2 = UpNRGrid(2, a).Value
                    DateGrab = RemGidVal0
                    SQL = ""
                    SQL = SQL & "Select * From 20_Miscellaneous_Salary "
                    SQL = SQL & "Where Nik = ('" & RemGidVal1 & "') "
                    SQL = SQL & "And Date = ('" & DateGrab.ToString("yyyy-MM-dd") & "') "
                    SQL = SQL & "And TypeCtrl = ('" & "Old" & "') "

                    OpenTbl(ADb, Atbl30, SQL)

                    If Atbl30.RecordCount > 0 Then

                        Atbl30.Delete()
                        UpNRGrid(3, a).Value = "Has Been Deleted"

                    End If
                Next

            End If

        ElseIf RmChk3.Checked = True Then

            For a = 0 To UpNRGrid.Rows.Count - 1

                RemGidVal0 = UpNRGrid(0, a).Value
                RemGidVal1 = UpNRGrid(1, a).Value

                SQL = ""
                SQL = SQL & "Select * from SalarySync1_Table "
                SQL = SQL & "Where Nik = ('" & RemGidVal0 & "') "
                SQL = SQL & "And Name = ('" & RemGidVal1 & "') "
                SQL = SQL & "And Periode = ('" & RemCmb2.Text & "') "
                SQL = SQL & "And PeriodeRange = ('" & RemCmb3.Text & "') "
                OpenTbl(CBb, Ctbl52, SQL)

                If Ctbl52.RecordCount > 0 Then
                    Ctbl52.Delete()
                    UpNRGrid(21, a).Value = "Has Been Deleted"
                End If

            Next


        End If

    End Sub

    Sub LookRefresh()

        RemGidVal0 = Nothing
        RemGidVal1 = Nothing
        RemGidVal2 = Nothing
        RemGidVal3 = Nothing
        RemGidVal4 = Nothing
        RemGidVal5 = Nothing
        RemGidVal6 = Nothing
        RemGidVal7 = Nothing
        RemGidVal8 = Nothing
        RemGidVal9 = Nothing
        RemGidVal10 = Nothing
        RemGidVal11 = Nothing
        RemGidVal12 = Nothing
        RemGidVal13 = Nothing
        RemGidVal14 = Nothing
        RemGidVal15 = Nothing
        RemGidVal16 = Nothing
        RemGidVal17 = Nothing
        RemGidVal18 = Nothing
        RemGidVal19 = Nothing
        RemGidVal20 = Nothing


    End Sub

#End Region

#Region "GUI Text Generator"

    'Sub FillDateCmb()

    '    CmbDater = Format(Now, "yyyy")
    '    CmbDater2 = CmbDater - 1
    '    CmbDater3 = CmbDater + 1

    '    With RemCmb3

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

    Sub RemGridHeader()

        UpNRGrid.Rows.Clear()
        UpNRGrid.Columns.Clear()

        With UpNRGrid

            If RmChk1.Checked = True Then

                If RemCmb1.Text = "Conveyour" Then
                    .Columns.Add("Col1", "Process ID")
                    .Columns.Add("Col2", "Date")
                    .Columns.Add("Col3", "Time")
                    .Columns.Add("Col4", "Nik")
                    .Columns.Add("Col5", "Pieces")
                    .Columns.Add("Col6", "Target")
                    .Columns.Add("Col7", "Salary")
                    .Columns.Add("Col8", "Status")

                    .Columns(0).ReadOnly = True
                    .Columns(1).ReadOnly = True
                    .Columns(2).ReadOnly = True
                    .Columns(3).ReadOnly = True
                    .Columns(7).ReadOnly = True

                ElseIf RemCmb1.Text = "Packing" Then

                    .Columns.Add("Col1", "Process ID")
                    .Columns.Add("Col2", "Date")
                    .Columns.Add("Col3", "Time")
                    .Columns.Add("Col4", "Nik")
                    .Columns.Add("Col5", "Carton")
                    .Columns.Add("Col6", "Target")
                    .Columns.Add("Col7", "Coupon")
                    .Columns.Add("Col8", "Salary")
                    .Columns.Add("Col9", "Status")

                    .Columns(0).ReadOnly = True
                    .Columns(1).ReadOnly = True
                    .Columns(2).ReadOnly = True
                    .Columns(3).ReadOnly = True
                    .Columns(6).ReadOnly = True
                    .Columns(8).ReadOnly = True

                ElseIf RemCmb1.Text = "Wallet" Then

                    .Columns.Add("Col1", "Process ID")
                    .Columns.Add("Col2", "Date")
                    .Columns.Add("Col3", "Time")
                    .Columns.Add("Col4", "Nik")
                    .Columns.Add("Col5", "Pieces")
                    .Columns.Add("Col6", "Target")
                    .Columns.Add("Col7", "Coupon")
                    .Columns.Add("Col8", "Salary")
                    .Columns.Add("Col9", "Status")

                    .Columns(0).ReadOnly = True
                    .Columns(1).ReadOnly = True
                    .Columns(2).ReadOnly = True
                    .Columns(3).ReadOnly = True
                    .Columns(6).ReadOnly = True

                ElseIf RemCmb1.Text = "Mutu II" Then

                    .Columns.Add("Col1", "Process ID")
                    .Columns.Add("Col2", "Date")
                    .Columns.Add("Col3", "Time")
                    .Columns.Add("Col4", "Nik")
                    .Columns.Add("Col5", "Pieces")
                    .Columns.Add("Col6", "Target")
                    .Columns.Add("Col7", "Coupon")
                    .Columns.Add("Col8", "Salary")
                    .Columns.Add("Col9", "Status")

                    .Columns(0).ReadOnly = True
                    .Columns(1).ReadOnly = True
                    .Columns(2).ReadOnly = True
                    .Columns(3).ReadOnly = True
                    .Columns(8).ReadOnly = True

                ElseIf RemCmb1.Text = "Miscellaneous" Then

                    .Columns.Add("Col1", "Process ID")
                    .Columns.Add("Col2", "Date")
                    .Columns.Add("Col3", "Time")
                    .Columns.Add("Col4", "Nik")
                    .Columns.Add("Col5", "Salary")
                    .Columns.Add("Col6", "Status")

                    .Columns(0).ReadOnly = True
                    .Columns(1).ReadOnly = True
                    .Columns(2).ReadOnly = True
                    .Columns(3).ReadOnly = True
                    .Columns(5).ReadOnly = True

                ElseIf RemCmb1.Text = "Sortasi" Then

                    .Columns.Add("Col1", "Process ID")
                    .Columns.Add("Col2", "Date")
                    .Columns.Add("Col3", "Time")
                    .Columns.Add("Col4", "Nik")
                    .Columns.Add("Col5", "Pieces")
                    .Columns.Add("Col6", "NoKg")
                    .Columns.Add("Col7", "NoBag")
                    .Columns.Add("Col8", "NoGr")
                    .Columns.Add("Col9", "Coupon")
                    .Columns.Add("Col10", "Salary")
                    .Columns.Add("Col11", "Status")

                    .Columns(0).ReadOnly = True
                    .Columns(1).ReadOnly = True
                    .Columns(2).ReadOnly = True
                    .Columns(3).ReadOnly = True
                    .Columns(8).ReadOnly = True
                    .Columns(10).ReadOnly = True

                End If

            ElseIf RmChk2.Checked = True Then

                If RemCmb1.Text = "Miscellaneous" Then

                    .Columns.Add("Col1", "Date")
                    .Columns.Add("Col2", "Nik")
                    .Columns.Add("Col3", "Salary")
                    .Columns.Add("Col4", "Status")



                ElseIf RemCmb1.Text = "Sortasi" Then

                    .Columns.Add("Col1", "Date")
                    .Columns.Add("Col2", "Nik")
                    .Columns.Add("Col3", "Salary")
                    .Columns.Add("Col4", "Status")


                Else

                    .Columns.Add("Col1", "Date")
                    .Columns.Add("Col2", "Nik")
                    .Columns.Add("Col3", "Salary")
                    .Columns.Add("Col4", "Status")

                    .Columns(0).ReadOnly = True
                    .Columns(1).ReadOnly = True

                End If

            ElseIf RmChk2.Checked = True Then

                .Columns.Add("Col1", "Process ID")
                .Columns.Add("Col2", "Date")
                .Columns.Add("Col3", "Time")
                .Columns.Add("Col4", "Nik")
                .Columns.Add("Col5", "Pieces")
                .Columns.Add("Col6", "NoKg")
                .Columns.Add("Col7", "NoBag")
                .Columns.Add("Col8", "NoGr")
                .Columns.Add("Col9", "Coupon")
                .Columns.Add("Col10", "Salary")
                .Columns.Add("Col11", "Status")
                .Columns.Add("Col12", "Pay")
                .Columns.Add("Col13", "Astek Val")

                .Columns(0).ReadOnly = True
                .Columns(1).ReadOnly = True
                .Columns(2).ReadOnly = True
                .Columns(3).ReadOnly = True
                .Columns(8).ReadOnly = True
                .Columns(10).ReadOnly = True
                .Columns(11).ReadOnly = True
                .Columns(12).ReadOnly = True

            ElseIf RmChk3.Checked = True Then


                .Columns.Add("Col0", "Nik")
                .Columns.Add("Col1", "Name")
                .Columns.Add("Col2", "Salary 1")
                .Columns.Add("Col3", "Salary 2")
                .Columns.Add("Col4", "Salary 3")
                .Columns.Add("Col5", "Salary 4")
                .Columns.Add("Col6", "Salary 5")
                .Columns.Add("Col7", "Salary 6")
                .Columns.Add("Col8", "Salary 7")
                .Columns.Add("Col9", "Salary 8")
                .Columns.Add("Col10", "Salary 9")
                .Columns.Add("Col11", "Salary 10")
                .Columns.Add("Col12", "Salary 11")
                .Columns.Add("Col13", "Salary 12")
                .Columns.Add("Col14", "Salary 13")
                .Columns.Add("Col15", "Salary 14")
                .Columns.Add("Col16", "Salary 15")
                .Columns.Add("Col17", "Salary 16")
                .Columns.Add("Col18", "Pay")
                .Columns.Add("Col19", "Astek Value")
                .Columns.Add("Col20", "Pot Lain")
                .Columns.Add("Col21", "Status")


                .Columns(0).ReadOnly = True
                .Columns(1).ReadOnly = True
                .Columns(2).ReadOnly = True
                .Columns(3).ReadOnly = True
                .Columns(4).ReadOnly = True
                .Columns(5).ReadOnly = True
                .Columns(6).ReadOnly = True
                .Columns(7).ReadOnly = True
                .Columns(8).ReadOnly = True
                .Columns(9).ReadOnly = True
                .Columns(10).ReadOnly = True
                .Columns(11).ReadOnly = True
                .Columns(12).ReadOnly = True
                .Columns(13).ReadOnly = True
                .Columns(14).ReadOnly = True
                .Columns(15).ReadOnly = True
                .Columns(16).ReadOnly = True
                .Columns(17).ReadOnly = True
                .Columns(18).ReadOnly = True
                .Columns(21).ReadOnly = True

            End If

        End With

    End Sub

#End Region

    Private Sub RemCmb3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)

        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True

    End Sub

    Private Sub RmChk3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RmChk3.CheckedChanged

        RemCmb1.Enabled = False
        RemCmb2.Enabled = True
        RemCmb3.Enabled = True
        RmBtn2.Enabled = True
        RemDP.Enabled = False

    End Sub

    Private Sub RmBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RmBtn1.Click

        RemGridHeader()
        LookValue()

    End Sub

    Private Sub RemTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles RemTbx1.KeyPress

        If e.KeyChar.ToString = ChrW(Keys.Enter) Then

            LookData()
            e.Handled = True

        End If

    End Sub
   
    Private Sub RmBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RmBtn2.Click
        LookSave()
        LookRefresh()
    End Sub

    Private Sub RmBtn3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RmBtn3.Click

        Dim result As Integer = MessageBox.Show("Are you Sure?", "Codex" + BuildCounter, MessageBoxButtons.YesNo)

        If result = DialogResult.Yes Then
            LookDelete()
            LookRefresh()
        End If


    End Sub

End Class