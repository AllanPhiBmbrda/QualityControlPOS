Public Class MainTotBlock
    Dim SyMaDater As String
    Dim SyMaDater2 As String
    Dim SyMaDater3 As String
    Dim LookTot As String
    Dim LookDate As String
    Dim LookPay As String
    Dim LookAstek As String
    Dim MotNik As String
    Dim MotName As String
    Dim MotSalary As String
    Dim MotDate As String
    Dim MotPay As String
    Dim MotAst As String
    Dim UploadCount As String = 0
    Dim SaveLoadCount As String = 0



    Private Sub DokLook_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim NewMDIChild As New DateBlock()
        DateBlock.MdiParent = MainMenu
        DateBlock.Show()
        Me.Refresh()
        MainMenu.Refresh()

    End Sub

    Private Sub MainTotBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MTRanSet()
        LoadDB()
        LoadDB2()
        LoadDB4()


    End Sub

    Sub EmpLookMe()
        SQL = ""
        SQL = SQL & "Select * from 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & NewGSal(0) & "') "
        OpenTbl(ADb, Atb3, SQL)

        If Atb3.RecordCount > 0 Then

            LookPay = IIf(IsDBNull(Atb3("Pay").Value), "", Atb3("Pay").Value)
            LookAstek = IIf(IsDBNull(Atb3("Jamsostek").Value), "", Atb3("Jamsostek").Value)

        End If
    End Sub

    Sub MTHeader()
        With MoTGrid1
            .Columns.Add("Col1", "Nik")
            .Columns.Add("Col2", "Name")
            .Columns.Add("Col3", "Main Salary")
            .Columns.Add("Col4", "Main Date")
            .Columns.Add("Col5", "Pay")
            .Columns.Add("Col6", "Astek")
            .Columns.Add("Col7", "Status")
        End With
    End Sub

    Sub AstPayLoad()


    End Sub

    Sub MTRanSet()

        SyMaDater = Format(Now, "yyyy")
        SyMaDater2 = SyMaDater - 1
        SyMaDater3 = SyMaDater + 1

        With MTCmb1

            .Items.Add("Dec " + SyMaDater2)
            .Items.Add("Jan " + Format(Now, "yyyy"))
            .Items.Add("Feb " + Format(Now, "yyyy"))
            .Items.Add("Mar " + Format(Now, "yyyy"))
            .Items.Add("Apr " + Format(Now, "yyyy"))
            .Items.Add("May " + Format(Now, "yyyy"))
            .Items.Add("Jun " + Format(Now, "yyyy"))
            .Items.Add("Jul " + Format(Now, "yyyy"))
            .Items.Add("Aug " + Format(Now, "yyyy"))
            .Items.Add("Sep " + Format(Now, "yyyy"))
            .Items.Add("Oct " + Format(Now, "yyyy"))
            .Items.Add("Nov " + Format(Now, "yyyy"))
            .Items.Add("Dec " + Format(Now, "yyyy"))
            .Items.Add("Jan " + SyMaDater3)

        End With

    End Sub
    Sub FillRadio()
        SQL = ""
        SQL = SQL & "Select * From DateCounter2Table "
        SQL = SQL & "Where Periode = ('" & MTCmb2.Text & "') "
        SQL = SQL & "And PeriodeRange = ('" & MTCmb1.Text & "') "
        OpenTbl(FBb, Ftbl24, SQL)


        If Ftbl24.RecordCount > 0 Then

            NewSald(0) = IIf(IsDBNull(Ftbl24("Date1").Value), "", Ftbl24("Date1").Value)
            NewSald(1) = IIf(IsDBNull(Ftbl24("Date2").Value), "", Ftbl24("Date2").Value)
            NewSald(2) = IIf(IsDBNull(Ftbl24("Date3").Value), "", Ftbl24("Date3").Value)
            NewSald(3) = IIf(IsDBNull(Ftbl24("Date4").Value), "", Ftbl24("Date4").Value)
            NewSald(4) = IIf(IsDBNull(Ftbl24("Date5").Value), "", Ftbl24("Date5").Value)
            NewSald(5) = IIf(IsDBNull(Ftbl24("Date6").Value), "", Ftbl24("Date6").Value)
            NewSald(6) = IIf(IsDBNull(Ftbl24("Date7").Value), "", Ftbl24("Date7").Value)
            NewSald(7) = IIf(IsDBNull(Ftbl24("Date8").Value), "", Ftbl24("Date8").Value)
            NewSald(8) = IIf(IsDBNull(Ftbl24("Date9").Value), "", Ftbl24("Date9").Value)
            NewSald(9) = IIf(IsDBNull(Ftbl24("Date10").Value), "", Ftbl24("Date10").Value)
            NewSald(10) = IIf(IsDBNull(Ftbl24("Date11").Value), "", Ftbl24("Date11").Value)
            NewSald(11) = IIf(IsDBNull(Ftbl24("Date12").Value), "", Ftbl24("Date12").Value)
            NewSald(12) = IIf(IsDBNull(Ftbl24("Date13").Value), "", Ftbl24("Date13").Value)
            NewSald(13) = IIf(IsDBNull(Ftbl24("Date14").Value), "", Ftbl24("Date14").Value)
            NewSald(14) = IIf(IsDBNull(Ftbl24("Date15").Value), "", Ftbl24("Date15").Value)
            NewSald(15) = IIf(IsDBNull(Ftbl24("Date16").Value), "", Ftbl24("Date16").Value)

        Else

            For i = 0 To 15
                NewSald(i) = Nothing
            Next

        End If

        RdDate1.Text = NewSald(0)
        RdDate2.Text = NewSald(1)
        RdDate3.Text = NewSald(2)
        RdDate4.Text = NewSald(3)
        RdDate5.Text = NewSald(4)
        RdDate6.Text = NewSald(5)
        RdDate7.Text = NewSald(6)
        RdDate8.Text = NewSald(7)
        RdDate9.Text = NewSald(8)
        RdDate10.Text = NewSald(9)
        RdDate11.Text = NewSald(10)
        RdDate12.Text = NewSald(11)
        RdDate13.Text = NewSald(12)
        RdDate14.Text = NewSald(13)
        RdDate15.Text = NewSald(14)
        RdDate16.Text = NewSald(15)



    End Sub

    Sub FillSalary()

        MoTGrid1.Rows.Clear()

        SQL = ""
        SQL = SQL & "Select * from SalarySync1_Table "
        SQL = SQL & "Where Periode = ('" & MTCmb2.Text & "') "
        SQL = SQL & "And PeriodeRange = ('" & MTCmb1.Text & "') "
        SQL = SQL & "Order by Nik Asc"
        OpenTbl(FBb, Ftbl25, SQL)


        If RdDate1.Checked = True Then
            MTHeader()

            If Ftbl25.RecordCount <> 0 Then
                Ftbl25.MoveFirst()
                Do While Not Ftbl25.EOF
                    NewGSal(0) = Ftbl25("Nik").Value
                    NewGSal(1) = Ftbl25("Name").Value
                    LookTot = IIf(IsDBNull(Ftbl25("Salary1").Value), "", Ftbl25("Salary1").Value)
                    LookDate = RdDate1.Text
                    EmpLookMe()
                    MoTGrid1.Rows.Add(NewGSal(0), NewGSal(1), LookTot, LookDate, LookPay, LookAstek)
                    UploadCount = UploadCount + 1
                    RecTbx1.Text = UploadCount
                    Ftbl25.MoveNext()
                Loop
            End If



        ElseIf RdDate2.Checked = True Then
            MTHeader()

            If Ftbl25.RecordCount <> 0 Then
                Ftbl25.MoveFirst()
                Do While Not Ftbl25.EOF
                    NewGSal(0) = Ftbl25("Nik").Value
                    NewGSal(1) = Ftbl25("Name").Value
                    LookTot = IIf(IsDBNull(Ftbl25("Salary2").Value), "", Ftbl25("Salary2").Value)
                    LookDate = RdDate2.Text
                    EmpLookMe()
                    MoTGrid1.Rows.Add(NewGSal(0), NewGSal(1), LookTot, LookDate, LookPay, LookAstek)
                    UploadCount = UploadCount + 1
                    RecTbx1.Text = UploadCount
                    Ftbl25.MoveNext()
                Loop
            End If

        ElseIf RdDate3.Checked = True Then

            MTHeader()
            If Ftbl25.RecordCount <> 0 Then
                Ftbl25.MoveFirst()
                Do While Not Ftbl25.EOF
                    NewGSal(0) = Ftbl25("Nik").Value
                    NewGSal(1) = Ftbl25("Name").Value
                    LookTot = IIf(IsDBNull(Ftbl25("Salary3").Value), "", Ftbl25("Salary3").Value)
                    LookDate = RdDate3.Text
                    EmpLookMe()
                    MoTGrid1.Rows.Add(NewGSal(0), NewGSal(1), LookTot, LookDate, LookPay, LookAstek)
                    UploadCount = UploadCount + 1
                    RecTbx1.Text = UploadCount
                    Ftbl25.MoveNext()
                Loop
            End If

        ElseIf RdDate4.Checked = True Then
            MTHeader()
            If Ftbl25.RecordCount <> 0 Then
                Ftbl25.MoveFirst()

                Do While Not Ftbl25.EOF
                    NewGSal(0) = Ftbl25("Nik").Value
                    NewGSal(1) = Ftbl25("Name").Value
                    LookTot = IIf(IsDBNull(Ftbl25("Salary4").Value), "", Ftbl25("Salary4").Value)
                    LookDate = RdDate4.Text
                    EmpLookMe()
                    MoTGrid1.Rows.Add(NewGSal(0), NewGSal(1), LookTot, LookDate, LookPay, LookAstek)
                    UploadCount = UploadCount + 1
                    RecTbx1.Text = UploadCount
                    Ftbl25.MoveNext()
                Loop
            End If


        ElseIf RdDate5.Checked = True Then
            MTHeader()

            If Ftbl25.RecordCount <> 0 Then
                Ftbl25.MoveFirst()

                Do While Not Ftbl25.EOF
                    NewGSal(0) = Ftbl25("Nik").Value
                    NewGSal(1) = Ftbl25("Name").Value
                    LookTot = IIf(IsDBNull(Ftbl25("Salary5").Value), "", Ftbl25("Salary5").Value)
                    LookDate = RdDate5.Text
                    EmpLookMe()
                    MoTGrid1.Rows.Add(NewGSal(0), NewGSal(1), LookTot, LookDate, LookPay, LookAstek)
                    UploadCount = UploadCount + 1
                    RecTbx1.Text = UploadCount
                    Ftbl25.MoveNext()
                Loop
            End If


        ElseIf RdDate6.Checked = True Then
            MTHeader()

            If Ftbl25.RecordCount <> 0 Then
                Ftbl25.MoveFirst()
                Do While Not Ftbl25.EOF
                    NewGSal(0) = Ftbl25("Nik").Value
                    NewGSal(1) = Ftbl25("Name").Value
                    LookTot = IIf(IsDBNull(Ftbl25("Salary6").Value), "", Ftbl25("Salary6").Value)
                    LookDate = RdDate6.Text
                    EmpLookMe()
                    MoTGrid1.Rows.Add(NewGSal(0), NewGSal(1), LookTot, LookDate, LookPay, LookAstek)
                    UploadCount = UploadCount + 1
                    RecTbx1.Text = UploadCount
                    Ftbl25.MoveNext()
                Loop
            End If

        ElseIf RdDate7.Checked = True Then
            MTHeader()

            If Ftbl25.RecordCount <> 0 Then
                Ftbl25.MoveFirst()
                Do While Not Ftbl25.EOF
                    NewGSal(0) = Ftbl25("Nik").Value
                    NewGSal(1) = Ftbl25("Name").Value
                    LookTot = IIf(IsDBNull(Ftbl25("Salary7").Value), "", Ftbl25("Salary7").Value)
                    LookDate = RdDate7.Text
                    EmpLookMe()
                    MoTGrid1.Rows.Add(NewGSal(0), NewGSal(1), LookTot, LookDate, LookPay, LookAstek)
                    UploadCount = UploadCount + 1
                    RecTbx1.Text = UploadCount
                    Ftbl25.MoveNext()
                Loop
            End If

        ElseIf RdDate8.Checked = True Then
            MTHeader()

            If Ftbl25.RecordCount <> 0 Then
                Ftbl25.MoveFirst()

                Do While Not Ftbl25.EOF
                    NewGSal(0) = Ftbl25("Nik").Value
                    NewGSal(1) = Ftbl25("Name").Value
                    LookTot = IIf(IsDBNull(Ftbl25("Salary8").Value), "", Ftbl25("Salary8").Value)
                    LookDate = RdDate8.Text
                    EmpLookMe()
                    MoTGrid1.Rows.Add(NewGSal(0), NewGSal(1), LookTot, LookDate, LookPay, LookAstek)
                    UploadCount = UploadCount + 1
                    RecTbx1.Text = UploadCount
                    Ftbl25.MoveNext()
                Loop
            End If

        ElseIf RdDate9.Checked = True Then

            MTHeader()

            If Ftbl25.RecordCount <> 0 Then
                Ftbl25.MoveFirst()

                Do While Not Ftbl25.EOF
                    NewGSal(0) = Ftbl25("Nik").Value
                    NewGSal(1) = Ftbl25("Name").Value
                    LookTot = IIf(IsDBNull(Ftbl25("Salary9").Value), "", Ftbl25("Salary9").Value)
                    LookDate = RdDate9.Text
                    EmpLookMe()
                    MoTGrid1.Rows.Add(NewGSal(0), NewGSal(1), LookTot, LookDate, LookPay, LookAstek)
                    UploadCount = UploadCount + 1
                    RecTbx1.Text = UploadCount
                    Ftbl25.MoveNext()
                Loop
            End If

        ElseIf RdDate10.Checked = True Then
            MTHeader()
            If Ftbl25.RecordCount <> 0 Then
                Ftbl25.MoveFirst()
                Do While Not Ftbl25.EOF
                    NewGSal(0) = Ftbl25("Nik").Value
                    NewGSal(1) = Ftbl25("Name").Value
                    LookTot = IIf(IsDBNull(Ftbl25("Salary10").Value), "", Ftbl25("Salary10").Value)
                    LookDate = RdDate10.Text
                    EmpLookMe()
                    MoTGrid1.Rows.Add(NewGSal(0), NewGSal(1), LookTot, LookDate, LookPay, LookAstek)
                    UploadCount = UploadCount + 1
                    RecTbx1.Text = UploadCount
                    Ftbl25.MoveNext()
                Loop
            End If

        ElseIf RdDate11.Checked = True Then
            MTHeader()
            If Ftbl25.RecordCount <> 0 Then
                Ftbl25.MoveFirst()
                Do While Not Ftbl25.EOF
                    NewGSal(0) = Ftbl25("Nik").Value
                    NewGSal(1) = Ftbl25("Name").Value
                    LookTot = IIf(IsDBNull(Ftbl25("Salary11").Value), "", Ftbl25("Salary11").Value)
                    LookDate = RdDate11.Text
                    EmpLookMe()
                    MoTGrid1.Rows.Add(NewGSal(0), NewGSal(1), LookTot, LookDate, LookPay, LookAstek)
                    UploadCount = UploadCount + 1
                    RecTbx1.Text = UploadCount
                    Ftbl25.MoveNext()
                Loop
            End If

        ElseIf RdDate12.Checked = True Then
            MTHeader()

            If Ftbl25.RecordCount <> 0 Then
                Ftbl25.MoveFirst()
                Do While Not Ftbl25.EOF
                    NewGSal(0) = Ftbl25("Nik").Value
                    NewGSal(1) = Ftbl25("Name").Value
                    LookTot = IIf(IsDBNull(Ftbl25("Salary12").Value), "", Ftbl25("Salary12").Value)
                    LookDate = RdDate12.Text
                    EmpLookMe()
                    MoTGrid1.Rows.Add(NewGSal(0), NewGSal(1), LookTot, LookDate, LookPay, LookAstek)
                    UploadCount = UploadCount + 1
                    RecTbx1.Text = UploadCount
                    Ftbl25.MoveNext()
                Loop
            End If

        ElseIf RdDate13.Checked = True Then
            MTHeader()

            If Ftbl25.RecordCount <> 0 Then
                Ftbl25.MoveFirst()
                Do While Not Ftbl25.EOF
                    NewGSal(0) = Ftbl25("Nik").Value
                    NewGSal(1) = Ftbl25("Name").Value
                    LookTot = IIf(IsDBNull(Ftbl25("Salary13").Value), "", Ftbl25("Salary13").Value)
                    LookDate = RdDate13.Text
                    EmpLookMe()
                    MoTGrid1.Rows.Add(NewGSal(0), NewGSal(1), LookTot, LookDate, LookPay, LookAstek)
                    UploadCount = UploadCount + 1
                    RecTbx1.Text = UploadCount
                    Ftbl25.MoveNext()
                Loop
            End If

        ElseIf RdDate14.Checked = True Then

            MTHeader()
            If Ftbl25.RecordCount <> 0 Then
                Ftbl25.MoveFirst()
                Do While Not Ftbl25.EOF
                    NewGSal(0) = Ftbl25("Nik").Value
                    NewGSal(1) = Ftbl25("Name").Value
                    LookTot = IIf(IsDBNull(Ftbl25("Salary14").Value), "", Ftbl25("Salary14").Value)
                    LookDate = RdDate14.Text
                    EmpLookMe()
                    MoTGrid1.Rows.Add(NewGSal(0), NewGSal(1), LookTot, LookDate, LookPay, LookAstek)
                    UploadCount = UploadCount + 1
                    RecTbx1.Text = UploadCount
                    Ftbl25.MoveNext()
                Loop
            End If

        ElseIf RdDate15.Checked = True Then
            MTHeader()

            If Ftbl25.RecordCount <> 0 Then
                Ftbl25.MoveFirst()
                Do While Not Ftbl25.EOF
                    NewGSal(0) = Ftbl25("Nik").Value
                    NewGSal(1) = Ftbl25("Name").Value
                    LookTot = IIf(IsDBNull(Ftbl25("Salary15").Value), "", Ftbl25("Salary15").Value)
                    LookDate = RdDate15.Text
                    EmpLookMe()
                    MoTGrid1.Rows.Add(NewGSal(0), NewGSal(1), LookTot, LookDate, LookPay, LookAstek)
                    UploadCount = UploadCount + 1
                    RecTbx1.Text = UploadCount
                    Ftbl25.MoveNext()
                Loop
            End If

        ElseIf RdDate16.Checked = True Then

            MTHeader()
            If Ftbl25.RecordCount <> 0 Then
                Ftbl25.MoveFirst()
                Do While Not Ftbl25.EOF
                    NewGSal(0) = Ftbl25("Nik").Value
                    NewGSal(1) = Ftbl25("Name").Value
                    LookTot = IIf(IsDBNull(Ftbl25("Salary16").Value), "", Ftbl25("Salary16").Value)
                    LookDate = RdDate16.Text
                    EmpLookMe()
                    MoTGrid1.Rows.Add(NewGSal(0), NewGSal(1), LookTot, LookDate, LookPay, LookAstek)
                    UploadCount = UploadCount + 1
                    RecTbx1.Text = UploadCount
                    Ftbl25.MoveNext()
                Loop
            End If

        Else
            MsgBox("Please Select the date for data inquiry")

        End If





    End Sub

    Private Sub PerCmb1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MTCmb1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub PerCmb2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MTCmb2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub PerCmb2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MTCmb2.SelectedIndexChanged
        FillRadio()
    End Sub

    Private Sub SynchBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SynchBtn1.Click
        FillSalary()
    End Sub

    Private Sub SynchBtn3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SynchBtn3.Click
        MoTGrid1.Rows.Clear()
        MoTGrid1.Columns.Clear()
        RecTbx1.Text = ""
        UploadCount = 0
        SaveLoadCount = 0
    End Sub

    Sub SynchSaver()

        For i = 0 To MoTGrid1.Rows.Count - 1

            MotNik = MoTGrid1(0, i).Value
            MotName = MoTGrid1(1, i).Value
            MotSalary = MoTGrid1(2, i).Value
            MotDate = MoTGrid1(3, i).Value
            MotPay = MoTGrid1(4, i).Value
            MotAst = MoTGrid1(5, i).Value

            SQL = ""
            SQL = SQL & "Select * From SalarySync1_Table "
            SQL = SQL & "Where Nik = ('" & MotNik & "') "
            SQL = SQL & "And Periode = ('" & MTCmb2.Text & "') "
            SQL = SQL & "And PeriodeRange = ('" & MTCmb1.Text & "') "
            OpenTbl(CBb, Ctbl1, SQL)


            If Not Ctbl1.RecordCount <> 0 Then
                Ctbl1.AddNew()
            End If

            Ctbl1("Nik").Value = MotNik
            Ctbl1("Name").Value = MotName
            Ctbl1("Periode").Value = MTCmb2.Text
            Ctbl1("PeriodeRange").Value = MTCmb1.Text
            Ctbl1("Pay").Value = MotPay
            Ctbl1("AstekVal").Value = MotAst


            If RdDate1.Checked = True Then

                Ctbl1("Salary1").Value = MotSalary
                Ctbl1("Date1").Value = MotDate

                Ctbl1.Update()
                MoTGrid1(6, i).Value = "Has Been Saved"
                SaveLoadCount = SaveLoadCount + 1
                RecTbx2.Text = SaveLoadCount
                Me.Refresh()

            ElseIf RdDate2.Checked = True Then

                Ctbl1("Salary2").Value = MotSalary
                Ctbl1("Date2").Value = MotDate

                Ctbl1.Update()
                MoTGrid1(6, i).Value = "Has Been Saved"
                SaveLoadCount = SaveLoadCount + 1
                RecTbx2.Text = SaveLoadCount
                Me.Refresh()

            ElseIf RdDate3.Checked = True Then

                Ctbl1("Salary3").Value = MotSalary
                Ctbl1("Date3").Value = MotDate
                Ctbl1.Update()
                MoTGrid1(6, i).Value = "Has Been Saved"
                SaveLoadCount = SaveLoadCount + 1
                RecTbx2.Text = SaveLoadCount
                Me.Refresh()

            ElseIf RdDate4.Checked = True Then

                Ctbl1("Salary4").Value = MotSalary
                Ctbl1("Date4").Value = MotDate

                Ctbl1.Update()
                MoTGrid1(6, i).Value = "Has Been Saved"
                SaveLoadCount = SaveLoadCount + 1
                RecTbx2.Text = SaveLoadCount
                Me.Refresh()

            ElseIf RdDate5.Checked = True Then

                Ctbl1("Salary5").Value = MotSalary
                Ctbl1("Date5").Value = MotDate

                Ctbl1.Update()
                MoTGrid1(6, i).Value = "Has Been Saved"
                SaveLoadCount = SaveLoadCount + 1
                RecTbx2.Text = SaveLoadCount
                Me.Refresh()

            ElseIf RdDate6.Checked = True Then

                Ctbl1("Salary6").Value = MotSalary
                Ctbl1("Date6").Value = MotDate

                Ctbl1.Update()
                MoTGrid1(6, i).Value = "Has Been Saved"
                SaveLoadCount = SaveLoadCount + 1
                RecTbx2.Text = SaveLoadCount
                Me.Refresh()

            ElseIf RdDate7.Checked = True Then

                Ctbl1("Salary7").Value = MotSalary
                Ctbl1("Date7").Value = MotDate

                Ctbl1.Update()
                MoTGrid1(6, i).Value = "Has Been Saved"
                SaveLoadCount = SaveLoadCount + 1
                RecTbx2.Text = SaveLoadCount
                Me.Refresh()

            ElseIf RdDate8.Checked = True Then

                Ctbl1("Salary8").Value = MotSalary
                Ctbl1("Date8").Value = MotDate

                Ctbl1.Update()
                MoTGrid1(6, i).Value = "Has Been Saved"
                SaveLoadCount = SaveLoadCount + 1
                RecTbx2.Text = SaveLoadCount
                Me.Refresh()

            ElseIf RdDate9.Checked = True Then

                Ctbl1("Salary9").Value = MotSalary
                Ctbl1("Date9").Value = MotDate

                Ctbl1.Update()
                MoTGrid1(6, i).Value = "Has Been Saved"
                SaveLoadCount = SaveLoadCount + 1
                RecTbx2.Text = SaveLoadCount
                Me.Refresh()

            ElseIf RdDate10.Checked = True Then

                Ctbl1("Salary10").Value = MotSalary
                Ctbl1("Date10").Value = MotDate
                Ctbl1.Update()
                MoTGrid1(6, i).Value = "Has Been Saved"
                SaveLoadCount = SaveLoadCount + 1
                RecTbx2.Text = SaveLoadCount
                Me.Refresh()

            ElseIf RdDate11.Checked = True Then

                Ctbl1("Salary11").Value = MotSalary
                Ctbl1("Date11").Value = MotDate

                Ctbl1.Update()
                MoTGrid1(6, i).Value = "Has Been Saved"
                SaveLoadCount = SaveLoadCount + 1
                RecTbx2.Text = SaveLoadCount
                Me.Refresh()

            ElseIf RdDate12.Checked = True Then

                Ctbl1("Salary12").Value = MotSalary
                Ctbl1("Date12").Value = MotDate

                Ctbl1.Update()
                MoTGrid1(6, i).Value = "Has Been Saved"
                SaveLoadCount = SaveLoadCount + 1
                RecTbx2.Text = SaveLoadCount
                Me.Refresh()

            ElseIf RdDate13.Checked = True Then

                Ctbl1("Salary13").Value = MotSalary
                Ctbl1("Date13").Value = MotDate

                Ctbl1.Update()
                MoTGrid1(6, i).Value = "Has Been Saved"
                SaveLoadCount = SaveLoadCount + 1
                RecTbx2.Text = SaveLoadCount
                Me.Refresh()


            ElseIf RdDate14.Checked = True Then

                Ctbl1("Salary14").Value = MotSalary
                Ctbl1("Date14").Value = MotDate

                Ctbl1.Update()
                MoTGrid1(6, i).Value = "Has Been Saved"
                SaveLoadCount = SaveLoadCount + 1
                RecTbx2.Text = SaveLoadCount
                Me.Refresh()

            ElseIf RdDate15.Checked = True Then

                Ctbl1("Salary15").Value = MotSalary
                Ctbl1("Date15").Value = MotDate

                Ctbl1.Update()
                MoTGrid1(6, i).Value = "Has Been Saved"
                SaveLoadCount = SaveLoadCount + 1
                RecTbx2.Text = SaveLoadCount
                Me.Refresh()

            ElseIf RdDate16.Checked = True Then

                Ctbl1("Salary16").Value = MotSalary
                Ctbl1("Date16").Value = MotDate

                Ctbl1.Update()
                MoTGrid1(6, i).Value = "Has Been Saved"
                SaveLoadCount = SaveLoadCount + 1
                RecTbx2.Text = SaveLoadCount
                Me.Refresh()


            End If

        Next

        MsgBox("Done")

    End Sub

    Private Sub SynchBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SynchBtn2.Click
        SynchSaver()
    End Sub

    Private Sub MTCmb1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MTCmb1.SelectedIndexChanged

    End Sub
End Class