Public Class EmpDatusBlock

    Dim EmployeeNum As String
    Dim EmployeeDig As String
    Dim OldEmployeeNum As String
    Dim OldEmployeeDig As String
    Public LogNumber As String
    Public LogNumber2 As String
    Dim EmpActive As Boolean
    Dim EmpPay As String
    Dim AstActive As Boolean

    Private Sub EmpDatusBlock_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        MainMenu.Refresh()
        AntiDupActuator = 0
    End Sub

    Private Sub EmpDatusBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadDB()
        LoadDBPPh21()
        EmployeeNum = LogNumber
        OldEmployeeNum = LogNumber2
    End Sub

    Sub GenEmployeeCode() ' Auto Generating Number

        SQL = ""
        SQL = SQL & "Select * From Emp_Table001 "
        SQL = SQL & "Order by PD_Id Desc"
        OpenTbl(PPhDB, PPhTb1, SQL)
        If PPhTb1.RecordCount > 0 Then
            EmployeeDig = PPhTb1("PD_Id").Value
            EmployeeNum = Format(EmployeeDig + 1, "00000000")
        Else
            EmployeeNum = "00000001"
        End If

    End Sub

    Sub OldGenEmployeeCode() ' Auto Generating Number

        SQL = ""
        SQL = SQL & "Select * From 02_Name_Table "
        SQL = SQL & "Order by ID_Number Desc"
        OpenTbl(ADb, DbTbl5, SQL)
        If DbTbl5.RecordCount > 0 Then
            OldEmployeeDig = DbTbl5("ID_Number").Value
            OldEmployeeNum = Format(OldEmployeeDig + 1, "00000000")
        Else
            OldEmployeeNum = "00000001"
        End If

    End Sub

    Sub EmpNewSave()

        SQL = ""
        SQL = SQL & "Select * From Emp_Table001 "
        SQL = SQL & "Where Nik = ('" & EDTbx1.Text & "') "
        SQL = SQL & "And PD_Id = ('" & LogNumber & "')"
        OpenTbl(PPhDB, PPhTb1, SQL)

        If Not PPhTb1.RecordCount <> 0 Then
            PPhTb1.AddNew()
        End If

        PPhTb1("PD_Id").Value = EmployeeNum
        PPhTb1("Nik").Value = EDTbx1.Text
        PPhTb1("Nama").Value = EDTbx2.Text
        PPhTb1("Tempat_Lahir").Value = EDTbx3.Text
        PPhTb1("Alamat").Value = EDTbx4.Text
        PPhTb1("TelNum").Value = EDTbx5.Text
        PPhTb1("Pendidikan").Value = EDTbx6.Text
        PPhTb1("NoKTP").Value = EDTbx7.Text
        PPhTb1("NoNPWP").Value = EDTbx8.Text
        PPhTb1("NoKPJ").Value = EDTbx9.Text
        PPhTb1("NoRek").Value = EDTbx11.Text
        PPhTb1("GajiMin").Value = EDTbx12.Text
        PPhTb1("Astek").Value = EDTbx13.Text
        PPhTb1("JKKJMM").Value = EDTbx14.Text

        PPhTb1("Dept").Value = EDCmb1.Text
        PPhTb1("Jabatan").Value = EDCmb2.Text
        PPhTb1("HariLim").Value = EDCmb3.Text
        PPhTb1("MasukKer").Value = EDDate1.Text
        PPhTb1("EfKer").Value = EDDate2.Text

        If EDRb1.Checked = True Then
            PPhTb1("PayAs").Value = "BTN"
        ElseIf EDRb2.Checked = True Then
            PPhTb1("PayAs").Value = "Cash"
        End If

        If EDCB1.Checked = True Then
            PPhTb1("Active").Value = False
            PPhTb1("TglKel").Value = EDDate3.Text
        ElseIf EDCB1.Checked = False Then
            PPhTb1("Active").Value = True
        End If

        If EDCB2.Checked = True Then
            PPhTb1("AstekOn").Value = True
        End If


        PPhTb1.Update()

        Me.Refresh()

    End Sub

    Sub LoadAstek()

        SQL = ""
        SQL = SQL & "Select * From 08_Standard_Table "
        SQL = SQL & "Where Original = ('" & "Jamsostek" & "') "
        OpenTbl(ADb, Atbl22, SQL)
        If Atbl22.RecordCount <> 0 Then

            EDTbx13.Text = Atbl22("Standard_Wage").Value
            Atbl22.MoveNext()

        End If
    End Sub

    Sub AntiDupNik()

        If AntiDupActuator = 0 Then

            SQL = ""
            SQL = SQL & "Select * from Emp_Table001 where "
            SQL = SQL & "Nik = ('" & EDTbx1.Text & "') "
            OpenTbl(PPhDB, PPhTb3, SQL)

            If PPhTb3.RecordCount > 0 Then
                MsgBox("NIK number is already taken", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

            Else
                EmpNewSave()
            End If

        ElseIf AntiDupActuator = 1 Then

            EmpNewSave()


        End If
    End Sub

    Sub AntiDupNik2()

        If AntiDupActuator = 0 Then

            SQL = ""
            SQL = SQL & "Select * from 02_Name_Table where "
            SQL = SQL & "Nik = ('" & EDTbx1.Text & "') "
            OpenTbl(ADb, DBTbl2, SQL)

            If DBTbl2.RecordCount > 0 Then
                MsgBox("NIK number is already taken", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")


            Else
                EmpSaveOld()
            End If

        ElseIf AntiDupActuator = 1 Then
            EmpSaveOld()

        End If
    End Sub

    Sub EmpSaveOld()


        SQL = ""
        SQL = SQL & "Select * From 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & EDTbx1.Text & "') "
        SQL = SQL & "and ID_Number = ('" & LogNumber2 & "')"
        OpenTbl(ADb, Atb1, SQL)

        If Not Atb1.RecordCount <> 0 Then
            Atb1.AddNew()
        End If

        Atb1("ID_Number").Value = OldEmployeeNum
        Atb1("Nik").Value = EDTbx1.Text
        Atb1("Name").Value = EDTbx2.Text
        Atb1("DateStart").Value = EDDate1.Text
        Atb1("NPWP").Value = EDTbx8.Text
        Atb1("NKTP").Value = EDTbx7.Text
        Atb1("JabData").Value = EDCmb1.Text

        If EDCB2.Checked = True Then

            Atb1("Jamsostek").Value = EDTbx13.Text

        ElseIf EDCB2.Checked = False Then

            Atb1("Jamsostek").Value = "0"

        End If


        If EDRb1.Checked = True Then
            Atb1("Pay").Value = "BTN"

        ElseIf EDRb2.Checked = True Then
            Atb1("Pay").Value = "CASH"
        End If

        If EDCB1.Checked = True Then

            Atb1("Active").Value = "No"
        ElseIf EDCB1.Checked = False Then
            Atb1("Active").Value = "Yes"

        End If


        Atb1.Update()

    End Sub

    Sub ActiveTranslate()

        EmpActive = IIf(IsDBNull(PPhTb4("Active").Value), "", PPhTb4("Active").Value)
        EmpPay = IIf(IsDBNull(PPhTb4("PayAs").Value), "", PPhTb4("PayAs").Value)
        AstActive = IIf(IsDBNull(PPhTb4("AstekOn").Value), "", PPhTb4("AstekOn").Value)

        If EmpActive = True Then
            EDCB1.Checked = False
        ElseIf EmpActive = False Then
            EDCB1.Checked = True
        End If

        If AstActive = True Then
            EDCB2.Checked = True
        ElseIf AstActive = False Then
            EDCB2.Checked = False
        End If


        If EmpPay = "BTN" Then
            EDRb1.Checked = True
        ElseIf EmpPay = "CASH" Then
            EDRb2.Checked = True
        End If


    End Sub

    Sub EmpDatusLoad()

        SQL = ""
        SQL = SQL & "Select * from Emp_Table001 "
        SQL = SQL & "Where Nik = ('" & EDTbx1.Text & "') "
        SQL = SQL & "Order by Nik"
        OpenTbl(PPhDB, PPhTb4, SQL)

        If PPhTb4.RecordCount > 0 Then

            EDTbx1.Text = PPhTb4("Nik").Value
            EDTbx2.Text = PPhTb4("Nama").Value
            LogNumber = PPhTb4("PD_Id").Value
            EDTbx7.Text = IIf(IsDBNull(PPhTb4("NoKTP").Value), "", PPhTb4("NoKTP").Value)
            EDTbx8.Text = IIf(IsDBNull(PPhTb4("NoNPWP").Value), "", PPhTb4("NoNPWP").Value)
            EDTbx9.Text = IIf(IsDBNull(PPhTb4("NoKPJ").Value), "", PPhTb4("NoKPJ").Value)
            EDTbx14.Text = IIf(IsDBNull(PPhTb4("JKKJMM").Value), "", PPhTb4("JKKJMM").Value)
            EDTbx3.Text = IIf(IsDBNull(PPhTb4("Tempat_Lahir").Value), "", PPhTb4("Tempat_Lahir").Value)
            EDTbx4.Text = IIf(IsDBNull(PPhTb4("Agama").Value), "", PPhTb4("Agama").Value)
            EDTbx4.Text = IIf(IsDBNull(PPhTb4("Alamat").Value), "", PPhTb4("Alamat").Value)
            EDTbx5.Text = IIf(IsDBNull(PPhTb4("TelNum").Value), "", PPhTb4("TelNum").Value)
            EDTbx6.Text = IIf(IsDBNull(PPhTb4("Pendidikan").Value), "", PPhTb4("Pendidikan").Value)
            EDCmb1.Text = IIf(IsDBNull(PPhTb4("Dept").Value), "", PPhTb4("Dept").Value)
            EDCmb2.Text = IIf(IsDBNull(PPhTb4("Jabatan").Value), "", PPhTb4("Jabatan").Value)
            EDDate1.Text = IIf(IsDBNull(PPhTb4("MasukKer").Value), "", PPhTb4("MasukKer").Value)
            EDDate2.Text = IIf(IsDBNull(PPhTb4("EfKer").Value), "", PPhTb4("EfKer").Value)
            EDDate3.Text = IIf(IsDBNull(PPhTb4("TglKel").Value), "", PPhTb4("TglKel").Value)
            EDTbx11.Text = IIf(IsDBNull(PPhTb4("NoRek").Value), "", PPhTb4("NoRek").Value)
            EDCmb3.Text = IIf(IsDBNull(PPhTb4("HariLim").Value), "", PPhTb4("HariLim").Value)
            EDTbx12.Text = IIf(IsDBNull(PPhTb4("GajiMin").Value), "", PPhTb4("GajiMin").Value)
            EDTbx13.Text = IIf(IsDBNull(PPhTb4("Astek").Value), "", PPhTb4("Astek").Value)


            ActiveTranslate()
            Log1.Visible = True


        End If

    End Sub

    Sub EmpDatusLoad2()


        SQL = ""
        SQL = SQL & "Select * from 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & EDTbx1.Text & "') "
        SQL = SQL & "Order by Nik"
        OpenTbl(ADb, Atb3, SQL)

        If Atb3.RecordCount > 0 Then

            LogNumber2 = Atb3("ID_Number").Value
            Log2.Visible = True


        End If
    End Sub

    Private Sub EDCmb1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles EDCmb1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub EDCmb2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles EDCmb2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub EDCmb3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles EDCmb3.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub EDCB2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EDCB2.CheckedChanged

        If EDCB2.Checked = True Then
            LoadAstek()

        ElseIf EDCB2.Checked = False Then
            EDTbx13.Text = ""

        End If

    End Sub

    Private Sub EDTbx13_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles EDTbx13.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub EDCB1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EDCB1.CheckedChanged
        If EDCB1.Checked = True Then
            EDDate3.Enabled = True
        ElseIf EDCB1.Checked = False Then
            EDDate3.Enabled = False

        End If
    End Sub

    Private Sub EmpNewBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmpNewBtn1.Click
        GenEmployeeCode()
        OldGenEmployeeCode()
        Log1.Visible = True
        Log2.Visible = True

        Log1.Text = "Employee ID Number 2.5: " + EmployeeNum
        Log2.Text = "Employee ID Number: " + OldEmployeeNum
    End Sub

    Private Sub EmpNewBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmpNewBtn2.Click

        If EDTbx1.Text = "" Or EDTbx2.Text = "" Then
            MsgBox("Please Insert Nik or Name")

        ElseIf EmployeeNum = "" Then
            MsgBox("Click Add Personnel")
        Else
            AntiDupNik()

        End If
    End Sub

    Private Sub EmpNewBtn3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmpNewBtn3.Click

        If EDTbx1.Text = "" Or EDTbx2.Text = "" Then
            MsgBox("Please Insert Nik or Name")
        ElseIf OldEmployeeNum = "" Then
            MsgBox("Click Add Personnel")
        Else
            AntiDupNik2()
        End If

    End Sub

    Private Sub GlassButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GlassButton1.Click
        If EDTbx1.Text = "" Or EDTbx2.Text = "" Then
            MsgBox("Please Insert Nik or Name")
        ElseIf OldEmployeeNum = "" Then
            MsgBox("Click Add Personnel")
        Else
            AntiDupNik()
            AntiDupNik2()
        End If
    End Sub

    Private Sub EDTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles EDTbx1.KeyPress
        EDTbx1.CharacterCasing = CharacterCasing.Upper
    End Sub

End Class