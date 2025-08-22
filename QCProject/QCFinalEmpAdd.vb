Public Class QCFinalEmpAdd

    Dim EmpFNum As String
    Dim EmpFDig As String


    ' Variable for each Record
    Dim StringCaller As String = Nothing
    Dim PID As String = Nothing
    Dim PNik As String = Nothing
    Dim PName As String = Nothing
    Dim PNoKTP As String = Nothing
    Dim PNoNPWP As String = Nothing
    Dim PNoKPJ As String = Nothing
    Dim PJKK As String = Nothing
    Dim PEstate As String = Nothing
    Dim PAlamat As String = Nothing
    Dim PLahir As String = Nothing
    Dim PAgama As String = Nothing
    Dim PTelNum As String = Nothing
    Dim PPendi As String = Nothing
    Dim PDept As String = Nothing
    Dim PJabatan As String = Nothing
    Dim PAstek As String = Nothing
    Dim PDStart As String = Nothing
    Dim PPay As String = Nothing
    Dim PNoRek As String = Nothing
    Dim PActive As String = Nothing
    Dim PNoATM As String = Nothing
    ' End for Variable of Each Record

    Dim Jamsoscode As String = Nothing

    Private Sub QCFinalEmpAdd_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LoadDB()
        LoadDB2()
        DisplayByLoad()
        LoadJamsostekCombo()
        AstekCmbItems()

    End Sub

    Sub FinalEmpGrid1Header()

        FinalEmpGrid1.Rows.Clear()
        FinalEmpGrid1.Columns.Clear()

        With FinalEmpGrid1

            .Columns.Add("col19", "ID")
            .Columns.Add("col0", "Nik")
            .Columns.Add("col1", "Name")
            .Columns.Add("col2", "NoKTP")
            .Columns.Add("col3", "NoNPWP")
            .Columns.Add("col4", "NoKPJ")
            .Columns.Add("col5", "Alamat")
            .Columns.Add("col6", "JKKJLM")
            .Columns.Add("col7", "Estate")
            .Columns.Add("col8", "Tempat Lahir")
            .Columns.Add("col9", "Agama")
            .Columns.Add("col10", "TelNum")
            .Columns.Add("col11", "Pendidikan")
            .Columns.Add("col12", "Dept")
            .Columns.Add("col13", "Jabatan")
            .Columns.Add("col14", "Astek")
            .Columns.Add("col15", "Date Start")
            .Columns.Add("col16", "PayAs")
            .Columns.Add("col17", "NoRek")
            .Columns.Add("col18", "No ATM")
            .Columns.Add("col119", "Active")
            .Columns(0).Width = 130
            .Columns(1).Width = 130

        End With

    End Sub

    Sub FinalGenerateNum()

        SQL = ""
        SQL = SQL & "Select `ID_Number` From 02_Name_Table "
        SQL = SQL & "Order by ID_Number Desc"
        OpenTbl(ADb, DbTbl5, SQL)
        If DbTbl5.RecordCount > 0 Then
            EmpFDig = DbTbl5("ID_Number").Value
            EmpFNum = Format(EmpFDig + 1, "00000000")
        Else
            EmpFNum = "00000001"
        End If

    End Sub

    Sub AddEmpMode()

        If FTbx02.Enabled = False Then
            FTbx01.Enabled = True
            FTbx02.Enabled = True
            FTbx03.Enabled = True
            FTbx04.Enabled = True
            FTbx05.Enabled = True
            FTbx06.Enabled = True
            FTbx07.Enabled = True
            FTbx08.Enabled = True
            FTbx09.Enabled = True
            FTbx10.Enabled = True
            FTbx11.Enabled = True
            FTbx12.Enabled = True
            FTbx13.Enabled = True
            FCmb01.Enabled = True
            FCmb02.Enabled = True
            FCmb03.Enabled = True
            FCmb04.Enabled = True
            FDP01.Enabled = True
            EDRb1.Enabled = True
            EDRb2.Enabled = True

        Else

            FTbx01.Enabled = False
            FTbx02.Enabled = False
            FTbx03.Enabled = False
            FTbx04.Enabled = False
            FTbx05.Enabled = False
            FTbx06.Enabled = False
            FTbx07.Enabled = False
            FTbx08.Enabled = False
            FTbx09.Enabled = False
            FTbx10.Enabled = False
            FTbx11.Enabled = False
            FTbx12.Enabled = False
            FTbx13.Enabled = False
            FCmb01.Enabled = False
            FCmb02.Enabled = False
            FCmb03.Enabled = False
            FCmb04.Enabled = False
            FDP01.Enabled = False
            EDRb1.Enabled = False
            EDRb2.Enabled = False

            FTbx01.Clear()
            FTbx02.Clear()
            FTbx03.Clear()
            FTbx04.Clear()
            FTbx05.Clear()
            FTbx06.Clear()
            FTbx07.Clear()
            FTbx08.Clear()
            FTbx09.Clear()
            FTbx10.Clear()
            FTbx11.Clear()
            FTbx12.Clear()
            FTbx13.Clear()
            FCmb01.Text = ""
            FCmb02.Text = ""
            FCmb03.Text = ""
            FCmb04.Text = ""

        End If

        If FBtn04.Enabled = False Then
            FBtn04.Enabled = True

        End If


    End Sub

    Sub NikValidator()

        SQL = ""
        SQL = SQL & "Select * From 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & FTbx01.Text & "') "
        SQL = SQL & "And ID_Number = ('" & EmpFNum & "')"
        OpenTbl(ADb, Atb2, SQL)

        If Atb2.RecordCount > 0 Then

            Dim ProceedNum As Integer = MessageBox.Show("This is an Edit Mode, Do you want to Overwrite your Employee's Data?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If ProceedNum = DialogResult.Yes Then
                SaveFEmp()
            End If

        Else

            SQL = ""
            SQL = SQL & "Select * From 02_Name_Table "
            SQL = SQL & "Where Nik = ('" & FTbx01.Text & "') "
            SQL = SQL & "And Not ID_Number = ('" & EmpFNum & "')"
            OpenTbl(ADb, Atb2, SQL)

            If Atb2.RecordCount > 0 Then
                MessageBox.Show("Nik is already taken", "Warning", MessageBoxButtons.OK)

            Else

                SaveFEmp()

            End If

        End If

    End Sub

    Sub OutEmp()

        SQL = ""
        SQL = SQL & "Select * From 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & FTbx01.Text & "') "
        SQL = SQL & "And ID_Number = ('" & EmpFNum & "') "
        OpenTbl(ADb, Atb1, SQL)

        If Atb1.RecordCount > 0 Then

            Atb1("Active").Value = "No"

        End If

        Atb1.Update()

    End Sub

    Sub DeleteEmp()

        SQL = ""
        SQL = SQL & "Select * From 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & FTbx01.Text & "') "
        SQL = SQL & "And ID_Number = ('" & EmpFNum & "') "
        OpenTbl(ADb, Atb1, SQL)

        If Atb1.RecordCount <> 0 Then

            Atb1.Delete()

        End If

    End Sub

    Sub SaveFEmp()

        SQL = ""
        SQL = SQL & "Select * From 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & FTbx01.Text & "') "
        OpenTbl(ADb, Atb1, SQL)

        If Not Atb1.RecordCount <> 0 Then
            Atb1.AddNew()
        End If
        StringCaller = FTbx02.Text.Replace("'", "?")
        Atb1("ID_Number").Value = EmpFNum
        Atb1("Nik").Value = FTbx01.Text
        Atb1("Name").Value = StringCaller
        Atb1("DateStart").Value = FDP01.Text
        Atb1("Active").Value = "Yes"
        Atb1("Jamsostek").Value = FCmb03.Text

        If EDRb1.Checked = True Then

            Atb1("Pay").Value = "BTN"

        ElseIf EDRb2.Checked = True Then

            Atb1("Pay").Value = "CASH"

        End If

        Atb1("Bank_Ctrl").Value = FTbx12.Text
        Atb1("Jamsostek").Value = FCmb03.Text
        Atb1("NPWP").Value = FTbx08.Text
        Atb1("NoRek").Value = FTbx10.Text
        Atb1("NKTP").Value = FTbx07.Text
        Atb1("NoKPJ").Value = FTbx09.Text
        Atb1("Lahir").Value = FTbx03.Text
        Atb1("JabData").Value = FCmb02.Text
        Atb1("Estate").Value = FTbx13.Text
        Atb1("Agama").Value = FCmb04.Text
        Atb1("Alamat").Value = FTbx04.Text
        Atb1("TelNum").Value = FTbx05.Text
        Atb1("Pendi").Value = FTbx06.Text
        Atb1("Dept").Value = FCmb01.Text

        Atb1.Update()

        AddEmpMode()
        DisplayByLoad()

    End Sub

    Private Sub FBtn01_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FBtn01.Click

        If FTbx01.Text = Nothing Then
            MessageBox.Show("Please Apply the Nik Number of the new employee", "Warning", MessageBoxButtons.OK)

        ElseIf FTbx02.Text = Nothing Then
            MessageBox.Show("Please Apply the Name of the new employee", "Warning", MessageBoxButtons.OK)

        Else

            NikValidator()

        End If

    End Sub

#Region "Load My Item"

    Sub Variablepart()

        PID = IIf(IsDBNull(Atb2("ID_Number").Value), "", Atb2("ID_Number").Value) '1
        PNik = IIf(IsDBNull(Atb2("Nik").Value), "", Atb2("Nik").Value) '2
        PName = IIf(IsDBNull(Atb2("Name").Value), "", Atb2("Name").Value) '3
        PNoKTP = IIf(IsDBNull(Atb2("NKTP").Value), "", Atb2("NKTP").Value) '4
        PNoNPWP = IIf(IsDBNull(Atb2("NPWP").Value), "", Atb2("NPWP").Value) '5
        PNoKPJ = IIf(IsDBNull(Atb2("NoKPJ").Value), "", Atb2("NoKPJ").Value) '6
        PAlamat = IIf(IsDBNull(Atb2("Alamat").Value), "", Atb2("Alamat").Value) '7
        PJKK = IIf(IsDBNull(Atb2("JKKJKM").Value), "", Atb2("JKKJKM").Value) '8
        PEstate = IIf(IsDBNull(Atb2("Estate").Value), "", Atb2("Estate").Value) '9
        PLahir = IIf(IsDBNull(Atb2("Lahir").Value), "", Atb2("Lahir").Value) '10
        PAgama = IIf(IsDBNull(Atb2("Agama").Value), "", Atb2("Agama").Value) '11
        PTelNum = IIf(IsDBNull(Atb2("TelNum").Value), "", Atb2("TelNum").Value) '12
        PPendi = IIf(IsDBNull(Atb2("Pendi").Value), "", Atb2("Pendi").Value) '13
        PDept = IIf(IsDBNull(Atb2("Dept").Value), "", Atb2("Dept").Value) '14
        PJabatan = IIf(IsDBNull(Atb2("JabData").Value), "", Atb2("JabData").Value) ' 15
        PAstek = IIf(IsDBNull(Atb2("Jamsostek").Value), "", Atb2("Jamsostek").Value) ' 16
        PDStart = IIf(IsDBNull(Atb2("DateStart").Value), "", Atb2("DateStart").Value) ' 17
        PPay = IIf(IsDBNull(Atb2("Pay").Value), "", Atb2("Pay").Value) ' 18
        PNoRek = IIf(IsDBNull(Atb2("NoRek").Value), "", Atb2("NoRek").Value) ' 19 
        PNoATM = IIf(IsDBNull(Atb2("Bank_Ctrl").Value), "", Atb2("Bank_Ctrl").Value) ' 20
        PActive = IIf(IsDBNull(Atb2("Active").Value), "", Atb2("Active").Value) ' 21

    End Sub

    Sub SearchbyText()

        FinalEmpGrid1Header()

        If FinalEmpCmb01.Text = "Nik" Then

            SQL = ""
            SQL = SQL & "Select * From 02_Name_Table "
            SQL = SQL & "Where Nik like '" & FinalEmpTbx01.Text & "%' "
            If Not FinalEmpCmb02.Text = "Both" Then
                SQL = SQL & "And Active = ('" & FinalEmpCmb02.Text & " ') "
            End If

            OpenTbl(ADb, Atb2, SQL)
            If Atb2.RecordCount <> 0 Then

                Atb2.MoveFirst()
                Do While Not Atb2.EOF

                    Variablepart()
                    PName = PName.Replace("?", "'")
                    FinalEmpGrid1.Rows.Add(PID, PNik, PName, PNoKTP, PNoNPWP, PNoKPJ, PAlamat, PJKK, PEstate, PLahir, PAgama, PTelNum, PPendi, PDept, PJabatan, PAstek, PDStart, PPay, PNoRek, PNoATM, PActive)
                    Atb2.MoveNext()

                Loop

            End If

        ElseIf FinalEmpCmb01.Text = "Name" Then
            StringCaller = FinalEmpTbx01.Text.Replace("'", "?")
            SQL = ""
            SQL = SQL & "Select * From 02_Name_Table "
            SQL = SQL & "Where Name like '" & StringCaller & "%' "
            If Not FinalEmpCmb02.Text = "Both" Then
                SQL = SQL & "And Active = ('" & FinalEmpCmb02.Text & " ') "
            End If
            OpenTbl(ADb, Atb2, SQL)
            If Atb2.RecordCount <> 0 Then

                Atb2.MoveFirst()
                Do While Not Atb2.EOF
                    Variablepart()
                    PName = PName.Replace("?", "'")
                    FinalEmpGrid1.Rows.Add(PID, PNik, PName, PNoKTP, PNoNPWP, PNoKPJ, PAlamat, PJKK, PEstate, PLahir, PAgama, PTelNum, PPendi, PDept, PJabatan, PAstek, PDStart, PPay, PNoRek, PNoATM, PActive)
                    Atb2.MoveNext()

                Loop

            End If

        End If

    End Sub

    Sub DisplayByLoad()

        FinalEmpGrid1Header()
        SQL = ""
        SQL = SQL & "Select * From 02_Name_Table "
        OpenTbl(ADb, Atb2, SQL)
        If Atb2.RecordCount <> 0 Then

            Atb2.MoveFirst()
            Do While Not Atb2.EOF

                Variablepart()
                PName = PName.Replace("?", "'")
                FinalEmpGrid1.Rows.Add(PID, PNik, PName, PNoKTP, PNoNPWP, PNoKPJ, PAlamat, PJKK, PEstate, PLahir, PAgama, PTelNum, PPendi, PDept, PJabatan, PAstek, PDStart, PPay, PNoRek, PNoATM, PActive)

                Atb2.MoveNext()

            Loop

        End If

    End Sub

#End Region

    Private Sub FinalEmpGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles FinalEmpGrid1.DoubleClick

        FTbx01.Enabled = True
        FTbx02.Enabled = True
        FTbx03.Enabled = True
        FTbx04.Enabled = True
        FTbx05.Enabled = True
        FTbx06.Enabled = True
        FTbx07.Enabled = True
        FTbx08.Enabled = True
        FTbx09.Enabled = True
        FTbx10.Enabled = True
        FTbx11.Enabled = True
        FTbx12.Enabled = True
        FTbx13.Enabled = True
        FCmb01.Enabled = True
        FCmb02.Enabled = True
        FCmb03.Enabled = True
        FCmb04.Enabled = True
        FDP01.Enabled = True
        EDRb1.Enabled = True
        EDRb2.Enabled = True

        With FinalEmpGrid1

            EmpFNum = .CurrentRow.Cells(0).Value.ToString ' ID
            FTbx01.Text = .CurrentRow.Cells(1).Value.ToString ' Nik
            FTbx02.Text = .CurrentRow.Cells(2).Value.ToString ' Name
            FTbx07.Text = .CurrentRow.Cells(3).Value.ToString ' NoKTP
            FTbx08.Text = .CurrentRow.Cells(4).Value.ToString ' NoNPWP
            FTbx09.Text = .CurrentRow.Cells(5).Value.ToString ' NoKP
            FTbx04.Text = .CurrentRow.Cells(6).Value.ToString ' Alamat
            FTbx11.Text = .CurrentRow.Cells(7).Value.ToString ' JKKJLM
            FTbx13.Text = .CurrentRow.Cells(8).Value.ToString ' Estate
            FTbx03.Text = .CurrentRow.Cells(9).Value.ToString ' Tempat Lahir
            FCmb04.Text = .CurrentRow.Cells(10).Value.ToString ' Agama
            FTbx05.Text = .CurrentRow.Cells(11).Value.ToString ' TelNum
            FTbx06.Text = .CurrentRow.Cells(12).Value.ToString ' Pendidikan
            FCmb01.Text = .CurrentRow.Cells(13).Value.ToString ' Dept
            FCmb02.Text = .CurrentRow.Cells(14).Value.ToString ' Jabatan
            FCmb03.Text = .CurrentRow.Cells(15).Value.ToString ' Astek
            FDP01.Text = .CurrentRow.Cells(16).Value.ToString ' Date Start
            FTbx12.Text = .CurrentRow.Cells(19).Value.ToString ' Date Start

            If .CurrentRow.Cells(17).Value.ToString = "BTN" Then

                EDRb1.Checked = True

            Else

                EDRb2.Checked = True

            End If

            FTbx10.Text = .CurrentRow.Cells(18).Value.ToString ' No Rek
            If .CurrentRow.Cells(20).Value.ToString = "YES" Or .CurrentRow.Cells(20).Value.ToString = "Yes" Then ' Active

                StatLabel01.Text = "Active"
                StatLabel01.ForeColor = Color.Green

            Else

                StatLabel01.Text = "Inactive"
                StatLabel01.ForeColor = Color.Red

            End If

        End With

        FBtn04.Enabled = False
        EmpTabCtrl.SelectTab(1)
        FTbx01.Enabled = False
    End Sub

    Private Sub FBtn04_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FBtn04.Click
        FinalGenerateNum()
        AddEmpMode()
    End Sub

    Sub LoadJamsostekCombo()

        SQL = ""
        SQL = SQL & "Select * From 08_Standard_Table "
        SQL = SQL & "Where Original = ('" & "Jamsostek" & "') "
        OpenTbl(ADb, Atbl20, SQL)
        If Atbl20.RecordCount > 0 Then
            Atbl20.MoveFirst()
            Do While Not Atbl20.EOF

                Jamsoscode = IIf(IsDBNull(Atbl20("Standard_Wage").Value), "", Atbl20("Standard_Wage").Value)
                Atbl20.MoveNext()

            Loop
        End If
    End Sub

    Sub AstekCmbItems()
        With FCmb03
            .Items.Add(Jamsoscode)
            .Items.Add("0.00")
        End With

    End Sub

#Region "Keypress Border"

    Private Sub FCmb01_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles FCmb01.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub FCmb02_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles FCmb02.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub FCmb03_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles FCmb03.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub FCmb04_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles FCmb04.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

#End Region

    Private Sub FBtn02_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FBtn02.Click

        Dim OutResult As Integer = MessageBox.Show("Are you sure to inactive this person?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If OutResult = DialogResult.Yes Then
            OutEmp()
        End If
        AddEmpMode()
        DisplayByLoad()

    End Sub

    Private Sub FBtn03_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FBtn03.Click

        Dim DeleResult As Integer = MessageBox.Show("Are you sure to delete this person?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If DeleResult = DialogResult.Yes Then
            DeleteEmp()
        End If

        AddEmpMode()
        DisplayByLoad()

    End Sub

    Private Sub FBtn05_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FBtn05.Click

        Dim WorkResult As Integer = MessageBox.Show("Are you sure to delete this person?", "Information", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If WorkResult = DialogResult.Yes Then
            DeepDelete()
        End If

        AddEmpMode()
        DisplayByLoad()

    End Sub

    Sub DeepDelete()

        ' Per Job Delete
        ' For Conveyour

        SQL = ""
        SQL = SQL & "Select * From 03_Conveyour_Table "
        SQL = SQL & "Where Nik = ('" & FTbx01.Text & "') "
        OpenTbl(ADb, Atb1, SQL)

        If Atb1.RecordCount <> 0 Then
            Atb1.MoveFirst()
            Do While Not Atb1.EOF
                Atb1.Delete()
                Atb1.MoveNext()
            Loop

        End If


        ' For Mutu II

        SQL = ""
        SQL = SQL & "Select * From 04_MutuII_Table "
        SQL = SQL & "Where Nik = ('" & FTbx01.Text & "') "
        OpenTbl(ADb, Atb1, SQL)

        If Atb1.RecordCount <> 0 Then
            Atb1.MoveFirst()
            Do While Not Atb1.EOF
                Atb1.Delete()
                Atb1.MoveNext()
            Loop

        End If
        ' For  Packing

        SQL = ""
        SQL = SQL & "Select * From 05_Packing_Table "
        SQL = SQL & "Where Nik = ('" & FTbx01.Text & "') "
        OpenTbl(ADb, Atb1, SQL)

        If Atb1.RecordCount <> 0 Then
            Atb1.MoveFirst()
            Do While Not Atb1.EOF
                Atb1.Delete()
                Atb1.MoveNext()
            Loop

        End If

        ' for Wallet

        SQL = ""
        SQL = SQL & "Select * From 06_Wallet_Table "
        SQL = SQL & "Where Nik = ('" & FTbx01.Text & "') "
        OpenTbl(ADb, Atb1, SQL)

        If Atb1.RecordCount <> 0 Then
            Atb1.MoveFirst()
            Do While Not Atb1.EOF
                Atb1.Delete()
                Atb1.MoveNext()
            Loop

        End If

        ' for Sortasi

        SQL = ""
        SQL = SQL & "Select * From 21_NewMiscellaneous_Table "
        SQL = SQL & "Where Nik = ('" & FTbx01.Text & "') "
        OpenTbl(ADb, Atb1, SQL)

        If Atb1.RecordCount <> 0 Then
            Atb1.MoveFirst()
            Do While Not Atb1.EOF
                Atb1.Delete()
                Atb1.MoveNext()
            Loop

        End If

        ' For Miscellaneous

        SQL = ""
        SQL = SQL & "Select * From 19_Miscellaneous_Table "
        SQL = SQL & "Where Nik = ('" & FTbx01.Text & "') "
        OpenTbl(ADb, Atb1, SQL)
        If Atb1.RecordCount <> 0 Then
            Atb1.MoveFirst()
            Do While Not Atb1.EOF
                Atb1.Delete()
                Atb1.MoveNext()
            Loop

        End If

        ' Per Salary Delete

        ' Conveyour


        SQL = ""
        SQL = SQL & "Select * From 13_Conveyour_Salary "
        SQL = SQL & "Where Nik = ('" & FTbx01.Text & "') "
        OpenTbl(ADb, Atb2, SQL)

        If Atb2.RecordCount <> 0 Then
            Atb2.MoveFirst()
            Do While Not Atb2.EOF
                Atb2.Delete()
                Atb2.MoveNext()
            Loop

        End If

        ' Mutu II

        SQL = ""
        SQL = SQL & "Select * From 14_MutuII_Salary "
        SQL = SQL & "Where Nik = ('" & FTbx01.Text & "') "
        OpenTbl(ADb, Atb2, SQL)

        If Atb2.RecordCount <> 0 Then
            Atb2.MoveFirst()
            Do While Not Atb2.EOF
                Atb2.Delete()
                Atb2.MoveNext()
            Loop

        End If

        ' Packing

        SQL = ""
        SQL = SQL & "Select * From 16_Packing_Salary "
        SQL = SQL & "Where Nik = ('" & FTbx01.Text & "') "
        OpenTbl(ADb, Atb2, SQL)

        If Atb2.RecordCount <> 0 Then
            Atb2.MoveFirst()
            Do While Not Atb2.EOF
                Atb2.Delete()
                Atb2.MoveNext()
            Loop

        End If

        ' Wallet

        SQL = ""
        SQL = SQL & "Select * From 15_Wallet_Salary "
        SQL = SQL & "Where Nik = ('" & FTbx01.Text & "') "
        OpenTbl(ADb, Atb2, SQL)

        If Atb2.RecordCount <> 0 Then
            Atb2.MoveFirst()
            Do While Not Atb2.EOF
                Atb2.Delete()
                Atb2.MoveNext()
            Loop

        End If


        ' Sortasi and Miscellaneous


        SQL = ""
        SQL = SQL & "Select * From 20_Miscellaneous_Salary "
        SQL = SQL & "Where Nik = ('" & FTbx01.Text & "') "
        OpenTbl(ADb, Atb2, SQL)


        If Atb2.RecordCount <> 0 Then
            Atb2.MoveFirst()
            Do While Not Atb2.EOF
                Atb2.Delete()
                Atb2.MoveNext()
            Loop

        End If

        ' For Periode Delete Semua 

        SQL = ""
        SQL = SQL & "Select * from SalarySync1_Table "
        SQL = SQL & "Where Nik = ('" & FTbx01.Text & "') "
        OpenTbl(CBb, Ctbl1, SQL)

        If Ctbl1.RecordCount <> 0 Then
            Ctbl1.MoveFirst()
            Do While Not Ctbl1.EOF
                Ctbl1.Delete()
                Ctbl1.MoveNext()
            Loop

        End If

        MessageBox.Show("Deletion is Done", "Success")

    End Sub

    Private Sub FinalEmpTbx01_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles FinalEmpTbx01.KeyPress
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            SearchbyText()
            e.Handled = True
        End If
    End Sub

 
 
    Private Sub FinalEmpTbx01_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FinalEmpTbx01.TextChanged

    End Sub
End Class