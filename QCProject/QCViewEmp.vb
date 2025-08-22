Option Explicit On


Public Class Employee2Block

    Dim PDvalue As String
    Dim EmpActive As Boolean
    Dim EmpActiveTrans As String
    Dim LoadNik As String
    Dim LoadName As String
    Dim LoadKTP As String
    Dim LoadNPWP As String
    Dim LoadKPJ As String
    Dim LoadJKK As String
    Dim LoadEstate As String
    Dim LoadTempatLahir As String
    Dim LoadAgama As String
    Dim LoadTelNum As String
    Dim LoadPendidikan As String
    Dim LoadDept As String
    Dim LoadJabatan As String
    Dim LoadMasukKer As String
    Dim LoadEfikKer As String
    Dim LoadTangKel As String
    Dim LoadNoRek As String
    Dim LoadDOff As String
    Dim LoadGajiMin As String
    Dim LoadAstekV As String
    Dim LoadPayV As String
    Dim LoadAlamt As String


    Private Sub Employee2Block_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        MainMenu.Refresh()

    End Sub

    Private Sub Employee2Block_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadDB()
        LoadDBPPh21()
        AutoSearchEnabler()
        LoadEmployee2()


    End Sub


    Sub AutoSearchEnabler()

        If VECmb1.Checked = True Then

            SQL = ""
            SQL = SQL & "Select * From 08_Standard_Table "
            SQL = SQL & "Where Original = ('" & "iAutoSearch" & "') "
            OpenTbl(ADb, Atbl40, SQL)

            If Not Atbl40.RecordCount <> 0 Then
                Atbl40.AddNew()
            End If

            Atbl40("Original").Value = "iAutoSearch"
            Atbl40("Standard_Wage").Value = "Yes"


            Atbl40.Update()

        ElseIf VECmb1.Checked = False Then

            SQL = ""
            SQL = SQL & "Select * From 08_Standard_Table "
            SQL = SQL & "Where Original = ('" & "iAutoSearch" & "') "
            OpenTbl(ADb, Atbl40, SQL)

            If Not Atbl40.RecordCount <> 0 Then
                Atbl40.AddNew()
            End If

            Atbl40("Original").Value = "iAutoSearch"
            Atbl40("Standard_Wage").Value = "No"
            Atbl40.Update()


        End If

    End Sub

    Sub ActiveTranslate()

        If EmpActive = True Then

            EmpActiveTrans = "Yes"

        ElseIf EmpActive = False Then

            EmpActiveTrans = "No"


        End If

    End Sub

    Sub LoadEmployee(ByVal LookData As String)

        ViewGrid1.Rows.Clear()
        If ViewCmb1.Text = "Nik" Then

            SQL = ""
            SQL = SQL & "Select * from Emp_Table001 "
            SQL = SQL & "Where Nik like '" & LookData & "%' "
            SQL = SQL & "Order by Nik"
            OpenTbl(PPhDB, PPhTb4, SQL)

            If PPhTb4.RecordCount <> 0 Then

                PPhTb4.MoveFirst()
                Do While Not PPhTb4.EOF

                    DataLocator()
                    ViewGrid1.Rows.Add(LoadNik, LoadName, LoadPayV, LoadKTP, LoadNPWP, LoadKPJ, LoadJKK, LoadEstate, LoadTempatLahir, LoadAgama, LoadAlamt, LoadTelNum, LoadPendidikan, LoadDept, LoadJabatan, LoadMasukKer, LoadEfikKer, LoadTangKel, LoadNoRek, LoadDOff, LoadGajiMin, LoadAstekV, EmpActiveTrans)

                    PPhTb4.MoveNext()

                Loop
            End If
        Else

            If ViewCmb1.Text = "Name" Then

                SQL = ""
                SQL = SQL & "Select * from Emp_Table001 "
                SQL = SQL & "Where Nama like '" & LookData & "%' "
                SQL = SQL & "Order by Nik"
                OpenTbl(PPhDB, PPhTb4, SQL)


                If PPhTb4.RecordCount <> 0 Then

                    PPhTb4.MoveFirst()
                    Do While Not PPhTb4.EOF

                        DataLocator()
                        ViewGrid1.Rows.Add(LoadNik, LoadName, LoadPayV, LoadKTP, LoadNPWP, LoadKPJ, LoadJKK, LoadEstate, LoadTempatLahir, LoadAgama, LoadAlamt, LoadTelNum, LoadPendidikan, LoadDept, LoadJabatan, LoadMasukKer, LoadEfikKer, LoadTangKel, LoadNoRek, LoadDOff, LoadGajiMin, LoadAstekV, EmpActiveTrans)

                        PPhTb4.MoveNext()

                    Loop
                End If
            End If
        End If

    End Sub

    Sub LoadEmployee2()


        SQL = ""
        SQL = SQL & "Select * from Emp_Table001 "
        SQL = SQL & "Order by Nik"
        OpenTbl(PPhDB, PPhTb4, SQL)

        If PPhTb4.RecordCount <> 0 Then
            PPhTb4.MoveFirst()
            Do While Not PPhTb4.EOF


                DataLocator()
                ViewGrid1.Rows.Add(LoadNik, LoadName, LoadPayV, LoadKTP, LoadNPWP, LoadKPJ, LoadJKK, LoadEstate, LoadTempatLahir, LoadAgama, LoadAlamt, LoadTelNum, LoadPendidikan, LoadDept, LoadJabatan, LoadMasukKer, LoadEfikKer, LoadTangKel, LoadNoRek, LoadDOff, LoadGajiMin, LoadAstekV, EmpActiveTrans)

                PPhTb4.MoveNext()

            Loop
        End If


    End Sub

    Sub DataLocator()

        LoadNik = PPhTb4("Nik").Value
        LoadName = PPhTb4("Nama").Value
        LoadPayV = IIf(IsDBNull(PPhTb4("PayAs").Value), "", PPhTb4("PayAs").Value)
        LoadKTP = IIf(IsDBNull(PPhTb4("NoKTP").Value), "", PPhTb4("NoKTP").Value)
        LoadNPWP = IIf(IsDBNull(PPhTb4("NoNPWP").Value), "", PPhTb4("NoNPWP").Value)
        LoadKPJ = IIf(IsDBNull(PPhTb4("NoKPJ").Value), "", PPhTb4("NoKPJ").Value)
        LoadJKK = IIf(IsDBNull(PPhTb4("JKKJMM").Value), "", PPhTb4("JKKJMM").Value)
        LoadEstate = IIf(IsDBNull(PPhTb4("Estate").Value), "", PPhTb4("Estate").Value)
        LoadTempatLahir = IIf(IsDBNull(PPhTb4("Tempat_Lahir").Value), "", PPhTb4("Tempat_Lahir").Value)
        LoadAgama = IIf(IsDBNull(PPhTb4("Agama").Value), "", PPhTb4("Agama").Value)
        LoadAlamt = IIf(IsDBNull(PPhTb4("Alamat").Value), "", PPhTb4("Alamat").Value)
        LoadTelNum = IIf(IsDBNull(PPhTb4("TelNum").Value), "", PPhTb4("TelNum").Value)
        LoadPendidikan = IIf(IsDBNull(PPhTb4("Pendidikan").Value), "", PPhTb4("Pendidikan").Value)
        LoadDept = IIf(IsDBNull(PPhTb4("Dept").Value), "", PPhTb4("Dept").Value)
        LoadJabatan = IIf(IsDBNull(PPhTb4("Jabatan").Value), "", PPhTb4("Jabatan").Value)
        LoadMasukKer = IIf(IsDBNull(PPhTb4("MasukKer").Value), "", PPhTb4("MasukKer").Value)
        LoadEfikKer = IIf(IsDBNull(PPhTb4("EfKer").Value), "", PPhTb4("EfKer").Value)
        LoadTangKel = IIf(IsDBNull(PPhTb4("TglKel").Value), "", PPhTb4("TglKel").Value)
        LoadNoRek = IIf(IsDBNull(PPhTb4("NoRek").Value), "", PPhTb4("NoRek").Value)
        LoadDOff = IIf(IsDBNull(PPhTb4("HariLim").Value), "", PPhTb4("HariLim").Value)
        LoadGajiMin = IIf(IsDBNull(PPhTb4("GajiMin").Value), "", PPhTb4("GajiMin").Value)
        LoadAstekV = IIf(IsDBNull(PPhTb4("Astek").Value), "", PPhTb4("Astek").Value)

        EmpActive = IIf(IsDBNull(PPhTb4("Active").Value), "", PPhTb4("Active").Value)
        ActiveTranslate()

    End Sub

    Private Sub ViewCmb1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ViewCmb1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub ViewTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ViewTbx1.KeyPress
        LoadEmployee(ViewTbx1.Text)

        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            LoadEmployee(ViewTbx1.Text)
            e.Handled = True
        End If
    End Sub

    Private Sub ViewGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ViewGrid1.DoubleClick

        Dim NewMDIChild As New EmpDatusBlock()

        EmpDatusBlock.MdiParent = MainMenu

        EmpDatusBlock.EDTbx1.Text = ViewGrid1.CurrentRow.Cells(0).Value.ToString
        EmpDatusBlock.EmpDatusLoad()
        EmpDatusBlock.EmpDatusLoad2()

        EmpDatusBlock.Log1.Text = "Employee ID Number 2.5: " + EmpDatusBlock.LogNumber
        EmpDatusBlock.Log2.Text = "Employee ID Number: " + EmpDatusBlock.LogNumber2

        Me.Dispose()
        EmpDatusBlock.Show()
        AntiDupActuator = 1
        MainMenu.Refresh()
    End Sub


 
    Private Sub ViewGrid1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles ViewGrid1.CellContentClick

    End Sub

    Private Sub ViewTbx1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewTbx1.TextChanged

    End Sub
End Class