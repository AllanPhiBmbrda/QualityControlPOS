Public Class PermitBlock

    Dim DokDig As String
    Dim DokNum As String
    Dim InceRange As String
    Dim InceCount As String
    Dim InceTotCount As String




    Private Sub PermitBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadDB()
        LoadDB2()
        LoadDB3()
        IncentivesControlRange()
    End Sub

#Region "Incentive Dokter"

    Sub IncentivesControlSave()


        InceTotCount = InceCount + 1


        SQL = ""
        SQL = SQL & "Select * From 12_Incentives_Ctrl "
        SQL = SQL & "Where Nik = ('" & DokTbx1.Text & "') "
        SQL = SQL & "And MonthPeriodeRange = ('" & InceRange & "') "
        OpenTbl(ADb, Atbl31, SQL)
        If Not Atbl31.RecordCount <> 0 Then
            Atbl31.AddNew()
        End If

        Atbl31("MonthPeriodeRange").Value = InceRange
        Atbl31("Nik").Value = DokTbx1.Text
        Atbl31("Count").Value = InceTotCount

        Atbl31.Update()




    End Sub

    Sub IncentivesControlLoad()
        SQL = ""
        SQL = SQL & "Select * From 12_Incentives_Ctrl "
        SQL = SQL & "Where Nik = ('" & DokTbx1.Text & "') "
        SQL = SQL & "And MonthPeriodeRange = ('" & InceRange & "') "
        OpenTbl(ADb, Atbl32, SQL)

        If Atbl32.RecordCount <> 0 Then
            InceCount = Atbl32("Count").Value

        End If

    End Sub


    Sub IncentivesControlRange()
        SQL = ""
        SQL = SQL & "Select * From 22_Incentives_Setup "
        SQL = SQL & "Where Actives = ('" & "Yes" & "') "
        OpenTbl(ADb, Atbl33, SQL)

        If Atbl33.RecordCount > 0 Then

            InceRange = Atbl33("MonthPeriodeRange").Value

        End If
        Me.Refresh()
    End Sub
#End Region

    Sub DokNumGen()

        SQL = ""
        SQL = SQL & "Select `Doc_Num` From 23_Surat_Dokter "
        SQL = SQL & "Order by Doc_Num Desc "
        OpenTbl(ADb, DbTbl6, SQL)
        If DbTbl6.RecordCount <> 0 Then
            DokDig = DbTbl6("Doc_Num").Value
            DokNum = Format(DokDig + 1, "0000000000")
        Else
            DokNum = "0000000001"
        End If
    End Sub
    Dim GetDate As Date
    Sub DokSuratSave()
        GetDate = DokCal.Text
        SQL = ""
        SQL = SQL & "Select * From 23_Surat_Dokter "
        SQL = SQL & "Where Nik = ('" & DokTbx1.Text & "') "
        SQL = SQL & "And Date = ('" & GetDate.ToString("yyyy-MM-dd") & "') "
        OpenTbl(ADb, Atbl43, SQL)
        If Not Atbl43.RecordCount <> 0 Then
            Atbl43.AddNew()

        End If


        Atbl43("Doc_Num").Value = DokNum
        Atbl43("Nik").Value = DokTbx1.Text
        Atbl43("Nama").Value = DokTbx2.Text
        Atbl43("Date").Value = DokCal.Text
        Atbl43("Time").Value = DokTime.Text
        Atbl43("Remark").Value = DokTbx3.Text

        Atbl43.Update()

        MsgBox("Done")

    End Sub

    Sub EmpLookDoktor()
        SQL = ""
        SQL = SQL & "Select `Nik`, `Name` from 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & DokTbx1.Text & "') "
        SQL = SQL & "Order by Nik"
        OpenTbl(ADb, Atb3, SQL)

        If Atb3.RecordCount > 0 Then

            DokTbx2.Text = Atb3("Name").Value

        Else
            MsgBox("Employee Not Found", MsgBoxStyle.Information, "Codex ~ QC Build " & BuildCounter & " Warning!!")

        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        DokTime.Text = TimeOfDay
    End Sub

    Private Sub DokAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DokAdd.Click
        DokNumGen()
        DokTbx1.Enabled = True
        DokTbx2.Enabled = True
        DokCal.Enabled = True
        DokTbx3.Enabled = True
        DokLook.Enabled = True

    End Sub

    Private Sub DokTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DokTbx1.KeyPress

        DokTbx1.CharacterCasing = CharacterCasing.Upper
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            EmpLookDoktor()
            IncentivesControlLoad()
            e.Handled = True

        End If


    End Sub

  
    Private Sub DokSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DokSave.Click
        If DokTbx1.Text = "" Or DokTbx2.Text = "" Or DokTbx3.Text = "" Then
            MsgBox("Please complete the required data for Doktor Permit")

        Else
            IncentivesControlSave()
            DokSuratSave()
            Me.Dispose()
            MainMenu.SuratDoktorToolStripMenuItem.PerformClick()
        End If

    End Sub
End Class