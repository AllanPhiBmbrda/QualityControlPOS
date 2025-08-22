Public Class LoginDoor

    Private Sub LoginDoor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadDB()
        LoginTbx1.Focus()
        LoginBox.Text = "Build " & BuildCounter

    End Sub

    Private Sub LogCmd01_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogCmd01.Click

        SQL = ""
        SQL = SQL & "select * from 01_account_table where "
        SQL = SQL & "Username = ('" & LoginTbx1.Text & "') and "
        SQL = SQL & "Userpass = ('" & LoginTbx2.Text & "') "
        OpenTbl(ADb, DBTbl1, SQL)
        If DBTbl1.RecordCount > 0 Then
            With UserRec
                .QcName = DBTbl1("UserAccName").Value
                .QcUserName = DBTbl1("UserName").Value
                .QcPassword = DBTbl1("Userpass").Value
                .QcLevel = DBTbl1("Userlevel").Value

                User = DBTbl1("UserAccName").Value
                UsFieCode = DBTbl1("UserFieldCode").Value

                Sidelogon()

            End With

            Me.Hide()
            MainMenu.Show()

        ElseIf LogTimer.Interval >= 1000 Then
            MsgBox("Only 5 time attempt will be accepted at this moment", vbExclamation, "Codex Build 1.0")
            Me.Close()

        Else

            LogTimer.Interval = LogTimer.Interval + 200
            MsgBox("Incorrect Username/Password", vbCritical)
            LoginTbx1.Text = ""
            LoginTbx2.Text = ""

        End If

    End Sub

    Sub Sidelogon()

        SQL = ""
        SQL = SQL & "select * from 08_Standard_Table where "
        SQL = SQL & "Original = ('" & "OriginalWage" & "') "
        OpenTbl(ADb, DBTbl3, SQL)

        If DBTbl3.RecordCount > 0 Then

            StandardsSalary = DBTbl3("Standard_Wage").Value

        End If

        '----------- For Subsidi-------------------------

        SQL = ""
        SQL = SQL & "select * from 08_Standard_Table where "
        SQL = SQL & "Original = ('" & "Subsidi" & "') "
        OpenTbl(ADb, Dbtb38, SQL)

        If Dbtb38.RecordCount > 0 Then

            SubsidiSalary = Dbtb38("Standard_Wage").Value

        End If

        '---------- For Gaji Ctrl------------------------

        SQL = ""
        SQL = SQL & "select * from 08_Standard_Table where "
        SQL = SQL & "Original = ('" & "GajiCtrl" & "') "
        OpenTbl(ADb, Dbtb39, SQL)

        If Dbtb39.RecordCount > 0 Then

            GajiCtrlSalary = Dbtb39("Standard_Wage").Value

        End If
    End Sub

    Private Sub LoginTbx2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles LoginTbx2.KeyPress
        If Not InValid4.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            LogCmd01.Focus()
            e.Handled = True
        End If
    End Sub

    Private Sub LoginTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles LoginTbx1.KeyPress
        If Not InValid4.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then

            LoginTbx2.Focus()
            e.Handled = True

        End If

    End Sub

    Private Sub LogCmd02_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogCmd02.Click

        Me.Dispose()

    End Sub

End Class