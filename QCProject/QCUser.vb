
Option Explicit On


Public Class UserBlock

    Dim AUser As String
    Dim UserNum As String
    Dim UserDig As String

    Private Sub UserBlock_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        MainMenu.Refresh()
    End Sub


    Private Sub UserBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadDB()
        UserGridPop()
    End Sub

    Private Sub UserTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles UserTbx1.KeyPress
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            UserTbx2.Focus()
            e.Handled = True
        End If
    End Sub

    Private Sub UserTbx2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles UserTbx2.KeyPress
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            UserTbx3.Focus()
            e.Handled = True
        End If

    End Sub

    Private Sub UserTbx3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles UserTbx3.KeyPress
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            UserCmb1.Focus()
            e.Handled = True
        End If
    End Sub

    Sub GenUserCode()

        SQL = ""
        SQL = SQL & "Select * From 01_Account_Table "
        SQL = SQL & "Order by UserNumber Desc"
        OpenTbl(ADb, Dbtb26, SQL)
        If Dbtb26.RecordCount <> 0 Then
            UserDig = DBTb26("UserNumber").Value
            UserNum = Format(UserDig + 1, "00000")
        Else
            UserNum = "00001"
        End If
    End Sub

    Sub SaveUser()


        SQL = ""
        SQL = SQL & "Select * From 01_Account_Table "
        SQL = SQL & "Where Username = ('" & UserTbx1.Text & "') "

        OpenTbl(ADb, Atbl17, SQL)

        If Not Atbl17.RecordCount <> 0 Then
            Atbl17.AddNew()
        End If

        Atbl17("Username").Value = UserTbx1.Text
        Atbl17("Userpass").Value = UserTbx2.Text
        Atbl17("UserAccName").Value = UserTbx3.Text
        Atbl17("Userlevel").Value = UserCmb1.Text
        Atbl17("UserNumber").Value = UserNum
        Atbl17("UserFieldCode").Value = UserCmb2.Text

        Atbl17.Update()

        MsgBox("User Account Saved", MsgBoxStyle.Information, "Codex ~ QC Build " & BuildCounter & " Success")
        Me.Refresh()
        UserGridPop()

    End Sub


    Private Sub UserCmb1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserCmb1.SelectedIndexChanged
        UserBtn1.Focus()
    End Sub

    Private Sub UserBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserBtn1.Click
        If UserTbx1.Text = "" Then
            MsgBox("Please insert the desired Username", MsgBoxStyle.Information, "Codex ~ Build 1.00")
        Else
            If UserTbx2.Text = "" Then
                MsgBox("Please insert the Desired Password", MsgBoxStyle.Information, "Codex ~ Build 1.00")

            Else
                If UserTbx3.Text = "" Then
                    MsgBox("Please insert the Required Name", MsgBoxStyle.Information, "Codex ~ Build 1.00")

                Else
                    If UserCmb1.Text = "" Then
                        MsgBox("Please Select the authority level for new account", MsgBoxStyle.Information, "Codex ~ Build 1.00")

                    Else
                        SaveUser()
                    End If
            End If
        End If
        End If


    End Sub

    Sub UserGridPop()

        UserGrid.Rows.Clear()

        sql = ""
        SQL = SQL & "Select * From 01_Account_Table "
        OpenTbl(ADb, Atbl17, SQL)
        If Atbl17.RecordCount <> 0 Then
            Atbl17.MoveFirst()
            Do While Not Atbl17.EOF

                UserGrid.Rows.Add(Atbl17("Username").Value, Atbl17("Userpass").Value, Atbl17("Userlevel").Value, Atbl17("UserAccName").Value)

                Atbl17.MoveNext()
            Loop

        End If

    End Sub


    Private Sub UserBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserBtn2.Click
        GenUserCode()
        UserTbx1.Enabled = True
        UserTbx2.Enabled = True
        UserTbx3.Enabled = True
        UserCmb1.Enabled = True
        UserCmb2.Enabled = True

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub


End Class
