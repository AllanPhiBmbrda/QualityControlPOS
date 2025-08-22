Option Explicit On


Public Class EmployeeBlock


    Dim EmployeeNum As String
    Dim EmployeeDig As String
    Dim JamsosCode As String

    Private Sub EmployeeBlock_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        MainMenu.Refresh()
    End Sub

    Private Sub EmployeeBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadDB()
        EmpGridPop1()
        GenEmployeeCode()
        LoadJamsostekCombo()
        RegCmb3Items()

    End Sub

    Sub EmpGridPop1()

        EmpGrid1.Rows.Clear()

        SQL = ""
        SQL = SQL & "Select * From 02_Name_Table "
        OpenTbl(ADb, Atb2, SQL)
        If Atb2.RecordCount <> 0 Then

            Atb2.MoveFirst()
            Do While Not Atb2.EOF

                EmpGrid1.Rows.Add(Atb2("ID_Number").Value, Atb2("Nik").Value, Atb2("Name").Value, Atb2("Active").Value, Atb2("DateStart").Value, Atb2("Pay").Value, Atb2("Jamsostek").Value)
                Atb2.MoveNext()
            Loop

        End If

    End Sub

    Sub EmpGridPop2()
        EmpGrid1.Rows.Clear()

        If RegCmb4.Text = "Nik" Then

            SQL = ""
            SQL = SQL & "Select * From 02_Name_Table "
            SQL = SQL & "Where Nik = ('" & RegTbx5.Text & " ') "
            SQL = SQL & "And Active = ('" & RegCmb5.Text & " ') "
            OpenTbl(ADb, Atb2, SQL)
            If Atb2.RecordCount <> 0 Then

                Atb2.MoveFirst()
                Do While Not Atb2.EOF

                    EmpGrid1.Rows.Add(Atb2("ID_Number").Value, Atb2("Nik").Value, Atb2("Name").Value, Atb2("Active").Value, Atb2("DateStart").Value, Atb2("Pay").Value, Atb2("Jamsostek").Value)
                    Atb2.MoveNext()
                Loop

            End If
        Else
            If RegCmb4.Text = "Name" Then

                SQL = ""
                SQL = SQL & "Select * From 02_Name_Table "
                SQL = SQL & "Where Name = ('" & RegTbx5.Text & " ') "
                SQL = SQL & "and Active = ('" & RegCmb5.Text & "') "
                OpenTbl(ADb, Atb2, SQL)
                If Atb2.RecordCount <> 0 Then

                    Atb2.MoveFirst()
                    Do While Not Atb2.EOF

                        EmpGrid1.Rows.Add(Atb2("ID_Number").Value, Atb2("Nik").Value, Atb2("Name").Value, Atb2("Active").Value, Atb2("DateStart").Value, Atb2("Pay").Value, Atb2("Jamsostek").Value)
                        Atb2.MoveNext()

                    Loop

                End If
            End If
        End If

    End Sub

    Private Sub RegCmb5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles RegCmb5.KeyPress

        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True

    End Sub

    Private Sub RegCmb5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegCmb5.SelectedIndexChanged
        EmpGridPop2()
    End Sub

    Private Sub RegTbx5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles RegTbx5.KeyPress

        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            EmpGridPop2()
            e.Handled = True
        End If

    End Sub

    Private Sub RegCmb4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles RegCmb4.KeyPress

        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True

    End Sub

    Private Sub EmpBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmpBtn2.Click

        RegTbx1.Text = EmployeeNum
        ControlEnabler()

    End Sub

    Sub ControlEnabler()

        RegTbx1.Enabled = True
        RegTbx2.Enabled = True
        RegTbx3.Enabled = True
        MaskRegTbx1.Enabled = True
        RegCmb1.Enabled = True
        RegCmb2.Enabled = True
        RegCmb3.Enabled = True

    End Sub

    Private Sub MaskRegTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MaskRegTbx1.KeyPress

        MaskRegTbx1.Mask = "##/##/####"

    End Sub

    Sub GenEmployeeCode() ' Auto Generating Number

        SQL = ""
        SQL = SQL & "Select * From 02_Name_Table "
        SQL = SQL & "Order by ID_Number Desc"
        OpenTbl(ADb, DbTbl5, SQL)
        If DbTbl5.RecordCount > 0 Then
            EmployeeDig = DbTbl5("ID_Number").Value
            EmployeeNum = Format(EmployeeDig + 1, "00000000")
        Else
            EmployeeNum = "00000001"
        End If

    End Sub

    Private Sub RegTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles RegTbx1.KeyPress

        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True

    End Sub

    Private Sub EmpBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmpBtn1.Click

        If RegTbx1.Text = "" Then
            MsgBox("Please Click the Add Button First", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

        Else

            If RegTbx2.Text = "" Then
                MsgBox("Please Insert the Required Nik Number", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

            Else

                If RegTbx3.Text = "" Then
                    MsgBox("Please Insert the Required Name", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

                Else

                    If MaskRegTbx1.Text = "" Then
                        MsgBox("Please Insert the Starting Date", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

                    Else

                        If RegCmb1.Text = "" Then
                            MsgBox("Please Select the Working Status ", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

                        Else

                            If RegCmb2.Text = "" Then
                                MsgBox("Please Select the Paying Option", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

                            Else

                                If RegCmb3.Text = "" Then
                                    MsgBox("Please Select the Insurance Value", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

                                Else

                                    AntiDupNik()

                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

    End Sub

    Sub EmployeeDataSave()

        SQL = ""
        SQL = SQL & "Select * From 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & RegTbx2.Text & "') "
        SQL = SQL & "And ID_Number = ('" & RegTbx1.Text & "')"
        OpenTbl(ADb, Atb1, SQL)

        If Not Atb1.RecordCount <> 0 Then
            Atb1.AddNew()
        End If

        Atb1("ID_Number").Value = RegTbx1.Text
        Atb1("Nik").Value = RegTbx2.Text
        Atb1("Name").Value = RegTbx3.Text
        Atb1("DateStart").Value = MaskRegTbx1.Text
        Atb1("Active").Value = RegCmb1.Text
        Atb1("Pay").Value = RegCmb2.Text
        Atb1("Jamsostek").Value = RegCmb3.Text

        Atb1.Update()

        MsgBox("Success")
        Me.Refresh()

        RegTbx1.Text = ""
        RegTbx2.Text = ""
        RegTbx3.Text = ""
        MaskRegTbx1.Text = ""
        RegCmb1.Text = ""
        RegCmb2.Text = ""
        RegCmb3.Text = ""

        ControlDisabler()
        EmpGridPop1()

        Me.Dispose()
        Me.Close()

    End Sub

    Sub ControlDisabler()

        RegTbx1.Enabled = False
        RegTbx2.Enabled = False
        RegTbx3.Enabled = False
        MaskRegTbx1.Enabled = False
        RegCmb1.Enabled = False
        RegCmb2.Enabled = False
        RegCmb3.Enabled = False

    End Sub

    Sub AntiDupNik()

        SQL = ""
        SQL = SQL & "Select * from 02_Name_Table where "
        SQL = SQL & "Nik = ('" & RegTbx2.Text & "') "
        OpenTbl(ADb, DBTbl2, SQL)

        If DBTbl2.RecordCount > 0 Then
            MsgBox("NIK number is already taken", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

        Else
            EmployeeDataSave()

        End If

    End Sub
    Sub LoadJamsostekCombo()

        SQL = ""
        SQL = SQL & "Select * From 08_Standard_Table "
        SQL = SQL & "Where Original = ('" & "Jamsostek" & "') "
        OpenTbl(ADb, Atbl20, SQL)
        If Atbl20.RecordCount > 0 Then
            Atbl20.MoveFirst()
            Do While Not Atbl20.EOF

                JamsosCode = IIf(IsDBNull(Atbl20("Standard_Wage").Value), "", Atbl20("Standard_Wage").Value)

                Atbl20.MoveNext()

            Loop

        End If

    End Sub

    Sub RegCmb3Items()

        With RegCmb3

            .Items.Add(JamsosCode)
            .Items.Add("0.00")

        End With
    End Sub

    Private Sub RegCmb1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles RegCmb1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub RegCmb2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles RegCmb2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub RegCmb3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles RegCmb3.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub EmpGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles EmpGrid1.DoubleClick

        RegTbx1.Text = EmpGrid1.CurrentRow.Cells(0).Value.ToString
        RegTbx2.Text = EmpGrid1.CurrentRow.Cells(1).Value.ToString
        RegTbx3.Text = EmpGrid1.CurrentRow.Cells(2).Value.ToString
        RegCmb1.Text = EmpGrid1.CurrentRow.Cells(3).Value.ToString
        MaskRegTbx1.Text = EmpGrid1.CurrentRow.Cells(4).Value.ToString
        RegCmb2.Text = EmpGrid1.CurrentRow.Cells(5).Value.ToString
        RegCmb3.Text = EmpGrid1.CurrentRow.Cells(6).Value.ToString

        ControlEnabler()
        EmpBtn3.Visible = True

    End Sub

    Private Sub EmpBtn3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmpBtn3.Click
        If RegTbx1.Text = "" Then
            MsgBox("Please Click the Add Button First", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

        Else
            If RegTbx2.Text = "" Then
                MsgBox("Please Insert the Required Nik Number", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

            Else
                If RegTbx3.Text = "" Then
                    MsgBox("Please Insert the Required Name", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

                Else
                    If MaskRegTbx1.Text = "" Then
                        MsgBox("Please Insert the Starting Date", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

                    Else
                        If RegCmb1.Text = "" Then
                            MsgBox("Please Select the Working Status ", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

                        Else
                            If RegCmb2.Text = "" Then
                                MsgBox("Please Select the Paying Option", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

                            Else
                                If RegCmb3.Text = "" Then
                                    MsgBox("Please Select the Insurance Value", MsgBoxStyle.Critical, "Codex ~ QC Build " & BuildCounter & " Warning!!")

                                Else
                                    EmployeeDataSave()
                                    EmpBtn3.Visible = True

                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub RegTbx2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles RegTbx2.KeyPress
        RegTbx2.CharacterCasing = CharacterCasing.Upper
    End Sub

    Private Sub EmpGrid1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles EmpGrid1.CellContentClick

    End Sub

    Private Sub RegTbx5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegTbx5.TextChanged

    End Sub

    Private Sub RegCmb2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RegCmb2.SelectedIndexChanged

    End Sub
End Class
