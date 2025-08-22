Option Explicit On


Public Class EmployeeOldBlock


    Dim AutoS As String


    Private Sub Employee2Block_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        MainMenu.Refresh()
    End Sub

    Private Sub Employee2Block_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadDB()
        LoadEmployee2()
        AutoSearchEnabler()

    End Sub



    Sub LoadEmployee(ByVal LookData As String)

        ViewGrid1.Rows.Clear()

        If ViewCmb1.Text = "Nik" Then

            SQL = ""
            SQL = SQL & "Select `Nik`, `Name`, `Pay` from 02_Name_Table "
            SQL = SQL & "Where Nik like '" & LookData & "%' "
            SQL = SQL & "Order by Nik"
            OpenTbl(ADb, Atb3, SQL)

            If Atb3.RecordCount <> 0 Then

                Atb3.MoveFirst()
                Do While Not Atb3.EOF

                    ViewGrid1.Rows.Add(Atb3("Nik").Value, Atb3("Name").Value, Atb3("Pay").Value)

                    Atb3.MoveNext()

                Loop


            End If

        Else

            If ViewCmb1.Text = "Name" Then

                SQL = ""
                SQL = SQL & "Select `Nik`, `Name`, `Pay` from 02_Name_Table "
                SQL = SQL & "Where Name like '" & LookData & "%' "
                SQL = SQL & "Order by Nik"
                OpenTbl(ADb, Atb3, SQL)

                If Atb3.RecordCount <> 0 Then

                    Atb3.MoveFirst()
                    Do While Not Atb3.EOF

                        ViewGrid1.Rows.Add(Atb3("Nik").Value, Atb3("Name").Value, Atb3("Pay").Value)

                        Atb3.MoveNext()

                    Loop


                End If


            End If
        End If


    End Sub

    Sub LoadEmployee2()


        SQL = ""
        SQL = SQL & "Select `Nik`, `Name`, `Pay` from 02_Name_Table "
        SQL = SQL & "Order by Nik"
        OpenTbl(ADb, Atb3, SQL)

        If Atb3.RecordCount <> 0 Then

            Atb3.MoveFirst()
            Do While Not Atb3.EOF

                ViewGrid1.Rows.Add(Atb3("Nik").Value, Atb3("Name").Value, Atb3("Pay").Value)

                Atb3.MoveNext()

            Loop


        End If

    End Sub

    Private Sub ViewCmb1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ViewCmb1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub ViewTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ViewTbx1.KeyPress

        If VECmb1.Checked = True Then
            LoadEmployee(ViewTbx1.Text)
        End If

        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            LoadEmployee(ViewTbx1.Text)
            e.Handled = True
        End If
    End Sub

    Private Sub ViewGrid1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles ViewGrid1.CellDoubleClick



        If EmpRadBtn1.Checked = True Then

            Dim NewMDIChild As New WorkBlock()
            WorkBlock.MdiParent = MainMenu

            WorkBlock.PanelTbx1.Text = ViewGrid1.CurrentRow.Cells(0).Value.ToString
            Me.Dispose()
            WorkBlock.Show()
            WorkBlock.EmpLookup()
            WorkBlock.PanelTbx1.Focus()
            WorkBlock.Enabled = True
            WorkBlock.IncentivesControlLoad()

        ElseIf EmpRadBtn2.Checked = True Then

            Dim NewMDIChild As New WorkBlock()
            WorkFastBlock.MdiParent = MainMenu

            WorkFastBlock.WPTbx1.Text = ViewGrid1.CurrentRow.Cells(0).Value.ToString
            Me.Dispose()
            WorkFastBlock.Show()
            WorkFastBlock.FEmpLookup()
            WorkFastBlock.ErrorPeriode()
            WorkFastBlock.ErrorHoliday()
            WorkFastBlock.WPTbx1.Focus()
            'WorkFastBlock.IncentivesControlLoad()

        End If

        MainMenu.Refresh()


    End Sub



    Private Sub VECmb1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VECmb1.CheckedChanged


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


    Sub AutoSearchEnabler()
        SQL = ""
        SQL = SQL & "Select * From 08_Standard_Table "
        SQL = SQL & "Where Original = ('" & "iAutoSearch" & "') "
        OpenTbl(ADb, Atbl35, SQL)

        If Atbl35.RecordCount > 0 Then

            AutoS = Atbl35("Standard_Wage").Value

        End If
        Me.Refresh()


        If AutoS = "Yes" Then

            VECmb1.Checked = True

        End If
    End Sub



    Private Sub ViewGrid1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles ViewGrid1.CellContentClick

    End Sub

    Private Sub ViewTbx1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewTbx1.TextChanged

    End Sub
End Class