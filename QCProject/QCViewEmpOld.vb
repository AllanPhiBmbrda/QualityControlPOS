Option Explicit On


Public Class EmployeeOldBlock

    Private Sub Employee2Block_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        MainMenu.Refresh()
    End Sub

    Private Sub Employee2Block_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadDB()
        LoadEmployee2()

    End Sub

    Sub LoadEmployee()

        ViewGrid1.Rows.Clear()

        If ViewCmb1.Text = "Nik" Then

            SQL = ""
            SQL = SQL & "Select * from 02_Name_Table "
            SQL = SQL & "Where Nik = cstr('" & ViewTbx1.Text & "') "
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
                SQL = SQL & "Select * from 02_Name_Table "
                SQL = SQL & "Where Name = cstr('" & ViewTbx1.Text & "') "
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
        SQL = SQL & "Select * from 02_Name_Table "
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
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            LoadEmployee()
            e.Handled = True
        End If
    End Sub


    Private Sub ViewGrid1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ViewGrid1.DoubleClick


    End Sub


    Private Sub Employee2Block_Load_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
    Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'Employee2Block
        '
        Me.ClientSize = New System.Drawing.Size(284, 262)
        Me.Name = "Employee2Block"
        Me.ResumeLayout(False)

    End Sub
End Class