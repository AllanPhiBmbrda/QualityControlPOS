Option Explicit On


Public Class StandardBlock

    Dim StandardValue As String
    Dim ItemAdd As Integer



    Sub LoadStandard2()
        SQL = ""
        SQL = SQL & "Select * From 08_Standard_Table "
        SQL = SQL & "Where Original = ('" & StanCmb1.Text & "') "
        OpenTbl(ADb, Atbl22, SQL)
        If Atbl22.RecordCount > 0 Then
            Atbl22.MoveFirst()
            Do While Not Atbl22.EOF

                StandardValue = IIf(IsDBNull(Atbl22("Standard_Wage").Value), "", Atbl22("Standard_Wage").Value)
                Atbl22.MoveNext()
            Loop
        End If
    End Sub

    Sub StandardSave()

        SQL = ""
        SQL = SQL & "Select * From 08_Standard_Table "
        SQL = SQL & "Where Original = ('" & StanCmb1.Text & "') "
        OpenTbl(ADb, Atbl23, SQL)

        If Not Atbl23.RecordCount <> 0 Then
            Atbl23.AddNew()
        End If

        Atbl23("Original").Value = StanCmb1.Text
        Atbl23("Standard_Wage").Value = StanTbx1.Text

        Atbl23.Update()

        MsgBox("Success")
        Me.Refresh()

    End Sub

    Private Sub ComboBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles StanCmb1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub StanCmb1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StanCmb1.SelectedIndexChanged
        LoadStandard2()
        StanTbx1.Text = StandardValue
    End Sub

    Private Sub StanBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StanBtn1.Click
        StandardSave()
    End Sub

    Private Sub StandardBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadDB()
        StandardValue = 0
    End Sub
End Class