Public Class NoRekBlock

    Dim NoRekCmbDater As String
    Dim NoRekCmbDater2 As String
    Dim NoRekCmbDater3 As String

    Dim UpNRNik As String
    Dim UpNRName As String
    Dim UpNRPerRa As String
    Dim UpNRPer As String
    Dim UpNRNoRek As String
    Dim UpNRPay As String

    Dim UpNRSyNik As String
    Dim UpNRSyPer As String
    Dim UpNRSyPerR As String
    Dim UpNRSyPay As String
    Dim UpNRSyRek As String
    Dim PaySeperator As String

    Private Sub NoRekBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        UpCmbFiller()
        LoadDB2()
        LoadDB()

    End Sub

    Sub UpCmbFiller()
        NoRekCmbDater = Format(Now, "yyyy")
        NoRekCmbDater2 = NoRekCmbDater - 1
        NoRekCmbDater3 = NoRekCmbDater + 1

        With NRCmb1

            .Items.Add("Dec " + NoRekCmbDater2)
            .Items.Add("Jan " + Format(Now, "yyyy"))
            .Items.Add("Feb " + Format(Now, "yyyy"))
            .Items.Add("Mar " + Format(Now, "yyyy"))
            .Items.Add("Apr " + Format(Now, "yyyy"))
            .Items.Add("May " + Format(Now, "yyyy"))
            .Items.Add("Jun " + Format(Now, "yyyy"))
            .Items.Add("Jul " + Format(Now, "yyyy"))
            .Items.Add("Aug " + Format(Now, "yyyy"))
            .Items.Add("Sep " + Format(Now, "yyyy"))
            .Items.Add("Oct " + Format(Now, "yyyy"))
            .Items.Add("Nov " + Format(Now, "yyyy"))
            .Items.Add("Dec " + Format(Now, "yyyy"))
            .Items.Add("Jan " + NoRekCmbDater3)

        End With
    End Sub
    Sub UpNRHeader()
        With UpNRGrid
            .Columns.Add("Col1", "Periode")
            .Columns.Add("Col2", "Periode Range")
            .Columns.Add("Col3", "Nik")
            .Columns.Add("Col4", "No Rek")
            .Columns.Add("Col5", "Pay(Current)")
            .Columns.Add("Col6", "Status")
            .Columns(0).Width = 125
            .Columns(1).Width = 125
            .Columns(2).Width = 125
            .Columns(3).Width = 125
            .Columns(4).Width = 125
            .Columns(5).Width = 125



        End With
    End Sub


    Private Sub NRCmb1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles NRCmb1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub


    Private Sub NRCmb2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles NRCmb2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Sub UpNRLook()

        SQL = ""
        SQL = SQL & "Select * from SalarySync1_Table "
        SQL = SQL & "Where Periode = ('" & NRCmb2.Text & "') "
        SQL = SQL & "And PeriodeRange = ('" & NRCmb1.Text & "') "
        SQL = SQL & "And Pay = ('" & PaySeperator & "') "
        SQL = SQL & "Order by Nik "
        OpenTbl(CBb, Ctbl25, SQL)

        If Ctbl25.RecordCount > 0 Then
            Ctbl25.MoveFirst()

            Do While Not Ctbl25.EOF

                UpNRNik = IIf(IsDBNull(Ctbl25("Nik").Value), "", Ctbl25("Nik").Value)
                UpNRPerRa = IIf(IsDBNull(Ctbl25("PeriodeRange").Value), "", Ctbl25("PeriodeRange").Value)
                UpNRPer = IIf(IsDBNull(Ctbl25("Periode").Value), "", Ctbl25("Periode").Value)

                NoRekLoad()

                UpNRGrid.Rows.Add(UpNRPer, UpNRPerRa, UpNRNik, UpNRNoRek, UpNRPay)

                Ctbl25.MoveNext()
            Loop

            MsgBox("Done")

        End If

    End Sub
    Sub UPNRSaver()

        For i = 0 To UpNRGrid.Rows.Count - 1

            UpNRSyPer = UpNRGrid(0, i).Value
            UpNRSyPerR = UpNRGrid(1, i).Value
            UpNRSyNik = UpNRGrid(2, i).Value
            UpNRSyRek = UpNRGrid(3, i).Value
            UpNRSyPay = UpNRGrid(4, i).Value

            SQL = ""
            SQL = SQL & "Select * From SalarySync1_Table "
            SQL = SQL & "Where Nik = ('" & UpNRSyNik & "') "
            SQL = SQL & "And Periode = ('" & UpNRSyPer & "') "
            SQL = SQL & "And PeriodeRange = ('" & UpNRSyPerR & "') "
            OpenTbl(CBb, Ctbl1, SQL)

            If Not Ctbl1.RecordCount <> 0 Then
                Ctbl1.AddNew()
            End If

            If NRChkBox1.Checked = True Then
                Ctbl1("PNoRek").Value = UpNRSyRek
            End If

            If NRChkBox2.Checked = True Then
                Ctbl1("Pay").Value = "BTN"
            End If

            Ctbl1.Update()

            UpNRGrid(5, i).Value = "Has Been Updated"
        Next

        MsgBox("Done")
    End Sub

    Sub NoRekLoad()
        SQL = ""
        SQL = SQL & "Select * From 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & UpNRNik & "') "
        OpenTbl(ADb, Atbl37, SQL)

        If Atbl37.RecordCount > 0 Then
            UpNRNoRek = IIf(IsDBNull(Atbl37("NoRek").Value), "", Atbl37("NoRek").Value)
            UpNRPay = IIf(IsDBNull(Atbl37("Pay").Value), "", Atbl37("Pay").Value)
        End If
    End Sub

    Private Sub UpNRBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpNRBtn1.Click
        UpNRBtn3.PerformClick()
        If NRChkBox1.Checked = False And NRChkBox2.Checked = False Then
            MsgBox("Please Select of which item do you want to Generate")
        Else

            If NRChkBox1.Checked = True Then
                PaySeperator = "BTN"

            ElseIf NRChkBox2.Checked = True Then
                PaySeperator = "CASH"

            End If

            UpNRHeader()
            UpNRLook()

        End If

    End Sub

    Private Sub UpNRBtn3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpNRBtn3.Click
        UpNRGrid.Rows.Clear()
        UpNRGrid.Columns.Clear()
        PaySeperator = ""
    End Sub

    Private Sub UpNRBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpNRBtn2.Click
        If NRChkBox1.Checked = False And NRChkBox2.Checked = False Then
            MsgBox("Please Select of which item do you want to Updated")
        Else
            UPNRSaver()
        End If

    End Sub


    Private Sub NRChkBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NRChkBox1.Click

        If NRChkBox2.Checked = True Then
            NRChkBox2.Checked = False
        End If
    End Sub


    Private Sub NRChkBox2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NRChkBox2.Click
        If NRChkBox1.Checked = True Then
            NRChkBox1.Checked = False
        End If

    End Sub
End Class