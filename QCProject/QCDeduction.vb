



Public Class DedBlock

 

    Dim OpenDedPath As String
    Dim DedCmbDater As String
    Dim DedCmbDater2 As String
    Dim DedCmbDater3 As String
    Dim XlArrC1(10000) As String
    Dim XlArrC2(10000) As String
    Dim DedDataNik As String
    Dim DedDataDeduc As String
    Dim DedRead As String = 0
    Dim DedSave As String = 0
    Dim DedGridRdNik As String
    Dim DedGridRdDeduc As String


    Private Sub DedBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DedDater()
        LoadDB2()
    End Sub

    Private Sub DedBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DedBtn1.Click
        OpenDedExcel.Filter = "Excel File(*.xls,*.xlsx ) |*.xls; *.xlsx |All files (*.*)|*.*"
        OpenDedExcel.InitialDirectory = Application.StartupPath
        If OpenDedExcel.ShowDialog <> Windows.Forms.DialogResult.Cancel Then

            UpDedTbx1.Text = System.IO.Path.GetFileName(OpenDedExcel.FileName)
            OpenDedPath = OpenDedExcel.FileName

        End If
    End Sub

    Sub DedGridHeader()
        With DedUpGrid

            .Columns.Add("Col1", "NIK")
            .Columns.Add("Col2", "Deduction")
            .Columns.Add("Col3", "Status")
            .Columns(0).Width = 250
            .Columns(1).Width = 250
            .Columns(2).Width = 200


        End With

    End Sub

    Sub DedDater()

        DedCmbDater = Format(Now, "yyyy")
        DedCmbDater2 = DedCmbDater - 1
        DedCmbDater3 = DedCmbDater + 1

        With DedPerCmb1

            .Items.Add("Dec " + DedCmbDater2)
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
            .Items.Add("Jan " + DedCmbDater3)

        End With

    End Sub

    Sub PotSaver()

        For i = 0 To DedUpGrid.Rows.Count - 1

            DedGridRdNik = DedUpGrid(0, i).Value
            DedGridRdDeduc = DedUpGrid(1, i).Value


            SQL = ""
            SQL = SQL & "Select * From SalarySync1_Table "
            SQL = SQL & "Where Nik = ('" & DedGridRdNik & "') "
            SQL = SQL & "And Periode = ('" & DedPerCmb2.Text & "') "
            SQL = SQL & "And PeriodeRange = ('" & DedPerCmb1.Text & "') "
            OpenTbl(CBb, Ctbl1, SQL)


            If Not Ctbl1.RecordCount <> 0 Then
                Ctbl1.AddNew()
            End If

            Ctbl1("Nik").Value = DedGridRdNik
            Ctbl1("PotLain").Value = DedGridRdDeduc
            Ctbl1("Periode").Value = DedPerCmb2.Text
            Ctbl1("PeriodeRange").Value = DedPerCmb1.Text
            Ctbl1.Update()
            DedSave = DedSave + 1
            UpDedTbx3.Text = DedSave
            DedUpGrid(2, i).Value = "Has Been Saved"
            Me.Refresh()

        Next

        MsgBox("Done")
    End Sub

    Sub ExcelReader()
        DedRead = 0

        Dim xlRow As Long, xlCtr As Long

        StartExcel()
        OpenExlWbk(OpenDedPath)

        OpenExlWsh(1)
        xlCtr = 0


        ReDim XlArrC1(10000)
        ReDim XlArrC2(10000)

        For xlRow = 4 To 10000

            If ExcelWSh.Cells(xlRow, 1).Value = "END" Then

                Exit For

            Else

                xlCtr = xlCtr + 1
                XlArrC1(xlCtr) = ExcelWSh.Cells(xlRow, 1).Value
                XlArrC2(xlCtr) = ExcelWSh.Cells(xlRow, 2).Value


            End If

        Next xlRow

        For xlRow = 1 To xlCtr


            DedDataNik = XlArrC1(xlRow)
            DedDataDeduc = XlArrC2(xlRow)
            DedUpGrid.Rows.Add(DedDataNik, DedDataDeduc)

            DedRead = DedRead + 1
            UpDedTbx2.Text = DedRead


        Next xlRow

        CloseWorkSheet()
    End Sub

    Private Sub DedBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DedBtn2.Click
        DedUpGrid.Rows.Clear()
        DedUpGrid.Columns.Clear()

        DedGridHeader()

        If DedPerCmb1.Text = "" Or DedPerCmb2.Text = "" Or UpDedTbx1.Text = "" Then
            MsgBox("Please Select the Periode Range or Excel File", vbExclamation)

        Else
            ExcelReader()
        End If
    End Sub

    Private Sub DedBtn3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DedBtn3.Click
        PotSaver()
    End Sub

    Private Sub DedPerCmb1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DedPerCmb1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub DedPerCmb2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DedPerCmb2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub UpDedTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles UpDedTbx1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub UpDedTbx2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles UpDedTbx2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub UpDedTbx3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles UpDedTbx3.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

 
End Class