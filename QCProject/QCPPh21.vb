Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel

Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.ComponentModel

Public Class PPH21Block

    Dim ExcelAP As Excel.Application
    Dim ExcelWB As Excel.Workbook
    Dim ExcelWS As Excel.Worksheet
    Dim ExcelName As String

    Dim PPhFill1 As String
    Dim PPhFill2 As String
    Dim PPhFill3 As String
    Dim Portion As String = "'"

    Dim Pph21Nik As String
    Dim Pph21Name As String
    Dim pph21add As String
    Dim pph21KTP As String
    Dim pph21NPWP As String
    Dim pph21gaji1 As String
    Dim pph21gaji2 As String
    Dim pph21gaji3 As String
    Dim pph21astek As String
    Dim pph21incen As String
    Dim pph21nikUp As String

    Dim a As Integer
    Dim b As Integer

    Private Sub PPH21Block_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        UPandSlipFil()
        LoadDB()
        LoadDB2()
        LoadDBPPh21()


    End Sub

    Sub UPandSlipFil()

        PPhFill1 = Format(Now, "yyyy")
        PPhFill2 = PPhFill1 - 1
        PPhFill3 = PPhFill2 + 1

        With PPh21cmb1

            .Items.Add("Dec " + PPhFill2 + " - " + "Jan " + Format(Now, "yyyy"))
            .Items.Add("Jan " + Format(Now, "yyyy") + " - " + "Feb " + Format(Now, "yyyy"))
            .Items.Add("Feb " + Format(Now, "yyyy") + " - " + "Mar " + Format(Now, "yyyy"))
            .Items.Add("Mar " + Format(Now, "yyyy") + " - " + "Apr " + Format(Now, "yyyy"))
            .Items.Add("Apr " + Format(Now, "yyyy") + " - " + "May " + Format(Now, "yyyy"))
            .Items.Add("May " + Format(Now, "yyyy") + " - " + "Jun " + Format(Now, "yyyy"))
            .Items.Add("Jun " + Format(Now, "yyyy") + " - " + "Jul " + Format(Now, "yyyy"))
            .Items.Add("Jul " + Format(Now, "yyyy") + " - " + "Aug " + Format(Now, "yyyy"))
            .Items.Add("Aug " + Format(Now, "yyyy") + " - " + "Sep " + Format(Now, "yyyy"))
            .Items.Add("Sep " + Format(Now, "yyyy") + " - " + "Oct " + Format(Now, "yyyy"))
            .Items.Add("Oct " + Format(Now, "yyyy") + " - " + "Nov " + Format(Now, "yyyy"))
            .Items.Add("Nov " + Format(Now, "yyyy") + " - " + "Dec " + Format(Now, "yyyy"))
            .Items.Add("Dec " + Format(Now, "yyyy") + " - " + "Jan " + PPhFill3)

        End With
    End Sub

    Sub LoadPPh21ASTEKfromName()
        pph21astek = Nothing
        SQL = ""
        SQL = SQL & "Select `Nik`, `Jamsostek` From 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & Pph21Nik & "') "
        OpenTbl(ADb, Atb2, SQL)
        If Atb2.RecordCount > 0 Then
            pph21astek = IIf(IsDBNull(Atb2("Jamsostek").Value), "", Atb2("Jamsostek").Value)
        End If

    End Sub

    Sub PPh21Loader()

        If PPh21cmb2.Text = "BTN" Then

            SQL = ""
            SQL = SQL & "Select * from Emp_PPHTable "
            SQL = SQL & "Where PeriodeGajian = ('" & PPh21cmb1.Text & "') "
            SQL = SQL & "And Pay = ('" & "BTN" & "') "
            SQL = SQL & "Order by Nik "
            OpenTbl(PPhDB, PPhTb6, SQL)
            If PPhTb6.RecordCount > 0 Then
                PPhTb6.MoveFirst()
                Do While Not PPhTb6.EOF

                    Pph21Nik = IIf(IsDBNull(PPhTb6("Nik").Value), "", PPhTb6("Nik").Value)
                    Pph21Name = IIf(IsDBNull(PPhTb6("Name").Value), "", PPhTb6("Name").Value.ToString.Replace("?", ""))
                    pph21KTP = IIf(IsDBNull(PPhTb6("KTP").Value), "", PPhTb6("KTP").Value)
                    pph21NPWP = IIf(IsDBNull(PPhTb6("NPWP").Value), "", PPhTb6("NPWP").Value)
                    LoadPPh21ASTEKfromName()
                    pph21gaji1 = IIf(IsDBNull(PPhTb6("MainSalary1").Value), "", PPhTb6("MainSalary1").Value)
                    pph21gaji2 = IIf(IsDBNull(PPhTb6("MainSalary2").Value), "", PPhTb6("MainSalary2").Value)
                    pph21gaji3 = IIf(IsDBNull(PPhTb6("MainSalary3").Value), "", PPhTb6("MainSalary3").Value)
                    pph21incen = IIf(IsDBNull(PPhTb6("Incentif").Value), "", PPhTb6("Incentif").Value)
                    pph21add = IIf(IsDBNull(PPhTb6("EmAdd").Value), "", PPhTb6("EmAdd").Value)

                    pph21nikUp = Pph21Nik.ToUpper
                    PPhGrid1.Rows.Add(pph21nikUp, Pph21Name, pph21add, Portion + pph21KTP, pph21NPWP, pph21incen, pph21astek, pph21gaji1, pph21gaji2, pph21gaji3)
                    DataNester()

                    PPhTb6.MoveNext()

                Loop

            End If

        ElseIf PPh21cmb2.Text = "CASH" Then

            SQL = ""
            SQL = SQL & "Select * from Emp_PPHTable "
            SQL = SQL & "Where PeriodeGajian = ('" & PPh21cmb1.Text & "') "
            SQL = SQL & "And Pay = ('" & "CASH" & "') "
            SQL = SQL & "Order by Nik "
            OpenTbl(PPhDB, PPhTb6, SQL)
            If PPhTb6.RecordCount > 0 Then
                PPhTb6.MoveFirst()
                Do While Not PPhTb6.EOF

                    Pph21Nik = IIf(IsDBNull(PPhTb6("Nik").Value), "", PPhTb6("Nik").Value)
                    Pph21Name = IIf(IsDBNull(PPhTb6("Name").Value), "", PPhTb6("Name").Value.ToString.Replace("?", "'"))
                    pph21KTP = IIf(IsDBNull(PPhTb6("KTP").Value), "", PPhTb6("KTP").Value)
                    pph21NPWP = IIf(IsDBNull(PPhTb6("NPWP").Value), "", PPhTb6("NPWP").Value)
                    LoadPPh21ASTEKfromName()
                    pph21gaji1 = IIf(IsDBNull(PPhTb6("MainSalary1").Value), "", PPhTb6("MainSalary1").Value)
                    pph21gaji2 = IIf(IsDBNull(PPhTb6("MainSalary2").Value), "", PPhTb6("MainSalary2").Value)
                    pph21gaji3 = IIf(IsDBNull(PPhTb6("MainSalary3").Value), "", PPhTb6("MainSalary3").Value)
                    pph21incen = IIf(IsDBNull(PPhTb6("Incentif").Value), "", PPhTb6("Incentif").Value)
                    pph21add = IIf(IsDBNull(PPhTb6("EmAdd").Value), "", PPhTb6("EmAdd").Value)

                    pph21nikUp = Pph21Nik.ToUpper
                    PPhGrid1.Rows.Add(pph21nikUp, Pph21Name, pph21add, Portion + pph21KTP, pph21NPWP, pph21incen, pph21astek, pph21gaji1, pph21gaji2, pph21gaji3)
                    DataNester()
                    PPhTb6.MoveNext()

                Loop

            End If

        Else

            SQL = ""
            SQL = SQL & "Select * from Emp_PPHTable "
            SQL = SQL & "Where PeriodeGajian = ('" & PPh21cmb1.Text & "') "
            SQL = SQL & "Order by Nik "
            OpenTbl(PPhDB, PPhTb6, SQL)
            If PPhTb6.RecordCount > 0 Then
                PPhTb6.MoveFirst()
                Do While Not PPhTb6.EOF

                    Pph21Nik = IIf(IsDBNull(PPhTb6("Nik").Value), "", PPhTb6("Nik").Value)
                    Pph21Name = IIf(IsDBNull(PPhTb6("Name").Value), "", PPhTb6("Name").Value.ToString.Replace("?", "'"))
                    pph21KTP = IIf(IsDBNull(PPhTb6("KTP").Value), "", PPhTb6("KTP").Value)
                    pph21NPWP = IIf(IsDBNull(PPhTb6("NPWP").Value), "", PPhTb6("NPWP").Value)
                    LoadPPh21ASTEKfromName()
                    pph21gaji1 = IIf(IsDBNull(PPhTb6("MainSalary1").Value), "", PPhTb6("MainSalary1").Value)
                    pph21gaji2 = IIf(IsDBNull(PPhTb6("MainSalary2").Value), "", PPhTb6("MainSalary2").Value)
                    pph21gaji3 = IIf(IsDBNull(PPhTb6("MainSalary3").Value), "", PPhTb6("MainSalary3").Value)
                    pph21incen = IIf(IsDBNull(PPhTb6("Incentif").Value), "", PPhTb6("Incentif").Value)
                    pph21add = IIf(IsDBNull(PPhTb6("EmAdd").Value), "", PPhTb6("EmAdd").Value)

                    pph21nikUp = Pph21Nik.ToUpper
                    PPhGrid1.Rows.Add(pph21nikUp, Pph21Name, pph21add, Portion + pph21KTP, pph21NPWP, pph21incen, pph21astek, pph21gaji1, pph21gaji2, pph21gaji3)
                    DataNester()
                    PPhTb6.MoveNext()

                Loop

            End If

        End If

    End Sub

    Sub DataNester()

        Pph21Nik = Nothing
        Pph21Name = Nothing
        pph21add = Nothing
        pph21KTP = Nothing
        pph21NPWP = Nothing
        pph21gaji1 = Nothing
        pph21gaji2 = Nothing
        pph21astek = Nothing
        pph21incen = Nothing
        pph21nikUp = Nothing

    End Sub

    Private Sub PPh21cmb1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PPh21cmb1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub PPh21cmb2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles PPh21cmb2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub PPh21Btn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PPh21Btn1.Click
        PPh21Loader()
    End Sub

    Private Sub PPh21Btn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PPh21Btn2.Click
        SaveFileLink()
        OnClickTheWorker()
        'ExportExcel()
    End Sub

#Region "Excel Codes"

    Sub GenExcel()

        ExcelName = "Sortasi Report" & "_" & PPh21cmb1.Text & "_" & Format(Now, "dd.MM.yyyy Hmmss")

        KillExcel()
        StartExcel()
        CreateWorkSheet()
        PopWorkSheet()
        SaveWorkSheet()
        CloseWorkSheet()
        OpenMe()

        If Dir("C:\Program Files\Microsoft Office\Office12\excel.exe", vbDirectory) <> "" Then
            Shell("C:\Program Files\Microsoft Office\Office12\Excel " & Application.StartupPath & "\Report Excel\" & ExcelName & ".xls", vbMaximizedFocus)

        ElseIf Dir("C:\Program Files\Microsoft Office\OFFICE11\excel.exe", vbDirectory) <> "" Then
            Shell("C:\C:\Program Files\Microsoft Office\OFFICE11\Excel " & Application.StartupPath & "\Report Excel\" & ExcelName & ".xls", vbMaximizedFocus)

        ElseIf Dir("C:\Program Files\Microsoft Office\Office10\excel.exe", vbDirectory) <> "" Then
            Shell("C:\Program Files\Microsoft Office\Office11\Excel " & Application.StartupPath & "\Report Excel\" & ExcelName & ".xls", vbMaximizedFocus)

        ElseIf Dir("C:\Program Files\Microsoft Office\Office\excel.exe", vbDirectory) <> "" Then
            Shell("C:\Program Files\Microsoft Office\Office11\Excel " & Application.StartupPath & "\Report Excel\" & ExcelName & ".xls", vbMaximizedFocus)

        Else
            MsgBox("Microsoft Excel has not been found.", vbOKOnly + 64, "")
        End If

    End Sub

    Sub KillExcel()

        If Dir(Application.StartupPath & "\Reports Excel\" & ExcelName & ".xls") <> "" Then
            Kill(Application.StartupPath & "\Reports Excel\" & ExcelName & ".xls")
        End If

    End Sub

    Sub StartExcel()
        On Error GoTo Err
        ExcelAP = GetObject("Excel.Application")
        Exit Sub
Err:
        ExcelAP = CreateObject("Excel.Application")
    End Sub

    Sub CreateWorkSheet()
        ExcelWB = ExcelAP.Workbooks.Add
        ExcelWS = ExcelWB.Worksheets(1)
    End Sub

    Sub PopWorkSheet()

        ExcelWS.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperLegal
        ExcelWS.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape
        ExcelWS.PageSetup.PrintTitleRows = "A7"
        ExcelWS.PageSetup.Zoom = 85

        With ExcelAP.Range("A1:AA1")

            .Merge()
            .Cells.Value = "PT. UNIVERSAL GLOVES"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2

        End With

        With ExcelAP.Range("A2:AA2")

            .Merge()
            .Cells.Value = "JL. Pertahanan No. 17 Patumbak 20361 Deli Serdang  - Indonesia"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2

        End With

        With ExcelAP.Range("A3:AA3")

            .Merge()
            .Cells.Value = "DAFTAR GAJI BORONGAN PER PERIODE"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2

        End With

        With ExcelAP.Range("A4:AA4")

            .Merge()
            .Cells.Value = "BAGIAN : SORTASI"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2

        End With

        With ExcelAP.Range("A5:AA5")

            .Merge()
            .Font.Bold = True
            .Cells.Value = PPh21cmb1.Text
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .Font.Name = "Calibri"
            .Font.Size = 10

        End With

        '-----------------------------------------------------------------------------------------------

        With ExcelAP.Range("A7:A7")

            .Cells.Value = "NIK"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 15

        End With

        With ExcelAP.Range("B7:B7")

            .Cells.Value = "NAMA"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 20

        End With

        With ExcelAP.Range("C7:C7")

            .Cells.Value = "ADDRESS"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 20

        End With

        With ExcelAP.Range("D7:D7")

            .Cells.Value = "KTP"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 20

        End With

        With ExcelAP.Range("E7:E7")

            .Cells.Value = "NPWP"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 20

        End With

        With ExcelAP.Range("F7:F7")

            .Cells.Value = "Incentif"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 20

        End With

        With ExcelAP.Range("G7:G7")

            .Cells.Value = "ASTEK"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 20

        End With

        With ExcelAP.Range("H7:H7")

            .Cells.Value = "Gaji 1"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 20

        End With

        With ExcelAP.Range("I7:I7")

            .Cells.Value = "Gaji 2"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 20

        End With

        With ExcelAP.Range("J7:J7")

            .Cells.Value = "Gaji 3"
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Borders.LineStyle = 1
            .Font.Size = 10
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = 2
            .ColumnWidth = 20

        End With

        For Me.a = 0 To PPhGrid1.Rows.Count - 1
            For Me.b = 0 To PPhGrid1.ColumnCount - 1
                ExcelWS.Cells(a + 8, b + 1) = PPhGrid1(b, a).Value
                ExcelWS.Cells(a + 8, b + 1).Borders.LineStyle = 1
            Next
        Next

        a = a + 2

    End Sub

    Sub SaveWorkSheet()
        On Error GoTo Err
        SaveWorkBook()
Err:
    End Sub

    Sub SaveWorkBook()
        ExcelWB.SaveAs(Application.StartupPath & "\Report Excel\" & ExcelName & ".xls")
    End Sub

    Sub CloseWorkSheet()
        ExcelAP.Workbooks.Close()
        ExcelAP.Quit()
    End Sub

    Sub OpenMe()

        Dim oXLApp As Object, oXLWorkbook As Object

        oXLApp = CreateObject("Excel.Application")


        oXLWorkbook = oXLApp.Workbooks.Open(FileName:=Application.StartupPath & "\Report Excel\" & ExcelName & ".xls")

        oXLApp.Visible = True


    End Sub

#End Region


#Region "NEW Excel CALL"
    Dim SaveName As String
    Sub SaveFileLink()

        Dim SaveFileName As New SaveFileDialog
        SaveFileName.Filter = "Excel File (*.xlsx)|*.xlsx"
        SaveFileName.FilterIndex = 1
        If SaveFileName.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            SaveName = SaveFileName.FileName
        End If

    End Sub

    Sub ExportExcel()

        Try

            Dim NewFile As New FileInfo(SaveName)
            If NewFile.Exists Then
                NewFile.Delete()
            End If

            Using ExcelModPkg = New ExcelPackage(NewFile)

                ' Create Work Sheet

                Dim ExcelNewWSH As ExcelWorksheet = ExcelModPkg.Workbook.Worksheets.Add("GAJIAN")

                ExcelNewWSH.PrinterSettings.PaperSize = ePaperSize.Legal
                ExcelNewWSH.PrinterSettings.Orientation = eOrientation.Landscape

                With ExcelNewWSH.Cells("A1:J1")

                    .Merge = True
                    .Value = "PT. UNIVERSAL GLOVES"
                    .Style.Font.Bold = True
                    .Style.Font.Name = "Calibri"
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

                End With

                With ExcelNewWSH.Cells("A2:J2")

                    .Merge = True
                    .Value = "JL. Pertahanan No. 17 Patumbak 20361 Deli Serdang  - Indonesia"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

                End With

                With ExcelNewWSH.Cells("A3:J3")

                    .Merge = True
                    .Value = "DAFTAR GAJI BORONGAN PER PERIODE"
                    .Style.Font.Name = "Calibri"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

                End With

                With ExcelNewWSH.Cells("A4:J4")
                    .Merge = True
                    .Value = "BAGIAN : SORTASI " + PPh21cmb1.Text
                    .Style.Font.Name = "Calibri"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

                End With

                With ExcelNewWSH.Cells("A6")

                    .Value = "NIK"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With

                ' Block

                With ExcelNewWSH.Cells("B6")

                    .Value = "Name"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()

                End With
                With ExcelNewWSH.Cells("C6")
                    .Value = "Alamat"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("D6")

                    .Value = "No. KTP"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("E6")

                    .Value = "No. NPWP"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("F6")

                    .Value = "Incentif"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("G6")

                    .Value = "ASTEK"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("H6")

                    .Value = "GAJI I"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("I6")

                    .Value = "GAJI II"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With
                With ExcelNewWSH.Cells("J6")

                    .Value = "GAJI III"
                    .Style.Font.Bold = True
                    .Style.Font.Size = 10
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin)
                    .AutoFitColumns()
                End With

                For i = 0 To PPhGrid1.Rows.Count - 1
                    For j = 0 To PPhGrid1.ColumnCount - 1
                        ExcelNewWSH.Cells(i + 7, j + 1).Value = PPhGrid1(j, i).Value
                        ExcelNewWSH.Cells(i + 7, j + 1).Style.Border.BorderAround(ExcelBorderStyle.Thin)
                        ExcelNewWSH.Cells.Style.Font.Size = 8
                    Next

                    ExcelNewWSH.Cells.AutoFitColumns()
                Next

                'i = i + 2


                ExcelModPkg.Save()
                Dim LookMe As New ProcessStartInfo(SaveName)
                Process.Start(LookMe)

            End Using

        Catch ex As Exception

        End Try

    End Sub

#End Region

#Region "BGM on MODE"
    Private BGWorkMode() As BackgroundWorker
    Private i = 0
    Sub OnClickTheWorker()

        i += 1
        ReDim BGWorkMode(i)
        BGWorkMode(i) = New BackgroundWorker
        BGWorkMode(i).WorkerReportsProgress = True
        BGWorkMode(i).WorkerSupportsCancellation = True
        AddHandler BGWorkMode(i).DoWork, AddressOf WorkerDoWork
        AddHandler BGWorkMode(i).ProgressChanged, AddressOf WorkerProgressChanged
        AddHandler BGWorkMode(i).RunWorkerCompleted, AddressOf WorkerCompleted
        BGWorkMode(i).RunWorkerAsync()

    End Sub
    Private Sub WorkerDoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs)
        ExportExcel()
    End Sub

    Private Sub WorkerProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs)

    End Sub

    Private Sub WorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)
 
    End Sub

#End Region


    Private Sub PPh21Btn3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PPh21Btn3.Click
        PPhGrid1.Rows.Clear()
    End Sub
End Class