Imports System.Threading

Imports System.ComponentModel


Public Class UpEmpBlock

    Dim OpenMyPath As String
    Dim XlArrC1(10000) As String
    Dim XlArrC2(10000) As String
    Dim XlArrC3(10000) As String
    Dim XlArrC4(10000) As String
    Dim XlArrC5(10000) As String
    Dim XlArrC6(10000) As String
    Dim XlArrC7(10000) As String
    Dim XlArrC8(10000) As String
    Dim XlArrC9(10000) As String
    Dim XlArrC10(10000) As String
    Dim XlArrC11(10000) As String
    Dim XlArrC12(10000) As String
    Dim XlArrC13(10000) As String
    Dim XlArrC14(10000) As String
    Dim XlArrC15(10000) As String
    Dim XlArrC16(10000) As String
    Dim XlArrC17(10000) As String
    Dim XlArrC18(10000) As String
    Dim XlArrC19(10000) As String
    Dim XlArrC20(10000) As String
    Dim XlArrC21(10000) As String
    Dim XlArrC22(10000) As String
    Dim XlArrC23(10000) As String
    Dim XlArrC24(10000) As String
    Dim XlArrC25(10000) As String
    Dim CodeA As String = 0 ' Actuator in Saver1
    Dim CodeB As String = 0 ' Actuator in Saver2
    Dim CodeC As String = 0 ' Actuator in Both

    'Dim EmployeeDig As String
    'Dim EmployeeNum As String
    Dim OldEmployeeDig As String
    Dim OldEmployeeNum As String

    Dim UpIDSave As String
    Dim UpNikSave As String
    Dim UpNameSave As String
    Dim UpKTPSave As String
    Dim UpNPWPSave As String
    Dim UpKPJSave As String
    Dim UpJKKSave As String
    Dim UpEstSave As String
    Dim UpTemLahSave As String
    Dim UpAgSave As String
    Dim UpAlamSave As String
    Dim UpTelSave As String
    Dim UpPenSave As String
    Dim UpDeptSave As String
    Dim UpJabSave As String
    Dim UpAstSave As String
    Dim UpEfiktifSave As String
    Dim UpMaskerSave As String
    Dim UpPaySave As String
    Dim UpRekSave As String
    Dim UpHariLimSave As String
    Dim UpGajiMinSave As String
    Dim UpActiveSave As String

    Dim DataPD As String
    Dim DataID As String
    Dim DataNik As String
    Dim DataName As String
    Dim DataKTP As String
    Dim DataNPWP As String
    Dim DataKPJ As String
    Dim DataJKK As String
    Dim DataEst As String
    Dim DataTemLah As String
    Dim DataAg As String
    Dim DataAlam As String
    Dim DataTel As String
    Dim DataPen As String
    Dim DataDept As String
    Dim DataJab As String
    Dim DataAst As String
    Dim DataMasKer As String
    Dim DataEfiktif As String
    Dim DataPay As String
    Dim DataRek As String
    Dim DataHariLim As String
    Dim DataGajiMin As String
    Dim DataActive As String

    Dim UpEmpRead As String = 0
    Dim UpEmpSave As String = 0

#Region "Auto Number"

    'Sub GenEmployeeCode() ' Auto Generating Number

    '    SQL = ""
    '    SQL = SQL & "Select * From Emp_Table001 "
    '    SQL = SQL & "Order by PD_Id Desc"
    '    OpenTbl(PPhDB, PPhTb1, SQL)
    '    If PPhTb1.RecordCount <> 0 Then
    '        EmployeeDig = PPhTb1("PD_Id").Value
    '        EmployeeNum = Format(EmployeeDig + 1, "00000000")
    '    Else
    '        EmployeeNum = "00000001"
    '    End If

    'End Sub

    Sub OldGenEmployeeCode() ' Auto Generating Number

        SQL = ""
        SQL = SQL & "Select * From 02_Name_Table "
        SQL = SQL & "Order by ID_Number Desc"
        OpenTbl(ADb, DbTbl5, SQL)
        If DbTbl5.RecordCount > 0 Then
            OldEmployeeDig = DbTbl5("ID_Number").Value
            OldEmployeeNum = Format(OldEmployeeDig + 1, "00000000")
        Else
            OldEmployeeNum = "00000001"
        End If
    End Sub

#End Region

    Private Sub DateBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPEmpBtn1.Click

        OpenExcel.Filter = "Excel File(*.xls,*.xlsx ) |*.xls; *.xlsx |All files (*.*)|*.*"
        OpenExcel.InitialDirectory = Application.StartupPath
        If OpenExcel.ShowDialog <> Windows.Forms.DialogResult.Cancel Then

            OpenMyPath = OpenExcel.FileName
            UpEmpTbx1.Text = System.IO.Path.GetFileName(OpenExcel.FileName)

        End If

    End Sub

    Private Sub UpEmpBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadDBPPh21()
        LoadDB()

    End Sub

    Sub UpEmpGridRead()

        With UpEmpGrid

            .Columns.Add("Col0", "ID Number")
            .Columns.Add("Col1", "Nik")
            .Columns.Add("Col2", "Name")
            .Columns.Add("Col3", "No.KTP")
            .Columns.Add("Col4", "No.NPWP")
            .Columns.Add("Col5", "No.KPJ")
            .Columns.Add("Col6", "JKKJMM")
            .Columns.Add("Col7", "Estate")
            .Columns.Add("Col8", "Tempat Lahir")
            .Columns.Add("Col9", "Agama")
            .Columns.Add("Col10", "Alamat")
            .Columns.Add("Col11", "Tel Num")
            .Columns.Add("Col12", "Pendidikan")
            .Columns.Add("Col13", "Dept.")
            .Columns.Add("Col14", "Jabatan")
            .Columns.Add("Col15", "Astek")
            .Columns.Add("Col16", "Masuk Kerja")
            .Columns.Add("Col17", "Pay")
            .Columns.Add("Col18", "No.Rek")
            .Columns.Add("Col19", "Active")
            .Columns.Add("Col20", "Status")

            .Columns(0).Width = 200
            .Columns(1).Width = 200
            .Columns(2).Width = 200
            .Columns(3).Width = 200
            .Columns(4).Width = 200
            .Columns(5).Width = 200
            .Columns(6).Width = 200
            .Columns(7).Width = 200
            .Columns(8).Width = 200
            .Columns(9).Width = 200
            .Columns(10).Width = 200
            .Columns(11).Width = 200
            .Columns(12).Width = 200
            .Columns(13).Width = 200
            .Columns(14).Width = 200
            .Columns(15).Width = 200
            .Columns(16).Width = 200
            .Columns(17).Width = 200
            .Columns(18).Width = 200
            .Columns(19).Width = 200
            .Columns(20).Width = 200

        End With

    End Sub

    Sub GenerateNumber()

        OldGenEmployeeCode()

        For a = 0 To UpEmpGrid.Rows.Count - 1
            UpEmpGrid.Invoke(DirectCast(Sub() UpEmpGrid(0, a).Value = Format(OldEmployeeNum + a, "00000000"), MethodInvoker))
        Next


    End Sub
#Region "BGW on Mode"
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
        UpEmpAllocator()
    End Sub

    Private Sub WorkerProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs)

    End Sub

    Private Sub WorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)
        If e.Error IsNot Nothing Then
            MessageBox.Show(e.Error, Me.Text)
        Else
            MessageBox.Show("Upload Done", Me.Text)
        End If
    End Sub

    Sub OnClickTheWorker2()

        i += 1
        ReDim BGWorkMode(i)
        BGWorkMode(i) = New BackgroundWorker
        BGWorkMode(i).WorkerReportsProgress = True
        BGWorkMode(i).WorkerSupportsCancellation = True
        AddHandler BGWorkMode(i).DoWork, AddressOf WorkerDoWork2
        AddHandler BGWorkMode(i).ProgressChanged, AddressOf WorkerProgressChanged2
        AddHandler BGWorkMode(i).RunWorkerCompleted, AddressOf WorkerCompleted2
        BGWorkMode(i).RunWorkerAsync()

    End Sub

    Private Sub WorkerDoWork2(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs)
        EmpExcelReader()
    End Sub

    Private Sub WorkerProgressChanged2(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs)

    End Sub

    Private Sub WorkerCompleted2(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)
    End Sub

#End Region

    Sub UpEmpAllocator()

        For i = 0 To UpEmpGrid.Rows.Count - 1

            UpIDSave = UpEmpGrid(0, i).Value
            UpNikSave = UpEmpGrid(1, i).Value
            UpNameSave = UpEmpGrid(2, i).Value
            UpKTPSave = UpEmpGrid(3, i).Value
            UpNPWPSave = UpEmpGrid(4, i).Value
            UpKPJSave = UpEmpGrid(5, i).Value
            UpJKKSave = UpEmpGrid(6, i).Value
            UpEstSave = UpEmpGrid(7, i).Value
            UpTemLahSave = UpEmpGrid(8, i).Value
            UpAgSave = UpEmpGrid(9, i).Value
            UpAlamSave = UpEmpGrid(10, i).Value
            UpTelSave = UpEmpGrid(11, i).Value
            UpPenSave = UpEmpGrid(12, i).Value
            UpDeptSave = UpEmpGrid(13, i).Value
            UpJabSave = UpEmpGrid(14, i).Value
            UpAstSave = UpEmpGrid(15, i).Value
            UpMaskerSave = UpEmpGrid(16, i).Value
            UpPaySave = UpEmpGrid(17, i).Value
            UpRekSave = UpEmpGrid(18, i).Value
            UpActiveSave = UpEmpGrid(19, i).Value

 
            GenerateNumber()
            UpEmpSaver2()
            UpEmpGrid.Invoke(DirectCast(Sub() UpEmpGrid(20, i).Value = "Has been Save", MethodInvoker))
            UpEmpSave += 1
            UpEmpTbx3.Invoke(DirectCast(Sub() UpEmpTbx3.Text = UpEmpSave, MethodInvoker))


        Next

    End Sub

    Sub UpEmpSaver2()

        SQL = ""
        SQL = SQL & "Select * From 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & UpNikSave & "') "
        OpenTbl(ADb, Atb1, SQL)

        If Not Atb1.RecordCount <> 0 Then
            Atb1.AddNew()
        End If


        UpNameSave = UpNameSave.Replace("'", "?")
        Atb1("ID_Number").Value = UpIDSave
        Atb1("Nik").Value = UpNikSave
        Atb1("Name").Value = UpNameSave

        If Not UpActiveSave = Nothing Then
            Atb1("Active").Value = UpActiveSave
        End If

        If Not UpMaskerSave = Nothing Then
            Atb1("DateStart").Value = UpMaskerSave
        End If

        If Not UpPaySave = Nothing Then
            Atb1("Pay").Value = UpPaySave
        End If

        If Not UpAstSave = Nothing Then
            Atb1("Jamsostek").Value = UpAstSave
        End If

        If Not UpNPWPSave = Nothing Then
            Atb1("NPWP").Value = UpNPWPSave
        End If

        If Not UpRekSave = Nothing Then
            Atb1("NoRek").Value = UpRekSave
        End If

        If Not UpKTPSave = Nothing Then
            Atb1("NKTP").Value = UpKTPSave
        End If

        ' Extended

        If Not UpJKKSave = Nothing Then
            Atb1("JKKJKM").Value = UpJKKSave
        End If

        If Not UpTemLahSave = Nothing Then
            Atb1("Lahir").Value = UpTemLahSave
        End If

        If Not UpAgSave = Nothing Then
            Atb1("Agama").Value = UpAgSave
        End If

        If Not UpAlamSave = Nothing Then
            Atb1("Alamat").Value = UpAlamSave
        End If

        If Not UpTelSave = Nothing Then
            Atb1("TelNum").Value = UpTelSave
        End If

        If Not UpPenSave = Nothing Then
            Atb1("Pendi").Value = UpPenSave
        End If

        If Not UpDeptSave = Nothing Then
            Atb1("Dept").Value = UpDeptSave
        End If

        If Not UpJabSave = Nothing Then
            Atb1("JabData").Value = UpJabSave
        End If

        Atb1.Update()

    End Sub

    Sub EmpExcelReader()
        Dim xlRow As Long, xlCtr As Long

        StartExcel()
        OpenExlWbk(OpenMyPath)

        OpenExlWsh(1)
        xlCtr = 0

        ReDim XlArrC1(10000)
        ReDim XlArrC2(10000)
        ReDim XlArrC3(10000)
        ReDim XlArrC4(10000)
        ReDim XlArrC5(10000)
        ReDim XlArrC6(10000)
        ReDim XlArrC7(10000)
        ReDim XlArrC8(10000)
        ReDim XlArrC9(10000)
        ReDim XlArrC10(10000)
        ReDim XlArrC11(10000)
        ReDim XlArrC12(10000)
        ReDim XlArrC13(10000)
        ReDim XlArrC14(10000)
        ReDim XlArrC15(10000)
        ReDim XlArrC16(10000)
        ReDim XlArrC17(10000)
        ReDim XlArrC18(10000)
        ReDim XlArrC19(10000)
        ReDim XlArrC20(10000)

        For xlRow = 4 To 10000

            If ExcelWSh.Cells(xlRow, 4).Value = "END" Or ExcelWSh.Cells(xlRow, 3).Value = Nothing Then

                Exit For

            Else

                xlCtr = xlCtr + 1
                XlArrC2(xlCtr) = ExcelWSh.Cells(xlRow, 2).Value ' 2
                XlArrC3(xlCtr) = ExcelWSh.Cells(xlRow, 3).Value ' 3
                XlArrC4(xlCtr) = ExcelWSh.Cells(xlRow, 4).Value ' 4
                XlArrC5(xlCtr) = ExcelWSh.Cells(xlRow, 5).Value ' 5 
                XlArrC6(xlCtr) = ExcelWSh.Cells(xlRow, 6).Value ' 6
                XlArrC7(xlCtr) = ExcelWSh.Cells(xlRow, 7).Value ' 7
                XlArrC8(xlCtr) = ExcelWSh.Cells(xlRow, 8).Value ' 8
                XlArrC9(xlCtr) = ExcelWSh.Cells(xlRow, 9).Value ' 9
                XlArrC10(xlCtr) = ExcelWSh.Cells(xlRow, 10).Value ' 10
                XlArrC11(xlCtr) = ExcelWSh.Cells(xlRow, 11).Value ' 11
                XlArrC12(xlCtr) = ExcelWSh.Cells(xlRow, 12).Value ' 12
                XlArrC13(xlCtr) = ExcelWSh.Cells(xlRow, 13).Value ' 13
                XlArrC14(xlCtr) = ExcelWSh.Cells(xlRow, 14).Value ' 14
                XlArrC15(xlCtr) = ExcelWSh.Cells(xlRow, 15).Value ' 15
                XlArrC16(xlCtr) = ExcelWSh.Cells(xlRow, 16).Value ' 16
                XlArrC17(xlCtr) = ExcelWSh.Cells(xlRow, 17).Value ' 17
                XlArrC18(xlCtr) = ExcelWSh.Cells(xlRow, 18).Value ' 18
                XlArrC19(xlCtr) = ExcelWSh.Cells(xlRow, 19).Value ' 19
                XlArrC20(xlCtr) = ExcelWSh.Cells(xlRow, 20).Value ' 20

            End If

        Next xlRow

        For xlRow = 1 To xlCtr

            DataNik = XlArrC2(xlRow)
            DataName = XlArrC3(xlRow)
            DataKTP = XlArrC4(xlRow)
            DataNPWP = XlArrC5(xlRow)
            DataKPJ = XlArrC6(xlRow)
            DataJKK = XlArrC7(xlRow)
            DataEst = XlArrC8(xlRow)
            DataTemLah = XlArrC9(xlRow)
            DataAg = XlArrC10(xlRow)
            DataAlam = XlArrC11(xlRow)
            DataTel = XlArrC12(xlRow)
            DataPen = XlArrC13(xlRow)
            DataDept = XlArrC14(xlRow)
            DataJab = XlArrC15(xlRow)
            DataAst = XlArrC16(xlRow)
            DataMasKer = XlArrC17(xlRow)
            DataPay = XlArrC18(xlRow)
            DataRek = XlArrC19(xlRow)
            DataActive = XlArrC20(xlRow)

            UpEmpGrid.Invoke(DirectCast(Sub() UpEmpGrid.Rows.Add("", DataNik, DataName, DataKTP, DataNPWP, DataKPJ, DataJKK, DataEst, DataTemLah,
                               DataAg, DataAlam, DataTel, DataPen, DataDept, DataJab, DataAst, DataMasKer, DataPay, DataRek,
                               DataActive), MethodInvoker))
            UpEmpRead += 1
            UpEmpTbx2.Invoke(DirectCast(Sub() UpEmpTbx2.Text = UpEmpRead, MethodInvoker))

        Next xlRow

        CloseWorkSheet()

    End Sub

    Private Sub UPEmpBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPEmpBtn2.Click

        UpEmpGrid.Rows.Clear()
        UpEmpGrid.Columns.Clear()
        UpEmpGridRead()
        If UpEmpTbx1.Text = "" Then
            MsgBox("Please Select your File", vbExclamation)
        Else
            'EmpExcelReader()
            OnClickTheWorker2()
        End If
    End Sub

    Private Sub UpEmpTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles UpEmpTbx1.KeyPress

        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True

    End Sub

    Private Sub UpEmpTbx2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles UpEmpTbx2.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True

    End Sub

    Private Sub UpEmpTbx3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles UpEmpTbx3.KeyPress

        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True

    End Sub

    Private Sub UPEmpBtn3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPEmpBtn3.Click
       
        'UpEmpAllocator()
        OnClickTheWorker()
    End Sub

    Private Sub UPEmpBtn6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPEmpBtn6.Click

        UpEmpGrid.Rows.Clear()
        UpEmpGrid.Columns.Clear()
        UpEmpSave = 0

    End Sub

End Class