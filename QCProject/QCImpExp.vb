Imports System.IO
Imports System.ComponentModel



Public Class QCImpExp
    Dim PartExport As String = Nothing
    Dim path As String = Nothing
    Dim DateStringMode As String
    Dim DateStringMode2 As String
    Dim AddStringforPath As String
    Dim ExportDateFrom As Date
    Dim ExportDateUntil As Date
    Dim ExportDateGet As Date

    Private Sub QCImpExp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadDB()
        LoadDB2()
        LoadDBPPh21()
        LookCreateFolder()
    End Sub

    Sub LookCreateFolder()
        If (Not System.IO.Directory.Exists(Application.StartupPath + "\CSVFolder")) Then
            System.IO.Directory.CreateDirectory(Application.StartupPath + "\CSVFolder")
        End If
    End Sub

    Sub AccountExport()

        PartExport = "QCACCOUNTINFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(Path) Then
            File.Delete(Path)
        End If

        If Not File.Exists(Path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(Path)
                sw.WriteLine(Encrypt("UID;UN;UP;UL;UAN;UFC"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From 01_account_table "
        OpenTbl(ADb, Atb1, SQL)
        If Atb1.RecordCount <> 0 Then
            QCIEPrg01.Maximum = Atb1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                Atb1.MoveFirst()
                Do While Not Atb1.EOF

                    QCIEPrg01.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(Atb1("UserNumber").Value), "", Atb1("UserNumber").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Username").Value), "", Atb1("Username").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Userpass").Value), "", Atb1("Userpass").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Userlevel").Value), "", Atb1("Userlevel").Value) + ";" + _
                    IIf(IsDBNull(Atb1("UserAccName").Value), "", Atb1("UserAccName").Value) + ";" + _
                    IIf(IsDBNull(Atb1("UserFieldCode").Value), "", Atb1("UserFieldCode").Value)))
                    Atb1.MoveNext()
                Loop
            End Using

        End If


    End Sub

    Sub NameExport()

        PartExport = "QCEMPLOYEEINFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(path) Then
            File.Delete(path)
        End If

        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Encrypt("NID;NIK;NAME;ACTIVE;DATESTART;PAY;JAMSOSTEK;NPWP;NOREK;NKTP;JABDATA;NOKPJ;BANKCTRL;ESTATE;LAHIR;AGAMA;ALAMAT;TELNUM;PENDI;DEPT;JKKJKM"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From 02_name_table "
        OpenTbl(ADb, Atb1, SQL)
        If Atb1.RecordCount <> 0 Then
            QCIEPrg02.Maximum = Atb1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                Atb1.MoveFirst()
                Do While Not Atb1.EOF

                    QCIEPrg02.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(Atb1("ID_Number").Value), "", Atb1("ID_Number").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Nik").Value), "", Atb1("Nik").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Name").Value), "", Atb1("Name").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Active").Value), "", Atb1("Active").Value) + ";" + _
                    IIf(IsDBNull(Atb1("DateStart").Value), "", Atb1("DateStart").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Pay").Value), "", Atb1("Pay").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Jamsostek").Value), "", Atb1("Jamsostek").Value) + ";" + _
                    IIf(IsDBNull(Atb1("NPWP").Value), "", Atb1("NPWP").Value) + ";" + _
                    IIf(IsDBNull(Atb1("NoRek").Value), "", Atb1("NoRek").Value) + ";" + _
                    IIf(IsDBNull(Atb1("NKTP").Value), "", Atb1("NKTP").Value) + ";" + _
                    IIf(IsDBNull(Atb1("JabData").Value), "", Atb1("JabData").Value) + ";" + _
                    IIf(IsDBNull(Atb1("NoKPJ").Value), "", Atb1("NoKPJ").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Bank_Ctrl").Value), "", Atb1("Bank_Ctrl").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Estate").Value), "", Atb1("Estate").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Lahir").Value), "", Atb1("Lahir").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Agama").Value), "", Atb1("Agama").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Alamat").Value), "", Atb1("Alamat").Value) + ";" + _
                    IIf(IsDBNull(Atb1("TelNum").Value), "", Atb1("TelNum").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Pendi").Value), "", Atb1("Pendi").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Dept").Value), "", Atb1("Dept").Value) + ";" + _
                    IIf(IsDBNull(Atb1("JKKJKM").Value), "", Atb1("JKKJKM").Value)))
                    Atb1.MoveNext()
                Loop
            End Using

        End If

    End Sub

    Sub ConveyourExport()

        PartExport = "QCCONVEYOURINFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(path) Then
            File.Delete(path)
        End If

        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Encrypt("PID;DATE;TIME;NIK;PIECES;TARGET;SALARY;COUPON"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From 03_conveyour_table "
        If QCIEChkBoxDate01.Checked = True Then
            SQL = SQL & "Where Date between ('" & ExportDateFrom.ToString("yyyy-MM-dd") & "') "
            SQL = SQL & "And ('" & ExportDateUntil.ToString("yyyy-MM-dd") & "') "
        End If
        OpenTbl(ADb, Atb1, SQL)
        If Atb1.RecordCount <> 0 Then
            QCIEPrg03.Maximum = Atb1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                Atb1.MoveFirst()
                Do While Not Atb1.EOF

                    QCIEPrg03.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(Atb1("Process_ID").Value), "", Atb1("Process_ID").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Date").Value), "", Atb1("Date").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Time").Value), "", Atb1("Time").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Nik").Value), "", Atb1("Nik").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Pieces").Value), "", Atb1("Pieces").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Target").Value), "", Atb1("Target").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Salary").Value), "", Atb1("Salary").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Coupon").Value), "", Atb1("Coupon").Value)))
                    Atb1.MoveNext()
                Loop
            End Using

        End If

    End Sub

    Sub MutuIIExport()

        PartExport = "QCMUTU2INFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(path) Then
            File.Delete(path)
        End If

        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Encrypt("PID;DATE;TIME;NIK;COUPON;PIECES;TARGET;SALARY"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From 04_mutuii_table "
        If QCIEChkBoxDate01.Checked = True Then
            SQL = SQL & "Where Date between ('" & ExportDateFrom.ToString("yyyy-MM-dd") & "') "
            SQL = SQL & "And ('" & ExportDateUntil.ToString("yyyy-MM-dd") & "') "
        End If
        OpenTbl(ADb, Atb1, SQL)
        If Atb1.RecordCount <> 0 Then
            QCIEPrg04.Maximum = Atb1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                Atb1.MoveFirst()
                Do While Not Atb1.EOF

                    QCIEPrg04.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(Atb1("Process_ID").Value), "", Atb1("Process_ID").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Date").Value), "", Atb1("Date").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Time").Value), "", Atb1("Time").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Nik").Value), "", Atb1("Nik").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Coupon").Value), "", Atb1("Coupon").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Pieces").Value), "", Atb1("Pieces").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Target").Value), "", Atb1("Target").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Salary").Value), "", Atb1("Salary").Value)))

                    Atb1.MoveNext()
                Loop
            End Using

        End If

    End Sub

    Sub PackingExport()
        PartExport = "QCPACKINGINFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(path) Then
            File.Delete(path)
        End If

        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Encrypt("PID;DATE;TIME;NIK;COUPON;CARTON;TARGET;SALARY"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From 05_packing_table "
        If QCIEChkBoxDate01.Checked = True Then
            SQL = SQL & "Where Date between ('" & ExportDateFrom.ToString("yyyy-MM-dd") & "') "
            SQL = SQL & "And ('" & ExportDateUntil.ToString("yyyy-MM-dd") & "') "
        End If
        OpenTbl(ADb, Atb1, SQL)
        If Atb1.RecordCount <> 0 Then
            QCIEPrg05.Maximum = Atb1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                Atb1.MoveFirst()
                Do While Not Atb1.EOF

                    QCIEPrg05.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(Atb1("Process_ID").Value), "", Atb1("Process_ID").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Date").Value), "", Atb1("Date").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Time").Value), "", Atb1("Time").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Nik").Value), "", Atb1("Nik").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Coupon").Value), "", Atb1("Coupon").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Carton").Value), "", Atb1("Carton").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Target").Value), "", Atb1("Target").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Salary").Value), "", Atb1("Salary").Value)))

                    Atb1.MoveNext()
                Loop
            End Using

        End If
    End Sub

    Sub WalletExport()
        PartExport = "QCWALLETINFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(path) Then
            File.Delete(path)
        End If

        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Encrypt("PID;DATE;TIME;NIK;COUPON;TARGET;PIECES;SALARY"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From 06_wallet_table "
        If QCIEChkBoxDate01.Checked = True Then
            SQL = SQL & "Where Date between ('" & ExportDateFrom.ToString("yyyy-MM-dd") & "') "
            SQL = SQL & "And ('" & ExportDateUntil.ToString("yyyy-MM-dd") & "') "
        End If
        OpenTbl(ADb, Atb1, SQL)
        If Atb1.RecordCount <> 0 Then
            QCIEPrg06.Maximum = Atb1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                Atb1.MoveFirst()
                Do While Not Atb1.EOF

                    QCIEPrg06.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(Atb1("Process_ID").Value), "", Atb1("Process_ID").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Date").Value), "", Atb1("Date").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Time").Value), "", Atb1("Time").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Nik").Value), "", Atb1("Nik").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Coupon").Value), "", Atb1("Coupon").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Target").Value), "", Atb1("Target").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Pieces").Value), "", Atb1("Pieces").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Salary").Value), "", Atb1("Salary").Value)))

                    Atb1.MoveNext()
                Loop
            End Using
        End If

    End Sub

    Sub StandardExport()
        PartExport = "QCSTANDARDINFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(path) Then
            File.Delete(path)
        End If

        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Encrypt("ORIGINAL;STANDARDWAGES"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From 08_standard_table "
        OpenTbl(ADb, Atb1, SQL)
        If Atb1.RecordCount <> 0 Then
            QCIEPrg07.Maximum = Atb1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                Atb1.MoveFirst()
                Do While Not Atb1.EOF

                    QCIEPrg07.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(Atb1("Original").Value), "", Atb1("Original").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Standard_Wage").Value), "", Atb1("Standard_Wage").Value)))

                    Atb1.MoveNext()
                Loop
            End Using
        End If

    End Sub

    Sub ConveSalaryExport()
        PartExport = "QCCONVSALARYINFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(path) Then
            File.Delete(path)
        End If

        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Encrypt("DATE;NIK;SALARY"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From 13_conveyour_salary "
        If QCIEChkBoxDate01.Checked = True Then
            SQL = SQL & "Where Date between ('" & ExportDateFrom.ToString("yyyy-MM-dd") & "') "
            SQL = SQL & "And ('" & ExportDateUntil.ToString("yyyy-MM-dd") & "') "
        End If
        OpenTbl(ADb, Atb1, SQL)
        If Atb1.RecordCount <> 0 Then
            QCIEPrg08.Maximum = Atb1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                Atb1.MoveFirst()
                Do While Not Atb1.EOF

                    QCIEPrg08.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(Atb1("Date").Value), "", Atb1("Date").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Nik").Value), "", Atb1("Nik").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Salary").Value), "", Atb1("Salary").Value)))

                    Atb1.MoveNext()
                Loop
            End Using
        End If

    End Sub

    Sub MutuSalaryExport()
        PartExport = "QCMUTU2SALARYINFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(path) Then
            File.Delete(path)
        End If

        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Encrypt("DATE;NIK;SALARY"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From 14_mutuii_salary "
        If QCIEChkBoxDate01.Checked = True Then
            SQL = SQL & "Where Date between ('" & ExportDateFrom.ToString("yyyy-MM-dd") & "') "
            SQL = SQL & "And ('" & ExportDateUntil.ToString("yyyy-MM-dd") & "') "
        End If
        OpenTbl(ADb, Atb1, SQL)
        If Atb1.RecordCount <> 0 Then
            QCIEPrg09.Maximum = Atb1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                Atb1.MoveFirst()
                Do While Not Atb1.EOF

                    QCIEPrg09.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(Atb1("Date").Value), "", Atb1("Date").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Nik").Value), "", Atb1("Nik").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Salary").Value), "", Atb1("Salary").Value)))

                    Atb1.MoveNext()
                Loop
            End Using
        End If
    End Sub

    Sub WalletSalaryExport()
        PartExport = "QCWALLETSALARYINFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(path) Then
            File.Delete(path)
        End If

        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Encrypt("DATE;NIK;SALARY"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From 15_wallet_salary "
        If QCIEChkBoxDate01.Checked = True Then
            SQL = SQL & "Where Date between ('" & ExportDateFrom.ToString("yyyy-MM-dd") & "') "
            SQL = SQL & "And ('" & ExportDateUntil.ToString("yyyy-MM-dd") & "') "
        End If
        OpenTbl(ADb, Atb1, SQL)
        If Atb1.RecordCount <> 0 Then
            QCIEPrg10.Maximum = Atb1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                Atb1.MoveFirst()
                Do While Not Atb1.EOF

                    QCIEPrg10.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(Atb1("Date").Value), "", Atb1("Date").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Nik").Value), "", Atb1("Nik").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Salary").Value), "", Atb1("Salary").Value)))

                    Atb1.MoveNext()
                Loop
            End Using
        End If
    End Sub

    Sub PackSalaryExport()
        PartExport = "QCPACKINGSALARYINFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(path) Then
            File.Delete(path)
        End If

        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Encrypt("DATE;NIK;SALARY"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From 16_packing_salary "
        If QCIEChkBoxDate01.Checked = True Then
            SQL = SQL & "Where Date between ('" & ExportDateFrom.ToString("yyyy-MM-dd") & "') "
            SQL = SQL & "And ('" & ExportDateUntil.ToString("yyyy-MM-dd") & "') "
        End If
        OpenTbl(ADb, Atb1, SQL)
        If Atb1.RecordCount <> 0 Then
            QCIEPrg11.Maximum = Atb1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                Atb1.MoveFirst()
                Do While Not Atb1.EOF

                    QCIEPrg11.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(Atb1("Date").Value), "", Atb1("Date").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Nik").Value), "", Atb1("Nik").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Salary").Value), "", Atb1("Salary").Value)))

                    Atb1.MoveNext()
                Loop
            End Using
        End If
    End Sub

    Sub HolidayExport()
        PartExport = "QCGHOLIDAYINFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(path) Then
            File.Delete(path)
        End If

        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Encrypt("DATE;HOLIDAYNAME;SALARYMOD"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From 17_holiday_table "
        OpenTbl(ADb, Atb1, SQL)
        If Atb1.RecordCount <> 0 Then
            QCIEPrg12.Maximum = Atb1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                Atb1.MoveFirst()
                Do While Not Atb1.EOF

                    QCIEPrg12.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(Atb1("Date").Value), "", Atb1("Date").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Holiday_Name").Value), "", Atb1("Holiday_Name").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Salary_Mod").Value), "", Atb1("Salary_Mod").Value)))

                    Atb1.MoveNext()
                Loop
            End Using
        End If
    End Sub

    Sub MiscExport()
        PartExport = "QCMISCLLINFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(path) Then
            File.Delete(path)
        End If

        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Encrypt("PID;DATE;TIME;NIK;PIECES;TARGET;SALARY;COUPON"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From 19_miscellaneous_table "
        If QCIEChkBoxDate01.Checked = True Then
            SQL = SQL & "Where Date between ('" & ExportDateFrom.ToString("yyyy-MM-dd") & "') "
            SQL = SQL & "And ('" & ExportDateUntil.ToString("yyyy-MM-dd") & "') "
        End If
        OpenTbl(ADb, Atb1, SQL)
        If Atb1.RecordCount <> 0 Then
            QCIEPrg13.Maximum = Atb1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                Atb1.MoveFirst()
                Do While Not Atb1.EOF

                    QCIEPrg13.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(Atb1("Process_ID").Value), "", Atb1("Process_ID").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Date").Value), "", Atb1("Date").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Time").Value), "", Atb1("Time").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Nik").Value), "", Atb1("Nik").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Pieces").Value), "", Atb1("Pieces").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Target").Value), "", Atb1("Target").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Salary").Value), "", Atb1("Salary").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Coupon").Value), "", Atb1("Coupon").Value)))

                    Atb1.MoveNext()
                Loop
            End Using
        End If
    End Sub

    Sub MiscSalaryExport()
        PartExport = "QCMISCLLSALARYINFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(path) Then
            File.Delete(path)
        End If

        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Encrypt("DATE;NIK;SALARY;TYPECTRL"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From 20_miscellaneous_salary "
        If QCIEChkBoxDate01.Checked = True Then
            SQL = SQL & "Where Date between ('" & ExportDateFrom.ToString("yyyy-MM-dd") & "') "
            SQL = SQL & "And ('" & ExportDateUntil.ToString("yyyy-MM-dd") & "') "
        End If
        OpenTbl(ADb, Atb1, SQL)
        If Atb1.RecordCount <> 0 Then
            QCIEPrg14.Maximum = Atb1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                Atb1.MoveFirst()
                Do While Not Atb1.EOF

                    QCIEPrg14.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(Atb1("Date").Value), "", Atb1("Date").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Nik").Value), "", Atb1("Nik").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Salary").Value), "", Atb1("Salary").Value) + ";" + _
                    IIf(IsDBNull(Atb1("TypeCtrl").Value), "", Atb1("TypeCtrl").Value)))

                    Atb1.MoveNext()
                Loop
            End Using
        End If
    End Sub

    Sub SortasiExport()
        PartExport = "QCSORTASIINFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(path) Then
            File.Delete(path)
        End If

        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Encrypt("PID;DATE;TIME;NIK;COUPON;NOKG;NOBAG;NOGR;PIECES;SALARY"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From 21_newmiscellaneous_table "
        If QCIEChkBoxDate01.Checked = True Then
            SQL = SQL & "Where Date between ('" & ExportDateFrom.ToString("yyyy-MM-dd") & "') "
            SQL = SQL & "And ('" & ExportDateUntil.ToString("yyyy-MM-dd") & "') "
        End If
        OpenTbl(ADb, Atb1, SQL)
        If Atb1.RecordCount <> 0 Then
            QCIEPrg15.Maximum = Atb1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                Atb1.MoveFirst()
                Do While Not Atb1.EOF

                    QCIEPrg15.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(Atb1("Process_ID").Value), "", Atb1("Process_ID").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Date").Value), "", Atb1("Date").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Time").Value), "", Atb1("Time").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Nik").Value), "", Atb1("Nik").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Coupon").Value), "", Atb1("Coupon").Value) + ";" + _
                    IIf(IsDBNull(Atb1("NoKg").Value), "", Atb1("NoKg").Value) + ";" + _
                    IIf(IsDBNull(Atb1("NoBag").Value), "", Atb1("NoBag").Value) + ";" + _
                    IIf(IsDBNull(Atb1("NoGr").Value), "", Atb1("NoGr").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Pieces").Value), "", Atb1("Pieces").Value) + ";" + _
                    IIf(IsDBNull(Atb1("Salary").Value), "", Atb1("Salary").Value)))

                    Atb1.MoveNext()
                Loop
            End Using
        End If
    End Sub

    Sub DateCounterExport()

        PartExport = "QCOUNTERINFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(path) Then
            File.Delete(path)
        End If

        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Encrypt("IDDATE;PERIODE;PERIODERANGE;PERIODEVALID;DATE1;DATE2;DATE3;DATE4;DATE5;DATE6;DATE7;DATE8;DATE9;DATE10;DATE11;DATE12;DATE13;DATE14;DATE15;DATE16"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From datecounter2table "
        OpenTbl(CBb, Ctbl1, SQL)
        If Ctbl1.RecordCount <> 0 Then
            QCIEPrg16.Maximum = Ctbl1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                Ctbl1.MoveFirst()
                Do While Not Ctbl1.EOF

                    QCIEPrg16.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(Ctbl1("IDDate").Value), "", Ctbl1("IDDate").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Periode").Value), "", Ctbl1("Periode").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("PeriodeRange").Value), "", Ctbl1("PeriodeRange").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("PeriodeValid").Value), "", Ctbl1("PeriodeValid").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date1").Value), "", Ctbl1("Date1").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date2").Value), "", Ctbl1("Date2").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date3").Value), "", Ctbl1("Date3").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date4").Value), "", Ctbl1("Date4").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date5").Value), "", Ctbl1("Date5").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date6").Value), "", Ctbl1("Date6").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date7").Value), "", Ctbl1("Date7").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date8").Value), "", Ctbl1("Date8").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date9").Value), "", Ctbl1("Date9").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date10").Value), "", Ctbl1("Date10").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date11").Value), "", Ctbl1("Date11").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date12").Value), "", Ctbl1("Date12").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date13").Value), "", Ctbl1("Date13").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date14").Value), "", Ctbl1("Date14").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date15").Value), "", Ctbl1("Date15").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date16").Value), "", Ctbl1("Date16").Value)))

                    Ctbl1.MoveNext()
                Loop
            End Using
        End If
    End Sub

    Sub PeriodeCounterExport()
        PartExport = "QCPERIODECOUNTERINFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(path) Then
            File.Delete(path)
        End If

        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Encrypt("DATE;COUNTER;PERIODE;PERIODERANGE"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From periode_countertable "
        OpenTbl(CBb, Ctbl1, SQL)
        If Ctbl1.RecordCount <> 0 Then
            QCIEPrg17.Maximum = Ctbl1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                Ctbl1.MoveFirst()
                Do While Not Ctbl1.EOF

                    QCIEPrg17.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(Ctbl1("Date").Value), "", Ctbl1("Date").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Counter").Value), "", Ctbl1("Counter").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Periode").Value), "", Ctbl1("Periode").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("PeriodeRange").Value), "", Ctbl1("PeriodeRange").Value)))

                    Ctbl1.MoveNext()
                Loop
            End Using
        End If
    End Sub

    Sub MainSalaryExport()

        PartExport = "QCMAINSALARYINFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(path) Then
            File.Delete(path)
        End If

        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Encrypt("PERIODE;PERIODERANGE;NIK;NAME;PAY;SALARY1;SALARY2;SALARY3;SALARY4;SAlARY5;SALARY6;SALARY7;SALARY8;SALARY9;SALARY10;SALARY11;SALARY12;SALARY13;SALARY14;SALARY15;SALARY16" & _
                                     ";DATE1;DATE2;DATE3;DATE4;DATE5;DATE6;DATE7;DATE8;DATE9;DATE10;DATE11;DATE12;DATE13;DATE14;DATE15;DATE16;ASTEKVAL;PNOREK;POTLAIN"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From salarysync1_table "
        If QCIEChkBoxDate01.Checked = True Then
            SQL = SQL & "Where PeriodeRange = ('" & AddStringforPath & "') "
        End If
        OpenTbl(CBb, Ctbl1, SQL)
        If Ctbl1.RecordCount <> 0 Then
            QCIEPrg18.Maximum = Ctbl1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                Ctbl1.MoveFirst()
                Do While Not Ctbl1.EOF

                    QCIEPrg18.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(Ctbl1("Periode").Value), "", Ctbl1("Periode").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("PeriodeRange").Value), "", Ctbl1("PeriodeRange").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Nik").Value), "", Ctbl1("Nik").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Name").Value), "", Ctbl1("Name").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Pay").Value), "", Ctbl1("Pay").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Salary1").Value), "", Ctbl1("Salary1").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Salary2").Value), "", Ctbl1("Salary2").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Salary3").Value), "", Ctbl1("Salary3").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Salary4").Value), "", Ctbl1("Salary4").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Salary5").Value), "", Ctbl1("Salary5").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Salary6").Value), "", Ctbl1("Salary6").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Salary7").Value), "", Ctbl1("Salary7").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Salary8").Value), "", Ctbl1("Salary8").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Salary9").Value), "", Ctbl1("Salary9").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Salary10").Value), "", Ctbl1("Salary10").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Salary11").Value), "", Ctbl1("Salary11").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Salary12").Value), "", Ctbl1("Salary12").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Salary13").Value), "", Ctbl1("Salary13").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Salary14").Value), "", Ctbl1("Salary14").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Salary15").Value), "", Ctbl1("Salary15").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Salary16").Value), "", Ctbl1("Salary16").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date1").Value), "", Ctbl1("Date1").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date2").Value), "", Ctbl1("Date2").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date3").Value), "", Ctbl1("Date3").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date4").Value), "", Ctbl1("Date4").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date5").Value), "", Ctbl1("Date5").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date6").Value), "", Ctbl1("Date6").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date7").Value), "", Ctbl1("Date7").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date8").Value), "", Ctbl1("Date8").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date9").Value), "", Ctbl1("Date9").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date10").Value), "", Ctbl1("Date10").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date11").Value), "", Ctbl1("Date11").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date12").Value), "", Ctbl1("Date12").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date13").Value), "", Ctbl1("Date13").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date14").Value), "", Ctbl1("Date14").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date15").Value), "", Ctbl1("Date15").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("Date16").Value), "", Ctbl1("Date16").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("AstekVal").Value), "", Ctbl1("AstekVal").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("PNoRek").Value), "", Ctbl1("PNoRek").Value) + ";" + _
                    IIf(IsDBNull(Ctbl1("PotLain").Value), "", Ctbl1("PotLain").Value)))

                    Ctbl1.MoveNext()
                Loop
            End Using
        End If

    End Sub

    Sub PPH21Export()
        PartExport = "QCPPH21INFO "
        path = Application.StartupPath + "\CSVFolder\" + PartExport + Today.ToString("ddMMMyyyy") + ".csv"

        If File.Exists(path) Then
            File.Delete(path)
        End If

        If Not File.Exists(path) Then
            ' Create a file to write to. 
            Using sw As StreamWriter = File.CreateText(path)
                sw.WriteLine(Encrypt("PERIODEGAJIAN;PERIODE;PERIODERANGE;NIK;NAME;EMADD;PAY;ASTEK;NPWP;KTP;MAINSALARY1;MAINSALARY2;MAINSALARY3;INCENTIF"))
            End Using

        End If

        SQL = Nothing
        SQL = SQL & "Select * From emp_pphtable "
        OpenTbl(PPhDB, PPhTb1, SQL)
        If PPhTb1.RecordCount <> 0 Then
            QCIEPrg19.Maximum = PPhTb1.RecordCount
            Using sw As StreamWriter = File.AppendText(path)
                PPhTb1.MoveFirst()
                Do While Not PPhTb1.EOF

                    QCIEPrg19.Value += 1
                    sw.WriteLine(Encrypt(IIf(IsDBNull(PPhTb1("PeriodeGajian").Value), "", PPhTb1("PeriodeGajian").Value) + ";" + _
                    IIf(IsDBNull(PPhTb1("Periode").Value), "", PPhTb1("Periode").Value) + ";" + _
                    IIf(IsDBNull(PPhTb1("PeriodeRange").Value), "", PPhTb1("PeriodeRange").Value) + ";" + _
                    IIf(IsDBNull(PPhTb1("Nik").Value), "", PPhTb1("Nik").Value) + ";" + _
                    IIf(IsDBNull(PPhTb1("Name").Value), "", PPhTb1("Name").Value) + ";" + _
                    IIf(IsDBNull(PPhTb1("EmAdd").Value), "", PPhTb1("EmAdd").Value) + ";" + _
                    IIf(IsDBNull(PPhTb1("Pay").Value), "", PPhTb1("Pay").Value) + ";" + _
                    IIf(IsDBNull(PPhTb1("Astek").Value), "", PPhTb1("Astek").Value) + ";" + _
                    IIf(IsDBNull(PPhTb1("NPWP").Value), "", PPhTb1("NPWP").Value) + ";" + _
                    IIf(IsDBNull(PPhTb1("KTP").Value), "", PPhTb1("KTP").Value) + ";" + _
                    IIf(IsDBNull(PPhTb1("MainSalary1").Value), "", PPhTb1("MainSalary1").Value) + ";" + _
                    IIf(IsDBNull(PPhTb1("MainSalary2").Value), "", PPhTb1("MainSalary2").Value) + ";" + _
                    IIf(IsDBNull(PPhTb1("MainSalary3").Value), "", PPhTb1("MainSalary3").Value) + ";" + _
                    IIf(IsDBNull(PPhTb1("Incentif").Value), "", PPhTb1("Incentif").Value)))

                    PPhTb1.MoveNext()
                Loop
            End Using
        End If
    End Sub

    Sub ResetProgBar(ByVal TabCall As TabPage)
        For Each Control In TabCall.Controls
            If TypeOf Control Is ProgressBar Then
                Control.Value = Nothing
                Control.Maximum = Nothing
            End If
        Next Control
    End Sub

    Private Sub QCIEBtn01_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QCIEBtn01.Click

        ExportDateFrom = QCImpExpDP01.Text
        ExportDateUntil = QCImpExpDP02.Text
        AddStringforPath = QCImpExpDP03.Text
        ResetProgBar(TabPage1)
        OnClickTheWorker()

    End Sub

#Region "BackGroundWorker on Dynamic"

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
        If QCIEChBox01.Checked = True Then
            AccountExport()
        End If
        If QCIEChBox02.Checked = True Then
            NameExport()
        End If

        If QCIEChBox03.Checked = True Then
            ConveyourExport()
        End If

        If QCIEChBox04.Checked = True Then
            MutuIIExport()
        End If

        If QCIEChBox05.Checked = True Then
            PackingExport()
        End If

        If QCIEChBox06.Checked = True Then
            WalletExport()
        End If

        If QCIEChBox07.Checked = True Then
            StandardExport()
        End If

        If QCIEChBox08.Checked = True Then
            ConveSalaryExport()
        End If

        If QCIEChBox09.Checked = True Then
            MutuSalaryExport()
        End If

        If QCIEChBox10.Checked = True Then
            WalletSalaryExport()
        End If

        If QCIEChBox11.Checked = True Then
            PackSalaryExport()
        End If

        If QCIEChBox12.Checked = True Then
            HolidayExport()
        End If

        If QCIEChBox13.Checked = True Then
            MiscExport()
        End If

        If QCIEChBox14.Checked = True Then
            MiscSalaryExport()
        End If

        If QCIEChBox15.Checked = True Then
            SortasiExport()
        End If


        If QCIEChBox16.Checked = True Then
            DateCounterExport()
        End If


        If QCIEChBox17.Checked = True Then
            PeriodeCounterExport()
        End If


        If QCIEChBox18.Checked = True Then
            MainSalaryExport()
        End If


        If QCIEChBox19.Checked = True Then
            PPH21Export()
        End If
    End Sub

    Private Sub WorkerProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs)

    End Sub

    Private Sub WorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)

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
        GetCSVItem()
    End Sub

    Private Sub WorkerProgressChanged2(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs)

    End Sub

    Private Sub WorkerCompleted2(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)

    End Sub

    Sub OnClickTheWorker3()

        i += 1
        ReDim BGWorkMode(i)
        BGWorkMode(i) = New BackgroundWorker
        BGWorkMode(i).WorkerReportsProgress = True
        BGWorkMode(i).WorkerSupportsCancellation = True
        AddHandler BGWorkMode(i).DoWork, AddressOf WorkerDoWork3
        AddHandler BGWorkMode(i).ProgressChanged, AddressOf WorkerProgressChanged3
        AddHandler BGWorkMode(i).RunWorkerCompleted, AddressOf WorkerCompleted3
        BGWorkMode(i).RunWorkerAsync()


    End Sub

    Private Sub WorkerDoWork3(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs)
        Select Case QCIECombox01.Text

            Case "Account Info"
                SaveAccountInfo()
            Case "Employee Info"
                SaveEmploInfo()
            Case "Conveyour Info"
                SaveConvInfo()
            Case "Mutu II Info"
                SaveMutInfo()
            Case "Packing Info"
                SavePackInfo()
            Case "Wallet Info"
                SaveWallInfo()
            Case "Standard Info"
                SaveStanInfo()
            Case "Conveyour Salary Info"
                SaveConvSal()
            Case "Mutu II Salary Info"
                SaveMutSal()
            Case "Packing Salary Info"
                SavePackSal()
            Case "Wallet Salary Info"
                SaveWalSal()
            Case "Misc / Sortasi Salary Info"
                SaveMiscSortSal()
            Case "Holiday"
                SaveHolInfo()
            Case "Misc Info"
                SaveMiscInfo()
            Case "Sortasi Info"
                SaveSortInfo()
            Case "Date Counter"
                SaveDateCount()
            Case "Periode Counter"
                SavePeriCount()
            Case "Main Salary"
                SaveMainSal()
            Case "PPH 21"
                SavePPh21Sal()

        End Select


    End Sub

    Private Sub WorkerProgressChanged3(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs)

    End Sub

    Private Sub WorkerCompleted3(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)
        If e.Error IsNot Nothing Then
            MessageBox.Show(e.Error, Me.Text)
        Else
            RCountLbl01.Invoke(DirectCast(Sub() RCountLbl01.Text = "Done", MethodInvoker))
        End If

    End Sub
#End Region

    Private Sub QCIEBtn02_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QCIEBtn02.Click
        Process.Start("explorer.exe", Application.StartupPath + "\CSVFolder")
    End Sub

#Region "Import CSV File"

    Sub QCGridHeader()

        With QCIEGrid01
            .Rows.Clear()
            .Columns.Clear()

            Select Case QCIECombox01.Text

                Case "Account Info"

                    .Columns.Add("col0", "User Number")
                    .Columns.Add("col1", "User Name")
                    .Columns.Add("col2", "User Password")
                    .Columns.Add("col3", "User Level")
                    .Columns.Add("col4", "User Account Name")
                    .Columns.Add("col5", "User Field Code")
                    .Columns.Add("col6", "Status")

                    For i = 0 To 6

                        .Columns(i).width = 150

                    Next

                Case "Employee Info"

                    .Columns.Add("col0", "Emp Number")
                    .Columns.Add("col1", "Nik")
                    .Columns.Add("col2", "Name")
                    .Columns.Add("col3", "Active")
                    .Columns.Add("col4", "Date Start")
                    .Columns.Add("col5", "Pay")
                    .Columns.Add("col6", "Astek")
                    .Columns.Add("col7", "NPWP")
                    .Columns.Add("col8", "No Rek")
                    .Columns.Add("col9", "No KTP")
                    .Columns.Add("col10", "Jab Data")
                    .Columns.Add("col11", "No KPJ")
                    .Columns.Add("col12", "Bank CTRL")
                    .Columns.Add("col13", "ESTATE")
                    .Columns.Add("col14", "Lahir")
                    .Columns.Add("col15", "Agama")
                    .Columns.Add("col16", "Alamat")
                    .Columns.Add("col17", "Tel Num")
                    .Columns.Add("col18", "Pendidikan")
                    .Columns.Add("col19", "Dept")
                    .Columns.Add("col20", "JKKJKM")
                    .Columns.Add("col21", "STATUS")

                    For i = 0 To 21
                        .Columns(i).width = 200

                    Next


                Case "Conveyour Info"

                    .Columns.Add("col0", "Process ID")
                    .Columns.Add("col1", "Date")
                    .Columns.Add("col2", "Time")
                    .Columns.Add("col3", "Nik")
                    .Columns.Add("col4", "Pieces")
                    .Columns.Add("col5", "Target")
                    .Columns.Add("col6", "Salary")
                    .Columns.Add("col7", "Coupon")
                    .Columns.Add("col8", "Status")

                    For i = 0 To 8
                        .Columns(i).width = 200

                    Next

                Case "Mutu II Info"

                    .Columns.Add("col0", "Process ID")
                    .Columns.Add("col1", "Date")
                    .Columns.Add("col2", "Time")
                    .Columns.Add("col3", "Nik")
                    .Columns.Add("col4", "Coupon")
                    .Columns.Add("col5", "Pieces")
                    .Columns.Add("col6", "Target")
                    .Columns.Add("col7", "Salary")
                    .Columns.Add("col8", "Status")

                    For i = 0 To 8
                        .Columns(i).width = 200

                    Next

                Case "Packing Info"


                    .Columns.Add("col0", "Process ID")
                    .Columns.Add("col1", "Date")
                    .Columns.Add("col2", "Time")
                    .Columns.Add("col3", "Nik")
                    .Columns.Add("col4", "Coupon")
                    .Columns.Add("col5", "Carton")
                    .Columns.Add("col6", "Target")
                    .Columns.Add("col7", "Salary")
                    .Columns.Add("col8", "Status")

                    For i = 0 To 8
                        .Columns(i).width = 200

                    Next

                Case "Wallet Info"

                    .Columns.Add("col0", "Process ID")
                    .Columns.Add("col1", "Date")
                    .Columns.Add("col2", "Time")
                    .Columns.Add("col3", "Nik")
                    .Columns.Add("col4", "Coupon")
                    .Columns.Add("col5", "Target")
                    .Columns.Add("col6", "Pieces")
                    .Columns.Add("col7", "Salary")
                    .Columns.Add("col8", "Status")

                    For i = 0 To 8
                        .Columns(i).width = 200

                    Next

                Case "Standard Info"

                    .Columns.Add("col0", "Original")
                    .Columns.Add("col1", "Standard Wages")
                    .Columns.Add("col2", "Status")

                    For i = 0 To 2
                        .Columns(i).width = 200
                    Next

                Case "Conveyour Salary Info", "Mutu II Salary Info", "Packing Salary Info", "Wallet Salary Info"

                    .Columns.Add("col0", "Date")
                    .Columns.Add("col1", "Nik")
                    .Columns.Add("col2", "Salary")
                    .Columns.Add("col3", "Status")

                    For i = 0 To 3
                        .Columns(i).width = 200
                    Next

                Case "Misc / Sortasi Salary Info"
                    .Columns.Add("col0", "Date")
                    .Columns.Add("col1", "Nik")
                    .Columns.Add("col2", "Salary")
                    .Columns.Add("col3", "Type Ctrl")
                    .Columns.Add("col4", "Status")

                    For i = 0 To 4
                        .Columns(i).width = 200
                    Next

                Case "Holiday"

                    .Columns.Add("col0", "Date")
                    .Columns.Add("col1", "Holiday Name")
                    .Columns.Add("col2", "Salary Mod")
                    .Columns.Add("col3", "Status")

                    For i = 0 To 3
                        .Columns(i).width = 200
                    Next

                Case "Misc Info"

                    .Columns.Add("col0", "Process Number")
                    .Columns.Add("col1", "Date")
                    .Columns.Add("col2", "Time")
                    .Columns.Add("col3", "Nik")
                    .Columns.Add("col4", "Pieces")
                    .Columns.Add("col5", "Target")
                    .Columns.Add("col6", "Salary")
                    .Columns.Add("col7", "Coupon")
                    .Columns.Add("col8", "Status")

                    For i = 0 To 8
                        .Columns(i).width = 200
                    Next

                Case "Sortasi Info"

                    .Columns.Add("col0", "Process Number")
                    .Columns.Add("col1", "Date")
                    .Columns.Add("col2", "Time")
                    .Columns.Add("col3", "Nik")
                    .Columns.Add("col4", "Coupon")
                    .Columns.Add("col5", "No KG")
                    .Columns.Add("col6", "No Bag")
                    .Columns.Add("col7", "No Gr")
                    .Columns.Add("col8", "Pieces")
                    .Columns.Add("col9", "Salary")
                    .Columns.Add("col10", "Status")

                    For i = 0 To 10
                        .Columns(i).width = 200
                    Next

                Case "Date Counter"


                    .Columns.Add("col0", "ID Date")
                    .Columns.Add("col1", "Periode")
                    .Columns.Add("col2", "PeriodeRange")
                    .Columns.Add("col3", "Periode Valid")
                    .Columns.Add("col4", "Date 1")
                    .Columns.Add("col5", "Date 2")
                    .Columns.Add("col6", "Date 3")
                    .Columns.Add("col7", "Date 4")
                    .Columns.Add("col8", "Date 5")
                    .Columns.Add("col9", "Date 6")
                    .Columns.Add("col10", "Date 7")
                    .Columns.Add("col11", "Date 8")
                    .Columns.Add("col12", "Date 9")
                    .Columns.Add("col13", "Date 10")
                    .Columns.Add("col14", "Date 11")
                    .Columns.Add("col15", "Date 12")
                    .Columns.Add("col16", "Date 13")
                    .Columns.Add("col17", "Date 14")
                    .Columns.Add("col18", "Date 15")
                    .Columns.Add("col19", "Date 16")
                    .Columns.Add("col20", "Status")

                    For i = 0 To 20
                        .Columns(i).width = 200
                    Next
                Case "Periode Counter"

                    .Columns.Add("col0", "Date")
                    .Columns.Add("col1", "Counter")
                    .Columns.Add("col2", "Periode")
                    .Columns.Add("col3", "Periode Range")
                    .Columns.Add("col4", "Status")
          
                    For i = 0 To 4
                        .Columns(i).width = 200
                    Next

                Case "Main Salary"

                    .Columns.Add("col0", "Periode")
                    .Columns.Add("col1", "PeriodeRange")
                    .Columns.Add("col2", "Nik")
                    .Columns.Add("col3", "Name")
                    .Columns.Add("col4", "Pay")
                    .Columns.Add("col5", "Salary 1")
                    .Columns.Add("col6", "Salary 2")
                    .Columns.Add("col7", "Salary 3")
                    .Columns.Add("col8", "Salary 4")
                    .Columns.Add("col9", "Salary 5")
                    .Columns.Add("col10", "Salary 6")
                    .Columns.Add("col11", "Salary 7")
                    .Columns.Add("col12", "Salary 8")
                    .Columns.Add("col13", "Salary 9")
                    .Columns.Add("col14", "Salary 10")
                    .Columns.Add("col15", "Salary 11")
                    .Columns.Add("col16", "Salary 12")
                    .Columns.Add("col17", "Salary 13")
                    .Columns.Add("col18", "Salary 14")
                    .Columns.Add("col19", "Salarye 15")
                    .Columns.Add("col20", "Salary 16")

                    .Columns.Add("col21", "Date 1")
                    .Columns.Add("col22", "Date 2")
                    .Columns.Add("col23", "Date 3")
                    .Columns.Add("col24", "Date 4")
                    .Columns.Add("col25", "Date 5")
                    .Columns.Add("col26", "Date 6")
                    .Columns.Add("col27", "Date 7")
                    .Columns.Add("col28", "Date 8")
                    .Columns.Add("col29", "Date 9")
                    .Columns.Add("col30", "Date 10")
                    .Columns.Add("col31", "Date 11")
                    .Columns.Add("col32", "Date 12")
                    .Columns.Add("col33", "Date 13")
                    .Columns.Add("col34", "Date 14")
                    .Columns.Add("col35", "Date 15")
                    .Columns.Add("col36", "Date 16")

                    .Columns.Add("col37", "Astek")
                    .Columns.Add("col38", "No. Rek")
                    .Columns.Add("col39", "Pot. Lain")
                    .Columns.Add("col40", "Status")

                    For i = 0 To 40
                        .Columns(i).width = 200
                    Next

                Case "PPH 21"

                    .Columns.Add("col0", "Periode Gajian")
                    .Columns.Add("col1", "Periode")
                    .Columns.Add("col2", "Periode Range")
                    .Columns.Add("col3", "Nik")
                    .Columns.Add("col4", "Name")
                    .Columns.Add("col5", "EMADD")
                    .Columns.Add("col6", "PAY")
                    .Columns.Add("col7", "ASTEK")
                    .Columns.Add("col8", "NPWP")
                    .Columns.Add("col9", "KTP")
                    .Columns.Add("col10", "MAIN SALARY 1")
                    .Columns.Add("col11", "MAIN SALARY 2")
                    .Columns.Add("col12", "MAIN SALARY 3")
                    .Columns.Add("col13", "INCENTIF")
                    .Columns.Add("col14", "Status")

                    For i = 0 To 14
                        .Columns(i).width = 200
                    Next

            End Select

        End With

    End Sub
    Dim TextLineRead As String = Nothing
    Dim SplitLine() As String
    Dim LinkName As String
    Sub GetCSVItem()
        Dim countme As Integer = 0

        If System.IO.File.Exists(LinkName) = True Then

            Dim objReader As New System.IO.StreamReader(LinkName)

            Do While objReader.Peek() <> -1

                TextLineRead = objReader.ReadLine()
                If QCExImRB01.Checked = True Then
                    SplitLine = Split(Decrypt(TextLineRead), ";")
                ElseIf QCExImRB02.Checked = True Then
                    SplitLine = Split(TextLineRead, ";")
                End If

                QCIEGrid01.Invoke(DirectCast(Sub() QCIEGrid01.Rows.Add(SplitLine), MethodInvoker))

            Loop
            QCIEGrid01.Rows.RemoveAt(0)

            objReader.Close()
            objReader = Nothing

        End If

   

    End Sub

    Sub OpenFileLink()
        Dim LinkFileName As New OpenFileDialog
        LinkFileName.InitialDirectory = Application.StartupPath + "\CSVFolder\"
        LinkFileName.Filter = " CSV File [.csv] |*.csv| All Files |*.*"
        If LinkFileName.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            LinkName = LinkFileName.FileName
            QCIETbx01.Text = LinkFileName.FileName
            QCIECombox01.Enabled = True
        End If
    End Sub

    Sub SaveAccountInfo()

        For i = 0 To QCIEGrid01.Rows.Count - 1

            SQL = ""
            SQL = SQL & "Select * From 01_Account_Table "
            SQL = SQL & "Where Username = ('" & QCIEGrid01(1, i).Value & "') "

            OpenTbl(ADb, Atb1, SQL)

            If Not Atb1.RecordCount <> 0 Then
                Atb1.AddNew()
            End If


            Atb1("Username").Value = QCIEGrid01(1, i).Value
            Atb1("Userpass").Value = QCIEGrid01(2, i).Value
            Atb1("UserAccName").Value = QCIEGrid01(4, i).Value
            Atb1("Userlevel").Value = QCIEGrid01(3, i).Value
            Atb1("UserNumber").Value = QCIEGrid01(0, i).Value
            Atb1("UserFieldCode").Value = QCIEGrid01(5, i).Value


            Atb1.Update()
            QCIECsvProg01.Value += 1
        Next
    End Sub
    Dim SetDate As Date
    Dim ReStringName As String

    Sub SaveEmploInfo()

        For i = 0 To QCIEGrid01.Rows.Count - 1

            ReStringName = QCIEGrid01(2, i).Value
            ReStringName = ReStringName.Replace("'", "?")
            SQL = ""
            SQL = SQL & "Select * From 02_Name_Table "
            SQL = SQL & "Where Nik = ('" & QCIEGrid01(1, i).Value & "') "
            SQL = SQL & "And Name = ('" & ReStringName & "') "
            OpenTbl(ADb, Atb1, SQL)

            If Not Atb1.RecordCount <> 0 Then
                Atb1.AddNew()
            End If


            SetDate = QCIEGrid01(4, i).Value
            Atb1("ID_Number").Value = QCIEGrid01(0, i).Value
            Atb1("Nik").Value = QCIEGrid01(1, i).Value
            Atb1("Name").Value = ReStringName
            Atb1("DateStart").Value = SetDate.ToString("yyyy-MM-dd")
            Atb1("Active").Value = QCIEGrid01(3, i).Value
            Atb1("Jamsostek").Value = QCIEGrid01(6, i).Value
            Atb1("Pay").Value = QCIEGrid01(5, i).Value

            Atb1("Bank_Ctrl").Value = QCIEGrid01(12, i).Value
            Atb1("Jamsostek").Value = QCIEGrid01(6, i).Value
            Atb1("NPWP").Value = QCIEGrid01(7, i).Value
            Atb1("NoRek").Value = QCIEGrid01(8, i).Value
            Atb1("NKTP").Value = QCIEGrid01(9, i).Value
            Atb1("NoKPJ").Value = QCIEGrid01(11, i).Value
            Atb1("Lahir").Value = QCIEGrid01(14, i).Value
            Atb1("JabData").Value = QCIEGrid01(10, i).Value
            Atb1("Estate").Value = QCIEGrid01(13, i).Value
            Atb1("Agama").Value = QCIEGrid01(15, i).Value
            Atb1("Alamat").Value = QCIEGrid01(16, i).Value
            Atb1("TelNum").Value = QCIEGrid01(17, i).Value
            Atb1("Pendi").Value = QCIEGrid01(18, i).Value
            Atb1("Dept").Value = QCIEGrid01(19, i).Value
            QCIEGrid01.Invoke(DirectCast(Sub() QCIEGrid01(20, i).Value = "Has been Saved", MethodInvoker))
            Atb1.Update()
            QCIECsvProg01.Value += 1
        Next

    End Sub

    Sub SaveConvInfo()

        For i = 0 To QCIEGrid01.Rows.Count - 1

            SQL = ""
            SQL = SQL & "Select * From 03_Conveyour_Table "
            SQL = SQL & "Where Nik = ('" & QCIEGrid01(3, i).Value & "') "
            SQL = SQL & "And Process_ID = ('" & QCIEGrid01(0, i).Value & "') "
            OpenTbl(ADb, Atb1, SQL)

            If Not Atb1.RecordCount <> 0 Then
                Atb1.AddNew()
            End If

            SetDate = QCIEGrid01(1, i).Value
            Atb1("Time").Value = QCIEGrid01(2, i).Value
            Atb1("Nik").Value = QCIEGrid01(3, i).Value
            Atb1("Target").Value = QCIEGrid01(5, i).Value
            Atb1("Pieces").Value = QCIEGrid01(4, i).Value
            Atb1("Salary").Value = QCIEGrid01(6, i).Value
            Atb1("Date").Value = SetDate.ToString("yyyy-MM-dd")
            Atb1("Process_ID").Value = QCIEGrid01(0, i).Value
            QCIEGrid01.Invoke(DirectCast(Sub() QCIEGrid01(8, i).Value = "Has been Saved", MethodInvoker))


            Atb1.Update()
            QCIECsvProg01.Value += 1
        Next

    End Sub

    Sub SaveMutInfo()

        For i = 0 To QCIEGrid01.Rows.Count - 1

            SQL = ""
            SQL = SQL & "Select * From 04_MutuII_Table "
            SQL = SQL & "Where Nik = ('" & QCIEGrid01(3, i).Value & "') "
            SQL = SQL & "And Process_ID = ('" & QCIEGrid01(0, i).Value & "') "
            OpenTbl(ADb, Atb1, SQL)

            If Not Atb1.RecordCount <> 0 Then
                Atb1.AddNew()
            End If
            SetDate = QCIEGrid01(1, i).Value
            Atb1("Time").Value = QCIEGrid01(2, i).Value
            Atb1("Nik").Value = QCIEGrid01(3, i).Value
            Atb1("Target").Value = QCIEGrid01(6, i).Value
            Atb1("Pieces").Value = QCIEGrid01(5, i).Value
            Atb1("Salary").Value = QCIEGrid01(7, i).Value
            Atb1("Date").Value = SetDate.ToString("yyyy-MM-dd")
            Atb1("Process_ID").Value = QCIEGrid01(0, i).Value
            Atb1("Coupon").Value = QCIEGrid01(4, i).Value
            QCIEGrid01.Invoke(DirectCast(Sub() QCIEGrid01(8, i).Value = "Has been Saved", MethodInvoker))
            Atb1.Update()

        Next

    End Sub

    Sub SavePackInfo()

        For i = 0 To QCIEGrid01.Rows.Count - 1

            SQL = ""
            SQL = SQL & "Select * From 05_Packing_Table "
            SQL = SQL & "Where Nik = ('" & QCIEGrid01(3, i).Value & "') "
            SQL = SQL & "And Process_ID = ('" & QCIEGrid01(0, i).Value & "') "
            OpenTbl(ADb, Atb1, SQL)

            If Not Atb1.RecordCount <> 0 Then
                Atb1.AddNew()
            End If
            SetDate = QCIEGrid01(1, i).Value
            Atb1("Time").Value = QCIEGrid01(2, i).Value
            Atb1("Nik").Value = QCIEGrid01(3, i).Value
            Atb1("Target").Value = QCIEGrid01(6, i).Value
            Atb1("Salary").Value = QCIEGrid01(7, i).Value
            Atb1("Date").Value = SetDate.ToString("yyyy-MM-dd")
            Atb1("Process_ID").Value = QCIEGrid01(0, i).Value
            Atb1("Carton").Value = QCIEGrid01(5, i).Value
            Atb1("Coupon").Value = QCIEGrid01(4, i).Value
            Atb1.Update()
            QCIEGrid01.Invoke(DirectCast(Sub() QCIEGrid01(8, i).Value = "Has been Saved", MethodInvoker))
            QCIECsvProg01.Value += 1
        Next
    End Sub

    Sub SaveWallInfo()

        For i = 0 To QCIEGrid01.Rows.Count - 1

            SQL = ""
            SQL = SQL & "Select * From 06_Wallet_Table "
            SQL = SQL & "Where Nik = ('" & QCIEGrid01(3, i).Value & "') "
            SQL = SQL & "And Process_ID = ('" & QCIEGrid01(0, i).Value & "') "
            OpenTbl(ADb, Atb1, SQL)

            If Not Atb1.RecordCount <> 0 Then
                Atb1.AddNew()
            End If

            SetDate = QCIEGrid01(1, i).Value
            Atb1("Time").Value = QCIEGrid01(1, i).Value
            Atb1("Nik").Value = QCIEGrid01(3, i).Value
            Atb1("Target").Value = QCIEGrid01(5, i).Value
            Atb1("Pieces").Value = QCIEGrid01(6, i).Value
            Atb1("Salary").Value = QCIEGrid01(7, i).Value
            Atb1("Date").Value = SetDate
            Atb1("Process_ID").Value = QCIEGrid01(0, i).Value
            Atb1("Coupon").Value = QCIEGrid01(4, i).Value

            Atb1.Update()
            QCIEGrid01.Invoke(DirectCast(Sub() QCIEGrid01(8, i).Value = "Has been Saved", MethodInvoker))
            QCIECsvProg01.Value += 1
        Next

    End Sub

    Sub SaveMiscInfo()
        For i = 0 To QCIEGrid01.Rows.Count - 1

            SQL = ""
            SQL = SQL & "Select * From 19_miscellaneous_table "
            SQL = SQL & "Where Nik = ('" & QCIEGrid01(3, i).Value & "') "
            SQL = SQL & "And Process_ID = ('" & QCIEGrid01(0, i).Value & "') "
            OpenTbl(ADb, Atb1, SQL)

            If Not Atb1.RecordCount <> 0 Then
                Atb1.AddNew()
            End If
            SetDate = QCIEGrid01(1, i).Value()
            Atb1("Time").Value = QCIEGrid01(2, i).Value
            Atb1("Nik").Value = QCIEGrid01(3, i).Value
            Atb1("Salary").Value = QCIEGrid01(6, i).Value
            Atb1("Date").Value = SetDate.ToString("yyyy-MM-dd")
            Atb1("Process_ID").Value = QCIEGrid01(0, i).Value

            Atb1.Update()
            QCIEGrid01.Invoke(DirectCast(Sub() QCIEGrid01(8, i).Value = "Has been Saved", MethodInvoker))
            QCIECsvProg01.Value += 1
        Next

    End Sub

    Sub SaveSortInfo()

        For i = 0 To QCIEGrid01.Rows.Count - 1

            SQL = ""
            SQL = SQL & "Select * From 21_newmiscellaneous_table "
            SQL = SQL & "Where Nik = ('" & QCIEGrid01(3, i).Value & "') "
            SQL = SQL & "And Process_ID = ('" & QCIEGrid01(0, i).Value & "') "
            OpenTbl(ADb, Atb1, SQL)

            If Not Atb1.RecordCount <> 0 Then
                Atb1.AddNew()
            End If
            SetDate = QCIEGrid01(1, i).Value
            Atb1("Time").Value = QCIEGrid01(2, i).Value
            Atb1("Nik").Value = QCIEGrid01(3, i).Value
            Atb1("NoKg").Value = QCIEGrid01(5, i).Value
            Atb1("NoBag").Value = QCIEGrid01(6, i).Value
            Atb1("NoGr").Value = QCIEGrid01(7, i).Value
            Atb1("Pieces").Value = QCIEGrid01(8, i).Value
            Atb1("Coupon").Value = QCIEGrid01(4, i).Value
            Atb1("Salary").Value = QCIEGrid01(9, i).Value
            Atb1("Date").Value = SetDate.ToString("yyyy-MM-dd")
            Atb1("Process_ID").Value = QCIEGrid01(0, i).Value
            QCIEGrid01.Invoke(DirectCast(Sub() QCIEGrid01(10, i).Value = "Has been Saved", MethodInvoker))
            Atb1.Update()
            QCIECsvProg01.Value += 1
        Next

    End Sub

    Sub SaveStanInfo()

        For i = 0 To QCIEGrid01.Rows.Count - 1

            SQL = ""
            SQL = SQL & "Select * From 08_Standard_Table "
            SQL = SQL & "Where Original = ('" & QCIEGrid01(0, i).Value & "') "
            OpenTbl(ADb, Atb1, SQL)

            If Not Atb1.RecordCount <> 0 Then
                Atb1.AddNew()
            End If

            Atb1("Original").Value = QCIEGrid01(0, i).Value
            Atb1("Standard_Wage").Value = QCIEGrid01(1, i).Value

            Atb1.Update()
            QCIECsvProg01.Value += 1
        Next

    End Sub

    Sub SaveConvSal()

        For i = 0 To QCIEGrid01.Rows.Count - 1
            SetDate = QCIEGrid01(0, i).Value
            SQL = ""
            SQL = SQL & "Select * From 13_Conveyour_Salary "
            SQL = SQL & "Where Nik = ('" & QCIEGrid01(1, i).Value & "') "
            SQL = SQL & "And Date = ('" & SetDate.ToString("yyyy-MM-dd") & "') "
            OpenTbl(ADb, Atb1, SQL)

            If Not Atb1.RecordCount <> 0 Then
                Atb1.AddNew()
            End If
            Atb1("Date").Value = SetDate.ToString("yyyy-MM-dd")
            Atb1("Nik").Value = QCIEGrid01(1, i).Value
            Atb1("Salary").Value = QCIEGrid01(2, i).Value
            Atb1.Update()
            QCIECsvProg01.Value += 1
            QCIEGrid01.Invoke(DirectCast(Sub() QCIEGrid01(3, i).Value = "Has been Saved", MethodInvoker))
        Next
    End Sub

    Sub SaveMutSal()

        For i = 0 To QCIEGrid01.Rows.Count - 1
            SetDate = QCIEGrid01(0, i).Value
            SQL = ""
            SQL = SQL & "Select * From 14_MutuII_Salary "
            SQL = SQL & "Where Nik = ('" & QCIEGrid01(1, i).Value & "') "
            SQL = SQL & "And Date = ('" & SetDate.ToString("yyyy-MM-dd") & "') "
            OpenTbl(ADb, Atb1, SQL)

            If Not Atb1.RecordCount <> 0 Then
                Atb1.AddNew()
            End If
            Atb1("Date").Value = SetDate.ToString("yyyy-MM-dd")
            Atb1("Nik").Value = QCIEGrid01(1, i).Value
            Atb1("Salary").Value = QCIEGrid01(2, i).Value
            Atb1.Update()
            QCIECsvProg01.Value += 1
            QCIEGrid01.Invoke(DirectCast(Sub() QCIEGrid01(3, i).Value = "Has been Saved", MethodInvoker))
        Next
    End Sub

    Sub SaveWalSal()
        For i = 0 To QCIEGrid01.Rows.Count - 1
            SetDate = QCIEGrid01(0, i).Value
            SQL = ""
            SQL = SQL & "Select * From 15_Wallet_Salary "
            SQL = SQL & "Where Nik = ('" & QCIEGrid01(1, i).Value & "') "
            SQL = SQL & "And Date = ('" & SetDate.ToString("yyyy-MM-dd") & "') "
            OpenTbl(ADb, Atb1, SQL)

            If Not Atb1.RecordCount <> 0 Then
                Atb1.AddNew()
            End If
            Atb1("Date").Value = SetDate.ToString("yyyy-MM-dd")
            Atb1("Nik").Value = QCIEGrid01(1, i).Value
            Atb1("Salary").Value = QCIEGrid01(2, i).Value
            QCIEGrid01.Invoke(DirectCast(Sub() QCIEGrid01(3, i).Value = "Has been Saved", MethodInvoker))
            Atb1.Update()
            QCIECsvProg01.Value += 1
        Next
    End Sub

    Sub SavePackSal()
        For i = 0 To QCIEGrid01.Rows.Count - 1
            SetDate = QCIEGrid01(0, i).Value
            SQL = ""
            SQL = SQL & "Select * From 16_Packing_Salary "
            SQL = SQL & "Where Nik = ('" & QCIEGrid01(1, i).Value & "') "
            SQL = SQL & "And Date = ('" & SetDate.ToString("yyyy-MM-dd") & "') "
            OpenTbl(ADb, Atb1, SQL)

            If Not Atb1.RecordCount <> 0 Then
                Atb1.AddNew()
            End If
            Atb1("Date").Value = SetDate.ToString("yyyy-MM-dd")
            Atb1("Nik").Value = QCIEGrid01(1, i).Value
            Atb1("Salary").Value = QCIEGrid01(2, i).Value
            QCIEGrid01.Invoke(DirectCast(Sub() QCIEGrid01(3, i).Value = "Has been Saved", MethodInvoker))
            Atb1.Update()
            QCIECsvProg01.Value += 1
        Next
    End Sub

    Sub SaveMiscSortSal()
        For i = 0 To QCIEGrid01.Rows.Count - 1
            SetDate = QCIEGrid01(0, i).Value
            SQL = ""
            SQL = SQL & "Select * From 20_Miscellaneous_Salary "
            SQL = SQL & "Where Nik = ('" & QCIEGrid01(1, i).Value & "') "
            SQL = SQL & "And Date = ('" & SetDate.ToString("yyyy-MM-dd") & "') "
            SQL = SQL & "And TypeCtrl = ('" & QCIEGrid01(3, i).Value & "') "
            OpenTbl(ADb, Atb1, SQL)

            If Not Atb1.RecordCount <> 0 Then
                Atb1.AddNew()
            End If
            Atb1("Date").Value = SetDate.ToString("yyyy-MM-dd")
            Atb1("Nik").Value = QCIEGrid01(1, i).Value
            Atb1("Salary").Value = QCIEGrid01(2, i).Value
            Atb1("TypeCtrl").Value = QCIEGrid01(3, i).Value
            Atb1.Update()
            QCIECsvProg01.Value += 1
            QCIEGrid01.Invoke(DirectCast(Sub() QCIEGrid01(4, i).Value = "Has been Saved", MethodInvoker))
        Next
    End Sub

    Sub SaveHolInfo()
        For i = 0 To QCIEGrid01.Rows.Count - 1
            SetDate = QCIEGrid01(0, i).Value
            SQL = ""
            SQL = SQL & "Select * From 17_Holiday_Table "
            SQL = SQL & "Where Date = ('" & SetDate.ToString("yyyy-MM-dd") & "') "
            OpenTbl(ADb, Atb1, SQL)

            If Not Atb1.RecordCount <> 0 Then
                Atb1.AddNew()
            End If
            Atb1("Date").Value = SetDate.ToString("yyyy-MM-dd")
            Atb1("Holiday_Name").Value = QCIEGrid01(1, i).Value
            Atb1("Salary_Mod").Value = QCIEGrid01(2, i).Value
            Atb1.Update()
            QCIECsvProg01.Value += 1
        Next
    End Sub

    Sub SaveDateCount()

        For i = 0 To QCIEGrid01.Rows.Count - 1
            SQL = Nothing
            SQL = SQL & "Select * From datecounter2table "
            SQL = SQL & "Where Periode = ('" & QCIEGrid01(1, i).Value & "') "
            SQL = SQL & "And PeriodeRange = ('" & QCIEGrid01(2, i).Value & "') "
            OpenTbl(CBb, Ctbl1, SQL)

            If Not Ctbl1.RecordCount <> 0 Then
                Ctbl1.AddNew()
            End If

            Ctbl1("IDDate").Value = QCIEGrid01(0, i).Value
            Ctbl1("Periode").Value = QCIEGrid01(1, i).Value
            Ctbl1("PeriodeRange").Value = QCIEGrid01(2, i).Value
            Ctbl1("PeriodeValid").Value = QCIEGrid01(3, i).Value
            Ctbl1("Date1").Value = QCIEGrid01(4, i).Value
            Ctbl1("Date2").Value = QCIEGrid01(5, i).Value
            Ctbl1("Date3").Value = QCIEGrid01(6, i).Value
            Ctbl1("Date4").Value = QCIEGrid01(7, i).Value
            Ctbl1("Date5").Value = QCIEGrid01(8, i).Value
            Ctbl1("Date6").Value = QCIEGrid01(9, i).Value
            Ctbl1("Date7").Value = QCIEGrid01(10, i).Value
            Ctbl1("Date8").Value = QCIEGrid01(11, i).Value
            Ctbl1("Date9").Value = QCIEGrid01(12, i).Value
            Ctbl1("Date10").Value = QCIEGrid01(13, i).Value
            Ctbl1("Date11").Value = QCIEGrid01(14, i).Value
            Ctbl1("Date12").Value = QCIEGrid01(15, i).Value
            Ctbl1("Date13").Value = QCIEGrid01(16, i).Value
            Ctbl1("Date14").Value = QCIEGrid01(17, i).Value
            Ctbl1("Date15").Value = QCIEGrid01(18, i).Value
            Ctbl1("Date16").Value = QCIEGrid01(19, i).Value

            Ctbl1.Update()

            QCIECsvProg01.Value += 1
        Next
    End Sub

    Sub SavePeriCount()

        For i = 0 To QCIEGrid01.Rows.Count - 1

            SQL = Nothing
            SQL = SQL & "Select * From datecounter2table "
            SQL = SQL & "Where Date = ('" & QCIEGrid01(0, i).Value & "') "
            OpenTbl(CBb, Ctbl1, SQL)

            If Not Ctbl1.RecordCount <> 0 Then
                Ctbl1.AddNew()
            End If

            Ctbl1("Date").Value = QCIEGrid01(0, i).Value
            Ctbl1("Counter").Value = QCIEGrid01(1, i).Value
            Ctbl1("Periode").Value = QCIEGrid01(2, i).Value
            Ctbl1("PeriodeRange").Value = QCIEGrid01(4, i).Value

            Ctbl1.Update()

        Next
    End Sub

    Sub SaveMainSal()

        For i = 0 To QCIEGrid01.Rows.Count - 1

            SQL = Nothing
            SQL = SQL & "Select * From salarysync1_table "
            SQL = SQL & "Where Periode = ('" & QCIEGrid01(0, i).Value & "') "
            SQL = SQL & "and PeriodeRange = ('" & QCIEGrid01(0, i).Value & "') "
            SQL = SQL & "and Nik = ('" & QCIEGrid01(0, i).Value & "') "

            OpenTbl(CBb, Ctbl1, SQL)

            If Not Ctbl1.RecordCount <> 0 Then
                Ctbl1.AddNew()
            End If

            ReStringName = QCIEGrid01(3, i).Value
            ReStringName = ReStringName.Replace("'", "?")
            Ctbl1("Periode").Value = QCIEGrid01(0, i).Value
            Ctbl1("PeriodeRange").Value = QCIEGrid01(1, i).Value
            Ctbl1("Nik").Value = QCIEGrid01(2, i).Value
            Ctbl1("Name").Value = ReStringName
            Ctbl1("Pay").Value = QCIEGrid01(4, i).Value
            Ctbl1("Salary1").Value = QCIEGrid01(5, i).Value
            Ctbl1("Salary2").Value = QCIEGrid01(6, i).Value
            Ctbl1("Salary3").Value = QCIEGrid01(7, i).Value
            Ctbl1("Salary4").Value = QCIEGrid01(8, i).Value
            Ctbl1("Salary5").Value = QCIEGrid01(9, i).Value
            Ctbl1("Salary6").Value = QCIEGrid01(10, i).Value
            Ctbl1("Salary7").Value = QCIEGrid01(11, i).Value
            Ctbl1("Salary8").Value = QCIEGrid01(12, i).Value
            Ctbl1("Salary9").Value = QCIEGrid01(13, i).Value
            Ctbl1("Salary10").Value = QCIEGrid01(14, i).Value
            Ctbl1("Salary11").Value = QCIEGrid01(15, i).Value
            Ctbl1("Salary12").Value = QCIEGrid01(16, i).Value
            Ctbl1("Salary13").Value = QCIEGrid01(17, i).Value
            Ctbl1("Salary14").Value = QCIEGrid01(18, i).Value
            Ctbl1("Salary15").Value = QCIEGrid01(19, i).Value
            Ctbl1("Salary16").Value = QCIEGrid01(20, i).Value
            Ctbl1("Date1").Value = QCIEGrid01(21, i).Value
            Ctbl1("Date2").Value = QCIEGrid01(22, i).Value
            Ctbl1("Date3").Value = QCIEGrid01(23, i).Value
            Ctbl1("Date4").Value = QCIEGrid01(24, i).Value
            Ctbl1("Date5").Value = QCIEGrid01(25, i).Value
            Ctbl1("Date6").Value = QCIEGrid01(26, i).Value
            Ctbl1("Date7").Value = QCIEGrid01(27, i).Value
            Ctbl1("Date8").Value = QCIEGrid01(28, i).Value
            Ctbl1("Date9").Value = QCIEGrid01(29, i).Value
            Ctbl1("Date10").Value = QCIEGrid01(30, i).Value
            Ctbl1("Date11").Value = QCIEGrid01(31, i).Value
            Ctbl1("Date12").Value = QCIEGrid01(32, i).Value
            Ctbl1("Date13").Value = QCIEGrid01(33, i).Value
            Ctbl1("Date14").Value = QCIEGrid01(34, i).Value
            Ctbl1("Date15").Value = QCIEGrid01(35, i).Value
            Ctbl1("Date16").Value = QCIEGrid01(36, i).Value
            Ctbl1("AstekVal").Value = QCIEGrid01(37, i).Value
            Ctbl1("PNoRek").Value = QCIEGrid01(38, i).Value
            Ctbl1("PotLain").Value = QCIEGrid01(39, i).Value

            Ctbl1.Update()
            QCIECsvProg01.Value += 1
            QCIEGrid01.Invoke(DirectCast(Sub() QCIEGrid01(40, i).Value = "Has been Saved", MethodInvoker))
        Next

    End Sub

    Sub SavePPh21Sal()
        For i = 0 To QCIEGrid01.Rows.Count - 1


            SQL = Nothing
            SQL = SQL & "Select * From emp_pphtable "
            SQL = SQL & "Where PeriodeGajian = ('" & QCIEGrid01(0, i).Value & "') "
            SQL = SQL & "and Periode = ('" & QCIEGrid01(1, i).Value & "') "
            SQL = SQL & "and PeriodeRange = ('" & QCIEGrid01(2, i).Value & "') "
            SQL = SQL & "and Nik = ('" & QCIEGrid01(3, i).Value & "') "
            OpenTbl(PPhDB, PPhTb1, SQL)

            If Not PPhTb1.RecordCount <> 0 Then
                PPhTb1.AddNew()
            End If

            ReStringName = QCIEGrid01(4, i).Value
            ReStringName = ReStringName.Replace("'", "?")
            PPhTb1("PeriodeGajian").Value = QCIEGrid01(0, i).Value
            PPhTb1("Periode").Value = QCIEGrid01(1, i).Value
            PPhTb1("PeriodeRange").Value = QCIEGrid01(2, i).Value
            PPhTb1("Nik").Value = QCIEGrid01(3, i).Value
            PPhTb1("Name").Value = QCIEGrid01(4, i).Value
            PPhTb1("EmAdd").Value = QCIEGrid01(5, i).Value
            PPhTb1("Pay").Value = QCIEGrid01(6, i).Value
            PPhTb1("Astek").Value = QCIEGrid01(7, i).Value
            PPhTb1("NPWP").Value = QCIEGrid01(8, i).Value
            PPhTb1("KTP").Value = QCIEGrid01(9, i).Value
            PPhTb1("MainSalary1").Value = QCIEGrid01(10, i).Value
            PPhTb1("MainSalary2").Value = QCIEGrid01(11, i).Value
            PPhTb1("MainSalary3").Value = QCIEGrid01(12, i).Value
            PPhTb1("Incentif").Value = QCIEGrid01(13, i).Value
            PPhTb1.Update()
            QCIECsvProg01.Value += 1
        Next

    End Sub


#End Region

    Private Sub QCIECombox01_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles QCIECombox01.KeyPress
        e.Handled = True
    End Sub

    Private Sub QCIETbx01_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles QCIETbx01.KeyPress
        e.Handled = True

    End Sub

    Private Sub QCIEBtn04_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QCIEBtn04.Click
        QCIECsvProg01.Maximum = Nothing
        QCIECsvProg01.Value = Nothing
        Try
            QCGridHeader()
            'GetCSVItem()
            OnClickTheWorker2()
        
        Catch ex As Exception

        End Try

        If Not QCIETbx01.Text = Nothing Then
            QCIECombox01.Enabled = False
        End If
    End Sub

    Private Sub QCIEBtn03_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QCIEBtn03.Click
        QCIEGrid01.Rows.Clear()
        QCIEGrid01.Columns.Clear()
        OpenFileLink()
    End Sub

    Private Sub QCIEBtn05_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QCIEBtn05.Click

        If QCIEGrid01.Rows.Count <> 0 Then
            QCIECsvProg01.Maximum = QCIEGrid01.Rows.Count
            RCountLbl01.Text = "Row Counts : " + QCIECsvProg01.Maximum.ToString
        End If
        'SaveConvInfo()

        OnClickTheWorker3()
        ''SaveEmploInfo()
    End Sub



End Class