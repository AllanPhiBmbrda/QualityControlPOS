Option Explicit On

Public Class UserSynchBlock


    Dim UserSaveSynchID As String
    Dim UserSaveSynchNik As String
    Dim UserSaveSynchName As String
    Dim UserSaveSynchActive As String
    Dim UserSaveSynchDate As String
    Dim UserSaveSynchPay As String
    Dim UserSaveSynchAstek As String
    Dim UserSaveSynchNPWP As String
    Dim UserSaveSynchBank As String
    Dim UserSaveSynchNoRek As String
    Dim UserSaveSynchNoKTP As String
    Dim UserSaveSynchNoKPJ As String
    Dim UserSaveSynchJab As String
    Dim UserSaveSynchEstate As String
    Dim UserSaveSynchLahir As String
    Dim UserSaveSynchAgama As String
    Dim UserSaveSynchTelNum As String
    Dim UserSaveSynchPendi As String
    Dim UserSaveSynchDept As String
    Dim UserSaveSynchJKK As String

    Dim UserSaveNikInc As String
    Dim UserSaveMonInc As String
    Dim UserSaveCountInc As String

    Private Sub UserSynchBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadDB()
        LoadDB3()
        UserSynchRadio1.Checked = True

    End Sub

#Region "Incentives Action"

#End Region

    Sub UserGridHeader()

        With UserSynchGrid1

            If UserSynchRadio1.Checked = True Then
                .Columns.Add("Col0", "ID Number")
                .Columns.Add("Col1", "Nik")
                .Columns.Add("Col2", "Name")
                .Columns.Add("Col3", "Active")
                .Columns.Add("Col4", "DateStart")
                .Columns.Add("Col5", "Pay")
                .Columns.Add("Col6", "Jamsostek")
                .Columns.Add("Col7", "NPWP")
                .Columns.Add("Col8", "Bank")
                .Columns.Add("Col9", "No Rek")
                .Columns.Add("Col10", "No KTP")
                .Columns.Add("Col11", "No KPJ")
                .Columns.Add("Col12", "Jabatan")
                .Columns.Add("Col13", "Estate")
                .Columns.Add("Col14", "Tempat Lahir")
                .Columns.Add("Col15", "Agama")
                .Columns.Add("Col16", "Tel Num")
                .Columns.Add("Col17", "Pendi")
                .Columns.Add("Col18", "Dept")
                .Columns.Add("Col19", "JKKLM")
                .Columns.Add("Col20", "Status")
            End If

        End With

    End Sub

    Sub SynchNest()

        UserSaveSynchID = Nothing
        UserSaveSynchNik = Nothing
        UserSaveSynchName = Nothing
        UserSaveSynchActive = Nothing
        UserSaveSynchDate = Nothing
        UserSaveSynchPay = Nothing
        UserSaveSynchAstek = Nothing
        UserSaveSynchNPWP = Nothing
        UserSaveSynchBank = Nothing
        UserSaveSynchNoRek = Nothing
        UserSaveSynchNoKTP = Nothing
        UserSaveSynchNoKPJ = Nothing
        UserSaveSynchJab = Nothing
        UserSaveSynchEstate = Nothing
        UserSaveSynchLahir = Nothing
        UserSaveSynchAgama = Nothing
        UserSaveSynchTelNum = Nothing
        UserSaveSynchPendi = Nothing
        UserSaveSynchDept = Nothing
        UserSaveSynchJKK = Nothing

    End Sub

    Sub UserSynchLoad()

        UserGridHeader()

        SQL = ""
        SQL = SQL & "Select * from 02_Name_Table "
        SQL = SQL & "Where DateStart between ('" & USDTPick01.Text & "') "
        SQL = SQL & "and ('" & USDTPick02.Text & "') "
        SQL = SQL & "Order by Nik Asc"
        OpenTbl(BBb, BBtbl3, SQL)

        If BBtbl3.RecordCount <> 0 Then

            BBtbl3.MoveFirst()
            Do While Not BBtbl3.EOF

                UserSynchGrid1.Rows.Add(IIf(IsDBNull(BBtbl3("ID_Number").Value), "", BBtbl3("ID_Number").Value),
                                        IIf(IsDBNull(BBtbl3("Nik").Value), "", BBtbl3("Nik").Value),
                                        IIf(IsDBNull(BBtbl3("Name").Value), "", BBtbl3("Name").Value),
                                        IIf(IsDBNull(BBtbl3("Active").Value), "", BBtbl3("Active").Value),
                                        IIf(IsDBNull(BBtbl3("DateStart").Value), "", BBtbl3("DateStart").Value),
                                        IIf(IsDBNull(BBtbl3("Pay").Value), "", BBtbl3("Pay").Value),
                                        IIf(IsDBNull(BBtbl3("Jamsostek").Value), "", BBtbl3("Jamsostek").Value),
                                        IIf(IsDBNull(BBtbl3("NPWP").Value), "", BBtbl3("NPWP").Value),
                                        IIf(IsDBNull(BBtbl3("Bank_Ctrl").Value), "", BBtbl3("Bank_Ctrl").Value),
                                        IIf(IsDBNull(BBtbl3("NoRek").Value), "", BBtbl3("NoRek").Value),
                                        IIf(IsDBNull(BBtbl3("NKTP").Value), "", BBtbl3("NKTP").Value),
                                        IIf(IsDBNull(BBtbl3("NoKPJ").Value), "", BBtbl3("NoKPJ").Value),
                                        IIf(IsDBNull(BBtbl3("JabData").Value), "", BBtbl3("JabData").Value),
                                        IIf(IsDBNull(BBtbl3("Estate").Value), "", BBtbl3("Estate").Value),
                                        IIf(IsDBNull(BBtbl3("Lahir").Value), "", BBtbl3("Lahir").Value),
                                        IIf(IsDBNull(BBtbl3("Agama").Value), "", BBtbl3("Agama").Value),
                                        IIf(IsDBNull(BBtbl3("TelNum").Value), "", BBtbl3("TelNum").Value),
                                        IIf(IsDBNull(BBtbl3("Pendi").Value), "", BBtbl3("Pendi").Value),
                                        IIf(IsDBNull(BBtbl3("Dept").Value), "", BBtbl3("Dept").Value),
                                        IIf(IsDBNull(BBtbl3("JKKJKM").Value), "", BBtbl3("JKKJKM").Value))

                BBtbl3.MoveNext()

            Loop

        End If

    End Sub

    Sub UserSynchSave()

        If UserSynchRadio1.Checked = True Then

            For i = 0 To UserSynchGrid1.Rows.Count - 1

                UserSaveSynchID = UserSynchGrid1(0, i).Value
                UserSaveSynchNik = UserSynchGrid1(1, i).Value
                UserSaveSynchName = UserSynchGrid1(2, i).Value
                UserSaveSynchActive = UserSynchGrid1(3, i).Value
                UserSaveSynchDate = UserSynchGrid1(4, i).Value
                UserSaveSynchPay = UserSynchGrid1(5, i).Value
                UserSaveSynchAstek = UserSynchGrid1(6, i).Value
                UserSaveSynchNPWP = UserSynchGrid1(7, i).Value
                UserSaveSynchBank = UserSynchGrid1(8, i).Value
                UserSaveSynchNoRek = UserSynchGrid1(9, i).Value
                UserSaveSynchNoKTP = UserSynchGrid1(10, i).Value
                UserSaveSynchNoKPJ = UserSynchGrid1(11, i).Value
                UserSaveSynchJab = UserSynchGrid1(12, i).Value
                UserSaveSynchEstate = UserSynchGrid1(13, i).Value
                UserSaveSynchLahir = UserSynchGrid1(14, i).Value
                UserSaveSynchAgama = UserSynchGrid1(15, i).Value
                UserSaveSynchTelNum = UserSynchGrid1(16, i).Value
                UserSaveSynchPendi = UserSynchGrid1(17, i).Value
                UserSaveSynchDept = UserSynchGrid1(18, i).Value
                UserSaveSynchJKK = UserSynchGrid1(19, i).Value

                SQL = ""
                SQL = SQL & "Select * From 02_Name_Table "
                SQL = SQL & "Where DateStart = ('" & UserSaveSynchDate & "') "
                SQL = SQL & "And Nik = ('" & UserSaveSynchNik & "') "

                OpenTbl(ADb, Atbl39, SQL)

                If Not Atbl39.RecordCount <> 0 Then

                    Atbl39.AddNew()

                    Atbl39("ID_Number").Value = UserSaveSynchID
                    Atbl39("Nik").Value = UserSaveSynchNik
                    Atbl39("Name").Value = UserSaveSynchName
                    Atbl39("Active").Value = UserSaveSynchActive
                    Atbl39("DateStart").Value = UserSaveSynchDate
                    Atbl39("Pay").Value = UserSaveSynchPay
                    Atbl39("Jamsostek").Value = UserSaveSynchAstek

                    ' New Update as of 13/12/2014

                    Atbl39("NPWP").Value = UserSaveSynchNPWP
                    Atbl39("Bank_Ctrl").Value = UserSaveSynchBank
                    Atbl39("NoRek").Value = UserSaveSynchNoRek
                    Atbl39("NKTP").Value = UserSaveSynchNoKTP
                    Atbl39("NoKPJ").Value = UserSaveSynchNoKPJ
                    Atbl39("JabData").Value = UserSaveSynchJab
                    Atbl39("Estate").Value = UserSaveSynchEstate
                    Atbl39("Lahir").Value = UserSaveSynchLahir
                    Atbl39("Agama").Value = UserSaveSynchAgama
                    Atbl39("TelNum").Value = UserSaveSynchTelNum
                    Atbl39("Pendi").Value = UserSaveSynchPendi
                    Atbl39("Dept").Value = UserSaveSynchDept
                    Atbl39("JKKJKM").Value = UserSaveSynchJKK

                    Atbl39.Update()

                    UserSynchGrid1(20, i).Value = "Has Been Saved"

                ElseIf Atbl39.RecordCount > 0 Then

                    UserSynchGrid1(20, i).Value = "Already Exist"

                End If

            Next

        End If

    End Sub

    Private Sub UserSynchBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserSynchBtn1.Click

        If UserSynchRadio1.Checked = True Then

            UserSynchLoad()

        End If

    End Sub

    Private Sub UserSynchBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserSynchBtn2.Click

        UserSynchSave()

    End Sub

End Class