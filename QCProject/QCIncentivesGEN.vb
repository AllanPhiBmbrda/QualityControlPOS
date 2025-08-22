Public Class QCIncentivesGEN

    Dim IncSlip1 As String
    Dim IncSlip2 As String
    Dim IncSlip3 As String
    Dim Count01 As String = Nothing
    Dim Count02 As String = Nothing
    Dim Count03 As String = Nothing
    Dim Count04 As String = Nothing
    Dim Count05 As String = Nothing


    Dim DateA As String ' 1
    Dim DateB As String ' 2
    Dim DateC As String ' 3
    Dim DateD As String ' 4
    Dim DateE As String ' 5
    Dim DateF As String ' 6
    Dim DateG As String ' 7
    Dim DateH As String ' 8
    Dim DateI As String ' 9
    Dim DateJ As String ' 10
    Dim DateK As String ' 11
    Dim DateL As String ' 12
    Dim DateM As String ' 13
    Dim DateN As String ' 14
    Dim DateO As String ' 15
    Dim DateP As String ' 16

    Dim DateQ As String ' 17
    Dim DateR As String ' 18
    Dim DateS As String ' 19
    Dim DateT As String ' 20
    Dim DateU As String ' 21 
    Dim DateV As String ' 22
    Dim DateW As String ' 23
    Dim DateX As String ' 24
    Dim DateY As String ' 25
    Dim DateZ As String ' 26
    Dim DateAA As String ' 27
    Dim DateAB As String ' 28
    Dim DateAC As String ' 29
    Dim DateAD As String ' 30
    Dim DateAE As String ' 31
    Dim DateAF As String ' 32

    Dim GetNik As String = Nothing
    Dim GetName As String = Nothing

    Dim PR1 As String = Nothing
    Dim NikV1 As String = Nothing
    Dim NameV1 As String = Nothing
    Dim CD1 As String = Nothing
    Dim Cc1 As String = Nothing
    Dim Cc2 As String = Nothing
    Dim Cc3 As String = Nothing
    Dim Fc1 As String = Nothing
    Dim Inv1 As String = Nothing

    Dim IncMoney As String
    Dim IncValPoints As String

    Dim OrangDate As Date
    Dim DateNow As Date = Today
    Dim DateDivider As Integer


    Private Sub QCIncentivesGEN_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadDB()
        LoadDB2()

    End Sub

    Sub IncenGridEHeader()
        IncGrid1.Rows.Clear()
        IncGrid1.Columns.Clear()

        With IncGrid1

            .Columns.Add("col1", "Periode Range")
            .Columns.Add("col2", "Nik")
            .Columns.Add("col3", "Name")

            .Columns.Add("col4", "Count Periode I")
            .Columns.Add("col5", "Count Periode II")
            .Columns.Add("col5", "Count Periode EX")
            .Columns.Add("col6", "Total")
            .Columns.Add("col7", "Incentives")
            .Columns.Add("col8", "Employee's Day Ages")
            .Columns.Add("col9", "Status")

        End With

    End Sub

    Sub UPandSlipFil()

        IncCmb01.Items.Clear()
        IncSlip1 = Format(Now, "yyyy")
        IncSlip2 = IncSlip1 - 1
        IncSlip3 = IncSlip2 + 1

        With IncCmb01

            .Items.Add("Dec " + IncSlip2)
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

        End With

    End Sub

    Sub LoadNumberDate()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Periode =  ('" & "Periode I" & "') "
        SQL = SQL & "and PeriodeRange = ('" & IncCmb01.Text & " ') "
        OpenTbl(CBb, Ctbl54, SQL)

        If Ctbl54.RecordCount <> 0 Then
            Ctbl54.MoveFirst()
            Do While Not Ctbl54.EOF

                Count01 = Val(Count01) + 1

                Ctbl54.MoveNext()

            Loop

        End If

    End Sub

    Sub LoadNumberDate2()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Periode =  ('" & "Periode II" & "') "
        SQL = SQL & "and PeriodeRange = ('" & IncCmb01.Text & " ') "
        OpenTbl(CBb, Ctbl55, SQL)

        If Ctbl55.RecordCount <> 0 Then
            Ctbl55.MoveFirst()

            Do While Not Ctbl55.EOF

                Count02 = Val(Count02) + 1

                Ctbl55.MoveNext()

            Loop

        End If

    End Sub

    Sub LoadNumberDate3()

        SQL = ""
        SQL = SQL & "Select * From Periode_CounterTable "
        SQL = SQL & "Where Periode =  ('" & "Periode II" & "') "
        SQL = SQL & "and PeriodeRange = ('" & IncCmb01.Text & " ') "
        OpenTbl(CBb, Ctbl55, SQL)

        If Ctbl55.RecordCount <> 0 Then
            Ctbl55.MoveFirst()

            Do While Not Ctbl55.EOF

                Count02 = Val(Count02) + 1

                Ctbl55.MoveNext()

            Loop

        End If

    End Sub

    Sub GenerateIncentives()

        If Not DateDivider <= 60 Then

            IncMoney = Val(IncGenTbx1.Text) - Val(Count05)

            If IncMoney = 4 Then

                IncValPoints = Val(StandardsSalary)

            ElseIf IncMoney = 3 Then

                IncValPoints = Val(StandardsSalary) * 2

            ElseIf IncMoney = 2 Then

                IncValPoints = Val(StandardsSalary) * 3

            ElseIf IncMoney = 1 Then

                IncValPoints = Val(StandardsSalary) * 4

            ElseIf IncMoney >= 5 Then

                IncValPoints = "0"

            End If

        Else

            IncMoney = Val(IncGenTbx1.Text) - Val(Count05)

            If IncMoney = 4 Then

                IncValPoints = Val(SubsidiSalary)

            ElseIf IncMoney = 3 Then

                IncValPoints = Val(SubsidiSalary) * 2

            ElseIf IncMoney = 2 Then

                IncValPoints = Val(SubsidiSalary) * 3

            ElseIf IncMoney = 1 Then

                IncValPoints = Val(SubsidiSalary) * 4

            ElseIf IncMoney >= 5 Then

                IncValPoints = "0"

            End If

        End If

    End Sub

    Sub UploadIncentives()

        For a = 0 To IncGrid1.Rows.Count - 1

            PR1 = IncGrid1(0, a).Value ' Periode Range
            NikV1 = IncGrid1(1, a).Value ' Nik
            NameV1 = IncGrid1(2, a).Value ' Name
            Cc1 = IncGrid1(3, a).Value ' Count PR I
            Cc2 = IncGrid1(4, a).Value ' Count PR II
            Fc1 = IncGrid1(5, a).Value ' Count Total
            Inv1 = IncGrid1(6, a).Value '  Incentives Money
            CD1 = IncGrid1(7, a).Value '  Days that distinguished of Anak Baru or Regular

            SQL = ""
            SQL = SQL & "Select * From IncentiveO_Table "
            SQL = SQL & "Where Periode_Range = ('" & PR1 & "') "
            SQL = SQL & "And Nik_Val =  ('" & NikV1 & "') "
            SQL = SQL & "And Name_Val =  ('" & NameV1 & "') "

            OpenTbl(CBb, Ctbl59, SQL)

            If Not Ctbl59.RecordCount <> 0 Then
                Ctbl59.AddNew()

                Ctbl59("Periode_Range").Value = PR1
                Ctbl59("Nik_Val").Value = NikV1
                Ctbl59("Name_Val").Value = NameV1

                Ctbl59("Count1").Value = Cc1
                Ctbl59("Count2").Value = Cc2
                Ctbl59("Count3").Value = Cc3
                Ctbl59("FinalCount").Value = Fc1
                Ctbl59("IncenValue").Value = Inv1
                Ctbl59("CountDate").Value = CD1
                Ctbl59.Update()

                IncGrid1(8, a).Value = "Has Been Saved"

            ElseIf Ctbl59.RecordCount > 0 Then

                IncGrid1(8, a).Value = "Data is Already Exist"

            End If

        Next

    End Sub

    Sub LoadCountForPer()

        SQL = ""
        SQL = SQL & "Select * From SalarySync1_Table "
        SQL = SQL & "Where PeriodeRange = ('" & IncCmb01.Text & " ') "
        SQL = SQL & "And Periode =  ('" & "Periode I" & "') "
        OpenTbl(CBb, Ctbl56, SQL)

        If Ctbl56.RecordCount <> 0 Then
            Ctbl56.MoveFirst()
            Do While Not Ctbl56.EOF

                GetNik = IIf(IsDBNull(Ctbl56("Nik").Value), "", Ctbl56("Nik").Value)
                GetName = IIf(IsDBNull(Ctbl56("Name").Value), "", Ctbl56("Name").Value)

                SQL = ""
                SQL = SQL & "Select * From SalarySync1_Table "
                SQL = SQL & "Where Periode =  ('" & "Periode I" & "') "
                SQL = SQL & "And PeriodeRange =  ('" & IncCmb01.Text & "') "
                SQL = SQL & "And Nik =  ('" & GetNik & "') "
                SQL = SQL & "And Name =  ('" & GetName & "') "
                OpenTbl(CBb, Ctbl57, SQL)

                If Ctbl57.RecordCount > 0 Then

                    DateA = IIf(IsDBNull(Ctbl57("Salary1").Value), "", Ctbl57("Salary1").Value)
                    DateB = IIf(IsDBNull(Ctbl57("Salary2").Value), "", Ctbl57("Salary2").Value)
                    DateC = IIf(IsDBNull(Ctbl57("Salary3").Value), "", Ctbl57("Salary3").Value)
                    DateD = IIf(IsDBNull(Ctbl57("Salary4").Value), "", Ctbl57("Salary4").Value)
                    DateE = IIf(IsDBNull(Ctbl57("Salary5").Value), "", Ctbl57("Salary5").Value)
                    DateF = IIf(IsDBNull(Ctbl57("Salary6").Value), "", Ctbl57("Salary6").Value)
                    DateG = IIf(IsDBNull(Ctbl57("Salary7").Value), "", Ctbl57("Salary7").Value)
                    DateH = IIf(IsDBNull(Ctbl57("Salary8").Value), "", Ctbl57("Salary8").Value)
                    DateI = IIf(IsDBNull(Ctbl57("Salary9").Value), "", Ctbl57("Salary9").Value)
                    DateJ = IIf(IsDBNull(Ctbl57("Salary10").Value), "", Ctbl57("Salary10").Value)
                    DateK = IIf(IsDBNull(Ctbl57("Salary11").Value), "", Ctbl57("Salary11").Value)
                    DateL = IIf(IsDBNull(Ctbl57("Salary12").Value), "", Ctbl57("Salary12").Value)
                    DateM = IIf(IsDBNull(Ctbl57("Salary13").Value), "", Ctbl57("Salary13").Value)
                    DateN = IIf(IsDBNull(Ctbl57("Salary14").Value), "", Ctbl57("Salary14").Value)
                    DateO = IIf(IsDBNull(Ctbl57("Salary15").Value), "", Ctbl57("Salary15").Value)
                    DateP = IIf(IsDBNull(Ctbl57("Salary16").Value), "", Ctbl57("Salary16").Value)

                End If

                SQL = ""
                SQL = SQL & "Select * From SalarySync1_Table "
                SQL = SQL & "Where Periode =  ('" & "Periode II" & "') "
                SQL = SQL & "And PeriodeRange =  ('" & IncCmb01.Text & "') "
                SQL = SQL & "And Nik =  ('" & GetNik & "') "
                SQL = SQL & "And Name =  ('" & GetName & "') "
                OpenTbl(CBb, Ctbl58, SQL)

                If Ctbl58.RecordCount > 0 Then

                    DateQ = IIf(IsDBNull(Ctbl58("Salary1").Value), "", Ctbl58("Salary1").Value)
                    DateR = IIf(IsDBNull(Ctbl58("Salary2").Value), "", Ctbl58("Salary2").Value)
                    DateS = IIf(IsDBNull(Ctbl58("Salary3").Value), "", Ctbl58("Salary3").Value)
                    DateT = IIf(IsDBNull(Ctbl58("Salary4").Value), "", Ctbl58("Salary4").Value)
                    DateU = IIf(IsDBNull(Ctbl58("Salary5").Value), "", Ctbl58("Salary5").Value)
                    DateV = IIf(IsDBNull(Ctbl58("Salary6").Value), "", Ctbl58("Salary6").Value)
                    DateW = IIf(IsDBNull(Ctbl58("Salary7").Value), "", Ctbl58("Salary7").Value)
                    DateX = IIf(IsDBNull(Ctbl58("Salary8").Value), "", Ctbl58("Salary8").Value)
                    DateY = IIf(IsDBNull(Ctbl58("Salary9").Value), "", Ctbl58("Salary9").Value)
                    DateZ = IIf(IsDBNull(Ctbl58("Salary10").Value), "", Ctbl58("Salary10").Value)
                    DateAA = IIf(IsDBNull(Ctbl58("Salary11").Value), "", Ctbl58("Salary11").Value)
                    DateAB = IIf(IsDBNull(Ctbl58("Salary12").Value), "", Ctbl58("Salary12").Value)
                    DateAC = IIf(IsDBNull(Ctbl58("Salary13").Value), "", Ctbl58("Salary13").Value)
                    DateAD = IIf(IsDBNull(Ctbl58("Salary14").Value), "", Ctbl58("Salary14").Value)
                    DateAE = IIf(IsDBNull(Ctbl58("Salary15").Value), "", Ctbl58("Salary15").Value)
                    DateAF = IIf(IsDBNull(Ctbl58("Salary16").Value), "", Ctbl58("Salary16").Value)

                End If

                DataCount()

                Ctbl56.MoveNext()

            Loop

        End If

    End Sub

    Sub DataCount()

        If Not DateA = Nothing Or DateA = "0" Then
            Count03 = Val(Count03) + 1
        End If

        If Not DateB = Nothing Or DateB = "0" Then
            Count03 = Val(Count03) + 1
        End If

        If Not DateC = Nothing Or DateC = "0" Then
            Count03 = Val(Count03) + 1
        End If

        If Not DateD = Nothing Or DateD = "0" Then
            Count03 = Val(Count03) + 1
        End If

        If Not DateE = Nothing Or DateE = "0" Then
            Count03 = Val(Count03) + 1
        End If

        If Not DateF = Nothing Or DateF = "0" Then
            Count03 = Val(Count03) + 1
        End If

        If Not DateG = Nothing Or DateG = "0" Then
            Count03 = Val(Count03) + 1
        End If

        If Not DateH = Nothing Or DateH = "0" Then
            Count03 = Val(Count03) + 1
        End If

        If Not DateI = Nothing Or DateI = "0" Then
            Count03 = Val(Count03) + 1
        End If

        If Not DateJ = Nothing Or DateJ = "0" Then
            Count03 = Val(Count03) + 1
        End If

        If Not DateK = Nothing Or DateK = "0" Then
            Count03 = Val(Count03) + 1
        End If

        If Not DateL = Nothing Or DateL = "0" Then
            Count03 = Val(Count03) + 1
        End If

        If Not DateM = Nothing Or DateM = "0" Then
            Count03 = Val(Count03) + 1
        End If

        If Not DateN = Nothing Or DateN = "0" Then
            Count03 = Val(Count03) + 1
        End If

        If Not DateO = Nothing Or DateO = "0" Then
            Count03 = Val(Count03) + 1
        End If

        If Not DateP = Nothing Or DateP = "0" Then
            Count03 = Val(Count03) + 1
        End If

        If Not DateQ = Nothing Or DateQ = "0" Then
            Count04 = Val(Count04) + 1
        End If

        If Not DateR = Nothing Or DateR = "0" Then
            Count04 = Val(Count04) + 1
        End If

        If Not DateS = Nothing Or DateS = "0" Then
            Count04 = Val(Count04) + 1
        End If

        If Not DateT = Nothing Or DateT = "0" Then
            Count04 = Val(Count04) + 1
        End If

        If Not DateU = Nothing Or DateU = "0" Then
            Count04 = Val(Count04) + 1
        End If

        If Not DateV = Nothing Or DateV = "0" Then
            Count04 = Val(Count04) + 1
        End If

        If Not DateW = Nothing Or DateW = "0" Then
            Count04 = Val(Count04) + 1
        End If

        If Not DateY = Nothing Or DateY = "0" Then
            Count04 = Val(Count04) + 1
        End If

        If Not DateZ = Nothing Or DateZ = "0" Then
            Count04 = Val(Count04) + 1
        End If

        If Not DateAA = Nothing Or DateAA = "0" Then
            Count04 = Val(Count04) + 1
        End If

        If Not DateAB = Nothing Or DateAB = "0" Then
            Count04 = Val(Count04) + 1
        End If

        If Not DateAC = Nothing Or DateAC = "0" Then
            Count04 = Val(Count04) + 1
        End If

        If Not DateAD = Nothing Or DateAD = "0" Then
            Count04 = Val(Count04) + 1
        End If

        If Not DateAE = Nothing Or DateAE = "0" Then
            Count04 = Val(Count04) + 1
        End If

        If Not DateAF = Nothing Or DateAF = "0" Then
            Count04 = Val(Count04) + 1
        End If

        Count05 = Val(Count03) + Val(Count04)
        LookCountMyDate()
        DateDivider = DateNow.Subtract(OrangDate).Days

        GenerateIncentives()

        IncGrid1.Rows.Add(IncCmb01.Text, GetNik, GetName, Count03, Count04, Count05, IncValPoints, DateDivider)

        Count05 = Nothing
        Count04 = Nothing
        Count03 = Nothing
        IncValPoints = Nothing
        IncMoney = Nothing

    End Sub

    Sub LookCountMyDate()

        SQL = ""
        SQL = SQL & "Select `Nik`, `Name`, `DateStart` From 02_Name_Table "
        SQL = SQL & "Where Nik = ('" & GetNik & " ')"
        SQL = SQL & "And Name = ('" & GetName & " ')"
        OpenTbl(ADb, Atb2, SQL)

        If Atb2.RecordCount > 0 Then

            OrangDate = IIf(IsDBNull(Atb2("DateStart").Value), "", Atb2("DateStart").Value)

        End If

    End Sub

    Private Sub IncCmb01_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles IncCmb01.Click

        UPandSlipFil()

    End Sub

    Private Sub InceTbx01_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InceTbx01.Click

        IncenGridEHeader()

        If Not IncCmb01.Text = Nothing And Not IncGenTbx1.Text = Nothing Then
            Count01 = Nothing
            Count02 = Nothing
            Count03 = Nothing

            LoadNumberDate()
            LoadNumberDate2()
            LoadCountForPer()

        Else

            MessageBox.Show("Please choose the Periode Range /  Fill the Number of Date for Days/Hari", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        End If

    End Sub

    Private Sub InceTbx02_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InceTbx02.Click
        UploadIncentives()
    End Sub

End Class