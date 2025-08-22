Option Explicit On



Public Class SynchBlock

    Dim SaveSynchID As String
    Dim SaveSynchDate As String
    Dim SaveSynchTime As String
    Dim SaveSynchNik As String
    Dim SaveSynchTarget As String
    Dim SaveSynchSalary As String
    Dim SaveSynchCoupon As String
    Dim SaveSynchPieces As String
    Dim SaveSynchCarton As String
    Dim SaveSynchNoKg As String
    Dim SaveSynchNoBag As String
    Dim SaveSynchNoGr As String
    Dim SaveBtmSynchDate As String
    Dim SaveBtmSynchNik As String
    Dim SaveBtmSynchSalary As String
    Dim SaveBtmSynchType As String
    Dim RecHigh As Integer
    Dim RecLow As Integer


    Private Sub MaskSynch1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MaskSynch1.KeyPress
        MaskSynch1.Mask = "##/##/####"
        If e.KeyChar.ToString = ChrW(Keys.Enter) Then
            MaskSynch2.Focus()
            e.Handled = True

        End If
    End Sub

    Private Sub MaskSynch2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MaskSynch2.KeyPress
        MaskSynch2.Mask = "##/##/####"

    End Sub


    Private Sub SynchBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SynchBtn1.Click
        DateUnknown()
    End Sub

    Sub DateUnknown()

        On Error GoTo Err

        If MaskSynch1.Text = "" Or MaskSynch2.Text = "" Then
            MsgBox("Kindly Input The Required Date")
        Else
            SynchHigh()
            SynchLow()
        End If

        Exit Sub
Err:
        SynchGrid1.Columns.Clear()
        SynchGrid2.Columns.Clear()
        MsgBox("Invalid Date")
    End Sub

#Region "Synch Load"

    Sub SynchCounterNest()
        RecHigh = 0
        RecLow = 0

    End Sub

    Sub SynchHigh()
        SynchCounterNest()

        If SynchRadio1.Checked = True Then

            SynchGrid1.Rows.Clear()
            GridHeader()

            SQL = ""
            SQL = SQL & "Select * from 03_Conveyour_Table "
            SQL = SQL & "Where Date between cdate ('" & MaskSynch1.Text & "') "
            SQL = SQL & "And cdate ('" & MaskSynch2.Text & "') "
            SQL = SQL & "Order by Nik Asc"
            OpenTbl(BBb, BBTbl1, SQL)

            If BBTbl1.RecordCount <> 0 Then

                BBTbl1.MoveFirst()
                Do While Not BBTbl1.EOF


                    SynchGrid1.Rows.Add(BBTbl1("Process_ID").Value, Format(BBTbl1("Date").Value, "dd/MM/yyyy"), BBTbl1("Time").Value, BBTbl1("Nik").Value, BBTbl1("Pieces").Value, BBTbl1("Target").Value, BBTbl1("Salary").Value)

                    BBTbl1.MoveNext()

                    RecHigh = RecHigh + 1
                    RecTbx1.Text = RecHigh

                Loop

            End If

        ElseIf SynchRadio2.Checked = True Then
            SynchGrid1.Rows.Clear()
            GridHeader()

            SQL = ""
            SQL = SQL & "Select * from 04_MutuII_Table  "
            SQL = SQL & "Where Date between cdate ('" & MaskSynch1.Text & "') "
            SQL = SQL & "And cdate ('" & MaskSynch2.Text & "') "
            SQL = SQL & "Order by Nik Asc"
            OpenTbl(BBb, BBTbl1, SQL)

            If BBTbl1.RecordCount <> 0 Then

                BBTbl1.MoveFirst()
                Do While Not BBTbl1.EOF

                    SynchGrid1.Rows.Add(BBTbl1("Process_ID").Value, Format(BBTbl1("Date").Value, "dd/MM/yyyy"), BBTbl1("Time").Value, BBTbl1("Nik").Value, BBTbl1("Pieces").Value, BBTbl1("Target").Value, BBTbl1("Salary").Value, BBTbl1("Coupon").Value)

                    BBTbl1.MoveNext()

                    RecHigh = RecHigh + 1
                    RecTbx1.Text = RecHigh

                Loop

            End If

        ElseIf SynchRadio3.Checked = True Then
            SynchGrid1.Rows.Clear()
            GridHeader()

            SQL = ""
            SQL = SQL & "Select * from 06_Wallet_Table "
            SQL = SQL & "Where Date between cdate ('" & MaskSynch1.Text & "') "
            SQL = SQL & "And cdate ('" & MaskSynch2.Text & "') "
            SQL = SQL & "Order by Nik Asc"
            OpenTbl(BBb, BBTbl1, SQL)

            If BBTbl1.RecordCount <> 0 Then

                BBTbl1.MoveFirst()
                Do While Not BBTbl1.EOF

                    SynchGrid1.Rows.Add(BBTbl1("Process_ID").Value, Format(BBTbl1("Date").Value, "dd/MM/yyyy"), BBTbl1("Time").Value, BBTbl1("Nik").Value, BBTbl1("Pieces").Value, BBTbl1("Target").Value, BBTbl1("Salary").Value, BBTbl1("Coupon").Value)

                    BBTbl1.MoveNext()

                    RecHigh = RecHigh + 1
                    RecTbx1.Text = RecHigh

                Loop

            End If

        ElseIf SynchRadio4.Checked = True Then
            SynchGrid1.Rows.Clear()
            GridHeader()

            SQL = ""
            SQL = SQL & "Select * from 05_Packing_Table "
            SQL = SQL & "Where Date between cdate ('" & MaskSynch1.Text & "') "
            SQL = SQL & "And cdate ('" & MaskSynch2.Text & "') "
            SQL = SQL & "Order by Nik Asc"
            OpenTbl(BBb, BBTbl1, SQL)

            If BBTbl1.RecordCount <> 0 Then

                BBTbl1.MoveFirst()
                Do While Not BBTbl1.EOF

                    SynchGrid1.Rows.Add(BBTbl1("Process_ID").Value, Format(BBTbl1("Date").Value, "dd/MM/yyyy"), BBTbl1("Time").Value, BBTbl1("Nik").Value, BBTbl1("Target").Value, BBTbl1("Salary").Value, BBTbl1("Coupon").Value, BBTbl1("Carton").Value)

                    BBTbl1.MoveNext()

                    RecHigh = RecHigh + 1
                    RecTbx1.Text = RecHigh
                Loop

            End If

        ElseIf SynchRadio5.Checked = True Then
            SynchGrid1.Rows.Clear()
            GridHeader()

            SQL = ""
            SQL = SQL & "Select * from 19_Miscellaneous_Table "
            SQL = SQL & "Where Date between cdate ('" & MaskSynch1.Text & "') "
            SQL = SQL & "And cdate ('" & MaskSynch2.Text & "') "
            SQL = SQL & "Order by Nik Asc"
            OpenTbl(BBb, BBTbl1, SQL)

            If BBTbl1.RecordCount <> 0 Then

                BBTbl1.MoveFirst()
                Do While Not BBTbl1.EOF

                    SynchGrid1.Rows.Add(BBTbl1("Process_ID").Value, Format(BBTbl1("Date").Value, "dd/MM/yyyy"), BBTbl1("Time").Value, BBTbl1("Nik").Value, BBTbl1("Salary").Value)

                    BBTbl1.MoveNext()

                    RecHigh = RecHigh + 1
                    RecTbx1.Text = RecHigh
                Loop

            End If

        ElseIf SynchRadio6.Checked = True Then
            SynchGrid1.Rows.Clear()
            GridHeader()

            SQL = ""
            SQL = SQL & "Select * from 21_NewMiscellaneous_Table "
            SQL = SQL & "Where Date between cdate ('" & MaskSynch1.Text & "') "
            SQL = SQL & "And cdate ('" & MaskSynch2.Text & "') "
            SQL = SQL & "Order by Nik Asc"
            OpenTbl(BBb, BBTbl1, SQL)

            If BBTbl1.RecordCount <> 0 Then

                BBTbl1.MoveFirst()
                Do While Not BBTbl1.EOF

                    SynchGrid1.Rows.Add(BBTbl1("Process_ID").Value, Format(BBTbl1("Date").Value, "dd/MM/yyyy"), BBTbl1("Time").Value, BBTbl1("Nik").Value, BBTbl1("Coupon").Value, BBTbl1("NoKg").Value, BBTbl1("NoBag").Value, BBTbl1("NoGr").Value, BBTbl1("Pieces").Value, BBTbl1("Salary").Value)

                    BBTbl1.MoveNext()

                    RecHigh = RecHigh + 1
                    RecTbx1.Text = RecHigh

                Loop

            End If
        Else
            MsgBox("Select Department First")

        End If
    End Sub

    Sub SynchLow()

        SynchCounterNest()

        If SynchRadio1.Checked = True Then
            SynchGrid2.Rows.Clear()
            GridHeader2()

            SQL = ""
            SQL = SQL & "Select * from 13_Conveyour_Salary "
            SQL = SQL & "Where Date between cdate('" & MaskSynch1.Text & "') "
            SQL = SQL & "And cdate('" & MaskSynch2.Text & "') "
            SQL = SQL & "Order by Nik Asc"
            OpenTbl(BBb, BBTbl2, SQL)
            If BBTbl2.RecordCount <> 0 Then

                BBTbl2.MoveFirst()
                Do While Not BBTbl2.EOF
                    SynchGrid2.Rows.Add(Format(BBTbl2("Date").Value, "dd/MM/yyyy"), BBTbl2("Nik").Value, BBTbl2("Salary").Value)
                    BBTbl2.MoveNext()

                    RecLow = RecLow + 1
                    RecTbx2.Text = RecLow
                Loop

            End If

        ElseIf SynchRadio2.Checked = True Then
            SynchGrid2.Rows.Clear()
            GridHeader2()

            SQL = ""
            SQL = SQL & "Select * from 14_MutuII_Salary "
            SQL = SQL & "Where Date between cdate('" & MaskSynch1.Text & "') "
            SQL = SQL & "And cdate('" & MaskSynch2.Text & "') "
            SQL = SQL & "Order by Nik Asc"
            OpenTbl(BBb, BBTbl2, SQL)
            If BBTbl2.RecordCount <> 0 Then

                BBTbl2.MoveFirst()
                Do While Not BBTbl2.EOF
                    SynchGrid2.Rows.Add(Format(BBTbl2("Date").Value, "dd/MM/yyyy"), BBTbl2("Nik").Value, BBTbl2("Salary").Value)
                    BBTbl2.MoveNext()

                    RecLow = RecLow + 1
                    RecTbx2.Text = RecLow
                Loop

            End If

        ElseIf SynchRadio3.Checked = True Then
            SynchGrid2.Rows.Clear()
            GridHeader2()

            SQL = ""
            SQL = SQL & "Select * from 15_Wallet_Salary "
            SQL = SQL & "Where Date between cdate('" & MaskSynch1.Text & "') "
            SQL = SQL & "And cdate('" & MaskSynch2.Text & "') "
            SQL = SQL & "Order by Nik Asc"
            OpenTbl(BBb, BBTbl2, SQL)
            If BBTbl2.RecordCount <> 0 Then

                BBTbl2.MoveFirst()
                Do While Not BBTbl2.EOF
                    SynchGrid2.Rows.Add(Format(BBTbl2("Date").Value, "dd/MM/yyyy"), BBTbl2("Nik").Value, BBTbl2("Salary").Value)
                    BBTbl2.MoveNext()

                    RecLow = RecLow + 1
                    RecTbx2.Text = RecLow
                Loop

            End If

        ElseIf SynchRadio4.Checked = True Then

            SynchGrid2.Rows.Clear()
            GridHeader2()

            SQL = ""
            SQL = SQL & "Select * from 16_Packing_Salary "
            SQL = SQL & "Where Date between cdate('" & MaskSynch1.Text & "') "
            SQL = SQL & "And cdate('" & MaskSynch2.Text & "') "
            SQL = SQL & "Order by Nik Asc"
            OpenTbl(BBb, BBTbl2, SQL)
            If BBTbl2.RecordCount <> 0 Then

                BBTbl2.MoveFirst()

                Do While Not BBTbl2.EOF

                    SynchGrid2.Rows.Add(Format(BBTbl2("Date").Value, "dd/MM/yyyy"), BBTbl2("Nik").Value, BBTbl2("Salary").Value)
                    BBTbl2.MoveNext()

                    RecLow = RecLow + 1
                    RecTbx2.Text = RecLow

                Loop

            End If

        ElseIf SynchRadio5.Checked = True Then

            SynchGrid2.Rows.Clear()
            GridHeader2()

            SQL = ""
            SQL = SQL & "Select * from 20_Miscellaneous_Salary "
            SQL = SQL & "Where Date between cdate('" & MaskSynch1.Text & "') "
            SQL = SQL & "And cdate('" & MaskSynch2.Text & "') "
            SQL = SQL & "And TypeCtrl = ('" & "Old" & "') "
            SQL = SQL & "Order by Nik Asc"
            OpenTbl(BBb, BBTbl2, SQL)
            If BBTbl2.RecordCount <> 0 Then
                BBTbl2.MoveFirst()

                Do While Not BBTbl2.EOF

                    SynchGrid2.Rows.Add(Format(BBTbl2("Date").Value, "dd/MM/yyyy"), BBTbl2("Nik").Value, BBTbl2("Salary").Value, BBTbl2("TypeCtrl").Value)
                    BBTbl2.MoveNext()

                    RecLow = RecLow + 1
                    RecTbx2.Text = RecLow

                Loop

            End If

        ElseIf SynchRadio6.Checked = True Then

            SynchGrid2.Rows.Clear()
            GridHeader2()

            SQL = ""
            SQL = SQL & "Select * from 20_Miscellaneous_Salary "
            SQL = SQL & "Where Date between cdate('" & MaskSynch1.Text & "') "
            SQL = SQL & "And cdate('" & MaskSynch2.Text & "') "
            SQL = SQL & "And TypeCtrl = ('" & "New" & "') "
            SQL = SQL & "Order by Nik Asc"
            OpenTbl(BBb, BBTbl2, SQL)
            If BBTbl2.RecordCount <> 0 Then

                BBTbl2.MoveFirst()
                Do While Not BBTbl2.EOF
                    SynchGrid2.Rows.Add(Format(BBTbl2("Date").Value, "dd/MM/yyyy"), BBTbl2("Nik").Value, BBTbl2("Salary").Value, BBTbl2("TypeCtrl").Value)
                    BBTbl2.MoveNext()

                    RecLow = RecLow + 1
                    RecTbx2.Text = RecLow
                Loop

            End If

        End If

    End Sub

#End Region

#Region "Grid Header Text"

    Sub GridHeader()
        With SynchGrid1

            If SynchRadio1.Checked = True Then
                .Columns.Add("Col1", "Process ID")
                .Columns.Add("Col2", "Date")
                .Columns.Add("Col3", "Time")
                .Columns.Add("Col4", "Nik")
                .Columns.Add("Col5", "Pieces")
                .Columns.Add("Col6", "Target")
                .Columns.Add("Col7", "Salary")
                .Columns.Add("Col8", "Remarks")

            ElseIf SynchRadio2.Checked = True Or SynchRadio3.Checked = True Then
                .Columns.Add("Col1", "Process ID")
                .Columns.Add("Col2", "Date")
                .Columns.Add("Col3", "Time")
                .Columns.Add("Col4", "Nik")
                .Columns.Add("Col5", "Pieces")
                .Columns.Add("Col6", "Target")
                .Columns.Add("Col7", "Salary")
                .Columns.Add("Col8", "Coupon")
                .Columns.Add("Col9", "Remarks")

            ElseIf SynchRadio4.Checked = True Then
                .Columns.Add("Col1", "Process ID")
                .Columns.Add("Col2", "Date")
                .Columns.Add("Col3", "Time")
                .Columns.Add("Col4", "Nik")
                .Columns.Add("Col5", "Target")
                .Columns.Add("Col6", "Salary")
                .Columns.Add("Col7", "Coupon")
                .Columns.Add("Col8", "Carton")
                .Columns.Add("Col9", "Remarks")

            ElseIf SynchRadio5.Checked = True Then
                .Columns.Add("Col1", "Process ID")
                .Columns.Add("Col2", "Date")
                .Columns.Add("Col3", "Time")
                .Columns.Add("Col4", "Nik")
                .Columns.Add("Col6", "Salary")
                .Columns.Add("Col9", "Remarks")

            ElseIf SynchRadio6.Checked = True Then
                .Columns.Add("Col1", "Process ID")
                .Columns.Add("Col2", "Date")
                .Columns.Add("Col3", "Time")
                .Columns.Add("Col4", "Nik")
                .Columns.Add("Col5", "Coupon")
                .Columns.Add("Col6", "No. of Kilogram")
                .Columns.Add("Col7", "No. of Bag")
                .Columns.Add("Col8", "No. of Gram")
                .Columns.Add("Col9", "Pieces")
                .Columns.Add("Col10", "Salary")
                .Columns.Add("Col11", "Remarks")
                .Columns(5).Width = 120

            End If

        End With

    End Sub

    Sub GridHeader2()

        With SynchGrid2

            If SynchRadio1.Checked = True Or SynchRadio2.Checked = True Or SynchRadio3.Checked = True Or SynchRadio4.Checked = True Then

                .Columns.Add("Col1b", "Date")
                .Columns.Add("Col2b", "Nik")
                .Columns.Add("Col3b", "Salary")
                .Columns.Add("Col34", "Remark")

            ElseIf SynchRadio5.Checked = True Or SynchRadio6.Checked = True Then

                .Columns.Add("Col1b", "Date")
                .Columns.Add("Col2b", "Nik")
                .Columns.Add("Col3b", "Salary")
                .Columns.Add("Col4b", "Control Type")
                .Columns.Add("Col34", "Remark")

            End If

        End With

    End Sub

#End Region

    Private Sub SynchBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LoadDB()
        LoadDB2()
        LoadDB3()
        MaskSynch1.Focus()

    End Sub

    Private Sub SynchBtn3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SynchBtn3.Click

        SynchGrid1.Rows.Clear()
        SynchGrid1.Columns.Clear()
        SynchGrid2.Rows.Clear()
        SynchGrid2.Columns.Clear()
        RecTbx1.Text = ""
        RecTbx2.Text = ""

    End Sub

#Region "Synch Update"

    Sub SaveSynchNest()

        SaveSynchID = ""
        SaveSynchDate = ""
        SaveSynchTime = ""
        SaveSynchNik = ""
        SaveSynchTarget = ""
        SaveSynchSalary = ""
        SaveSynchCoupon = ""
        SaveSynchPieces = ""
        SaveSynchCarton = ""
        SaveSynchNoKg = ""
        SaveSynchNoBag = ""
        SaveSynchNoGr = ""
        SaveBtmSynchDate = ""
        SaveBtmSynchNik = ""
        SaveBtmSynchSalary = ""
        SaveBtmSynchType = ""

    End Sub

    Sub UpConveyour()

        SaveSynchNest()

        For i = 0 To SynchGrid1.Rows.Count - 1

            SaveSynchID = SynchGrid1(0, i).Value
            SaveSynchDate = SynchGrid1(1, i).Value
            SaveSynchTime = SynchGrid1(2, i).Value
            SaveSynchNik = SynchGrid1(3, i).Value
            SaveSynchPieces = SynchGrid1(4, i).Value
            SaveSynchTarget = SynchGrid1(5, i).Value
            SaveSynchSalary = SynchGrid1(6, i).Value

            SQL = ""
            SQL = SQL & "Select * From 03_Conveyour_Table "
            SQL = SQL & "Where Date = ('" & SaveSynchDate & "') "
            SQL = SQL & "And Nik = ('" & SaveSynchNik & "') "
            SQL = SQL & "And Time = ('" & SaveSynchTime & "') "
            SQL = SQL & "And Process_ID = ('" & SaveSynchID & "') "
            OpenTbl(ADb, Atbl26, SQL)

            If Not Atbl26.RecordCount <> 0 Then

                Atbl26.AddNew()

                Atbl26("Process_ID").Value = SaveSynchID
                Atbl26("Date").Value = SaveSynchDate
                Atbl26("Time").Value = SaveSynchTime
                Atbl26("Nik").Value = SaveSynchNik
                Atbl26("Pieces").Value = SaveSynchPieces
                Atbl26("Target").Value = SaveSynchTarget
                Atbl26("Salary").Value = SaveSynchSalary
                Atbl26.Update()

                SynchGrid1(7, i).Value = "Has Been Saved"

            ElseIf Atbl26.RecordCount > 0 Then

                SynchGrid1(7, i).Value = "Already Exist"

            End If

        Next

    End Sub

    Sub UpMutuII()

        SaveSynchNest()

        For i = 0 To SynchGrid1.Rows.Count - 1

            SaveSynchID = SynchGrid1(0, i).Value
            SaveSynchDate = SynchGrid1(1, i).Value
            SaveSynchTime = SynchGrid1(2, i).Value
            SaveSynchNik = SynchGrid1(3, i).Value
            SaveSynchPieces = SynchGrid1(4, i).Value
            SaveSynchTarget = SynchGrid1(5, i).Value
            SaveSynchSalary = SynchGrid1(6, i).Value
            SaveSynchCoupon = SynchGrid1(7, i).Value

            SQL = ""
            SQL = SQL & "Select * From 04_MutuII_Table "
            SQL = SQL & "Where Date = ('" & SaveSynchDate & "') "
            SQL = SQL & "And Nik = ('" & SaveSynchNik & "') "
            SQL = SQL & "And Time = ('" & SaveSynchTime & "') "
            SQL = SQL & "And Process_ID = ('" & SaveSynchID & "') "
            OpenTbl(ADb, Atbl26, SQL)

            If Not Atbl26.RecordCount <> 0 Then
                Atbl26.AddNew()

                Atbl26("Process_ID").Value = SaveSynchID
                Atbl26("Date").Value = SaveSynchDate
                Atbl26("Time").Value = SaveSynchTime
                Atbl26("Nik").Value = SaveSynchNik
                Atbl26("Pieces").Value = SaveSynchPieces
                Atbl26("Target").Value = SaveSynchTarget
                Atbl26("Salary").Value = SaveSynchSalary
                Atbl26("Coupon").Value = SaveSynchCoupon
                Atbl26.Update()

                SynchGrid1(8, i).Value = "Has Been Saved"

            ElseIf Atbl26.RecordCount > 0 Then

                SynchGrid1(8, i).Value = "Already Exist"

            End If

        Next

    End Sub

    Sub UpWallet()

        SaveSynchNest()

        For i = 0 To SynchGrid1.Rows.Count - 1

            SaveSynchID = SynchGrid1(0, i).Value
            SaveSynchDate = SynchGrid1(1, i).Value
            SaveSynchTime = SynchGrid1(2, i).Value
            SaveSynchNik = SynchGrid1(3, i).Value
            SaveSynchPieces = SynchGrid1(4, i).Value
            SaveSynchTarget = SynchGrid1(5, i).Value
            SaveSynchSalary = SynchGrid1(6, i).Value
            SaveSynchCoupon = SynchGrid1(7, i).Value

            SQL = ""
            SQL = SQL & "Select * From 06_Wallet_Table "
            SQL = SQL & "Where Date = ('" & SaveSynchDate & "') "
            SQL = SQL & "And Nik = ('" & SaveSynchNik & "') "
            SQL = SQL & "And Time = ('" & SaveSynchTime & "') "
            SQL = SQL & "And Process_ID = ('" & SaveSynchID & "') "
            OpenTbl(ADb, Atbl26, SQL)

            If Not Atbl26.RecordCount <> 0 Then
                Atbl26.AddNew()

                Atbl26("Process_ID").Value = SaveSynchID
                Atbl26("Date").Value = SaveSynchDate
                Atbl26("Time").Value = SaveSynchTime
                Atbl26("Nik").Value = SaveSynchNik
                Atbl26("Pieces").Value = SaveSynchPieces
                Atbl26("Target").Value = SaveSynchTarget
                Atbl26("Salary").Value = SaveSynchSalary
                Atbl26("Coupon").Value = SaveSynchCoupon
                Atbl26.Update()

                SynchGrid1(8, i).Value = "Has Been Saved"

            ElseIf Atbl26.RecordCount > 0 Then

                SynchGrid1(8, i).Value = "Already Exist"

            End If

        Next

    End Sub

    Sub UpPacking()

        SaveSynchNest()

        For i = 0 To SynchGrid1.Rows.Count - 1

            SaveSynchID = SynchGrid1(0, i).Value
            SaveSynchDate = SynchGrid1(1, i).Value
            SaveSynchTime = SynchGrid1(2, i).Value
            SaveSynchNik = SynchGrid1(3, i).Value
            SaveSynchTarget = SynchGrid1(4, i).Value
            SaveSynchSalary = SynchGrid1(5, i).Value
            SaveSynchCoupon = SynchGrid1(6, i).Value
            SaveSynchCarton = SynchGrid1(7, i).Value

            SQL = ""
            SQL = SQL & "Select * From 05_Packing_Table "
            SQL = SQL & "Where Date = ('" & SaveSynchDate & "') "
            SQL = SQL & "And Nik = ('" & SaveSynchNik & "') "
            SQL = SQL & "And Time = ('" & SaveSynchTime & "') "
            SQL = SQL & "And Process_ID = ('" & SaveSynchID & "') "
            OpenTbl(ADb, Atbl26, SQL)

            If Not Atbl26.RecordCount <> 0 Then
                Atbl26.AddNew()

                Atbl26("Process_ID").Value = SaveSynchID
                Atbl26("Date").Value = SaveSynchDate
                Atbl26("Time").Value = SaveSynchTime
                Atbl26("Nik").Value = SaveSynchNik
                Atbl26("Target").Value = SaveSynchTarget
                Atbl26("Salary").Value = SaveSynchSalary
                Atbl26("Coupon").Value = SaveSynchCoupon
                Atbl26("Carton").Value = SaveSynchCarton
                Atbl26.Update()

                SynchGrid1(8, i).Value = "Has Been Saved"

            ElseIf Atbl26.RecordCount > 0 Then
                SynchGrid1(8, i).Value = "Already Exist"

            End If
        Next

    End Sub

    Sub UpMiscellaneous()

        SaveSynchNest()

        For i = 0 To SynchGrid1.Rows.Count - 1

            SaveSynchID = SynchGrid1(0, i).Value
            SaveSynchDate = SynchGrid1(1, i).Value
            SaveSynchTime = SynchGrid1(2, i).Value
            SaveSynchNik = SynchGrid1(3, i).Value
            SaveSynchSalary = SynchGrid1(4, i).Value

            SQL = ""
            SQL = SQL & "Select * From 19_Miscellaneous_Table "
            SQL = SQL & "Where Date = ('" & SaveSynchDate & "') "
            SQL = SQL & "And Nik = ('" & SaveSynchNik & "') "
            SQL = SQL & "And Time = ('" & SaveSynchTime & "') "
            SQL = SQL & "And Process_ID = ('" & SaveSynchID & "') "
            OpenTbl(ADb, Atbl26, SQL)

            If Not Atbl26.RecordCount <> 0 Then
                Atbl26.AddNew()


                Atbl26("Process_ID").Value = SaveSynchID
                Atbl26("Date").Value = SaveSynchDate
                Atbl26("Time").Value = SaveSynchTime
                Atbl26("Nik").Value = SaveSynchNik
                Atbl26("Salary").Value = SaveSynchSalary
                Atbl26.Update()

                SynchGrid1(5, i).Value = "Has Been Saved"

            ElseIf Atbl26.RecordCount > 0 Then
                SynchGrid1(5, i).Value = "Already Exist"


            End If

        Next
    End Sub

    Sub UpSortasi()


        SaveSynchNest()

        For i = 0 To SynchGrid1.Rows.Count - 1

            SaveSynchID = SynchGrid1(0, i).Value
            SaveSynchDate = SynchGrid1(1, i).Value
            SaveSynchTime = SynchGrid1(2, i).Value
            SaveSynchNik = SynchGrid1(3, i).Value
            SaveSynchCoupon = SynchGrid1(4, i).Value
            SaveSynchNoKg = SynchGrid1(5, i).Value
            SaveSynchNoBag = SynchGrid1(6, i).Value
            SaveSynchNoGr = SynchGrid1(7, i).Value
            SaveSynchPieces = SynchGrid1(8, i).Value
            SaveSynchSalary = SynchGrid1(9, i).Value

            SQL = ""
            SQL = SQL & "Select * From 21_NewMiscellaneous_Table "
            SQL = SQL & "Where Date = ('" & SaveSynchDate & "') "
            SQL = SQL & "And Nik = ('" & SaveSynchNik & "') "
            SQL = SQL & "And Time = ('" & SaveSynchTime & "') "
            SQL = SQL & "And Process_ID = ('" & SaveSynchID & "') "
            OpenTbl(ADb, Atbl26, SQL)

            If Not Atbl26.RecordCount <> 0 Then
                Atbl26.AddNew()


                Atbl26("Process_ID").Value = SaveSynchID
                Atbl26("Date").Value = SaveSynchDate
                Atbl26("Time").Value = SaveSynchTime
                Atbl26("Nik").Value = SaveSynchNik
                Atbl26("NoKg").Value = SaveSynchNoKg
                Atbl26("NoBag").Value = SaveSynchNoBag
                Atbl26("NoGr").Value = SaveSynchNoGr
                Atbl26("Pieces").Value = SaveSynchPieces
                Atbl26("Salary").Value = SaveSynchSalary
                Atbl26("Coupon").Value = SaveSynchCoupon
                Atbl26.Update()


                SynchGrid1(10, i).Value = "Has Been Saved"

            ElseIf Atbl26.RecordCount > 0 Then

                SynchGrid1(10, i).Value = "Already Exist"

            End If


        Next



    End Sub

    '---- Synch Salary Zone----

    Sub UpSalConveyour()


        For i = 0 To SynchGrid2.Rows.Count - 1

            SaveBtmSynchDate = SynchGrid2(0, i).Value
            SaveBtmSynchNik = SynchGrid2(1, i).Value
            SaveBtmSynchSalary = SynchGrid2(2, i).Value

            SQL = ""
            SQL = SQL & "Select * From 13_Conveyour_Salary "
            SQL = SQL & "Where Date = ('" & SaveBtmSynchDate & "') "
            SQL = SQL & "And Nik = ('" & SaveBtmSynchNik & "') "
            OpenTbl(ADb, Atbl27, SQL)

            If Not Atbl27.RecordCount <> 0 Then
                Atbl27.AddNew()

                Atbl27("Date").Value = SaveBtmSynchDate
                Atbl27("Nik").Value = SaveBtmSynchNik
                Atbl27("Salary").Value = SaveBtmSynchSalary
                Atbl27.Update()

                SynchGrid2(3, i).Value = "Has Been Saved"

            ElseIf Atbl27.RecordCount > 0 Then
                SynchGrid2(3, i).Value = "Already Exist"

            End If
        Next


    End Sub

    Sub UpSalMutuII()

        For i = 0 To SynchGrid2.Rows.Count - 1

            SaveBtmSynchDate = SynchGrid2(0, i).Value
            SaveBtmSynchNik = SynchGrid2(1, i).Value
            SaveBtmSynchSalary = SynchGrid2(2, i).Value

            SQL = ""
            SQL = SQL & "Select * From 14_MutuII_Salary "
            SQL = SQL & "Where Date = ('" & SaveBtmSynchDate & "') "
            SQL = SQL & "And Nik = ('" & SaveBtmSynchNik & "') "
            OpenTbl(ADb, Atbl27, SQL)

            If Not Atbl27.RecordCount <> 0 Then
                Atbl27.AddNew()

                Atbl27("Date").Value = SaveBtmSynchDate
                Atbl27("Nik").Value = SaveBtmSynchNik
                Atbl27("Salary").Value = SaveBtmSynchSalary
                Atbl27.Update()

                SynchGrid2(3, i).Value = "Has Been Saved"

            ElseIf Atbl27.RecordCount > 0 Then

                SynchGrid2(3, i).Value = "Already Exist"

            End If
        Next

    End Sub

    Sub UpSalWallet()


        For i = 0 To SynchGrid2.Rows.Count - 1

            SaveBtmSynchDate = SynchGrid2(0, i).Value
            SaveBtmSynchNik = SynchGrid2(1, i).Value
            SaveBtmSynchSalary = SynchGrid2(2, i).Value

            SQL = ""
            SQL = SQL & "Select * From 15_Wallet_Salary "
            SQL = SQL & "Where Date = ('" & SaveBtmSynchDate & "') "
            SQL = SQL & "And Nik = ('" & SaveBtmSynchNik & "') "
            OpenTbl(ADb, Atbl27, SQL)

            If Not Atbl27.RecordCount <> 0 Then
                Atbl27.AddNew()


                Atbl27("Date").Value = SaveBtmSynchDate
                Atbl27("Nik").Value = SaveBtmSynchNik
                Atbl27("Salary").Value = SaveBtmSynchSalary
                Atbl27.Update()


                SynchGrid2(3, i).Value = "Has Been Saved"

            ElseIf Atbl27.RecordCount > 0 Then

                SynchGrid2(3, i).Value = "Already Exist"


            End If
        Next


    End Sub

    Sub UpSalPacking()


        For i = 0 To SynchGrid2.Rows.Count - 1

            SaveBtmSynchDate = SynchGrid2(0, i).Value
            SaveBtmSynchNik = SynchGrid2(1, i).Value
            SaveBtmSynchSalary = SynchGrid2(2, i).Value

            SQL = ""
            SQL = SQL & "Select * From 16_Packing_Salary "
            SQL = SQL & "Where Date = ('" & SaveBtmSynchDate & "') "
            SQL = SQL & "And Nik = ('" & SaveBtmSynchNik & "') "
            OpenTbl(ADb, Atbl27, SQL)

            If Not Atbl27.RecordCount <> 0 Then
                Atbl27.AddNew()


                Atbl27("Date").Value = SaveBtmSynchDate
                Atbl27("Nik").Value = SaveBtmSynchNik
                Atbl27("Salary").Value = SaveBtmSynchSalary
                Atbl27.Update()

                SynchGrid2(3, i).Value = "Has Been Saved"

            ElseIf Atbl27.RecordCount > 0 Then

                SynchGrid2(3, i).Value = "Already Exist"


            End If
        Next

    End Sub

    Sub UpSalMiscellaneous() ' Miscellaneous and Sortasi Same Update Code 


        For i = 0 To SynchGrid2.Rows.Count - 1

            SaveBtmSynchDate = SynchGrid2(0, i).Value
            SaveBtmSynchNik = SynchGrid2(1, i).Value
            SaveBtmSynchSalary = SynchGrid2(2, i).Value
            SaveBtmSynchType = SynchGrid2(3, i).Value

            SQL = ""
            SQL = SQL & "Select * From 20_Miscellaneous_Salary "
            SQL = SQL & "Where Date = ('" & SaveBtmSynchDate & "') "
            SQL = SQL & "And Nik = ('" & SaveBtmSynchNik & "') "
            OpenTbl(ADb, Atbl27, SQL)

            If Not Atbl27.RecordCount <> 0 Then
                Atbl27.AddNew()


                Atbl27("Date").Value = SaveBtmSynchDate
                Atbl27("Nik").Value = SaveBtmSynchNik
                Atbl27("Salary").Value = SaveBtmSynchSalary
                Atbl27("TypeCtrl").Value = SaveBtmSynchType
                Atbl27.Update()

                SynchGrid2(4, i).Value = "Has Been Saved"

            ElseIf Atbl27.RecordCount > 0 Then

                SynchGrid2(4, i).Value = "Already Exist"


            End If
        Next

    End Sub



#End Region

    Private Sub SynchBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SynchBtn2.Click

        If SynchRadio1.Checked = True Then
            UpConveyour()
            UpSalConveyour()

        ElseIf SynchRadio2.Checked = True Then
            UpMutuII()
            UpSalMutuII()

        ElseIf SynchRadio3.Checked = True Then
            UpWallet()
            UpSalWallet()

        ElseIf SynchRadio4.Checked = True Then
            UpPacking()
            UpSalPacking()

        ElseIf SynchRadio5.Checked = True Then
            UpMiscellaneous()
            UpSalMiscellaneous()

        ElseIf SynchRadio6.Checked = True Then
            UpSortasi()
            UpSalMiscellaneous()

        End If

    End Sub






End Class