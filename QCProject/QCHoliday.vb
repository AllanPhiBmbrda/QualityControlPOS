Option Explicit On

Public Class DateCtrlBlock


    Dim TimerLimiter As Integer
    Dim DateCtrl As String

    Private Sub DateCtrlBlock_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        MainMenu.Refresh()
    End Sub

    Private Sub DateCtrlBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LoadDB()
        LoadHoliday()
        TimerLimiter = 0

    End Sub

    Sub LoadHoliday()

        HolGrid1.Rows.Clear()

        SQL = ""
        SQL = SQL & "Select * from 17_Holiday_Table "
        OpenTbl(ADb, Dbtb28, SQL)
        If Dbtb28.RecordCount > 0 Then

            If Dbtb28.RecordCount <> 0 Then

                Dbtb28.MoveFirst()
                Do While Not Dbtb28.EOF

                    Dim DateLoad As Date = IIf(IsDBNull(Dbtb28("Date").Value), Nothing, Dbtb28("Date").Value)

                    HolGrid1.Rows.Add(DateLoad.ToString("dd MMM yyyy"), Dbtb28("Holiday_Name").Value)

                    Dbtb28.MoveNext()

                Loop

            End If
        End If

    End Sub

    Sub DeleteHoliday()

        SQL = ""
        SQL = SQL & "Select * From 17_Holiday_Table "
        SQL = SQL & "Where Date = ('" & HolCal1.SelectionStart.ToString("yyyy-MM-dd") & "') "
        OpenTbl(ADb, Dbtb27, SQL)

        If Dbtb27.RecordCount > 0 Then
            Dbtb27.Delete()
            LoadHoliday()
        End If

    End Sub

    Sub SaveHoliday()
        SQL = ""
        SQL = SQL & "Select * From 17_Holiday_Table "
        SQL = SQL & "Where Date = ('" & HolCal1.SelectionStart.ToString("yyyy-MM-dd") & "') "

        OpenTbl(ADb, Dbtb27, SQL)

        If Not Dbtb27.RecordCount <> 0 Then
            Dbtb27.AddNew()
        End If

        Dbtb27("Date").Value = HolMaskTbx1.Text
        Dbtb27("Holiday_Name").Value = HolTbx1.Text
        Dbtb27("Salary_Mod").Value = "2"

        Dbtb27.Update()
        MsgBox("Holiday Saved", MsgBoxStyle.Information, "Codex ~ QC Build " & BuildCounter & " Success")
        Me.Refresh()
        HolTimer1.Enabled = True

    End Sub

    Private Sub HolBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HolBtn1.Click
        SaveHoliday()
    End Sub

    Private Sub HolMaskTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles HolMaskTbx1.KeyPress
        HolMaskTbx1.Mask = "##/##/####"
        e.Handled = True

    End Sub

    Private Sub HolCal1_DateChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles HolCal1.DateChanged
        HolMaskTbx1.Text = HolCal1.SelectionStart
        HolTbx1.Text = Nothing
        SQL = ""
        SQL = SQL & "Select * From 17_Holiday_Table "
        SQL = SQL & "Where Date = ('" & HolCal1.SelectionStart.ToString("yyyy-MM-dd") & "') "
        OpenTbl(ADb, Dbtb27, SQL)

        If Dbtb27.RecordCount > 0 Then
            HolTbx1.Text = IIf(IsDBNull(Dbtb27("Holiday_Name").Value), Nothing, Dbtb27("Holiday_Name").Value)
        End If


    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub


    Private Sub HolTimer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HolTimer1.Tick
        LoadHoliday()

        TimerLimiter = TimerLimiter + HolTimer1.Interval

        If TimerLimiter >= 3000 Then

            HolTimer1.Dispose()
        End If

    End Sub

    Private Sub HolBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HolBtn2.Click
        DeleteHoliday()
    End Sub


End Class