
Option Explicit On


Public Class IncentivesBlock



    Dim InceDate1 As String
    Dim InceDate2 As String
    Dim InceDate3 As String



    Private Sub IncentivesBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        InceSanSet()
        IncentivesLoad()

    End Sub


#Region "Incentives Action"

    Sub InceSanSet()

        InceDate1 = Format(Now, "yyyy")
        InceDate2 = InceDate1 - 1
        InceDate3 = InceDate2 + 1

        With InceDateCmb1

            .Items.Add("Dec " + InceDate2 + " - " + "Jan " + Format(Now, "yyyy"))
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
            .Items.Add("Dec " + Format(Now, "yyyy") + " - " + "Jan " + InceDate3)

        End With

    End Sub



    Sub IncenRangeSave()
        SQL = ""
        SQL = SQL & "Select * From 22_Incentives_Setup "
        SQL = SQL & "Where Actives = ('" & "Yes" & "') "
        SQL = SQL & "And  MonthPeriodeRange = ('" & InceDateCmb1.Text & "') "
        OpenTbl(ADb, Atbl34, SQL)

        If Not Atbl34.RecordCount <> 0 Then
            Atbl34.AddNew()
        End If


        Atbl34("MonthPeriodeRange").Value = InceDateCmb1.Text
        Atbl34("Day").Value = InceTbx1.Text
        Atbl34("Actives").Value = "Yes"

        Atbl34.Update()

        MsgBox("Done!")

    End Sub

    Sub IncenRangeSave2()

        If Not InceDateCmb1.Text = "" Then
            SQL = ""
            SQL = SQL & "Select * From 22_Incentives_Setup "
            SQL = SQL & "Where MonthPeriodeRange = ('" & InceDateCmb1.Text & "') "
            OpenTbl(ADb, Atbl34, SQL)

            If Not Atbl34.RecordCount <> 0 Then
                Atbl34.AddNew()
            End If

            Atbl34("MonthPeriodeRange").Value = InceDateCmb1.Text
            Atbl34("Day").Value = InceTbx1.Text
            Atbl34("Actives").Value = "No"

            Atbl34.Update()

        End If



    End Sub

    Sub IncentiveNest()
        InceDateCmb1.Text = ""
        InceTbx1.Text = ""
    End Sub

    Sub IncentivesLoad()

        SQL = ""
        SQL = SQL & "Select * From 22_Incentives_Setup "
        SQL = SQL & "Where Actives = ('" & "Yes" & "') "
        OpenTbl(ADb, Atbl33, SQL)

        If Atbl33.RecordCount > 0 Then

            InceDateCmb1.Text = Atbl33("MonthPeriodeRange").Value
            InceTbx1.Text = Atbl33("Day").Value

        End If

        Me.Refresh()
    End Sub

#End Region

#Region "GUI Action"

    Private Sub InceDateCmb1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles InceDateCmb1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.None) Then Exit Sub
        If Not InStr(InValid, Chr(Keys.None)) = 0 Then Chr(Keys.None = 0)
        e.Handled = True
    End Sub

    Private Sub InceTbx1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles InceTbx1.KeyPress
        If e.KeyChar.ToString = Chr(Keys.Back) Then Exit Sub
        If Not InValid2.IndexOf(e.KeyChar) = -1 Then
            e.Handled = True
        End If
    End Sub

    Private Sub InceBtn2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InceBtn2.Click
        IncenRangeSave()
        Me.Dispose()
    End Sub

    Private Sub InceBtn1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InceBtn1.Click
        InceDateCmb1.Enabled = True
        IncenRangeSave2()
        IncentiveNest()
    End Sub

#End Region



End Class