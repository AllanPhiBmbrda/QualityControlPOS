Imports CrystalDecisions

Public Class KeluarReportView

    Dim ReSourceReport As String = Application.StartupPath + "\CrystalReport\QCCrystal.rpt"

    Private Sub KeluarReportView_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        CryView01.ReportSource = ReSourceReport

    End Sub

    Private Sub reportDocument1_InitReport(ByVal sender As System.Object, ByVal e As System.EventArgs)



    End Sub

End Class