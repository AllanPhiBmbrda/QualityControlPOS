Public Class CouponBlock


    Dim ReSourceReport As String = Application.StartupPath + "\CrystalReport\QCCouDesign.rpt"
    Private Sub CouponBlock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CryCoup01.ReportSource = ReSourceReport



    End Sub

End Class