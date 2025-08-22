
Imports System
Imports System.Globalization
Imports System.Drawing
Imports Microsoft.Reporting.WinForms

Public Class KeluarNewKerja02

    Dim D01, D02, D03, D04, D05, D06, D07, D08, D09, D10, D11, D12, D13, D14, D15, D16 As Date
    Dim DD01, DD02, DD03, DD04, DD05, DD06, DD07, DD08, DD09, DD10, DD11, DD12, DD13, DD14, DD15, DD16 As String
    Dim G01, G02, G03, G04, G05, G06, G07, G08, G09, G10, G11, G12, G13, G14, G15, G16 As Double
    Dim GG01, GG02, GG03, GG04, GG05, GG06, GG07, GG08, GG09, GG10, GG11, GG12, GG13, GG14, GG15, GG16 As String

    Dim customNumberInfo As CultureInfo = CultureInfo.CreateSpecificCulture("id-ID")

    Dim AName, ANik, AAla, AKel, ATMK As String
    Dim ATot, AGTot, APot, ATun As Double
    Dim ATot2, AGTot2, APot2, ATun2 As String


    Dim SignA, SignB As String

    Dim ADate1 As String
    Dim ADate2 As String
    Private Sub KeluarNewKerja02_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LoadDB()
        LoadDB2()

        'ReControl()
        LocData()
        FillerParameter()
    End Sub

    Sub ReControl()

        Dim MyPageSetting As New System.Drawing.Printing.PageSettings()
        MyPageSetting.Margins = New System.Drawing.Printing.Margins(0.4, 0.1, 0, 0)


        Me.KeluarViewer.SetPageSettings(MyPageSetting)

    End Sub
    Sub FillerParameter()

        KeluarViewer.LocalReport.ReportPath = Application.StartupPath + "\KeluarFlowReport.rdlc"
        Dim g(45) As ReportParameter


        ' Salary Report
        g(0) = New ReportParameter("GV01", GG01)
        g(1) = New ReportParameter("GV02", GG02)
        g(2) = New ReportParameter("GV03", GG03)
        g(3) = New ReportParameter("GV04", GG04)
        g(4) = New ReportParameter("GV05", GG05)
        g(5) = New ReportParameter("GV06", GG06)
        g(6) = New ReportParameter("GV07", GG07)
        g(7) = New ReportParameter("GV08", GG08)
        g(8) = New ReportParameter("GV09", GG09)
        g(9) = New ReportParameter("GV10", GG10)
        g(10) = New ReportParameter("GV11", GG11)
        g(11) = New ReportParameter("GV12", GG12)
        g(12) = New ReportParameter("GV13", GG13)
        g(13) = New ReportParameter("GV14", GG14)
        g(14) = New ReportParameter("GV15", GG15)
        g(15) = New ReportParameter("GV16", GG16)

        ' Data of Person
        g(16) = New ReportParameter("VNik", ANik)
        g(17) = New ReportParameter("VName", AName)
        g(18) = New ReportParameter("VKel", AKel)
        g(19) = New ReportParameter("VTMK", ATMK)
        g(20) = New ReportParameter("VStart", ADate1)
        g(21) = New ReportParameter("VFin", ADate2)
        g(22) = New ReportParameter("VTot", ATot2)
        g(23) = New ReportParameter("VTun", ATun2)
        g(24) = New ReportParameter("VPot", APot2)
        g(25) = New ReportParameter("VGaDi", AGTot2)
        g(26) = New ReportParameter("VCrDa", "Patumbak, " + Format(Today, "dddd, dd MMMM yyyy"))
        g(27) = New ReportParameter("VFrom", SignA)
        g(28) = New ReportParameter("VTo", SignB)

        ' Date
        g(29) = New ReportParameter("DV01", DD01)
        g(30) = New ReportParameter("DV02", DD02)
        g(31) = New ReportParameter("DV03", DD03)
        g(32) = New ReportParameter("DV04", DD04)
        g(33) = New ReportParameter("DV05", DD05)
        g(34) = New ReportParameter("DV06", DD06)
        g(35) = New ReportParameter("DV07", DD07)
        g(36) = New ReportParameter("DV08", DD08)
        g(37) = New ReportParameter("DV09", DD09)
        g(38) = New ReportParameter("DV10", DD10)
        g(39) = New ReportParameter("DV11", DD11)
        g(40) = New ReportParameter("DV12", DD12)
        g(41) = New ReportParameter("DV13", DD13)
        g(42) = New ReportParameter("DV14", DD14)
        g(43) = New ReportParameter("DV15", DD15)
        g(44) = New ReportParameter("DV16", DD16)

        Dim NumForm As Double = IIf(KeluarBlock.KelAstek.Text = Nothing, 0, KeluarBlock.KelAstek.Text)
        g(45) = New ReportParameter("VAst", NumForm.ToString("N0", customNumberInfo))

        For V As Integer = 0 To 45
            Me.KeluarViewer.LocalReport.SetParameters(New ReportParameter() {g(V)})
        Next

        Me.KeluarViewer.RefreshReport()

    End Sub

    Sub LocData()

        SQL = ""
        SQL = SQL & "Select * From Keluar_Table "
        SQL = SQL & "Where Nik_Num = ('" & KeluarBlock.DelTbx1.Text & "') "
        SQL = SQL & "And Name = ('" & KeluarBlock.DelTbx2.Text & "')"
        OpenTbl(CBb, Ctbl61, SQL)
        If Ctbl61.RecordCount > 0 Then

            ANik = IIf(IsDBNull(Ctbl61("Nik_Num").Value), "", Ctbl61("Nik_Num").Value)
            AName = IIf(IsDBNull(Ctbl61("Name").Value), "", Ctbl61("Name").Value)
            ATMK = IIf(IsDBNull(Ctbl61("TglMsk").Value), "", Ctbl61("TglMsk").Value)
            AKel = IIf(IsDBNull(Ctbl61("KeloF").Value), "", Ctbl61("KeloF").Value)
            ADate1 = IIf(IsDBNull(Ctbl61("From_Date").Value), "", Ctbl61("From_Date").Value)
            ADate2 = If(IsDBNull(Ctbl61("To_Date").Value), "", Ctbl61("To_Date").Value)

            ATot = IIf(IsDBNull(Ctbl61("Total").Value), "", Ctbl61("Total").Value)
            AGTot = IIf(IsDBNull(Ctbl61("GajiDet").Value), "", Ctbl61("GajiDet").Value)
            APot = IIf(IsDBNull(Ctbl61("Potongan").Value), "", Ctbl61("Potongan").Value)
            ATun = IIf(IsDBNull(Ctbl61("Tunjangan").Value), "", Ctbl61("Tunjangan").Value)
            SignA = IIf(IsDBNull(Ctbl61("Sign1").Value), "", Ctbl61("Sign1").Value)
            SignB = IIf(IsDBNull(Ctbl61("Sign2").Value), "", Ctbl61("Sign2").Value)

            ATot2 = IIf(ATot = 0, "0,0", ATot.ToString("N1", customNumberInfo))
            AGTot2 = IIf(AGTot = 0, "0,0", AGTot.ToString("N1", customNumberInfo))
            APot2 = IIf(APot = 0, "0,0", APot.ToString("N1", customNumberInfo))
            ATun2 = IIf(ATun = 0, "0,0", ATun.ToString("N1", customNumberInfo))

            D01 = IIf(IsDBNull(Ctbl61("Date01").Value), Nothing, Ctbl61("Date01").Value)
            D02 = IIf(IsDBNull(Ctbl61("Date02").Value), Nothing, Ctbl61("Date02").Value)
            D03 = IIf(IsDBNull(Ctbl61("Date03").Value), Nothing, Ctbl61("Date03").Value)
            D04 = IIf(IsDBNull(Ctbl61("Date04").Value), Nothing, Ctbl61("Date04").Value)
            D05 = IIf(IsDBNull(Ctbl61("Date05").Value), Nothing, Ctbl61("Date05").Value)
            D06 = IIf(IsDBNull(Ctbl61("Date06").Value), Nothing, Ctbl61("Date06").Value)
            D07 = IIf(IsDBNull(Ctbl61("Date07").Value), Nothing, Ctbl61("Date07").Value)
            D08 = IIf(IsDBNull(Ctbl61("Date08").Value), Nothing, Ctbl61("Date08").Value)
            D09 = IIf(IsDBNull(Ctbl61("Date09").Value), Nothing, Ctbl61("Date09").Value)
            D10 = IIf(IsDBNull(Ctbl61("Date10").Value), Nothing, Ctbl61("Date10").Value)
            D11 = IIf(IsDBNull(Ctbl61("Date11").Value), Nothing, Ctbl61("Date11").Value)
            D12 = IIf(IsDBNull(Ctbl61("Date12").Value), Nothing, Ctbl61("Date12").Value)
            D13 = IIf(IsDBNull(Ctbl61("Date13").Value), Nothing, Ctbl61("Date13").Value)
            D14 = IIf(IsDBNull(Ctbl61("Date14").Value), Nothing, Ctbl61("Date14").Value)
            D15 = IIf(IsDBNull(Ctbl61("Date15").Value), Nothing, Ctbl61("Date15").Value)
            D16 = IIf(IsDBNull(Ctbl61("Date16").Value), Nothing, Ctbl61("Date16").Value)

            DD01 = IIf(D01 = Nothing, Nothing, D01.ToString("dd MMM yy"))
            DD02 = IIf(D02 = Nothing, Nothing, D02.ToString("dd MMM yy"))
            DD03 = IIf(D03 = Nothing, Nothing, D03.ToString("dd MMM yy"))
            DD04 = IIf(D04 = Nothing, Nothing, D04.ToString("dd MMM yy"))
            DD05 = IIf(D05 = Nothing, Nothing, D05.ToString("dd MMM yy"))
            DD06 = IIf(D06 = Nothing, Nothing, D06.ToString("dd MMM yy"))
            DD07 = IIf(D07 = Nothing, Nothing, D07.ToString("dd MMM yy"))
            DD08 = IIf(D08 = Nothing, Nothing, D08.ToString("dd MMM yy"))
            DD09 = IIf(D09 = Nothing, Nothing, D09.ToString("dd MMM yy"))
            DD10 = IIf(D10 = Nothing, Nothing, D10.ToString("dd MMM yy"))
            DD11 = IIf(D11 = Nothing, Nothing, D11.ToString("dd MMM yy"))
            DD12 = IIf(D12 = Nothing, Nothing, D12.ToString("dd MMM yy"))
            DD13 = IIf(D13 = Nothing, Nothing, D13.ToString("dd MMM yy"))
            DD14 = IIf(D14 = Nothing, Nothing, D14.ToString("dd MMM yy"))
            DD15 = IIf(D15 = Nothing, Nothing, D15.ToString("dd MMM yy"))
            DD16 = IIf(D16 = Nothing, Nothing, D16.ToString("dd MMM yy"))

            G01 = IIf(IsDBNull(Ctbl61("Sal01").Value), Nothing, Ctbl61("Sal01").Value)
            G02 = IIf(IsDBNull(Ctbl61("Sal02").Value), Nothing, Ctbl61("Sal02").Value)
            G03 = IIf(IsDBNull(Ctbl61("Sal03").Value), Nothing, Ctbl61("Sal03").Value)
            G04 = IIf(IsDBNull(Ctbl61("Sal04").Value), Nothing, Ctbl61("Sal04").Value)
            G05 = IIf(IsDBNull(Ctbl61("Sal05").Value), Nothing, Ctbl61("Sal05").Value)
            G06 = IIf(IsDBNull(Ctbl61("Sal06").Value), Nothing, Ctbl61("Sal06").Value)
            G07 = IIf(IsDBNull(Ctbl61("Sal07").Value), Nothing, Ctbl61("Sal07").Value)
            G08 = IIf(IsDBNull(Ctbl61("Sal08").Value), Nothing, Ctbl61("Sal08").Value)
            G09 = IIf(IsDBNull(Ctbl61("Sal09").Value), Nothing, Ctbl61("Sal09").Value)
            G10 = IIf(IsDBNull(Ctbl61("Sal10").Value), Nothing, Ctbl61("Sal10").Value)
            G11 = IIf(IsDBNull(Ctbl61("Sal11").Value), Nothing, Ctbl61("Sal11").Value)
            G12 = IIf(IsDBNull(Ctbl61("Sal12").Value), Nothing, Ctbl61("Sal12").Value)
            G13 = IIf(IsDBNull(Ctbl61("Sal13").Value), Nothing, Ctbl61("Sal13").Value)
            G14 = IIf(IsDBNull(Ctbl61("Sal14").Value), Nothing, Ctbl61("Sal14").Value)
            G15 = IIf(IsDBNull(Ctbl61("Sal15").Value), Nothing, Ctbl61("Sal15").Value)
            G16 = IIf(IsDBNull(Ctbl61("Sal16").Value), Nothing, Ctbl61("Sal16").Value)

            GG01 = IIf(G01 = Nothing, Nothing, G01.ToString("N1", customNumberInfo))
            GG02 = IIf(G02 = Nothing, Nothing, G02.ToString("N1", customNumberInfo))
            GG03 = IIf(G03 = Nothing, Nothing, G03.ToString("N1", customNumberInfo))
            GG04 = IIf(G04 = Nothing, Nothing, G04.ToString("N1", customNumberInfo))
            GG05 = IIf(G05 = Nothing, Nothing, G05.ToString("N1", customNumberInfo))
            GG06 = IIf(G06 = Nothing, Nothing, G06.ToString("N1", customNumberInfo))
            GG07 = IIf(G07 = Nothing, Nothing, G07.ToString("N1", customNumberInfo))
            GG08 = IIf(G08 = Nothing, Nothing, G08.ToString("N1", customNumberInfo))
            GG09 = IIf(G09 = Nothing, Nothing, G09.ToString("N1", customNumberInfo))
            GG10 = IIf(G10 = Nothing, Nothing, G10.ToString("N1", customNumberInfo))
            GG11 = IIf(G11 = Nothing, Nothing, G11.ToString("N1", customNumberInfo))
            GG12 = IIf(G12 = Nothing, Nothing, G12.ToString("N1", customNumberInfo))
            GG13 = IIf(G13 = Nothing, Nothing, G13.ToString("N1", customNumberInfo))
            GG14 = IIf(G14 = Nothing, Nothing, G14.ToString("N1", customNumberInfo))
            GG15 = IIf(G15 = Nothing, Nothing, G15.ToString("N1", customNumberInfo))
            GG16 = IIf(G16 = Nothing, Nothing, G16.ToString("N1", customNumberInfo))
      



        End If

    End Sub
   
 
End Class