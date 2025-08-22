Option Explicit On



Imports Microsoft.Office.Interop
Imports ExcelMe = Microsoft.Office.Interop.Excel
Imports System.Security.Cryptography
Imports System.Text
Imports System.Data
Imports System.IO

Imports System.Globalization

Module Support

    Public ExcelWM As ExcelMe.Application
    Public ExcelWBk As ExcelMe.Workbook
    Public ExcelWSh As ExcelMe.Worksheet

    Public CustomtoIndo As CultureInfo = CultureInfo.CreateSpecificCulture("id-ID")
    Public CustomtoUS As CultureInfo = CultureInfo.CreateSpecificCulture("en-US")

    Public PayAsSetup As String ' As Type of Pay

    Public NewSald(15) As String

    Public SDelAstek As String = ""
    Public SDelMasuk As String = ""
    Public SSigning As String = ""

    Public SDate As String = ""
    Public EDate As String = ""

    Public NewGSal(17) As String

    Public GetAstek As String
    Public GetAstek2 As String
    Public GetAstekFormat As String
    Public GetPotlain As String

    Public NewFormT(15) As String

    Public NewTotSalRow(18) As String

    Public TotAstek As String

    Public AntiDupActuator As Integer = 0

    Public Function CustomRound(ByVal RoundValue As Int64)

        Dim roundUpTo500 = (RoundValue Mod 1000) < 500

        If (roundUpTo500) Then
            Return Math.Floor(RoundValue / 1000) * 1000 + 500
        Else
            Return Math.Round(RoundValue / 1000) * 1000
        End If

        CustomRound = roundUpTo500

    End Function


#Region "Excel Control"
    Public Sub StartExcel()
        On Error GoTo Err
        ExcelWM = GetObject(, "Excel.Application")
        Exit Sub

Err:
        ExcelWM = CreateObject("Excel.Application")
    End Sub

    Public Sub OpenExlWbk(ByVal FileName As String)
        ExcelWM.Visible = False
        ExcelWBk = ExcelWM.Workbooks.Open(FileName)
    End Sub


    Public Sub OpenExlWsh(ByVal SheetIdx As Integer)
        ExcelWSh = ExcelWBk.Worksheets(SheetIdx)
    End Sub

    Public Sub CloseWorkSheet()
        ExcelWBk.Close()
        ExcelWM.Quit()
    End Sub
#End Region

#Region "KEYtoKEY"
    Public Function Encrypt(ByVal clearText As String) As String
        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        Dim clearBytes As Byte() = Encoding.Unicode.GetBytes(clearText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, _
             &H65, &H64, &H76, &H65, &H64, &H65, _
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                    cs.Write(clearBytes, 0, clearBytes.Length)
                    cs.Close()
                End Using
                clearText = Convert.ToBase64String(ms.ToArray())
            End Using
        End Using
        Return clearText
    End Function
    Public Function Decrypt(ByVal cipherText As String) As String
        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, _
             &H65, &H64, &H76, &H65, &H64, &H65, _
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write)
                    cs.Write(cipherBytes, 0, cipherBytes.Length)
                    cs.Close()
                End Using
                cipherText = Encoding.Unicode.GetString(ms.ToArray())
            End Using
        End Using
        Return cipherText
    End Function

#End Region

#Region "Clear Cache"
    Declare Function SetProcessWorkingSetSize Lib "kernel32.dll" (ByVal process As IntPtr, ByVal minimumWorkingSetSize As Integer, ByVal maximumWorkingSetSize As Integer) As Integer
    Public Sub FlushMemory()
        Try
            GC.Collect()
            GC.WaitForPendingFinalizers()
            If (Environment.OSVersion.Platform = PlatformID.Win32NT) Then
                SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, -1, -1)
                Dim myProcesses As Process() = Process.GetProcessesByName("Applica­tionName")
                Dim myProcess As Process 'Dim ProcessInfo As Process 
                For Each myProcess In myProcesses
                    SetProcessWorkingSetSize(myProcess.Handle, -1, -1)
                Next myProcess
            End If
        Catch ex As Exception
        End Try
    End Sub

#End Region

End Module
