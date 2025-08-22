
Option Explicit On


Module BackEnd

    Public ADb As New ADODB.Connection
    Public Atb1 As New ADODB.Recordset ' Save Employee
    Public Atb2 As New ADODB.Recordset ' Load/Search Employe
    Public Atb3 As New ADODB.Recordset ' Load Employee for Employee2block
    Public Atb4 As New ADODB.Recordset ' Load Employee / Load For Modification "RemBlock"
    Public Atb5 As New ADODB.Recordset ' Conveyour Save
    Public Atb6 As New ADODB.Recordset ' MutuII Save
    Public Atb7 As New ADODB.Recordset ' Wallet save
    Public Atb8 As New ADODB.Recordset ' Packing Save
    Public Atbl17 As New ADODB.Recordset ' Load User and Save User
    Public Atbl21 As New ADODB.Recordset ' Load Standard Combo
    Public Atbl22 As New ADODB.Recordset ' Load Standard2 Combo
    Public Atbl23 As New ADODB.Recordset ' Save StandardModification
    Public Atbl24 As New ADODB.Recordset ' Save Miscellaneous
    Public Atbl26 As New ADODB.Recordset ' Save Synch
    Public Atbl27 As New ADODB.Recordset ' Save Btm Synch
    Public Atbl20 As New ADODB.Recordset ' Load Jamsostek (Astek)
    Public Atbl28 As New ADODB.Recordset ' Save Sortasi
    Public Atbl29 As New ADODB.Recordset ' Load Astek
    Public Atbl30 As New ADODB.Recordset ' Save Multiple Salary Platform
    Public Atbl31 As New ADODB.Recordset ' Save Incentives for employee
    Public Atbl32 As New ADODB.Recordset ' Load Incentives counter for employee
    Public Atbl33 As New ADODB.Recordset ' Load Incentives Range for active range
    Public Atbl34 As New ADODB.Recordset ' Save Incentives Range for active range
    Public Atbl35 As New ADODB.Recordset ' Load Incentives Lock Boolean Rule
    Public Atbl36 As New ADODB.Recordset ' Save Incentives Lock Boolean Rule
    Public Atbl37 As New ADODB.Recordset ' Load Astek Looker (Astek Value)
    Public Atbl38 As New ADODB.Recordset ' Save Incentives Synch
    Public Atbl39 As New ADODB.Recordset ' Save Employee Synch
    Public Atbl40 As New ADODB.Recordset ' Deactivate Employee
    Public Atbl41 As New ADODB.Recordset ' Enabler Tot
    Public Atbl42 As New ADODB.Recordset ' Enabler
    Public Atbl43 As New ADODB.Recordset ' Save Dokter Note

    Public Atbl44 As New ADODB.Recordset '  New PPH loader as of 10-Dec-2014

    Public BBb As New ADODB.Connection
    Public BBTbl1 As New ADODB.Recordset ' Load Synch Lower Flex
    Public BBTbl2 As New ADODB.Recordset ' Load Synch Lower Flex
    Public BBtbl3 As New ADODB.Recordset ' Load User Synch Incentives
    Public BBtbl4 As New ADODB.Recordset ' Load User Synch Employee List

    Public DBTbl1 As New ADODB.Recordset ' User Login
    Public DBTbl2 As New ADODB.Recordset ' Anti-Dup Nik
    Public DBTbl3 As New ADODB.Recordset ' Load Standard Wage
    Public DbTbl4 As New ADODB.Recordset ' Load Conveyour
    Public DbTbl5 As New ADODB.Recordset ' Generate ID Name Process
    Public DbTbl6 As New ADODB.Recordset ' Gen Process ID
    Public DbTbl7 As New ADODB.Recordset ' Load MutuII
    Public DbTbl8 As New ADODB.Recordset ' Load Wallet
    Public DbTbl9 As New ADODB.Recordset ' Load Packing
    Public Dbtb20 As New ADODB.Recordset ' Load OverAll
    Public Dbtb21 As New ADODB.Recordset ' Load OverAll Conveyour
    Public Dbtb22 As New ADODB.Recordset ' Load OverAll MutuII
    Public Dbtb23 As New ADODB.Recordset ' Load Overall Wallet
    Public Dbtb24 As New ADODB.Recordset ' Load OverAll Packing
    Public DBTb26 As New ADODB.Recordset ' Generate User Number
    Public Dbtb27 As New ADODB.Recordset ' Save Holiday
    Public Dbtb28 As New ADODB.Recordset ' Load Holiday
    Public Dbtb29 As New ADODB.Recordset ' Load Holiday Mod
    Public Dbtb33 As New ADODB.Recordset ' YearHolMod
    Public Dbtb34 As New ADODB.Recordset ' Load Miscellaneous
    Public Dbtb35 As New ADODB.Recordset ' Misc Overall
    Public Dbtb37 As New ADODB.Recordset ' Load for Sortasi
    Public Dbtb38 As New ADODB.Recordset ' Load Subsidi
    Public Dtbb39a As New ADODB.Recordset ' Load Salary and Date 1 
    Public Dtbb39b As New ADODB.Recordset ' Load Salary and Date 2
    Public Dtbb39c As New ADODB.Recordset ' Load Salary and Date 3 
    Public Dtbb39d As New ADODB.Recordset ' Load Salary and Date 4 
    Public Dtbb39f As New ADODB.Recordset ' Load Salary and Date 5 
    Public Dtbb39g As New ADODB.Recordset ' Load Salary and Date 6 
    Public Dtbb39h As New ADODB.Recordset ' Load Salary and Date 7 
    Public Dtbb39i As New ADODB.Recordset ' Load Salary and Date 8 
    Public Dtbb39j As New ADODB.Recordset ' Load Salary and Date 9 
    Public Dtbb39k As New ADODB.Recordset ' Load Salary and Date 10 
    Public Dtbb39l As New ADODB.Recordset ' Load Salary and Date 11
    Public Dtbb39m As New ADODB.Recordset ' Load Salary and Date 12
    Public Dtbb39n As New ADODB.Recordset ' Load Salary and Date 13 
    Public Dtbb39o As New ADODB.Recordset ' Load Salary and Date 14 
    Public Dtbb39p As New ADODB.Recordset ' Load Salary and Date 15 
    Public Dtbb39q As New ADODB.Recordset ' Load Salary and Date 16 
    Public Dtbb40 As New ADODB.Recordset ' Load Employee on QC Report2
    Public Dbtb39 As New ADODB.Recordset '  Gaji Ctrl for Increase


    Public CBb As New ADODB.Connection
    Public Ctbl1 As New ADODB.Recordset 'Saving Full Salary
    Public Ctbl21 As New ADODB.Recordset 'DateDataSave
    Public Ctbl22 As New ADODB.Recordset 'LoadDateSetup
    Public Ctbl23 As New ADODB.Recordset 'DateDataSave2
    Public Ctbl24 As New ADODB.Recordset 'DateLoadSalary
    Public Ctbl25 As New ADODB.Recordset 'LoadEmployeeSalary
    Public Ctbl26 As New ADODB.Recordset 'Save Periode Counter 1
    Public Ctbl27 As New ADODB.Recordset 'Save Periode Counter 2
    Public Ctbl28 As New ADODB.Recordset 'Save Periode Counter 3
    Public Ctbl29 As New ADODB.Recordset 'Save Periode Counter 4
    Public Ctbl30 As New ADODB.Recordset 'Save Periode Counter 5
    Public Ctbl31 As New ADODB.Recordset 'Save Periode Counter 6
    Public Ctbl32 As New ADODB.Recordset 'Save Periode Counter 7
    Public Ctbl33 As New ADODB.Recordset 'Save Periode Counter 8
    Public Ctbl34 As New ADODB.Recordset 'Save Periode Counter 9
    Public Ctbl35 As New ADODB.Recordset 'Save Periode Counter 10
    Public Ctbl36 As New ADODB.Recordset 'Save Periode Counter 11
    Public Ctbl37 As New ADODB.Recordset 'Save Periode Counter 12
    Public Ctbl38 As New ADODB.Recordset 'Save Periode Counter 13
    Public Ctbl39 As New ADODB.Recordset 'Save Periode Counter 14
    Public Ctbl40 As New ADODB.Recordset 'Save Periode Counter 15
    Public Ctbl41 As New ADODB.Recordset 'Save Periode Counter 16
    Public Ctbl42 As New ADODB.Recordset 'LoadDayPeriodeCtrl
    Public Ctbl43 As New ADODB.Recordset 'Gen Date Auto Code
    Public Ctbl44 As New ADODB.Recordset 'Setting Date
    Public Ctbl45 As New ADODB.Recordset 'Keluar Kerja Nik and Data Look Up
    Public Ctbl49 As New ADODB.Recordset 'Delete Keluar
    Public Ctbl50 As New ADODB.Recordset 'Look FirstData
    Public Ctbl51 As New ADODB.Recordset 'Look LastData
    Public Ctbl52 As New ADODB.Recordset 'Save Data
    Public Ctbl53 As New ADODB.Recordset 'Save Keluar
    Public Ctbl54 As New ADODB.Recordset 'Re Count 1
    Public Ctbl55 As New ADODB.Recordset 'Re Count 2
    Public Ctbl56 As New ADODB.Recordset ' Obtaining Nik and Name 
    Public Ctbl57 As New ADODB.Recordset ' Obtaining Periode Stat 1
    Public Ctbl58 As New ADODB.Recordset ' Obtaining Periode Stat 2 
    Public Ctbl59 As New ADODB.Recordset ' Saving Incentives Value 
    Public Ctbl60 As New ADODB.Recordset ' Loading Incentives
    Public Ctbl61 As New ADODB.Recordset ' Looking For New Keluar Report

    Public FBb As New ADODB.Connection
    Public Ftbl24 As New ADODB.Recordset 'Sortasi2Synch Date Looker
    Public Ftbl25 As New ADODB.Recordset 'Sortasi2Synch Salary Looker

    Public PPhDB As New ADODB.Connection
    Public PPhTb1 As New ADODB.Recordset
    Public PPhTb2 As New ADODB.Recordset
    Public PPhTb3 As New ADODB.Recordset
    Public PPhTb4 As New ADODB.Recordset ' Use to Load my PPh21 Data
    Public PPhTb5 As New ADODB.Recordset ' Use to UP pph21 Data from  report block /  Save Coupon
    Public PPhTb6 As New ADODB.Recordset ' Load for PPh21 data
    Public PPhTb7 As New ADODB.Recordset ' Delete old Coupon
    Public PPhTb8 As New ADODB.Recordset ' Save employee
    Public PPhTb9 As New ADODB.Recordset ' Load Address

    Public PathStr As String
    Public SQL As String

    Public PeriodeCtrl As String
    Public PeriodeMonthCtrl As String
    Public PeriodeDayCtrl As String

    Public DeptCode As String
    Public StandardsSalary As String
    Public SubsidiSalary As String
    Public GajiCtrlSalary As String
    Public HolMod As Integer
    Public YearDate As String
    Public YearMod As Integer
    Public ConMainTot As Double
    Public MutuMainTot As Double
    Public WalletMainTot As Double
    Public PackingMainTot As Double
    Public SortMainTot As Double
    Public SortSpecialTot As Double
    Public SortPiecesTot As Double
    Public MiscMainTot As Double
    Public OverAllTot As Double

    Public UserRec As UserData
    Public DbPath_Config As String
    Public BuildCounter As String = " 2.9.1"

    Public Structure UserData
        Public QcName As String
        Public QcUserName As String
        Public QcPassword As String
        Public QcLevel As String
    End Structure

    Public User As String
    Public UsFieCode As String

    Public Const InValid As String = "QWERTYUIOPASDFGHJKLZXCVBNMqwertyuiopasdfghjklzxcvbnm 123456789~`'!@#$%^&*()_+|-=\/,.:;"
    Public Const InValid2 As String = "QWERTYUIOPASDFGHJKLZXCVBNMqwertyuiopasdfghjklzxcvbnm ~`'!@#$%^&*()_+|-=\/,:;"
    Public Const InValid3 As String = "QWERTYUIOPASDFGHJKLZXCVBNMqwertyuiopasdfghjklzxcvbnm ~`'!@#$%^&*()_+|-=\/:;"
    Public Const InValid4 As String = "~`'!@#$%^&*()_+|-=\/:;"

    Sub OpenDB(ByVal DBcon As ADODB.Connection, ByVal DBName As String, ByVal DBPass As String)

        PathStr = "DRIVER={MySQL ODBC 5.3 ANSI Driver};" _
                  & "SERVER=192.168.2.18;" _
                  & "Port=54444;" _
                  & "DATABASE=" + DBName + ";" _
                  & "UID=root; PWD=" + DBPass _
                  & "; OPTION = 3"

        If DBcon.State = 1 Then DBcon.Close()
        DBcon.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        DBcon.ConnectionString = PathStr
        DBcon.Open()

    End Sub

    'Sub OpenDB02(ByVal DBcon As ADODB.Connection, ByVal DBName As String, ByVal DBPass As String)

    '    PathStr = ""
    '    PathStr = "Provider=Microsoft.Jet.OLEDB.4.0;"
    '    PathStr = PathStr & "Data Source=" & Application.StartupPath + "\" & DBName & ";"
    '    PathStr = PathStr & "Persist Security Info=False;"
    '    PathStr = PathStr & "Jet OLEDB:database password= " & DBPass

    '    If DBcon.State = 1 Then DBcon.Close()
    '    DBcon.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '    DBcon.ConnectionString = PathStr
    '    DBcon.Open()

    'End Sub

    Sub OpenTbl(ByVal DB As ADODB.Connection, ByVal Tbl As ADODB.Recordset, ByVal SQL As String)

        If Tbl.State = 1 Then Tbl.Close()

        Tbl.Open(SQL, DB, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockPessimistic)

    End Sub

    Public Sub LoadDB()
        OpenDB(ADb, "sortasi1", "ug2015")

    End Sub

    Sub LoadDB2()
        OpenDB(CBb, "sortasi2", "ug2015")

    End Sub

    Sub LoadDB3()
        'OpenDB02(BBb, "SortasiSynch.mdb", "azure2013")
    End Sub

    Sub LoadDB4()
        'OpenDB02(FBb, "Sortasi2Synch.mdb", "azure2013")
    End Sub

    Sub LoadDBPPh21() ' Load DB fron PPH21
        OpenDB(PPhDB, "pph21sortasi", "ug2015")
    End Sub

End Module
