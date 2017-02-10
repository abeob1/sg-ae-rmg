Imports System.Data
Imports System.IO
Imports System.Data.OleDb


Public Class MainClass

    Public WithEvents oApplication As SAPbouiCOM.Application
    Public oCompany As New SAPbobsCOM.Company
    Public SboGuiApi As New SAPbouiCOM.SboGuiApi
    Public sConnectionString As String
    Dim orset_TM As SAPbobsCOM.Recordset
  

    ' Return Value Variable Control
    Dim sErrDesc As String = String.Empty


    Public Sub New()
        MyBase.New()

        Try
            Single_Signon()
            AddMenuItems()

            orset_TM = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset_TM.DoQuery("SELECT T0.[superuser], T0.[USER_CODE], T0.[U_NAME] FROM OUSR T0 WHERE T0.[USER_CODE] = '" & oCompany.UserName & "'")
            ' VT_Live_Timer()
            Company_Name = orset_TM.Fields.Item("U_NAME").Value
            p_bSuperUser = orset_TM.Fields.Item("superuser").Value

            Class_ChaufferDriver = New ChaufferDriver(oApplication, oCompany)
            Class_SelfDrive = New SelfDrive(oApplication, oCompany)
            Class_TrafficAccident = New TrafficAccident(oApplication, oCompany)
            Class_ServiceMaintenance = New ServiceMaintenance(oApplication, oCompany)
            Class_systemform = New SystemForm(oApplication, oCompany)
            Class_Report = New Report

            '  Class_TableCreation = New TableCreation(oApplication, oCompany)
            Dim orset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset.DoQuery("SELECT *  FROM [dbo].[@AE_NTIME]  T0")
            If orset.RecordCount = 0 Then
                oApplication.MessageBox("Kindly define the surcharge timing and restart ", 1, "Ok")
            Else
                Special_TimeStart = orset.Fields.Item("Code").Value
                Special_TimeEnd = orset.Fields.Item("Name").Value
            End If


            'If orset.RecordCount = 0 Then
            '    oApplication.MessageBox("Kindly define the GL Account and restart ", 1, "Ok")
            'Else
            orset.DoQuery("SELECT T0.[U_account] FROM [dbo].[@AE_GLACC]  T0 WHERE T0.[Code] = 'CD'")
            CD_GLAcc = orset.Fields.Item("U_account").Value
            orset.DoQuery("SELECT T0.[U_account] FROM [dbo].[@AE_GLACC]  T0 WHERE T0.[Code] = 'SD'")
            SD_GLACC = orset.Fields.Item("U_account").Value
            orset.DoQuery("SELECT T0.[U_account] FROM [dbo].[@AE_GLACC]  T0 WHERE T0.[Code] = 'SM'")
            SM_GLACC = orset.Fields.Item("U_account").Value
            orset.DoQuery("SELECT T0.[U_account] FROM [dbo].[@AE_GLACC]  T0 WHERE T0.[Code] = 'CDW'")
            CDW_GLACC = orset.Fields.Item("U_account").Value
            orset.DoQuery("SELECT T0.[U_account] FROM [dbo].[@AE_GLACC]  T0 WHERE T0.[Code] = 'PAI'")
            PAI_GLACC = orset.Fields.Item("U_account").Value

            orset.DoQuery("SELECT T0.[U_account] FROM [dbo].[@AE_GLACC]  T0 WHERE T0.[Code] = 'OC'")
            sOC_GLACC = orset.Fields.Item("U_account").Value
            orset.DoQuery("SELECT T0.[U_account] FROM [dbo].[@AE_GLACC]  T0 WHERE T0.[Code] = 'MRC'")
            sMR_GLACC = orset.Fields.Item("U_account").Value
            orset.DoQuery("SELECT T0.[U_account] FROM [dbo].[@AE_GLACC]  T0 WHERE T0.[Code] = 'PC'")
            sPet_GLACC = orset.Fields.Item("U_account").Value
            'End If

            orset.DoQuery("SELECT T0.[ItmsGrpCod] FROM OITM T0  INNER JOIN OITB T1 ON T0.ItmsGrpCod = T1.ItmsGrpCod WHERE T1.[ItmsGrpNam] = 'Vehicles'")
            ItemGroup = orset.Fields.Item("ItmsGrpCod").Value

            orset.DoQuery("SELECT getdate() as 'ServerDate'")
            Server_Date = orset.Fields.Item("ServerDate").Value
            crystalRepotconnect()
            AppConfigInfo(p_oCompDef, sErrDesc)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub VT_Live_Timer()
        Try
            CD_Timer.Interval = 60000 '600000cxzbnnbm n , j 
            AddHandler CD_Timer.Elapsed, AddressOf Class_Report.VT_Live_TimerScript
            Class_Report.oCompany = oCompany
            Class_Report.oApplication = oApplication
            CD_Timer.Start()
        Catch ex As Exception

        End Try
    End Sub

    Public Sub crystalRepotconnect()
        Dim crycon As SAPbobsCOM.Recordset = ocompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        crycon.DoQuery("SELECT T0.[Code], T0.[name] FROM [dbo].[@AE_CRYSTAL]  T0")
        Dim sa As String = crycon.Fields.Item("name").Value
        Dim connectionString As String = ""
        connectionString = "Provider=SQLOLEDB;"
        connectionString += "Server=" + ocompany.Server + ";Database=" + ocompany.CompanyDB + ";"
        connectionString += "User ID=" & ocompany.DbUserName & ";Password=" & sa & ""
        adoOleDbConnection = New OleDbConnection(connectionString)
    End Sub



    Public Sub Single_Signon()
        Dim sconn As String
        Dim ret As Integer
        Dim scook As String
        Dim str As String


        Try
            sconn = Environment.GetCommandLineArgs.GetValue(1)
            SboGuiApi.Connect(sconn)
            oApplication = SboGuiApi.GetApplication
            SboGuiApi = Nothing
            scook = oCompany.GetContextCookie
            str = oApplication.Company.GetConnectionContext(scook)
            ret = oCompany.SetSboLoginContext(str)
            oCompany.Connect()
            oCompany.GetLastError(ret, str)
            If ret <> 0 Then
                MsgBox(str)
            Else

                oApplication.StatusBar.SetText("Connected to the Company ........ " & oCompany.CompanyName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                If oCompany.UserName = "P002" Then
                    oApplication.Menus.Item("43531").Enabled = False
                Else
                    oApplication.Menus.Item("43531").Enabled = True
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ''Public Sub create_menu()

    ''    Dim omenus As SAPbouiCOM.Menus
    ''    Dim omenuitem As SAPbouiCOM.MenuItem
    ''    Dim omenuparams As SAPbouiCOM.MenuCreationParams

    ''    Dim sPath As String

    ''    'sPath = Application.StartupPath
    ''    sPath = IO.Directory.GetParent(Application.StartupPath).ToString
    ''    'sPath = sPath.Remove(sPath.Length - 8, 8)

    ''    omenus = SBO_Application.Menus
    ''    omenuparams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
    ''    omenuitem = SBO_Application.Menus.Item("43520")
    ''    omenus = omenuitem.SubMenus

    ''    'cassette maintenance module
    ''    omenuparams.Type = SAPbouiCOM.BoMenuType.mt_POPUP
    ''    omenuparams.UniqueID = "U1"
    ''    omenuparams.String = "Media"
    ''    omenuparams.Enabled = True
    ''    '  omenuparams.Image = sPath & "\Marketing\Sathiyam.bmp"  '"\Marketing\images7.bmp"
    ''    omenuparams.Position = 0
    ''    omenus.AddEx(omenuparams)


    ''End Sub

    Sub AddMenuItems()
        Dim oMenus As SAPbouiCOM.Menus
        Dim oMenuItem As SAPbouiCOM.MenuItem

        Dim sPath As String

        'sPath = Application.StartupPath
        sPath = IO.Directory.GetParent(Application.StartupPath).ToString
        'sPath = sPath.Remove(sPath.Length - 8, 8)

        oMenus = oApplication.Menus

        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = (oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams))
        oMenuItem = oApplication.Menus.Item("43520") 'Modules

        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
        oCreationPackage.UniqueID = "Fleet_Management"

        oCreationPackage.String = "Fleet Management"
        oCreationPackage.Enabled = True
        oCreationPackage.Image = sPath & "\AE_FleetMangement\Logo.bmp"  '"\Marketing\images7.bmp"
        oCreationPackage.Position = 2

        oMenus = oMenuItem.SubMenus

        Try
            'If the manu already exists this code will fail
            oMenus.AddEx(oCreationPackage)
        Catch
        End Try


        Try
            'Get the menu collection of the newly added pop-up item

            LoadFromXML("FleetManagement_menu.xml", oApplication)


           


            ' ''oMenuItem = oApplication.Menus.Item("Fleet_Management")
            ' ''oMenus = oMenuItem.SubMenus

            '' ''Create s sub menu
            ' ''oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            ' ''oCreationPackage.UniqueID = "VTYPE"
            ' ''oCreationPackage.String = "Vehicle Type"
            ' ''oMenus.AddEx(oCreationPackage)


            ' ''oMenuItem = oApplication.Menus.Item("Fleet_Management")
            ' ''oMenus = oMenuItem.SubMenus

            '' ''Create s sub menu
            ' ''oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            ' ''oCreationPackage.UniqueID = "CD"
            ' ''oCreationPackage.String = "Chauffer Driver"
            ' ''oMenus.AddEx(oCreationPackage)

            ' ''oMenuItem = oApplication.Menus.Item("CD")
            ' ''oMenus = oMenuItem.SubMenus

            '' ''Create s sub menu
            ' ''oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            ' ''oCreationPackage.UniqueID = "CDVB"
            ' ''oCreationPackage.String = "Vehicle Booking"
            ' ''oMenus.AddEx(oCreationPackage)

            ' ''oMenuItem = oApplication.Menus.Item("CD")
            ' ''oMenus = oMenuItem.SubMenus

            '' ''Create s sub menu
            ' ''oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            ' ''oCreationPackage.UniqueID = "CDSP"
            ' ''oCreationPackage.String = "Scheduling & Driver Assigning"
            ' ''oMenus.AddEx(oCreationPackage)

            ' ''oMenuItem = oApplication.Menus.Item("CD")
            ' ''oMenus = oMenuItem.SubMenus

            '' ''Create s sub menu
            ' ''oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            ' ''oCreationPackage.UniqueID = "CDPS"
            ' ''oCreationPackage.String = "Price Setup"
            ' ''oMenus.AddEx(oCreationPackage)

            ' ''oMenuItem = oApplication.Menus.Item("Fleet_Management")
            ' ''oMenus = oMenuItem.SubMenus

            '' ''Create s sub menu
            ' ''oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            ' ''oCreationPackage.UniqueID = "SD"
            ' ''oCreationPackage.String = "Self Drive"
            ' ''oMenus.AddEx(oCreationPackage)

            ' ''oMenuItem = oApplication.Menus.Item("SD")
            ' ''oMenus = oMenuItem.SubMenus

            '' ''Create s sub menu
            ' ''oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            ' ''oCreationPackage.UniqueID = "SDVB"
            ' ''oCreationPackage.String = "Vehicle Booking"
            ' ''oMenus.AddEx(oCreationPackage)

            ' ''oMenuItem = oApplication.Menus.Item("SD")
            ' ''oMenus = oMenuItem.SubMenus

            '' ''Create s sub menu
            ' ''oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            ' ''oCreationPackage.UniqueID = "SDB"
            ' ''oCreationPackage.String = "Billing"
            ' ''oMenus.AddEx(oCreationPackage)

            ' ''oMenuItem = oApplication.Menus.Item("Fleet_Management")
            ' ''oMenus = oMenuItem.SubMenus

            '' ''Create s sub menu
            ' ''oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            ' ''oCreationPackage.UniqueID = "TPO"
            ' ''oCreationPackage.String = "Traffic & Parking Offense"
            ' ''oMenus.AddEx(oCreationPackage)

            ' ''oMenuItem = oApplication.Menus.Item("Fleet_Management")
            ' ''oMenus = oMenuItem.SubMenus

            '' ''Create s sub menu
            ' ''oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            ' ''oCreationPackage.UniqueID = "VAC"
            ' ''oCreationPackage.String = "Vehicle Accident Claim"
            ' ''oMenus.AddEx(oCreationPackage)

            ' ''oMenuItem = oApplication.Menus.Item("Fleet_Management")
            ' ''oMenus = oMenuItem.SubMenus

            '' ''Create s sub menu
            ' ''oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            ' ''oCreationPackage.UniqueID = "SM"
            ' ''oCreationPackage.String = "Service & Maintenance"
            ' ''oMenus.AddEx(oCreationPackage)

            ' ''oMenuItem = oApplication.Menus.Item("SM")
            ' ''oMenus = oMenuItem.SubMenus

            '' ''Create s sub menu
            ' ''oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            ' ''oCreationPackage.UniqueID = "SMS"
            ' ''oCreationPackage.String = "Service & Maintenance Master Setup"
            ' ''oMenus.AddEx(oCreationPackage)

            ' ''oMenuItem = oApplication.Menus.Item("SM")
            ' ''oMenus = oMenuItem.SubMenus

            '' ''Create s sub menu
            ' ''oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            ' ''oCreationPackage.UniqueID = "SMSM"
            ' ''oCreationPackage.String = "Service & Maintenance"
            ' ''oMenus.AddEx(oCreationPackage)

            ' ''oMenuItem = oApplication.Menus.Item("Fleet_Management")
            ' ''oMenus = oMenuItem.SubMenus

            '' ''Create s sub menu
            ' ''oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            ' ''oCreationPackage.UniqueID = "VT"
            ' ''oCreationPackage.String = "Vehicle Tracking"
            ' ''oMenus.AddEx(oCreationPackage)


        Catch
            'Menu already exists
            oApplication.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub




    Private Sub oApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles oApplication.AppEvent
        If EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Then
            oApplication.StatusBar.SetText("Shutting Down Fleet Management addon", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Windows.Forms.Application.Exit()
        End If

        If EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
            oApplication.StatusBar.SetText("Shutting Down Fleet Management addon", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Windows.Forms.Application.Exit()
        End If

        If EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Then
            oApplication.StatusBar.SetText("Shutting Down Fleet Management addon", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Windows.Forms.Application.Exit()
        End If
    End Sub
End Class
