
Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System
Imports System.Windows.Forms


Module FleetManagement

    Public format1 As New System.Globalization.CultureInfo("fr-FR", True)
    Public Tmp As String
    Public FormType_Emp As Integer
    Public FormType_SM As Integer
    Public FormType_Invoice As Integer
    Public FormType_BP As Integer
    Public Special_TimeStart As Date
    Public Special_TimeEnd As Date
    Public Special_Hours As Integer
    Public Special_Mins As Integer
    Public Normal_Hours As Integer
    Public Normal_Mins As Integer
    Public CD_GLAcc As Integer
    Public SD_GLACC As Integer
    Public SM_GLACC As Integer
    Public CDW_GLACC As Integer
    Public PAI_GLACC As Integer

    Public Invoice_Type As String
    Public Server_Date As Date
    Public Company_Name As String
    Public Event_CD As Boolean = False
    Public stFilePathAndName As String
    Public SM As Boolean = False
    Public ItemGroup As String
    Public BP_PriceList As String
    Public CFL As Boolean = False
    Public Invoice_UDF As Boolean = False
    Public TAO_Flag As Boolean = False
    Public ACD_Flag As Boolean = False
    Public p_bSuperUser As String = String.Empty
    Public bSVH As Boolean = False

    Public Class_ChaufferDriver As ChaufferDriver
    Public Class_SelfDrive As SelfDrive
    Public Class_ServiceMaintenance As ServiceMaintenance
    Public Class_TrafficAccident As TrafficAccident
    Public Class_MainClass As MainClass
    Public Class_systemform As SystemForm
    Public Class_TableCreation As TableCreation
    Public Class_Report As Report

    Public CD_Timer As New System.Timers.Timer

    Public adoOleDbConnection As OleDbConnection
    Public adoOleDbDataAdapter As OleDbDataAdapter



    Public Sub main()
        Try

            Class_MainClass = New MainClass
            System.Windows.Forms.Application.Run()


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub


    Public Function dateconversion(ByVal _Date As String) As String
        Return _Date.ToString.Substring(3, 2) & "/" & _Date.ToString.Substring(0, 2) & "/" & _Date.ToString.Substring(6, 2)
    End Function
    

    Public Function NextSerialNo(ByRef oCompany As SAPbobsCOM.Company, ByRef oApplication As SAPbouiCOM.Application, ByVal ObjectCode As String) As String
        Try

            Dim orset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset.DoQuery("SELECT T0.Series FROM NNM1 T0 where  T0.ObjectCode = '" & ObjectCode & "'")
            Return orset.Fields.Item("Series").Value

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return -1
        End Try
    End Function


    Public Sub LoadFromXML(ByVal FileName As String, ByVal Sbo_application As SAPbouiCOM.Application)
        Try
            Dim oXmlDoc As New Xml.XmlDocument
            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.StartupPath).ToString
            oXmlDoc.Load(sPath & "\AE_FleetMangement\" & FileName)
            Sbo_application.LoadBatchActions(oXmlDoc.InnerXml)
        Catch ex As Exception
            MsgBox(ex)
        End Try

    End Sub

    Public Function Extended_Text(ByRef oform As SAPbouiCOM.Form, ByRef Oapplication_CD As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company _
                                  , ByVal Doc As String, ByVal Line As String, ByVal ColID As String, ByVal CD As String, ByVal Title As String)
        Try

            Dim obutton As SAPbouiCOM.Button
            Dim orset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            orset.DoQuery("SELECT T0.[code], T0.[U_AE_Rem], T0.[U_AE_Dno], T0.[U_AE_Lno], T0.[U_AE_Object] FROM [dbo].[@AE_EXTENDED]  T0 WHERE T0.[U_AE_Dno] = '" & Doc & "' and  T0.[U_AE_Lno] = '" & Line & "' and  T0.[U_AE_Object] = '" & CD & "' and T0.[U_AE_ColID] = '" & ColID & "'")
            If orset.RecordCount = 0 Then
                LoadFromXML("ExtendedBox.srf", Oapplication_CD)
                oform = Oapplication_CD.Forms.Item("Extended")
                oform.Freeze(True)
                oform.Title = Title
                oform.Items.Item("2").Specific.string = Doc
                oform.Items.Item("3").Specific.string = Line
                oform.Items.Item("5").Specific.string = CD
                oform.Items.Item("7").Specific.string = ColID
                obutton = oform.Items.Item("11").Specific
                obutton.Caption = "Add"
            Else
                LoadFromXML("ExtendedBox.srf", Oapplication_CD)
                oform = Oapplication_CD.Forms.Item("Extended")
                oform.Freeze(True)
                oform.Title = Title
                oform.Items.Item("2").Specific.string = Doc
                oform.Items.Item("3").Specific.string = Line
                oform.Items.Item("5").Specific.string = CD
                oform.Items.Item("10").Specific.string = orset.Fields.Item("U_AE_Rem").Value
                oform.Items.Item("6").Specific.string = orset.Fields.Item("code").Value
                oform.Items.Item("7").Specific.string = ColID
                obutton = oform.Items.Item("11").Specific
                obutton.Caption = "Update"
            End If
            oform.Freeze(False)
            oform.Visible = True
            Return True

        Catch ex As Exception
            oform.Freeze(False)
            Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function

    Public Function UDTform(ByVal n As String, ByRef SBO_Application As SAPbouiCOM.Application) As Integer 'function to load Udt forms
        Dim oMenu As SAPbouiCOM.MenuItem
        Dim i As Integer
        oMenu = SBO_Application.Menus.Item("51200")
        For i = 0 To oMenu.SubMenus.Count - 1
            If oMenu.SubMenus.Item(i).String = n Then
                Return oMenu.SubMenus.Item(i).UID
                Exit For
            End If
        Next
    End Function

    Public Function Udoform(ByVal n As String, ByRef SBO_Application As SAPbouiCOM.Application) As Integer 'function to load Udt forms
        Dim oMenu As SAPbouiCOM.MenuItem
        Dim i As Integer
        oMenu = SBO_Application.Menus.Item("47616")
        For i = 0 To oMenu.SubMenus.Count - 1
            If oMenu.SubMenus.Item(i).String = n Then
                Return oMenu.SubMenus.Item(i).UID
                Exit For
            End If
        Next
    End Function

    Public Sub CFL_Vehicle(ByRef oForm As SAPbouiCOM.Form, ByVal sbo_application As SAPbouiCOM.Application, ByVal ocompany As SAPbobsCOM.Company) ' Business Partner
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            ''   Dim oRecordSet As SAPbobsCOM.Recordset

            oCFLs = oForm.ChooseFromLists

            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = sbo_application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_Items
            oCFLCreationParams.UniqueID = "Vehicle"
            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_AE_Driver"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_IS_NULL
            ' oCon.CondVal = ""
            oCFL.SetConditions(oCons)
        Catch
            '  MsgBox(Err.Description)
        End Try

    End Sub

    Public Function DriverMAster(ByRef oform As SAPbouiCOM.Form, ByRef oApplication As SAPbouiCOM.Application) As Boolean
        Try

            If oform.Items.Item("4").Specific.String = "" Then
                oform.Items.Item("4").Specific.active = True
                oApplication.StatusBar.SetText("Driver Name should not be empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False
            End If

            ''If oform.Items.Item("14").Specific.String = "" Then
            ''    oform.Items.Item("14").Specific.active = True
            ''    oApplication.StatusBar.SetText("DOB should not be empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            ''    Return False
            ''End If

            ''If oform.Items.Item("23").Specific.String = "" Then
            ''    oform.Items.Item("23").Specific.active = True
            ''    oApplication.StatusBar.SetText("License No should not be empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            ''    Return False
            ''End If

            ''If oform.Items.Item("27").Specific.String = "" Then
            ''    oform.Items.Item("27").Specific.active = True
            ''    oApplication.StatusBar.SetText("License Expiry Date should not be empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            ''    Return False
            ''End If

            ' ''If oform.Items.Item("29").Specific.String = "" Then
            ' ''    oform.Items.Item("29").Specific.active = True
            ' ''    oApplication.StatusBar.SetText("Passport No should not be empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            ' ''    Return False
            ' ''End If

            ' ''If oform.Items.Item("33").Specific.String = "" Then
            ' ''    oform.Items.Item("33").Specific.active = True
            ' ''    oApplication.StatusBar.SetText("Passport Expiry Date should not be empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            ' ''    Return False
            ' ''End If

            ''If oform.Items.Item("36").Specific.String = "" Then
            ''    oform.Items.Item("36").Specific.active = True
            ''    oApplication.StatusBar.SetText("Handphone should not be empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            ''    Return False
            ''End If

            ''If oform.Items.Item("38").Specific.String = "" Then
            ''    oform.Items.Item("38").Specific.active = True
            ''    oApplication.StatusBar.SetText("Local Address should not be empty ........ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            ''    Return False
            ''End If

            Return True

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try


    End Function

    Public Function UDO_ADD_DuplicateRA(ByRef oform As SAPbouiCOM.Form, ByRef oCompany As SAPbobsCOM.Company, ByRef oApplication As SAPbouiCOM.Application) As Boolean

        Try

            '--------------------
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
            Dim oCompanyService As SAPbobsCOM.CompanyService
            Dim opt As SAPbouiCOM.OptionBtn = oform.Items.Item("195").Specific
            Dim opt1 As SAPbouiCOM.OptionBtn = oform.Items.Item("196").Specific
            Dim opt2 As SAPbouiCOM.OptionBtn

            oCompanyService = oCompany.GetCompanyService

            ' UDO Name
            oGeneralService = oCompanyService.GetGeneralService("AE_Sbooking")

            'Create data for Document field in main UDO
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            ' oGeneralData.SetProperty("CreateDate", System.DateTime.Parse(oform.Items.Item("16").Specific.string, format1, Globalization.DateTimeStyles.None))
            ' oGeneralData.SetProperty("U_period2", System.DateTime.Parse(oform.Items.Item("5").Specific.string, format1, Globalization.DateTimeStyles.None))
            oGeneralData.SetProperty("U_AE_Bcode", oform.Items.Item("Item_2").Specific.string)
            oGeneralData.SetProperty("U_AE_Bname", oform.Items.Item("Item_3").Specific.string)
            oGeneralData.SetProperty("U_AE_Cno", oform.Items.Item("Item_7").Specific.string)
            oGeneralData.SetProperty("U_AE_Address", oform.Items.Item("Item_5").Specific.string)
            oGeneralData.SetProperty("U_AE_co", oform.Items.Item("199").Specific.string)
            oGeneralData.SetProperty("U_AE_ct", Trim(oform.Items.Item("201").Specific.value))
            oGeneralData.SetProperty("U_AE_Contract", oform.Items.Item("Item_9").Specific.string)
            oGeneralData.SetProperty("U_AE_Atten", oform.Items.Item("235").Specific.value)

            If opt.Selected = True Then
                oGeneralData.SetProperty("U_AE_Term", "1")
                oGeneralData.SetProperty("U_AE_NXDT", oform.Items.Item("209").Specific.string)
            Else
                oGeneralData.SetProperty("U_AE_Term", "2")
            End If

            oGeneralData.SetProperty("U_AE_Status", "Open")
            oGeneralData.SetProperty("U_AE_Sdate", oform.Items.Item("203").Specific.string)
            oGeneralData.SetProperty("U_AE_Edate", oform.Items.Item("205").Specific.string)
            'oGeneralData.SetProperty("U_AE_Dcode", oform.Items.Item("Item_30").Specific.string)
            oGeneralData.SetProperty("U_AE_DName", oform.Items.Item("Item_31").Specific.string)
            oGeneralData.SetProperty("U_AE_Dadd", oform.Items.Item("Item_34").Specific.string)
            oGeneralData.SetProperty("U_AE_Dcno", oform.Items.Item("Item_36").Specific.string)
            oGeneralData.SetProperty("U_AE_Occuption", oform.Items.Item("Item_38").Specific.string)
            oGeneralData.SetProperty("U_AE_Nation", oform.Items.Item("Item_40").Specific.string)
            oGeneralData.SetProperty("U_AE_DOB", oform.Items.Item("Item_42").Specific.string)
            oGeneralData.SetProperty("U_AE_License", oform.Items.Item("Item_44").Specific.string)
            oGeneralData.SetProperty("U_AE_Pissue", oform.Items.Item("Item_49").Specific.string)
            oGeneralData.SetProperty("U_AE_Exdate", oform.Items.Item("Item_50").Specific.string)
            oGeneralData.SetProperty("U_AE_Passno", oform.Items.Item("Item_51").Specific.string)
            oGeneralData.SetProperty("U_AE_Pissuepno", oform.Items.Item("Item_52").Specific.string)
            oGeneralData.SetProperty("U_AE_Pexdate", oform.Items.Item("Item_38").Specific.string)
           
            'oGeneralData.SetProperty("U_AE_Dcode1", oform.Items.Item("Item_57").Specific.string)
            oGeneralData.SetProperty("U_AE_DName1", oform.Items.Item("Item_58").Specific.string)
            oGeneralData.SetProperty("U_AE_Dadd1", oform.Items.Item("Item_60").Specific.string)
            oGeneralData.SetProperty("U_AE_Dcno1", oform.Items.Item("Item_62").Specific.string)
            oGeneralData.SetProperty("U_AE_Occuption1", oform.Items.Item("Item_64").Specific.string)
            oGeneralData.SetProperty("U_AE_Nation1", oform.Items.Item("Item_66").Specific.string)
            oGeneralData.SetProperty("U_AE_DOB1", oform.Items.Item("Item_68").Specific.string)
            oGeneralData.SetProperty("U_AE_License1", oform.Items.Item("Item_70").Specific.string)
            oGeneralData.SetProperty("U_AE_Pissue1", oform.Items.Item("Item_75").Specific.string)
            oGeneralData.SetProperty("U_AE_Exdate1", oform.Items.Item("Item_76").Specific.string)
            oGeneralData.SetProperty("U_AE_Passno1", oform.Items.Item("Item_77").Specific.string)
            oGeneralData.SetProperty("U_AE_Pissuepno1", oform.Items.Item("Item_78").Specific.string)
            oGeneralData.SetProperty("U_AE_Pexdate1", oform.Items.Item("Item_80").Specific.string)

            oGeneralData.SetProperty("U_AE_Vregno", oform.Items.Item("Item_82").Specific.string)
            oGeneralData.SetProperty("U_AE_Vdes", oform.Items.Item("Item_101").Specific.string)
            oGeneralData.SetProperty("U_AE_Vmodel", oform.Items.Item("Item_84").Specific.string)
            'oGeneralData.SetProperty("U_AE_expecD", oform.Items.Item("Item_86").Specific.string)
            ' oGeneralData.SetProperty("U_AE_expecT", oform.Items.Item("187").Specific.string)
            oGeneralData.SetProperty("U_AE_Vexten", oform.Items.Item("Item_88").Specific.string)
            oGeneralData.SetProperty("U_AE_Vout", oform.Items.Item("Item_90").Specific.string)
            oGeneralData.SetProperty("U_AE_Vin", oform.Items.Item("Item_95").Specific.string)

            ' oGeneralData.SetProperty("U_AE_Vkmin", oform.Items.Item("Item_96").Specific.string)
            oGeneralData.SetProperty("U_AE_Vdatein", oform.Items.Item("Item_97").Specific.string)
            oGeneralData.SetProperty("U_AE_Vtimein", oform.Items.Item("Item_102").Specific.string)
            ' oGeneralData.SetProperty("U_AE_Vkmout", oform.Items.Item("Item_98").Specific.string)
            oGeneralData.SetProperty("U_AE_Vdatetout", oform.Items.Item("Item_100").Specific.string)
            oGeneralData.SetProperty("U_AE_Vtimeout", oform.Items.Item("Item_103").Specific.string)
            oGeneralData.SetProperty("U_AE_semp", oform.Items.Item("215").Specific.string)

            opt = oform.Items.Item("Item_104").Specific
            opt1 = oform.Items.Item("Item_105").Specific
            opt2 = oform.Items.Item("Item_106").Specific

            If opt.Selected = True Then
                oGeneralData.SetProperty("U_AE_charges", "1")
            ElseIf opt1.Selected = True Then
                oGeneralData.SetProperty("U_AE_charges", "2")
            Else
                oGeneralData.SetProperty("U_AE_charges", "3")
            End If

            oGeneralData.SetProperty("U_AE_rate", CDbl(oform.Items.Item("Item_109").Specific.string))
            oGeneralData.SetProperty("U_AE_dwm", oform.Items.Item("Item_111").Specific.string)
            oGeneralData.SetProperty("U_AE_stot", CDbl(oform.Items.Item("Item_113").Specific.string))
            oGeneralData.SetProperty("U_AE_PAI", CDbl(oform.Items.Item("Item_115").Specific.string))
            oGeneralData.SetProperty("U_AE_PAI1", Trim(oform.Items.Item("226").Specific.value))
            oGeneralData.SetProperty("U_AE_CDW", CDbl(oform.Items.Item("Item_123").Specific.string))
            oGeneralData.SetProperty("U_AE_CDW1", Trim(oform.Items.Item("228").Specific.value))
            oGeneralData.SetProperty("U_AE_Dcfees", CDbl(oform.Items.Item("Item_125").Specific.string))
            oGeneralData.SetProperty("U_AE_Ocharges", CDbl(oform.Items.Item("Item_118").Specific.string))
            oGeneralData.SetProperty("U_AE_Rcharg", CDbl(oform.Items.Item("Item_119").Specific.string))
            oGeneralData.SetProperty("U_AE_BGST", CDbl(oform.Items.Item("Item_129").Specific.string))
            oGeneralData.SetProperty("U_AE_GST", CDbl(oform.Items.Item("Item_131").Specific.string))
            oGeneralData.SetProperty("U_AE_Netc", CDbl(oform.Items.Item("Item_137").Specific.string))
            oGeneralData.SetProperty("U_AE_Exliability", CDbl(oform.Items.Item("Item_141").Specific.string))
            oGeneralData.SetProperty("U_AE_SPLIB", CDbl(oform.Items.Item("Item_145").Specific.string))
            oGeneralData.SetProperty("U_AE_GSTP", oform.Items.Item("188").Specific.string)
            oGeneralData.SetProperty("U_AE_petrol", oform.Items.Item("Item_127").Specific.string)
            oGeneralData.SetProperty("U_AE_Rdesc", oform.Items.Item("Item_139").Specific.string)
            oGeneralData.SetProperty("U_AE_Des", oform.Items.Item("Item_121").Specific.string)
            oGeneralData.SetProperty("U_AE_Pay", Trim(oform.Items.Item("Item_158").Specific.value))

            oGeneralData.SetProperty("U_AE_SPD", CDbl(oform.Items.Item("Item_143").Specific.string))
            oGeneralData.SetProperty("U_AE_SPT", CDbl(oform.Items.Item("Item_145").Specific.string))
            oGeneralData.SetProperty("U_AE_SPGST", CDbl(oform.Items.Item("Item_147").Specific.string))
            oGeneralData.SetProperty("U_AE_SPNET", CDbl(oform.Items.Item("Item_149").Specific.string))
            oGeneralData.SetProperty("U_AE_SPLIB", CDbl(oform.Items.Item("Item_154").Specific.string))

            oGeneralData.SetProperty("U_AE_RAPbyc", oform.Items.Item("217").Specific.string)
            oGeneralData.SetProperty("U_AE_RAPbyn", oform.Items.Item("218").Specific.string)

            oGeneralData.SetProperty("U_AE_Percode", oform.Items.Item("Item_156").Specific.string)
            oGeneralData.SetProperty("U_AE_Perpared", oform.Items.Item("189").Specific.string)
            oGeneralData.SetProperty("U_AE_Invcode", oform.Items.Item("Item_151").Specific.string)
            oGeneralData.SetProperty("U_AE_Invoice", oform.Items.Item("190").Specific.string)
            oGeneralData.SetProperty("U_AE_SPRemarks", oform.Items.Item("Item_152").Specific.String)




            'Add the new row, including children, to database
            oGeneralParams = oGeneralService.Add(oGeneralData)

            oApplication.MessageBox(" Rental Agreement No - " & oGeneralParams.GetProperty("DocEntry") & " Added Sucessfully Based On This Information ..............", 1, "Ok")
            '            MsgBox("Record added")
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try

    End Function


    Public Function AR_InvoiceTimeCalculation(ByVal mjs As Integer, ByRef oform As SAPbouiCOM.Form, ByRef oApplication As SAPbouiCOM.Application) As Boolean

        Try

            Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("39").Specific
            Dim Tmp_date, Tmp_date1, Tmp_date2, Tmp_date3 As Date
            Dim Tmp_sub As TimeSpan
            ' MsgBox(oMAtrix.Columns.Item("TimeOUT").Cells.Item(mjs).Specific.String.ToString.Substring(0, 2) & ":" & oMAtrix.Columns.Item("TimeOUT").Cells.Item(mjs).Specific.String.ToString.Substring(2, 2) & ":00")
            If oMAtrix.Columns.Item("TimeOUT").Cells.Item(mjs).Specific.String.length = 4 Then
                If IsDate(oMAtrix.Columns.Item("TimeOUT").Cells.Item(mjs).Specific.String.ToString.Substring(0, 2) & ":" & oMAtrix.Columns.Item("TimeOUT").Cells.Item(mjs).Specific.String.ToString.Substring(2, 2) & ":00") = False Then
                    oApplication.StatusBar.SetText("Invalid Time Format ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Return False
                Else

                    oMAtrix.Columns.Item("TimeOUT").Cells.Item(mjs).Specific.String = oMAtrix.Columns.Item("TimeOUT").Cells.Item(mjs).Specific.String.ToString.Substring(0, 2) & ":" & oMAtrix.Columns.Item("TimeOUT").Cells.Item(mjs).Specific.String.ToString.Substring(2, 2)
                    Tmp_date1 = oMAtrix.Columns.Item("TimeOUT").Cells.Item(mjs).Specific.String & ":00"
                End If

            Else
                If IsDate(oMAtrix.Columns.Item("TimeOUT").Cells.Item(mjs).Specific.String & ":00") = False Then
                    oApplication.StatusBar.SetText("Invalid Time Format ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Return False
                Else
                    Tmp_date1 = oMAtrix.Columns.Item("TimeOUT").Cells.Item(mjs).Specific.String
                End If
            End If

            Tmp_date = oMAtrix.Columns.Item("TimeIN").Cells.Item(mjs).Specific.String & ":00"

            Special_Hours = 0
            Special_Mins = 0
            Normal_Hours = 0
            Normal_Mins = 0

            If Tmp_date < Tmp_date1 Then

                Tmp_sub = Tmp_date1.Subtract(Tmp_date)
                Normal_Hours = Normal_Hours + Tmp_sub.Hours
                Normal_Mins = Normal_Mins + Tmp_sub.Minutes

                If Special_TimeStart < Tmp_date1 Then
                    Tmp_sub = Tmp_date1.Subtract(Special_TimeStart)
                    Special_Hours = Special_Hours + Tmp_sub.Hours
                    Special_Mins = Special_Mins + Tmp_sub.Minutes
                End If

            Else
                Tmp_date2 = "00:00:00" '"23:59:00"
                Tmp_date3 = "23:59:00"

                Tmp_sub = Tmp_date3.Subtract(Tmp_date) ' 0
                Normal_Hours = Normal_Hours + Math.Abs(Tmp_sub.Hours)
                Normal_Mins = Normal_Mins + Math.Abs(Tmp_sub.Minutes) + 1

                Tmp_sub = Tmp_date1.Subtract(Tmp_date2) '1
                Normal_Hours = Normal_Hours + Math.Abs(Tmp_sub.Hours)
                Normal_Mins = Normal_Mins + Math.Abs(Tmp_sub.Minutes)


                Tmp_sub = Tmp_date3.Subtract(Special_TimeStart)
                Special_Hours = Special_Hours + Math.Abs(Tmp_sub.Hours)
                Special_Mins = Special_Mins + Math.Abs(Tmp_sub.Minutes) + 1

                Tmp_sub = Tmp_date1.Subtract(Tmp_date2) '1
                Special_Hours = Special_Hours + Math.Abs(Tmp_sub.Hours)
                Special_Mins = Special_Mins + Math.Abs(Tmp_sub.Minutes)

            End If

            Tmp_sub = New TimeSpan(0, (Normal_Hours * 60) + Normal_Mins, 0)

            If Tmp_sub.Hours < 3 Then
                oMAtrix.Columns.Item("THOW").Cells.Item(mjs).Specific.String = 3

            Else

                Select Case Tmp_sub.Minutes
                    Case 1 To 39
                        oMAtrix.Columns.Item("THOW").Cells.Item(mjs).Specific.String = CDbl(Tmp_sub.Hours) + 0.5
                    Case 40 To 60
                        oMAtrix.Columns.Item("THOW").Cells.Item(mjs).Specific.String = Math.Ceiling(CDbl(Tmp_sub.Hours & "." & Tmp_sub.Minutes))
                    Case Else
                        oMAtrix.Columns.Item("THOW").Cells.Item(mjs).Specific.String = Tmp_sub.Hours
                End Select
            End If
            
            Tmp_sub = New TimeSpan(0, (Special_Hours * 60) + Special_Mins, 0)

            Select Case Tmp_sub.Minutes
                Case 1 To 60
                    oMAtrix.Columns.Item("EMH").Cells.Item(mjs).Specific.String = Math.Ceiling(CDbl(Tmp_sub.Hours & "." & Tmp_sub.Minutes))
                Case Else
                    oMAtrix.Columns.Item("EMH").Cells.Item(mjs).Specific.String = Tmp_sub.Hours
            End Select

            oMAtrix.Columns.Item("12").Cells.Item(mjs).Specific.String = (CDbl(oMAtrix.Columns.Item("THOW").Cells.Item(mjs).Specific.String) * CDbl(oMAtrix.Columns.Item("HR").Cells.Item(mjs).Specific.String)) + _
                                      (CDbl(oMAtrix.Columns.Item("EMH").Cells.Item(mjs).Specific.String) * CDbl(oMAtrix.Columns.Item("EMR").Cells.Item(mjs).Specific.String))

            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function


End Module
