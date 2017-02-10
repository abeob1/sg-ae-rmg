Imports System.IO

Imports System.Drawing.Printing

Public Class SelfDrive

    Dim WithEvents Oapplication_SD As SAPbouiCOM.Application
    Dim Ocompany_SD As New SAPbobsCOM.Company
    Private FileName As String


    Public Sub New(ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company)

        Oapplication_SD = oApplication
        Ocompany_SD = oCompany

    End Sub

    Private Sub Oapplication_SD_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles Oapplication_SD.FormDataEvent

        If BusinessObjectInfo.FormUID = "SDB" Then
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                Try
                    Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.Item("SDB")
                    Dim orset As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim VNO As String = oform.Items.Item("Item_82").Specific.String 'Trim(oform.DataSources.DBDataSources.Item(0).GetValue("U_AE_Vregno", 0))
                    Dim Mileage As String = oform.Items.Item("Item_179").Specific.String 'Trim(oform.DataSources.DBDataSources.Item(0).GetValue("U_AE_VMOdo", 0))
                    If VNO <> "" And Mileage <> "" Then
                        orset.DoQuery("update OITM set [U_AE_RKM] = '" & Mileage & "' WHERE [ItemCode]  = '" & VNO & "'")
                    End If
                Catch ex As Exception
                    Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False
                    Exit Try
                End Try

            End If
        End If
    End Sub

    Private Sub Oapplication_SD_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles Oapplication_SD.ItemEvent

        If pVal.FormUID = "SDB" Then
            If pVal.Before_Action = True Then

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.InnerEvent = False Then
                    Try

                        If pVal.ItemUID = "187" Or pVal.ItemUID = "Item_102" Or pVal.ItemUID = "Item_103" Or pVal.ItemUID = "Item_177" Or pVal.ItemUID = "Item_193" Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                                ' MsgBox((oform.Items.Item(pVal.ItemUID).Specific.String.ToString.Substring(0, 2) & ":" & oform.Items.Item(pVal.ItemUID).Specific.String.ToString.Substring(2, 2) & ":00"))
                                If oform.Items.Item(pVal.ItemUID).Specific.String <> "" Then
                                    If oform.Items.Item(pVal.ItemUID).Specific.String.ToString.Length = "4" Then
                                        If IsDate(oform.Items.Item(pVal.ItemUID).Specific.String.ToString.Substring(0, 2) & ":" & oform.Items.Item(pVal.ItemUID).Specific.String.ToString.Substring(2, 2) & ":00") = False Then
                                            oform.Items.Item(pVal.ItemUID).Specific.active = True
                                            Oapplication_SD.StatusBar.SetText("Invalid Time format ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Try
                                        Else
                                            oform.Items.Item(pVal.ItemUID).Specific.String = oform.Items.Item(pVal.ItemUID).Specific.String.ToString.Substring(0, 2) & ":" & oform.Items.Item(pVal.ItemUID).Specific.String.ToString.Substring(2, 2)
                                        End If
                                    ElseIf oform.Items.Item(pVal.ItemUID).Specific.String.ToString.Length = "5" Then
                                        If IsDate(oform.Items.Item(pVal.ItemUID).Specific.String & ":00") = False Then
                                            oform.Items.Item(pVal.ItemUID).Specific.active = True
                                            Oapplication_SD.StatusBar.SetText("Invalid Time format ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Try
                                        End If
                                    Else

                                        oform.Items.Item(pVal.ItemUID).Specific.active = True
                                        Oapplication_SD.StatusBar.SetText("Invalid Time format ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Try

                                    End If
                                End If
                            Catch ex As Exception
                            End Try
                            Exit Sub
                        End If


                        If pVal.ItemUID = "Item_115" Or pVal.ItemUID = "Item_109" Or pVal.ItemUID = "Item_97" Or pVal.ItemUID = "Item_111" Or pVal.ItemUID = "Item_123" Or pVal.ItemUID = "Item_125" Or pVal.ItemUID = "Item_118" Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                                Dim oCheck As SAPbouiCOM.CheckBox = oform.Items.Item("1000003").Specific
                                Dim ooption As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_104").Specific
                                Dim ooption1 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_105").Specific
                                Dim ooption2 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_106").Specific
                                Dim inoofdays As Integer
                                If pVal.ItemUID = "Item_97" Then


                                    If oform.Items.Item("Item_97").Specific.String <> "" Then


                                        Dim Sdate As Date = System.DateTime.Parse(oform.Items.Item("Item_100").Specific.String, format1, Globalization.DateTimeStyles.None)
                                        Dim Edate As Date = System.DateTime.Parse(oform.Items.Item("Item_97").Specific.String, format1, Globalization.DateTimeStyles.None)

                                        If ooption.Selected = True Then
                                            oform.Items.Item("Item_111").Specific.String = Edate.Subtract(Sdate).Days
                                            oform.Items.Item("Item_113").Specific.String = CDbl(oform.Items.Item("Item_109").Specific.String) * CDbl(oform.Items.Item("Item_111").Specific.String)
                                        ElseIf ooption1.Selected = True Then
                                            oform.Items.Item("Item_111").Specific.String = Edate.Subtract(Sdate).Days
                                            oform.Items.Item("Item_113").Specific.String = (CDbl(oform.Items.Item("Item_109").Specific.String) / 7) * CDbl(oform.Items.Item("Item_111").Specific.String)
                                        ElseIf ooption2.Selected = True Then
                                            oform.Items.Item("Item_111").Specific.String = Format(Edate.Subtract(Sdate).Days / 30, "0.0")
                                            oform.Items.Item("Item_113").Specific.String = CDbl(oform.Items.Item("Item_109").Specific.String) * CDbl(oform.Items.Item("Item_111").Specific.String)
                                        End If
                                    Else
                                        oform.Items.Item("Item_111").Specific.String = 0
                                        oform.Items.Item("Item_113").Specific.String = 0
                                    End If

                                End If

                                If Not String.IsNullOrEmpty(oform.Items.Item("Item_111").Specific.String) Then
                                    inoofdays = oform.Items.Item("Item_111").Specific.String
                                Else
                                    inoofdays = 0
                                End If
                                oform.Items.Item("Item_113").Specific.String = CDbl(oform.Items.Item("Item_109").Specific.String) * inoofdays

                                If oCheck.Checked = True Then
                                    oform.Items.Item("Item_129").Specific.String = CDbl(oform.Items.Item("Item_113").Specific.String) + (CDbl(oform.Items.Item("Item_123").Specific.String) * inoofdays) + (CDbl(oform.Items.Item("Item_115").Specific.String) * inoofdays) + CDbl(oform.Items.Item("Item_125").Specific.String) + CDbl(oform.Items.Item("Item_118").Specific.String)
                                    oform.Items.Item("Item_131").Specific.String = CDbl(oform.Items.Item("Item_129").Specific.String) * (7 / 100)
                                    oform.Items.Item("Item_137").Specific.String = CDbl(oform.Items.Item("Item_129").Specific.String) + CDbl(oform.Items.Item("Item_131").Specific.String)
                                Else
                                    oform.Items.Item("Item_129").Specific.String = CDbl(oform.Items.Item("Item_113").Specific.String) + (CDbl(oform.Items.Item("Item_123").Specific.String) * inoofdays) + (CDbl(oform.Items.Item("Item_115").Specific.String) * inoofdays) + CDbl(oform.Items.Item("Item_125").Specific.String)
                                    oform.Items.Item("Item_131").Specific.String = CDbl(oform.Items.Item("Item_129").Specific.String) * (7 / 100)
                                    oform.Items.Item("Item_137").Specific.String = CDbl(oform.Items.Item("Item_129").Specific.String) + CDbl(oform.Items.Item("Item_131").Specific.String) + CDbl(oform.Items.Item("Item_118").Specific.String)

                                End If

                            Catch ex As Exception

                            End Try
                            Exit Sub
                        End If

                        If pVal.ItemUID = "203" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm

                                oform.Items.Item("Item_100").Specific.String = oform.Items.Item("203").Specific.String

                            Catch ex As Exception

                            End Try
                        End If

                    Catch ex As Exception
                        Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False
                        Exit Try
                    End Try
                    Exit Sub
                End If


                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                    If pVal.ItemUID = "242" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Dim oform As SAPbouiCOM.Form
                        Try
                            oform = Oapplication_SD.Forms.ActiveForm
                            Dim oDate As Date = Nothing
                            Dim sMonth As String = String.Empty
                            Dim sDay As String = String.Empty
                            Dim sDocEntry As String = String.Empty
                            Dim opt As SAPbouiCOM.OptionBtn = Nothing
                            Dim opt1 As SAPbouiCOM.OptionBtn = Nothing
                            Dim oopt As SAPbouiCOM.OptionBtn = Nothing
                            Dim oopt1 As SAPbouiCOM.OptionBtn = Nothing
                            Dim oopt3 As SAPbouiCOM.OptionBtn = Nothing
                            Dim oCombo As SAPbouiCOM.ComboBox = Nothing

                            If CDbl(oform.Items.Item("241").Specific.String) <= 0 Then
                                oform.Items.Item("241").Specific.active = True
                                Oapplication_SD.SetStatusBarMessage("Invoice Amount Should not blank ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                BubbleEvent = False
                                Exit Sub
                            End If

                            oDate = DateTime.ParseExact(GateDate(oform.Items.Item("209").Specific.String, Ocompany_SD), "yyyyMMdd", Nothing)

                            sMonth = oDate.Month + 1
                            If sMonth = 13 Then
                                sMonth = 1
                                oDate = oDate.AddYears(1)
                            End If
                            sDay = Trim(oform.Items.Item("237").Specific.value)
                            sDocEntry = oform.Items.Item("Item_0").Specific.String

                            p_bSDBooking = True
                            oform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            p_bSDBooking = False

                            If SD_InvoiceGeneration(oform.Items.Item("Item_2").Specific.String, oform.Items.Item("Item_14").Specific.String, "Self Drive Billing for this Booking No : " & oform.Items.Item("Item_14").Specific.String, SD_GLACC, _
                                                  "SO", CDbl(oform.Items.Item("241").Specific.String), "Self Drive Billing for this Booking No : " & oform.Items.Item("Item_14").Specific.String, oform.Items.Item("209").Specific.String, Month(oDate), sDocEntry, oDate.Year & sMonth.PadLeft(2, "0"c) & sDay.PadLeft(2, "0"c)) = False Then
                                BubbleEvent = False
                                Exit Try
                            End If
                            oform.Close()
                            LoadFromXML("SelfDriving_Booking.srf", Oapplication_SD)
                            oform = Oapplication_SD.Forms.Item("SDB")
                            oform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oform.Items.Item("Item_0").Visible = True
                            oform.Items.Item("Item_0").Specific.String = sDocEntry
                            p_bSDBooking = True
                            oform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            p_bSDBooking = False
                            oform.Visible = True
                            opt = oform.Items.Item("195").Specific
                            opt1 = oform.Items.Item("196").Specific
                            oopt = oform.Items.Item("Item_104").Specific
                            oopt1 = oform.Items.Item("Item_105").Specific
                            oopt3 = oform.Items.Item("Item_106").Specific
                            opt1.GroupWith("195")
                            oopt1.GroupWith("Item_104")
                            oopt3.GroupWith("Item_104")
                            ''oCombo = oform.Items.Item("237").Specific
                            ''For imjs As Integer = 1 To 31
                            ''    oCombo.ValidValues.Add(imjs, imjs)
                            ''Next
                            ''oCombo.ValidValues.Add("", "")
                            p_bSDBooking = True
                            NavigationValidation_SelfDriver(oform, Ocompany_SD, Oapplication_SD)
                            p_bSDBooking = False
                            oform.Items.Item("Item_22").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Dim ocombobutton As SAPbouiCOM.ButtonCombo = oform.Items.Item("210").Specific
                            ocombobutton.ValidValues.Add("Copy To", "Copy To A/R Invoice")
                            oform.Items.Item("210").Enabled = False
                            oform.DataBrowser.BrowseBy = "Item_14"

                            '  oform.Items.Item("Item_2").Specific.active = True
                            '  oform.Items.Item("Item_0").Visible = False

                        Catch ex As Exception
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Try
                        End Try
                        Exit Sub

                    End If


                    If pVal.ItemUID = "Item_20" Or pVal.ItemUID = "Item_21" Or pVal.ItemUID = "Item_22" Or pVal.ItemUID = "Item_23" Or pVal.ItemUID = "Item_24" _
                        Or pVal.ItemUID = "Item_25" Or pVal.ItemUID = "Item_26" Or pVal.ItemUID = "206" Or pVal.ItemUID = "233" Then
                        Dim oform As SAPbouiCOM.Form
                        Try
                            oform = Oapplication_SD.Forms.ActiveForm
                            oform.Freeze(True)
                            Select Case pVal.ItemUID

                                Case "Item_20"
                                    oform.PaneLevel = 1
                                Case "Item_21"
                                    oform.PaneLevel = 2
                                Case "Item_22"
                                    oform.PaneLevel = 3
                                    'If oform.Items.Item("Item_141").Specific.String = "0.00" Then
                                    '    oform.Items.Item("Item_141").Specific.String = ""
                                    'End If

                                Case "Item_23"
                                    oform.PaneLevel = 4
                                Case "Item_24"
                                    oform.PaneLevel = 5
                                Case "Item_25"
                                    oform.PaneLevel = 6
                                Case "Item_26"
                                    oform.PaneLevel = 7
                                Case "206"
                                    oform.PaneLevel = 8
                                Case "233"
                                    Dim ogrid As SAPbouiCOM.Grid = oform.Items.Item("234").Specific
                                    Try
                                        oform.DataSources.DataTables.Add("VT")
                                    Catch ex As Exception

                                    End Try
                                    Dim jj = "SELECT T0.[U_AE_Date] as 'Date', T0.[U_AE_Time] as 'Time', T0.[U_AE_Vno] as 'Vehicle No', case when [U_AE_Petrol] = '1' then 'Empty' when [U_AE_Petrol] = '2' then '1/8' when [U_AE_Petrol] = '3' then '1/4' when [U_AE_Petrol] = '4' then '3/8' when [U_AE_Petrol] = '5' then '1/2' when [U_AE_Petrol] = '6' then '5/8' when [U_AE_Petrol] = '7' then '3/4' when [U_AE_Petrol] = '8' then '7/8' when [U_AE_Petrol] = '9' then 'Full' end as 'Petrol', replace(convert(varchar,convert(Money, T0.[U_AE_Mileage]),1),'.00','') as 'Mileage', T0.[U_AE_Loc] as 'Location', T0.[U_AE_Remark] as 'Remarks', T0.[U_AE_Name] as 'Employee Name' FROM [dbo].[@AE_VTRACK]  T0 where isnull(t0.[U_AE_RA],'') like '" & oform.Items.Item("Item_14").Specific.String & "' order by T0.[U_AE_Date] "
                                    oform.DataSources.DataTables.Item(0).ExecuteQuery("SELECT T0.[U_AE_Date] as 'Date', T0.[U_AE_Time] as 'Time', T0.[U_AE_Vno] as 'Vehicle No', case when [U_AE_Petrol] = '1' then 'Empty' when [U_AE_Petrol] = '2' then '1/8' when [U_AE_Petrol] = '3' then '1/4' when [U_AE_Petrol] = '4' then '3/8' when [U_AE_Petrol] = '5' then '1/2' when [U_AE_Petrol] = '6' then '5/8' when [U_AE_Petrol] = '7' then '3/4' when [U_AE_Petrol] = '8' then '7/8' when [U_AE_Petrol] = '9' then 'Full' end as 'Petrol', replace(convert(varchar,convert(Money, T0.[U_AE_Mileage]),1),'.00','') as 'Mileage', T0.[U_AE_Loc] as 'Location', T0.[U_AE_Remark] as 'Remarks', T0.[U_AE_Name] as 'Employee Name' FROM [dbo].[@AE_VTRACK]  T0 where isnull(t0.[U_AE_RA],'') like '" & oform.Items.Item("Item_14").Specific.String & "' order by T0.[U_AE_Date] ")
                                    ogrid.DataTable = oform.DataSources.DataTables.Item("VT")
                                    ogrid.AutoResizeColumns()
                                    oform.PaneLevel = 10
                            End Select

                            oform.Freeze(False)
                        Catch ex As Exception
                            oform.Freeze(False)
                        End Try
                        Exit Sub
                    End If

                    If pVal.ItemUID = "211" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Dim SDPC_Thread As System.Threading.Thread
                        Try

                            Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                            Dim Docnum As Integer = 0
                            Docnum = CInt(oform.Items.Item("Item_14").Specific.String)

                            If Docnum = 0 Then
                                Oapplication_SD.StatusBar.SetText("Document Number should not be empty  .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End If

                            SDPC_Thread = New System.Threading.Thread(AddressOf Class_Report.Report_CallingFunction)
                            Class_Report.oApplication = Oapplication_SD
                            Class_Report.oCompany = Ocompany_SD
                            Class_Report.Report_Name = "AE_RP007_Contract.rpt"
                            Class_Report.Report_Parameter = "@DocNum"
                            Class_Report.Docnum = Docnum
                            Class_Report.Report_Title = "Contract Report"
                            If SDPC_Thread.IsAlive Then
                                Oapplication_SD.MessageBox("Report is already open....")
                            Else
                                SDPC_Thread.TrySetApartmentState(Threading.ApartmentState.STA)
                                Oapplication_SD.StatusBar.SetText("Contract Report Opening in process ......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                SDPC_Thread.Start()
                            End If

                        Catch ex As Exception
                            SDPC_Thread.Abort()
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End Try
                        Exit Sub
                    End If

                    If pVal.ItemUID = "214" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Dim SDCL_Thread As System.Threading.Thread
                        Try

                            Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                            Dim Docnum As Integer = 0
                            Docnum = CInt(oform.Items.Item("Item_14").Specific.String)

                            If Docnum = 0 Then
                                Oapplication_SD.StatusBar.SetText("Document Number should not be empty  .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End If

                            SDCL_Thread = New System.Threading.Thread(AddressOf Class_Report.Report_ExportToPDF)
                            Class_Report.oApplication = Oapplication_SD
                            Class_Report.oCompany = Ocompany_SD
                            Class_Report.Report_Name = "AE_RP004_VehicleCheckList.rpt"
                            Class_Report.Report_Parameter = "@DocNum"
                            Class_Report.Docnum = Docnum
                            Class_Report.Report_Title = "Vehicle Check List"
                            Class_Report.FileName = "RA"
                            If SDCL_Thread.IsAlive Then
                                Oapplication_SD.MessageBox("Report is already open....")
                            Else
                                SDCL_Thread.TrySetApartmentState(Threading.ApartmentState.STA)
                                Oapplication_SD.StatusBar.SetText("Vehicle Check List Report Opening in process ......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                SDCL_Thread.Start()
                            End If

                        Catch ex As Exception
                            SDCL_Thread.Abort()
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End Try
                        Exit Sub
                    End If

                    If pVal.ItemUID = "213" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Dim SD_Thread As System.Threading.Thread
                        Try

                            Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                            Dim Docnum As Integer = 0
                            Docnum = CInt(oform.Items.Item("Item_14").Specific.String)

                            If Docnum = 0 Then
                                Oapplication_SD.StatusBar.SetText("Document Number should not be empty  .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End If

                            SD_Thread = New System.Threading.Thread(AddressOf Class_Report.Report_CallingFunction)
                            Class_Report.oApplication = Oapplication_SD
                            Class_Report.oCompany = Ocompany_SD
                            Class_Report.Report_Name = "AE_RP002_SupplementaryAgreement.rpt"
                            Class_Report.Report_Parameter = "@DocNum"
                            Class_Report.Docnum = Docnum
                            Class_Report.Report_Title = "Supplementary Agreement Report"


                            If SD_Thread.IsAlive Then
                                Oapplication_SD.MessageBox("Report is already open....")
                            Else
                                SD_Thread.TrySetApartmentState(Threading.ApartmentState.STA)
                                Oapplication_SD.StatusBar.SetText("Supplementary Agreement Report Opening in process ......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                SD_Thread.Start()
                            End If

                        Catch ex As Exception
                            SD_Thread.Abort()
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End Try
                        Exit Sub
                    End If

                    If pVal.ItemUID = "212" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Dim SD1_Thread As System.Threading.Thread
                        Try

                            Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                            Dim Docnum As Integer = 0
                            Docnum = CInt(oform.Items.Item("Item_14").Specific.String)

                            If Docnum = 0 Then
                                Oapplication_SD.StatusBar.SetText("Document Number should not be empty  .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End If

                            ' SD1_Thread = New System.Threading.Thread(AddressOf Class_Report.Report_CallingFunction)

                            SD1_Thread = New System.Threading.Thread(AddressOf Class_Report.Report_ExportToPDF)

                            Class_Report.oApplication = Oapplication_SD
                            Class_Report.oCompany = Ocompany_SD
                            Class_Report.Report_Name = "AE_RP003_RentalAgreement.rpt"
                            Class_Report.Report_Parameter = "@DocNum"
                            Class_Report.Docnum = Docnum
                            Class_Report.Report_Title = "Rental Agreement Report"
                            Class_Report.FileName = "RA"

                            If SD1_Thread.IsAlive Then
                                Oapplication_SD.MessageBox("Report is already open....")
                            Else
                                SD1_Thread.TrySetApartmentState(Threading.ApartmentState.STA)
                                Oapplication_SD.StatusBar.SetText("Rental Agreement Report Opening in process ......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                SD1_Thread.Start()
                            End If

                        Catch ex As Exception
                            SD1_Thread.Abort()
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End Try
                        Exit Sub
                    End If

                    If pVal.ItemUID = "Item_208" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.Item("SDB")
                            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_207").Specific
                            Dim rowcount As Integer = oMatrix.RowCount
                            Dim DeleteFlag As Boolean = False

                            For mjs As Integer = 1 To rowcount
                                If mjs <= oMatrix.RowCount Then
                                    If oMatrix.IsRowSelected(mjs) = True Then
                                        oMatrix.DeleteRow(mjs)
                                        DeleteFlag = True
                                        mjs = mjs - 1
                                    End If
                                Else
                                    Exit For
                                End If
                            Next mjs

                            If DeleteFlag = True Then
                                For mjs As Integer = 1 To oMatrix.RowCount
                                    oMatrix.Columns.Item("#").Cells.Item(mjs).Specific.string = mjs
                                Next
                                If oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If
                            End If

                        Catch ex As Exception
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End Try
                        Exit Sub
                    End If


                    If pVal.ItemUID = "Item_209" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.Item("SDB")
                            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_207").Specific

                            Dim orset As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            orset.DoQuery("SELECT attachpath from oadp")
                            showOpenFileDialog()
                            If FileName <> "" Then
                                Dim file = New FileInfo(stFilePathAndName)
                                file.CopyTo(Path.Combine(orset.Fields.Item("attachpath").Value, file.Name), True)

                                oMatrix = oform.Items.Item("Item_207").Specific
                                If oMatrix.RowCount = 0 Then
                                    oMatrix.AddRow()
                                Else
                                    If oMatrix.Columns.Item("Col_0").Cells.Item(oMatrix.RowCount).Specific.String <> "" And oMatrix.Columns.Item("Col_1").Cells.Item(oMatrix.RowCount).Specific.String <> "" Then
                                        oMatrix.AddRow()
                                    End If
                                End If

                                oMatrix.Columns.Item("#").Cells.Item(oMatrix.RowCount).Specific.string = oMatrix.RowCount
                                oMatrix = oform.Items.Item("Item_207").Specific
                                oMatrix.Columns.Item("Col_1").Cells.Item(oMatrix.RowCount).Specific.string = orset.Fields.Item("attachpath").Value & FileName
                                oMatrix.Columns.Item("Col_2").Cells.Item(oMatrix.RowCount).Specific.string = FileName
                                oMatrix.Columns.Item("Col_0").Cells.Item(oMatrix.RowCount).Specific.string = Format(Now.Date, "yyyyMMdd")
                                'oMatrix.AutoResizeColumns()
                                stFilePathAndName = ""
                            End If

                            If oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            End If

                            oform.Freeze(True)
                            oMatrix.AutoResizeColumns()
                            oform.Freeze(False)

                        Catch ex As Exception
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Try
                        End Try
                        Exit Sub
                    End If

                    If pVal.ItemUID = "1000002" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                            If UDO_ADD_DuplicateRA(oform, Ocompany_SD, Oapplication_SD) = False Then
                                BubbleEvent = False
                                Exit Try
                            End If


                        Catch ex As Exception

                        End Try
                        Exit Sub
                    End If

                    If pVal.ItemUID = "1" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                            Dim opt As SAPbouiCOM.OptionBtn = oform.Items.Item("195").Specific
                            Dim oMatrx As SAPbouiCOM.Matrix = oform.Items.Item("1000001").Specific


                            If oform.Items.Item("Item_2").Specific.String = "" Then
                                oform.Items.Item("Item_2").Specific.active = True
                                Oapplication_SD.StatusBar.SetText("Customer should not be empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Sub
                            End If

                            If Trim(oform.Items.Item("201").Specific.value) = "Y" Then
                                If oform.Items.Item("Item_9").Specific.String = "" Then
                                    oform.Items.Item("Item_9").Specific.active = True
                                    Oapplication_SD.StatusBar.SetText("Contract No should not be empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If


                            If oform.Items.Item("203").Specific.String = "" Then
                                oform.Items.Item("203").Specific.active = True
                                Oapplication_SD.StatusBar.SetText("Start Date should not be empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Sub
                            End If

                            If oform.Items.Item("205").Specific.String = "" Then
                                oform.Items.Item("205").Specific.active = True
                                Oapplication_SD.StatusBar.SetText("Expected End Date should not be empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Sub
                            End If


                            If oform.Items.Item("Item_82").Specific.String = "" Then
                                Oapplication_SD.StatusBar.SetText("Vehicle Number should not be empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Sub
                            End If

                            ''If oform.Items.Item("Item_115").Specific.String <> "" Then
                            ''    If CInt(oform.Items.Item("Item_115").Specific.String) > 0 Then
                            ''        If Trim(oform.Items.Item("226").Specific.value) = "No" Then
                            ''            Oapplication_SD.StatusBar.SetText("Kindly choose the PAI ...... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            ''            BubbleEvent = False
                            ''            Exit Sub
                            ''        End If
                            ''    End If
                            ''End If


                            If Trim(oform.Items.Item("228").Specific.value) = "-" Then
                                Oapplication_SD.StatusBar.SetText("CDW should not be empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Sub
                            End If

                            If oform.Items.Item("Item_141").Specific.String <> "" Then
                                If CInt(oform.Items.Item("Item_141").Specific.String) <= 0 Then
                                    Oapplication_SD.StatusBar.SetText("Excess Liability should not be empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Else
                                Oapplication_SD.StatusBar.SetText("Excess Liability should not be empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Sub
                            End If

                            If oMatrx.RowCount = 0 Then
                                If opt.Selected = True Then
                                    oform.Items.Item("209").Specific.String = oform.Items.Item("203").Specific.String
                                End If

                            End If

                            If oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                If Trim(oform.Items.Item("Item_18").Specific.value) = "Billing" Then
                                    If opt.Selected = False Then
                                        oform.Items.Item("210").Enabled = True
                                    Else
                                        oform.Items.Item("210").Enabled = False
                                    End If
                                Else
                                    oform.Items.Item("210").Enabled = False
                                End If
                            End If

                            If oform.Items.Item("Item_18").Specific.selected.value = "Closed" Then
                                oform.Items.Item("Item_18").Specific.active = True
                                Oapplication_SD.StatusBar.SetText("Can`t set status manually to Closed ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Sub
                            End If

                        Catch ex As Exception
                        End Try
                        Exit Sub
                    End If
                End If
            End If
            If pVal.Before_Action = False Then

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then

                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvento = pVal
                    Dim sCFL_ID As String
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    Dim oForm As SAPbouiCOM.Form = Oapplication_SD.Forms.Item(FormUID)
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                    Try

                        If oCFLEvento.BeforeAction = False Then
                            Dim oDataTable As SAPbouiCOM.DataTable
                            Dim oedit1 As SAPbouiCOM.EditText
                            Dim oCombo As SAPbouiCOM.ComboBox
                            oDataTable = oCFLEvento.SelectedObjects
                            If pVal.ItemUID = "Item_2" Then 'Billing Code

                                Try
                                    Dim orset As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim orset1 As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    Dim opt As SAPbouiCOM.OptionBtn = oForm.Items.Item("Item_104").Specific
                                    Dim opt1 As SAPbouiCOM.OptionBtn = oForm.Items.Item("Item_105").Specific
                                    Dim opt2 As SAPbouiCOM.OptionBtn = oForm.Items.Item("Item_106").Specific
                                    Dim opt3 As SAPbouiCOM.OptionBtn = oForm.Items.Item("195").Specific
                                    Dim opt4 As SAPbouiCOM.OptionBtn = oForm.Items.Item("196").Specific

                                    ''                                orset.DoQuery("SELECT T0.[DocEntry], T0.[U_AE_Dcode], T0.[U_AE_DName], T0.[U_AE_Dadd], T0.[U_AE_Dcno], T0.[U_AE_Occuption], T0.[U_AE_Nation], T0.[U_AE_DOB], T0.[U_AE_License], " & _
                                    ''                                              "T0.[U_AE_Pissue], T0.[U_AE_Exdate], T0.[U_AE_Passno], T0.[U_AE_Pissuepno], T0.[U_AE_Pexdate], " & _
                                    ''                                              "T0.[U_AE_Dcode1], T0.[U_AE_DName1], T0.[U_AE_Dadd1], T0.[U_AE_Dcno1], T0.[U_AE_Occuption1], T0.[U_AE_Nation1], T0.[U_AE_DOB1], " & _
                                    ''                                              "T0.[U_AE_License1], T0.[U_AE_Pissue1], T0.[U_AE_Exdate1], T0.[U_AE_Passno1], T0.[U_AE_Pissuepno1], T0.[U_AE_Pexdate1] " & _
                                    ''                                              ",T0.[U_AE_Vregno],T0.[U_AE_Vdes],T0.[U_AE_Vmodel],T0.[U_AE_expecD],T0.[U_AE_expecT], " & _
                                    ''"T0.[U_AE_Vexten],T0.[U_AE_Vout],T0.[U_AE_Vin],T0.[U_AE_Vkmin], " & _
                                    ''"T0.[U_AE_Vdatein],T0.[U_AE_Vtimein],T0.[U_AE_Vkmout],T0.[U_AE_Vtimeout],T0.[U_AE_Vdatetout] " & _
                                    ''",T0.[U_AE_rate],T0.[U_AE_dwm],T0.[U_AE_stot],T0.[U_AE_PAI], " & _
                                    ''"T0.[U_AE_CDW],T0.[U_AE_Dcfees],T0.[U_AE_Ocharges],T0.[U_AE_Des],T0.[U_AE_Rcharg], " & _
                                    ''"T0.[U_AE_Rdesc], T0.[U_AE_petrol], T0.[U_AE_BGST], T0.[U_AE_GST], T0.[U_AE_Netc], " & _
                                    ''"T0.[U_AE_Exliability], T0.[U_AE_Term], T0.[U_AE_charges] FROM [dbo].[@AE_SBOOKING]  T0 WHERE T0.[U_AE_Bcode] = '" & oDataTable.GetValue("CardCode", 0) & "' ORDER BY cast(T0.[DocEntry] as integer) desc")

                                    orset1.DoQuery("select T1.cardcode , (ISNULL(T2.Street,'') +  ', ' + ISNULL(T2.Block,'') + char(13) + ISNULL(T2.City,'') + ', ' + ISNULL(T2.StreetNo,'') + CHAR(13) + isnull(T3.[Name], '') + ' ' + isnull(T2.ZipCode,'')) as 'Address' " & _
    "from OCRD T1 " & _
    "LEFT JOIN CRD1 T2 ON T1.CardCode=T2.CardCode and [AdresType]='B' " & _
    "LEFT JOIN OCRY T3 ON T2.Country=T3.Code where T1.cardcode = '" & oDataTable.GetValue("CardCode", 0) & "' " & _
    "group by T1.cardcode , ISNULL(T2.Street,'') +  ', ' + ISNULL(T2.Block,'') + char(13) + ISNULL(T2.City,'') + ', ' + ISNULL(T2.StreetNo,'') + CHAR(13) + isnull(T3.[Name], '') + ' ' + isnull(T2.ZipCode,'') ")

                                    oForm.Items.Item("Item_3").Specific.string = oDataTable.GetValue("CardName", 0)
                                    oForm.Items.Item("Item_5").Specific.string = orset1.Fields.Item("Address").Value  'oDataTable.GetValue("Address", 0) & vbCrLf & oDataTable.GetValue("City", 0) & ", " & oDataTable.GetValue("Country", 0) & ", " & oDataTable.GetValue("ZipCode", 0)
                                    oForm.Items.Item("Item_7").Specific.string = oDataTable.GetValue("Cellular", 0)
                                    oForm.Items.Item("199").Specific.string = oDataTable.GetValue("CardFName", 0)


                                    orset.DoQuery("SELECT T0.[FirstName] + ' ' + T0.[LastName] as 'Name' FROM OCPR T0 WHERE T0.[CardCode]  = '" & oDataTable.GetValue("CardCode", 0) & "'")
                                    oCombo = oForm.Items.Item("235").Specific

                                    For mjs As Integer = oCombo.ValidValues.Count - 1 To 0 Step -1
                                        oCombo.ValidValues.Remove(mjs, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next mjs
                                    Try
                                        For mjs As Integer = 1 To orset.RecordCount
                                            oCombo.ValidValues.Add(orset.Fields.Item("Name").Value, "")
                                            orset.MoveNext()
                                        Next mjs
                                    Catch ex As Exception
                                    End Try

                                    oCombo.Select(0)

                                    oForm.Items.Item("Item_2").Specific.string = oDataTable.GetValue("CardCode", 0)

                                Catch ex As Exception

                                End Try

                            ElseIf pVal.ItemUID = "Item_156" Then 'Billing Code
                                oForm.Items.Item("189").Specific.string = oDataTable.GetValue("firstName", 0) & " " & oDataTable.GetValue("lastName", 0)
                                oForm.Items.Item("Item_156").Specific.string = oDataTable.GetValue("empID", 0)

                            ElseIf pVal.ItemUID = "215" Then ' Sale Employee
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If
                                oForm.Items.Item("215").Specific.string = oDataTable.GetValue("SlpName", 0)
                            ElseIf pVal.ItemUID = "Item_151" Then 'Billing Code
                                oForm.Items.Item("190").Specific.string = oDataTable.GetValue("firstName", 0) & " " & oDataTable.GetValue("lastName", 0)
                                oForm.Items.Item("Item_151").Specific.string = oDataTable.GetValue("empID", 0)

                            ElseIf pVal.ItemUID = "Item_185" Then 'Billing Code
                                oForm.Items.Item("191").Specific.string = oDataTable.GetValue("firstName", 0) & " " & oDataTable.GetValue("lastName", 0)
                                oForm.Items.Item("Item_185").Specific.string = oDataTable.GetValue("empID", 0)

                            ElseIf pVal.ItemUID = "Item_200" Then 'Billing Code
                                oForm.Items.Item("192").Specific.string = oDataTable.GetValue("firstName", 0) & " " & oDataTable.GetValue("lastName", 0)
                                oForm.Items.Item("Item_200").Specific.string = oDataTable.GetValue("empID", 0)

                            ElseIf pVal.ItemUID = "Item_30" Then 'Order By
                                oForm.Items.Item("Item_31").Specific.string = oDataTable.GetValue("U_AE_Dname", 0)
                                oForm.Items.Item("Item_34").Specific.string = oDataTable.GetValue("U_AE_Ladd", 0)
                                oForm.Items.Item("Item_36").Specific.string = oDataTable.GetValue("U_AE_Hphone", 0)
                                oForm.Items.Item("Item_38").Specific.string = oDataTable.GetValue("U_AE_Occ", 0)
                                oForm.Items.Item("Item_40").Specific.string = oDataTable.GetValue("U_AE_COB", 0)
                                oForm.Items.Item("Item_42").Specific.string = Format(oDataTable.GetValue("U_AE_DOB", 0), "yyyyMMdd")
                                oForm.Items.Item("Item_44").Specific.string = oDataTable.GetValue("U_AE_LicenseNo", 0)
                                oForm.Items.Item("Item_49").Specific.string = oDataTable.GetValue("U_AE_Lplace", 0)
                                oForm.Items.Item("Item_50").Specific.string = Format(oDataTable.GetValue("U_AE_LEdate", 0), "yyyyMMdd")
                                oForm.Items.Item("Item_51").Specific.string = oDataTable.GetValue("U_AE_Pno", 0)
                                oForm.Items.Item("Item_52").Specific.string = oDataTable.GetValue("U_AE_Pplace", 0)
                                oForm.Items.Item("Item_54").Specific.string = Format(oDataTable.GetValue("U_AE_PEdate", 0), "yyyyMMdd")
                                oForm.Items.Item("Item_30").Specific.string = oDataTable.GetValue("U_AE_Dcode", 0)

                            ElseIf pVal.ItemUID = "Item_57" Then 'Issued By
                                oForm.Items.Item("Item_58").Specific.string = oDataTable.GetValue("U_AE_Dname", 0)
                                oForm.Items.Item("Item_60").Specific.string = oDataTable.GetValue("U_AE_Ladd", 0)
                                oForm.Items.Item("Item_62").Specific.string = oDataTable.GetValue("U_AE_Hphone", 0)
                                oForm.Items.Item("Item_64").Specific.string = oDataTable.GetValue("U_AE_Occ", 0)
                                oForm.Items.Item("Item_66").Specific.string = oDataTable.GetValue("U_AE_COB", 0)
                                oForm.Items.Item("Item_68").Specific.string = Format(oDataTable.GetValue("U_AE_DOB", 0), "yyyyMMdd")
                                oForm.Items.Item("Item_70").Specific.string = oDataTable.GetValue("U_AE_LicenseNo", 0)
                                oForm.Items.Item("Item_75").Specific.string = oDataTable.GetValue("U_AE_Lplace", 0)
                                oForm.Items.Item("Item_76").Specific.string = Format(oDataTable.GetValue("U_AE_LEdate", 0), "yyyyMMdd")
                                oForm.Items.Item("Item_77").Specific.string = oDataTable.GetValue("U_AE_Pno", 0)
                                oForm.Items.Item("Item_78").Specific.string = oDataTable.GetValue("U_AE_Pplace", 0)
                                oForm.Items.Item("Item_80").Specific.string = Format(oDataTable.GetValue("U_AE_PEdate", 0), "yyyyMMdd")
                                oForm.Items.Item("Item_57").Specific.string = oDataTable.GetValue("U_AE_Dcode", 0)

                            ElseIf pVal.ItemUID = "Item_82" Then 'Issued By
                                oForm.Items.Item("Item_101").Specific.string = oDataTable.GetValue("ItemName", 0)
                                oForm.Items.Item("Item_84").Specific.string = oDataTable.GetValue("U_AE_MODEL", 0)
                                oForm.Items.Item("Item_82").Specific.string = oDataTable.GetValue("ItemCode", 0)

                            End If
                        End If
                    Catch ex As Exception
                        'Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        BubbleEvent = False
                        Exit Try
                    End Try
                End If


                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then

                    If pVal.ItemUID = "237" Then
                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                            Dim oDate As Date = Nothing
                            Dim oDate_BCC As Date = Nothing
                            Dim oMonth As String = 0
                            Dim syear As String = 0
                            Dim oDay As String = 0
                            Dim iNoDays As Integer = 0
                            Dim iCunDays As Integer = 0
                            Dim dMonthlyrental As Double = 0

                            If String.IsNullOrEmpty(oform.Items.Item("237").Specific.selected.value) Then
                                Exit Sub
                            End If

                            oDate = DateTime.ParseExact(GateDate(oform.Items.Item("209").Specific.String, Ocompany_SD), "yyyyMMdd", Nothing)
                            oMonth = oDate.Month + 1
                            If oMonth = 13 Then
                                oMonth = 1
                                syear = oDate.AddYears(1).Year
                            Else
                                syear = oDate.Year
                            End If
                            oDay = oform.Items.Item("237").Specific.selected.value
                            oDate_BCC = DateTime.ParseExact(syear & oMonth.ToString.PadLeft(2, "0"c) & oDay.ToString.PadLeft(2, "0"c), "yyyyMMdd", Nothing)
                            dMonthlyrental = oform.Items.Item("Item_109").Specific.String
                            iNoDays = DateDiff(DateInterval.Day, oDate, oDate.AddMonths(1))  'System.DateTime.DaysInMonth(oDate.Year, oDate.Month)
                            iCunDays = DateDiff(DateInterval.Day, oDate, oDate_BCC)
                            oform.Items.Item("239").Specific.String = "Calculation Amount" & Environment.NewLine & "---------------------------" & Environment.NewLine &
                                "Monthly Rental           : " & dMonthlyrental &
                                 Environment.NewLine & "Days in Billing Cycle    : " & iNoDays &
                                Environment.NewLine & "Per Day Cost              : " & Format(dMonthlyrental / iNoDays, "0.0000") &
                                Environment.NewLine & "Invoice Date Range    : " & Format(oDate, "dd-MMM-yyyy") & " To " & Format(oDate_BCC.AddDays(-1), "dd-MMM-yyyy") &
                                Environment.NewLine & "Consumed Days         : " & iCunDays &
                                Environment.NewLine & "Total Invoice Amount : " & Format((dMonthlyrental / iNoDays) * iCunDays, "0.0000")

                            oform.Items.Item("241").Specific.String = Format((dMonthlyrental / iNoDays) * iCunDays, "0.0000")
                            oform.Items.Item("241").Enabled = True
                            oform.Items.Item("242").Enabled = True

                        Catch ex As Exception
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Try
                        End Try
                    End If

                    If pVal.ItemUID = "226" Then
                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                            If oform.Items.Item("226").Specific.selected.value.ToString.Trim = "Yes" Then
                                oform.Items.Item("Item_115").Enabled = True
                            Else
                                oform.Items.Item("Item_115").Enabled = False
                            End If

                        Catch ex As Exception

                        End Try
                    End If


                    If pVal.ItemUID = "228" Then
                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                            If oform.Items.Item("228").Specific.selected.value.ToString.Trim = "Yes" Then
                                oform.Items.Item("Item_123").Enabled = True
                            Else
                                oform.Items.Item("Item_123").Enabled = False
                            End If

                        Catch ex As Exception

                        End Try
                    End If


                    If pVal.ItemUID = "201" Then
                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm

                            If oform.Items.Item("201").Specific.selected.value = "Y" Then
                                oform.Items.Item("Item_9").Enabled = True
                            Else
                                oform.Items.Item("Item_9").Specific.String = ""
                                oform.Items.Item("Item_2").Specific.active = True
                                oform.Items.Item("Item_9").Enabled = False
                            End If

                            If oform.Items.Item("Item_18").Specific.selected.value = "Billing" Then
                                If oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If
                            End If

                        Catch ex As Exception
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Try
                        End Try

                    End If

                    '
                    If pVal.ItemUID = "210" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        Dim oform As SAPbouiCOM.Form
                        Dim oform1 As SAPbouiCOM.Form = Oapplication_SD.Forms.Item("SDB")
                        Dim inoofdays, irowcount As Integer
                        Dim iPAI, iCDW, dOC, dMR, dPC As Double
                        Dim dSDate, dEDate As Date
                        Dim sDesc As String = String.Empty

                        Try
                            Dim orset As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            If Trim(oform1.Items.Item("Item_18").Specific.value) <> "Billing" Then
                                Oapplication_SD.StatusBar.SetText("Could not generate an invoice for this status " & oform1.Items.Item("Item_18").Specific.value, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End If

                            If Not String.IsNullOrEmpty(oform1.Items.Item("Item_111").Specific.String) Then
                                inoofdays = oform1.Items.Item("Item_111").Specific.String
                            Else
                                Oapplication_SD.StatusBar.SetText("Number of Days / Months should not be empty .........!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End If

                            If Trim(oform1.Items.Item("226").Specific.value) = "Yes" Then
                                iPAI = CDbl(oform1.Items.Item("Item_115").Specific.String) * inoofdays
                            Else
                                iPAI = 0
                            End If

                            If Trim(oform1.Items.Item("228").Specific.value) = "Yes" Then
                                iCDW = CDbl(oform1.Items.Item("Item_123").Specific.String) * inoofdays
                            Else
                                iCDW = 0
                            End If

                            If Not String.IsNullOrEmpty(oform1.Items.Item("Item_118").Specific.string) Then
                                If CInt(oform1.Items.Item("Item_118").Specific.string) > 0 Then
                                    dOC = CDbl(oform1.Items.Item("Item_118").Specific.string)
                                Else
                                    dOC = 0
                                End If
                            Else
                                dOC = 0
                            End If

                            If Not String.IsNullOrEmpty(oform1.Items.Item("Item_119").Specific.string) Then
                                If CInt(oform1.Items.Item("Item_119").Specific.string) > 0 Then
                                    dMR = CDbl(oform1.Items.Item("Item_119").Specific.string)
                                Else
                                    dMR = 0
                                End If
                            Else
                                dMR = 0
                            End If

                            If Not String.IsNullOrEmpty(oform1.Items.Item("Item_127").Specific.string) Then
                                If CInt(oform1.Items.Item("Item_127").Specific.string) > 0 Then
                                    dPC = CDbl(oform1.Items.Item("Item_127").Specific.string)
                                Else
                                    dPC = 0
                                End If
                            Else
                                dPC = 0
                            End If

                            If Not String.IsNullOrEmpty(oform1.Items.Item("203").Specific.string) Then
                                dSDate = DateTime.ParseExact(GateDate(oform1.Items.Item("203").Specific.string, Ocompany_SD), "yyyyMMdd", Nothing)
                            End If

                            If Not String.IsNullOrEmpty(oform1.Items.Item("205").Specific.string) Then
                                dEDate = DateTime.ParseExact(GateDate(oform1.Items.Item("205").Specific.string, Ocompany_SD), "yyyyMMdd", Nothing)
                            End If

                            Dim ContactP As String = ""
                            If Trim(oform1.Items.Item("235").Specific.value) <> "" Then
                                orset.DoQuery("SELECT T0.[Name] FROM OCPR T0 WHERE isnull(T0.[FirstName],'') + ' ' +  isnull(T0.[LastName],'')  = '" & Trim(oform1.Items.Item("235").Specific.value) & "' and T0.CardCode = '" & oform1.Items.Item("Item_2").Specific.String & "'")
                                ContactP = Trim(orset.Fields.Item("Name").Value)
                            End If
                            Dim ocombobuttom As SAPbouiCOM.ButtonCombo = oform1.Items.Item("210").Specific
                            If ocombobuttom.Selected.Description = "Copy To A/R Invoice" Then
                                Invoice_Type = "SD"
                                Invoice_UDF = True
                                Oapplication_SD.ActivateMenuItem("2053")
                                oform = Oapplication_SD.Forms.GetFormByTypeAndCount(133, FormType_Invoice)

                                Dim oMAtrix_IN As SAPbouiCOM.Matrix = oform.Items.Item("39").Specific
                                Dim oColumn As SAPbouiCOM.Column
                                oMAtrix_IN.Clear()
                                oform.Freeze(True)
                                oColumn = oMAtrix_IN.Columns.Add("INDate", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                oColumn.TitleObject.Caption = "IN Date"
                                oColumn.DataBind.SetBound(True, "INV1", "U_AE_IND")
                                oColumn.Editable = False

                                oColumn = oMAtrix_IN.Columns.Add("TimeIN", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                oColumn.TitleObject.Caption = "Time IN"
                                oColumn.DataBind.SetBound(True, "INV1", "U_AE_INT")

                                oColumn = oMAtrix_IN.Columns.Add("OUTDate", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                oColumn.TitleObject.Caption = "Out Date"
                                oColumn.DataBind.SetBound(True, "INV1", "U_AE_OTD")
                                oColumn.Editable = False

                                oColumn = oMAtrix_IN.Columns.Add("TimeOUT", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                oColumn.TitleObject.Caption = "Time Out"
                                oColumn.DataBind.SetBound(True, "INV1", "U_AE_OTT")

                                oform.Items.Item("3").Specific.select("S")
                                oMAtrix_IN.AddRow()
                                oform.Items.Item("4").Specific.String = oform1.Items.Item("Item_2").Specific.String
                                oform.Items.Item("14").Specific.String = "SD " & oform1.Items.Item("Item_14").Specific.String

                                '' oMAtrix_IN.Columns.Item("1").Cells.Item(1).Specific.String = "Self Drive Billing for this Booking No : " & oform1.Items.Item("Item_14").Specific.String
                                sDesc = "Rental period from " & Format(dSDate, "dd/MM/yyyy") & " to " & Format(dEDate, "dd/MM/yyyy") & " (" & inoofdays & "Days)"
                                oMAtrix_IN.Columns.Item("1").Cells.Item(1).Specific.String = sDesc

                                oMAtrix_IN.Columns.Item("12").Cells.Item(1).Specific.String = CDbl(oform1.Items.Item("Item_129").Specific.String) - (iPAI + iCDW)
                                oMAtrix_IN.Columns.Item("2").Cells.Item(1).Specific.String = SD_GLACC

                                oMAtrix_IN.Columns.Item("INDate").Cells.Item(1).Specific.String = oform1.Items.Item("Item_191").Specific.String
                                oMAtrix_IN.Columns.Item("TimeIN").Cells.Item(1).Specific.String = oform1.Items.Item("Item_193").Specific.String
                                oMAtrix_IN.Columns.Item("OUTDate").Cells.Item(1).Specific.String = oform1.Items.Item("Item_175").Specific.String
                                oMAtrix_IN.Columns.Item("TimeOUT").Cells.Item(1).Specific.String = oform1.Items.Item("Item_177").Specific.String

                                irowcount = oMAtrix_IN.RowCount

                                If oform1.Items.Item("228").Specific.value.ToString.Trim = "Yes" And CInt(oform1.Items.Item("Item_123").Specific.string) > 0 Then
                                    oMAtrix_IN.Columns.Item("1").Cells.Item(irowcount).Specific.String = "Collission Damage Waiver Fee"
                                    oMAtrix_IN.Columns.Item("12").Cells.Item(irowcount).Specific.String = iCDW
                                    oMAtrix_IN.Columns.Item("2").Cells.Item(irowcount).Specific.String = CDW_GLACC
                                End If

                                irowcount = oMAtrix_IN.RowCount

                                If oform1.Items.Item("226").Specific.value.ToString.Trim = "Yes" And CInt(oform1.Items.Item("Item_115").Specific.string) > 0 Then
                                    oMAtrix_IN.Columns.Item("1").Cells.Item(irowcount).Specific.String = "Personal Accidental Insurance"
                                    oMAtrix_IN.Columns.Item("12").Cells.Item(irowcount).Specific.String = iPAI
                                    oMAtrix_IN.Columns.Item("2").Cells.Item(irowcount).Specific.String = PAI_GLACC
                                End If

                                irowcount = oMAtrix_IN.RowCount

                                If CInt(dOC) > 0 Then
                                    oMAtrix_IN.Columns.Item("1").Cells.Item(irowcount).Specific.String = "Other Charges"
                                    oMAtrix_IN.Columns.Item("12").Cells.Item(irowcount).Specific.String = dOC
                                    oMAtrix_IN.Columns.Item("2").Cells.Item(irowcount).Specific.String = sOC_GLACC
                                End If

                                irowcount = oMAtrix_IN.RowCount

                                If CInt(dMR) > 0 Then
                                    oMAtrix_IN.Columns.Item("1").Cells.Item(irowcount).Specific.String = "Monthly Recurring Other Charges"
                                    oMAtrix_IN.Columns.Item("12").Cells.Item(irowcount).Specific.String = dMR
                                    oMAtrix_IN.Columns.Item("2").Cells.Item(irowcount).Specific.String = sMR_GLACC
                                End If

                                irowcount = oMAtrix_IN.RowCount

                                If CInt(dPC) > 0 Then
                                    oMAtrix_IN.Columns.Item("1").Cells.Item(irowcount).Specific.String = "Petrol Charges"
                                    oMAtrix_IN.Columns.Item("12").Cells.Item(irowcount).Specific.String = dPC
                                    oMAtrix_IN.Columns.Item("2").Cells.Item(irowcount).Specific.String = sPet_GLACC
                                End If


                                Dim ocombo As SAPbouiCOM.ComboBox
                                If ContactP <> "" Then
                                    ocombo = oform.Items.Item("85").Specific
                                    ocombo.Select(ContactP)
                                End If

                                ' oMAtrix_IN.Columns.Item("95").Cells.Item(1).Specific.String = "SO"
                                'Dim ocombo As SAPbouiCOM.ComboBox = oform.Items.Item("20").Specific
                                'ocombo.Select(salesemp)
                                oform.Items.Item("16").Specific.String = sDesc  ''"Self Drive Invoice Based on Booking No : " & oform1.Items.Item("Item_14").Specific.String
                                oform.Visible = True
                                oform.Freeze(False)
                            End If

                            Oapplication_SD.StatusBar.SetText("Opening the invoice for the booking no. " & oform1.Items.Item("Item_14").Specific.String & " ........... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        Catch ex As Exception
                            oform.Freeze(False)
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Try
                        End Try
                        Exit Sub
                    End If

                End If

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If pVal.Action_Success = True Then
                            Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm

                            Try
                                oform.Freeze(True)
                                Dim Tmp_val As String = oform.BusinessObject.GetNextSerialNumber(NextSerialNo(Ocompany_SD, Oapplication_SD, "AE_Sbooking"))
                                oform.Items.Item("Item_14").Specific.String = Tmp_val

                                oform.Items.Item("Item_20").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                oform.Items.Item("Item_16").Specific.String = Format(Now.Date, "yyyyMMdd") 'Now.Date 'Format(Now.Date, "dd MMM yyyy") 'Now.Date
                                oform.Items.Item("Item_14").Enabled = False
                                oform.Visible = True
                                Dim ocombo As SAPbouiCOM.ComboBox = oform.Items.Item("Item_18").Specific
                                ocombo.Select("Open")

                                ocombo = oform.Items.Item("226").Specific
                                ocombo.Select("No")
                                ocombo = oform.Items.Item("228").Specific
                                ocombo.Select("No")

                                oform.Items.Item("217").Specific.String = Ocompany_SD.UserName
                                oform.Items.Item("218").Specific.String = Company_Name

                                Dim opt As SAPbouiCOM.OptionBtn = oform.Items.Item("195").Specific
                                Dim opt1 As SAPbouiCOM.OptionBtn = oform.Items.Item("196").Specific
                                Dim opt2 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_104").Specific
                                Dim opt3 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_105").Specific
                                Dim opt4 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_106").Specific

                                opt1.GroupWith("195")
                                opt1.Selected = True
                                oform.PaneLevel = 3
                                opt3.GroupWith("Item_104")
                                opt4.GroupWith("Item_104")
                                opt3.Selected = True
                                opt2.Selected = True

                                oform.PaneLevel = 1
                                oform.DataBrowser.BrowseBy = "Item_14"
                                oform.Freeze(False)
                            Catch ex As Exception
                                oform.Freeze(False)
                                Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try

                            End Try


                        End If
                    End If


                    If pVal.ItemUID = "Item_104" Or pVal.ItemUID = "Item_105" Or pVal.ItemUID = "Item_106" Then
                        Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                        ' Dim oStatic As sa
                        Select Case pVal.ItemUID
                            Case "Item_104"
                                oform.Items.Item("Item_108").Specific.caption = "Daily Rates"
                                oform.Items.Item("Item_110").Specific.caption = "Number of Days"
                                oform.Items.Item("Item_114").Specific.caption = "PAI Per Day"
                                oform.Items.Item("Item_122").Specific.caption = "CDW Per Day"
                            Case "Item_105"
                                oform.Items.Item("Item_108").Specific.caption = "Weekly Rates"
                                oform.Items.Item("Item_110").Specific.caption = "Number of Day"
                                oform.Items.Item("Item_114").Specific.caption = "PAI Per Day"
                                oform.Items.Item("Item_122").Specific.caption = "CDW Per Day"
                            Case "Item_106"
                                oform.Items.Item("Item_108").Specific.caption = "Monthly Rates"
                                oform.Items.Item("Item_110").Specific.caption = "Number of Months"
                                oform.Items.Item("Item_114").Specific.caption = "PAI Per Month"
                                oform.Items.Item("Item_122").Specific.caption = "CDW Per Month"
                        End Select
                    End If

                    If pVal.ItemUID = "195" Or pVal.ItemUID = "196" Then
                        Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                        ' Dim oStatic As sa
                        Dim oopt As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_104").Specific
                        Dim oopt1 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_105").Specific
                        Dim oopt2 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_106").Specific
                        Select Case pVal.ItemUID
                            Case "195"
                                oform.Items.Item("203").Specific.String = ""
                                oform.Items.Item("205").Specific.String = ""
                                oform.Items.Item("Item_106").Enabled = True
                                oopt2.Selected = True
                                oform.Items.Item("Item_104").Enabled = False
                                oform.Items.Item("Item_105").Enabled = False
                                oform.PaneLevel = 3
                            Case "196"
                                oform.Items.Item("203").Specific.String = ""
                                oform.Items.Item("205").Specific.String = ""
                                oform.Items.Item("Item_104").Enabled = True
                                oform.Items.Item("Item_105").Enabled = True
                                oopt.Selected = True
                                oform.Items.Item("Item_106").Enabled = False
                                oform.PaneLevel = 3
                        End Select
                    End If

                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.Item("SDB")
                            If p_bSDBooking = False Then
                                NavigationValidation_SelfDriver(oform, Ocompany_SD, Oapplication_SD)
                            End If


                            ''Dim opt As SAPbouiCOM.OptionBtn = oform.Items.Item("195").Specific ' Long Term
                            ''Dim opt1 As SAPbouiCOM.OptionBtn = oform.Items.Item("196").Specific

                            ''If oform.Items.Item("226").Specific.value.ToString.Trim = "Yes" Then
                            ''    oform.Items.Item("Item_115").Enabled = True
                            ''Else
                            ''    oform.Items.Item("Item_115").Enabled = True
                            ''End If
                            ''oform.Items.Item("201").Enabled = True
                            ''If opt.Selected = True Then
                            ''    'oform.Items.Item("201").Enabled = False
                            ''Else
                            ''    If oform.Items.Item("Item_18").Specific.value.ToString.Trim = "Billing" Then
                            ''        'oform.Items.Item("201").Enabled = True
                            ''        oform.Items.Item("Item_18").Enabled = True
                            ''    ElseIf oform.Items.Item("Item_18").Specific.value.ToString.Trim = "Closed" Then
                            ''        ' oform.Items.Item("201").Enabled = False
                            ''        oform.Items.Item("Item_18").Enabled = False
                            ''    Else
                            ''        ' oform.Items.Item("201").Enabled = False
                            ''        oform.Items.Item("Item_18").Enabled = True
                            ''    End If


                            ''    oform.Items.Item("Item_14").Enabled = False
                            ''End If
                        Catch ex As Exception
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Try
                        End Try

                    End If



                    If pVal.ItemUID = "1000008" Then
                        Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                        Dim oCheck As SAPbouiCOM.CheckBox = oform.Items.Item("1000008").Specific
                        Dim oOPtion As SAPbouiCOM.OptionBtn = oform.Items.Item("195").Specific

                        Try
                            If String.IsNullOrEmpty(oform.Items.Item("209").Specific.string) Then
                                oCheck.Checked = False
                                Oapplication_SD.SetStatusBarMessage("Next Invoice date is empty, can`t avail this functionality  ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                BubbleEvent = False
                                Exit Try
                            End If

                            If oOPtion.Selected = False Then
                                oCheck.Checked = False
                                Oapplication_SD.SetStatusBarMessage("Can`t avail this functionality in short term ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                BubbleEvent = False
                                Exit Try
                            End If
                            If Trim(oform.Items.Item("Item_18").Specific.value) <> "Open" Then
                                oCheck.Checked = False
                                Oapplication_SD.SetStatusBarMessage("Status should be Open ", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                BubbleEvent = False
                                Exit Try
                            End If
                            If oCheck.Checked = True Then
                                oform.Items.Item("1000011").Specific.String = Format(Now.Date, "yyyyMMdd")
                                oform.Items.Item("237").Enabled = True
                            Else
                                oform.Items.Item("237").Specific.select("")
                                oform.Items.Item("1000011").Specific.String = String.Empty
                                oform.Items.Item("239").Specific.String = String.Empty
                                oform.Items.Item("241").Specific.String = String.Empty
                                oform.Items.Item("Item_2").Specific.active = True
                                oform.Items.Item("237").Enabled = False
                                oform.Items.Item("241").Enabled = False
                                oform.Items.Item("242").Enabled = False
                            End If
                        Catch ex As Exception
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Try
                        End Try
                    End If
                    Exit Sub
                End If
            End If
        End If

        If pVal.FormUID = "SDBR" Then
            If pVal.Before_Action = False Then

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvento = pVal
                    Dim sCFL_ID As String
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    Dim oForm As SAPbouiCOM.Form = Oapplication_SD.Forms.Item(FormUID)
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    oCFL = oForm.ChooseFromLists.Item(sCFL_ID)

                    If oCFLEvento.BeforeAction = False Then
                        Dim oDataTable As SAPbouiCOM.DataTable
                        Dim oedit1 As SAPbouiCOM.EditText
                        oDataTable = oCFLEvento.SelectedObjects

                        If pVal.ItemUID = "11" Then
                            Try
                                oForm.Items.Item("3").Specific.string = oDataTable.GetValue("CardName", 0)
                                oForm.Items.Item("11").Specific.string = oDataTable.GetValue("CardCode", 0)
                            Catch ex As Exception

                            End Try

                        End If
                    End If
                End If

            End If
            If pVal.Before_Action = True Then

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED Then
                    If pVal.ItemUID = "8" And pVal.ColUID = "Booking No" Then
                        Try
                            Dim oform_SDR As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                            Dim oform As SAPbouiCOM.Form
                            Dim ogrid As SAPbouiCOM.Grid = oform_SDR.Items.Item("8").Specific
                            Dim docentry As String = ogrid.DataTable.GetValue("Booking No", pVal.Row)
                            Dim orset As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim ocombo As SAPbouiCOM.ComboBox
                            Dim sAttention As String

                            LoadFromXML("SelfDriving_Booking.srf", Oapplication_SD)
                            oform = Oapplication_SD.Forms.Item("SDB")
                            oform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oform.Items.Item("Item_14").Enabled = True
                            oform.Items.Item("Item_14").Specific.String = docentry
                            oform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Dim ocombobutton As SAPbouiCOM.ButtonCombo = oform.Items.Item("210").Specific
                            oform.Items.Item("Item_5").Enabled = True
                            ocombobutton.ValidValues.Add("Copy To", "Copy To A/R Invoice")

                            orset.DoQuery("SELECT T0.[FirstName] + ' ' + T0.[LastName] as 'Name' FROM OCPR T0 WHERE T0.[CardCode]  = '" & oform.Items.Item("Item_2").Specific.string & "'")
                            ' " & _                                            "and T0.[FirstName] + ' ' + T0.[LastName] <> '" & oform.Items.Item("235").Specific.value.ToString.Trim & "'")
                            sAttention = oform.Items.Item("235").Specific.value.ToString.Trim

                            ocombo = oform.Items.Item("235").Specific

                            For mjs As Integer = ocombo.ValidValues.Count - 1 To 1 Step -1
                                ocombo.ValidValues.Remove(mjs, SAPbouiCOM.BoSearchKey.psk_Index)
                                orset.MoveNext()
                            Next mjs

                            Try
                                For mjs As Integer = 1 To orset.RecordCount
                                    ocombo.ValidValues.Add(orset.Fields.Item("Name").Value, "")
                                    orset.MoveNext()
                                Next mjs
                            Catch ex As Exception
                            End Try
                            ocombo.Select(sAttention, SAPbouiCOM.BoSearchKey.psk_ByValue)

                            '------------------- PAI 
                            If oform.Items.Item("226").Specific.value.ToString.Trim = "Yes" Then
                                oform.Items.Item("Item_115").Enabled = True
                            Else
                                oform.Items.Item("Item_115").Enabled = False
                            End If

                            '-------------------- CDW
                            If oform.Items.Item("228").Specific.selected.value.ToString.Trim = "Yes" Then
                                oform.Items.Item("Item_123").Enabled = True
                            Else
                                oform.Items.Item("Item_123").Enabled = False
                            End If

                            Dim ooption As SAPbouiCOM.OptionBtn = oform.Items.Item("195").Specific
                            Dim ooption1 As SAPbouiCOM.OptionBtn = oform.Items.Item("196").Specific
                            Dim ooption2 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_104").Specific
                            Dim ooption3 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_105").Specific
                            Dim ooption4 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_106").Specific

                            ooption1.GroupWith("195")
                            ooption3.GroupWith("Item_104")
                            ooption4.GroupWith("Item_104")

                            oform.Items.Item("Item_14").Enabled = True
                            If ooption2.Selected = True Then
                                oform.Items.Item("Item_108").Specific.caption = "Daily Rates"
                                oform.Items.Item("Item_110").Specific.caption = "Number of Days"
                                oform.Items.Item("Item_114").Specific.caption = "PAI Per Day"
                                oform.Items.Item("Item_122").Specific.caption = "CDW Per Day"
                            ElseIf ooption3.Selected = True Then
                                oform.Items.Item("Item_108").Specific.caption = "Weekly Rates"
                                oform.Items.Item("Item_110").Specific.caption = "Number of Days"
                                oform.Items.Item("Item_114").Specific.caption = "PAI Per Day"
                                oform.Items.Item("Item_122").Specific.caption = "CDW Per Day"
                            ElseIf ooption4.Selected = True Then
                                oform.Items.Item("Item_108").Specific.caption = "Monthly Rates"
                                oform.Items.Item("Item_110").Specific.caption = "Number of Months"
                                oform.Items.Item("Item_114").Specific.caption = "PAI Per Month"
                                oform.Items.Item("Item_122").Specific.caption = "CDW Per Month"
                            End If

                            '-----------------  Copy To Uneditable

                            If oform.Items.Item("Item_18").Specific.value.ToString.Trim <> "Billing" Then 'Or oform.Items.Item("Item_18").Specific.value.ToString.Trim = "Cancel" Then
                                oform.Items.Item("210").Enabled = False
                            Else
                                Dim opt As SAPbouiCOM.OptionBtn = oform.Items.Item("195").Specific
                                If opt.Selected = True Then
                                    oform.Items.Item("210").Enabled = False
                                Else
                                    oform.Items.Item("210").Enabled = True
                                End If
                            End If

                            oform.Items.Item("Item_22").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oform.Visible = True

                        Catch ex As Exception
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
                            BubbleEvent = False
                            Exit Try
                        End Try
                        Exit Sub
                    End If
                End If


                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.ItemUID = "9" Then
                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                            Dim customer, Vmodel, Sqlstring As String

                            If oform.Items.Item("11").Specific.String <> "" Then
                                customer = oform.Items.Item("11").Specific.String
                            Else
                                customer = "%"
                                oform.Items.Item("3").Specific.String = ""
                            End If

                            If oform.Items.Item("16").Specific.String <> "" Then
                                Vmodel = oform.Items.Item("16").Specific.String
                            Else
                                Vmodel = "%"
                            End If


                            If oform.Items.Item("1000001").Specific.String <> "" And oform.Items.Item("10").Specific.String <> "" And oform.Items.Item("12").Specific.String = "" And oform.Items.Item("14").Specific.String = "" Then
                                Sqlstring = "SELECT T0.[DocNum] as 'Booking No', T0.[U_AE_Vregno] as 'Vehicle Number', T0.[U_AE_Vmodel] as 'Vehicle Model', " & _
"T0.[U_AE_Bname] as 'Customer', T0.[U_AE_DName] as 'Driver', T0.[U_AE_Sdate] as 'Start Date', " & _
"T0.[U_AE_Edate] as 'End Date' FROM [dbo].[@AE_SBOOKING]  T0 where T0.U_AE_Sdate >= '" & GateDate(oform.Items.Item("1000001").Specific.String, Ocompany_SD) & "' and T0.U_AE_Sdate <= '" & GateDate(oform.Items.Item("10").Specific.String, Ocompany_SD) & "' " & _
"and T0.U_AE_Bcode like '" & customer & "' and isnull(T0.U_AE_Vmodel,'') like '" & Vmodel & "' "
                            ElseIf oform.Items.Item("1000001").Specific.String = "" And oform.Items.Item("10").Specific.String = "" And oform.Items.Item("12").Specific.String <> "" And oform.Items.Item("14").Specific.String <> "" Then

                                Sqlstring = "SELECT T0.[DocNum] as 'Booking No', T0.[U_AE_Vregno] as 'Vehicle Number', T0.[U_AE_Vmodel] as 'Vehicle Model', " & _
    "T0.[U_AE_Bname] as 'Customer', T0.[U_AE_DName] as 'Driver', T0.[U_AE_Sdate] as 'Start Date', " & _
    "T0.[U_AE_Edate] as 'End Date' FROM [dbo].[@AE_SBOOKING]  T0 where T0.U_AE_Edate >= '" & GateDate(oform.Items.Item("12").Specific.String, Ocompany_SD) & "' and T0.U_AE_Edate <= '" & GateDate(oform.Items.Item("14").Specific.String, Ocompany_SD) & "' " & _
    "and T0.U_AE_Bcode like '" & customer & "' and isnull(T0.U_AE_Vmodel,'') like '" & Vmodel & "' "

                            ElseIf oform.Items.Item("1000001").Specific.String <> "" And oform.Items.Item("10").Specific.String <> "" And oform.Items.Item("12").Specific.String <> "" And oform.Items.Item("14").Specific.String <> "" Then

                                Sqlstring = "SELECT T0.[DocNum] as 'Booking No', T0.[U_AE_Vregno] as 'Vehicle Number', T0.[U_AE_Vmodel] as 'Vehicle Model', " & _
    "T0.[U_AE_Bname] as 'Customer', T0.[U_AE_DName] as 'Driver', T0.[U_AE_Sdate] as 'Start Date', " & _
    "T0.[U_AE_Edate] as 'End Date' FROM [dbo].[@AE_SBOOKING]  T0 where T0.U_AE_Sdate >= '" & GateDate(oform.Items.Item("1000001").Specific.String, Ocompany_SD) & "' and T0.U_AE_Sdate <= '" & GateDate(oform.Items.Item("10").Specific.String, Ocompany_SD) & "' and " & _
" T0.U_AE_Edate >= '" & GateDate(oform.Items.Item("12").Specific.String, Ocompany_SD) & "' and T0.U_AE_Edate <= '" & GateDate(oform.Items.Item("14").Specific.String, Ocompany_SD) & "' " & _
    "and T0.U_AE_Bcode like '" & customer & "' and isnull(T0.U_AE_Vmodel,'') like '" & Vmodel & "' "

                            ElseIf oform.Items.Item("1000001").Specific.String = "" And oform.Items.Item("10").Specific.String = "" And oform.Items.Item("12").Specific.String = "" And oform.Items.Item("14").Specific.String = "" Then

                                Sqlstring = "SELECT T0.[DocNum] as 'Booking No', T0.[U_AE_Vregno] as 'Vehicle Number', T0.[U_AE_Vmodel] as 'Vehicle Model', " & _
    "T0.[U_AE_Bname] as 'Customer', T0.[U_AE_DName] as 'Driver', T0.[U_AE_Sdate] as 'Start Date', " & _
    "T0.[U_AE_Edate] as 'End Date' FROM [dbo].[@AE_SBOOKING]  T0 where  " & _
    "T0.U_AE_Bcode like '" & customer & "' and isnull(T0.U_AE_Vmodel,'') like '" & Vmodel & "' "
                            End If


                            Try
                                oform.DataSources.DataTables.Add("SDBR")
                            Catch ex As Exception

                            End Try
                            Dim ogrid As SAPbouiCOM.Grid = oform.Items.Item("8").Specific

                            oform.DataSources.DataTables.Item(0).ExecuteQuery(Sqlstring)
                            ogrid.DataTable = oform.DataSources.DataTables.Item("SDBR")
                            ogrid.AutoResizeColumns()

                            Dim ocol As SAPbouiCOM.EditTextColumn = ogrid.Columns.Item("Booking No")
                            ocol.LinkedObjectType = "AE_Sbooking"

                            ogrid.Columns.Item("Booking No").ForeColor = RGB(20, 20, 200)
                            ogrid.Columns.Item("Vehicle Number").ForeColor = RGB(20, 20, 200)
                            ogrid.Columns.Item("Booking No").TextStyle = 1
                            ogrid.Columns.Item("Vehicle Number").TextStyle = 1
                            ogrid.Columns.Item("Vehicle Model").TextStyle = 1
                            ogrid.Columns.Item("Customer").TextStyle = 1
                            ogrid.Columns.Item("Driver").TextStyle = 1
                            ogrid.Columns.Item("Start Date").TextStyle = 1
                            ogrid.Columns.Item("End Date").TextStyle = 1
                            oform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
                            oform.Visible = True

                        Catch ex As Exception
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Try
                        End Try

                    End If
                End If
            End If
        End If


        If pVal.FormUID = "SDINV" Then
            If pVal.Before_Action = False Then
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvento = pVal
                    Dim sCFL_ID As String
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    Dim oForm As SAPbouiCOM.Form = Oapplication_SD.Forms.Item(FormUID)
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                    Try
                        If oCFLEvento.BeforeAction = False Then
                            Dim oDataTable As SAPbouiCOM.DataTable
                            Dim oedit1 As SAPbouiCOM.EditText
                            Dim oCombo As SAPbouiCOM.ComboBox
                            oDataTable = oCFLEvento.SelectedObjects
                            If pVal.ItemUID = "15" Then 'Billing Code
                                oForm.Items.Item("16").Specific.string = oDataTable.GetValue("CardName", 0)
                                Dim ocheck As SAPbouiCOM.CheckBox = oForm.Items.Item("13").Specific
                                ocheck.Checked = False
                                oForm.Items.Item("15").Specific.string = oDataTable.GetValue("CardCode", 0)
                            End If
                        End If
                    Catch ex As Exception

                    End Try

                End If

                If pVal.ItemUID = "13" Then
                    Dim oform As SAPbouiCOM.Form
                    Try
                        oform = Oapplication_SD.Forms.ActiveForm
                        oform.Freeze(True)
                        Dim ocheck As SAPbouiCOM.CheckBox = oform.Items.Item("13").Specific
                        Dim ocheck1 As SAPbouiCOM.CheckBox
                        Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_6").Specific
                        Dim Flag As Boolean = False

                        If ocheck.Checked = True Then
                            Flag = True
                        Else
                            Flag = False
                        End If

                        For mjs As Integer = 1 To omatrix.RowCount
                            ocheck1 = omatrix.Columns.Item("Col_5m").Cells.Item(mjs).Specific
                            ocheck1.Checked = Flag
                        Next mjs
                        oform.Freeze(False)
                    Catch ex As Exception
                        oform.Freeze(False)
                        Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        BubbleEvent = False
                        Exit Try
                    End Try
                    Exit Sub
                End If

            ElseIf pVal.Before_Action = True Then

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.ItemUID = "Item_5" Then
                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                            Dim orset As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim Customer As String
                            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_6").Specific
                            Dim ocheck As SAPbouiCOM.CheckBox = oform.Items.Item("13").Specific
                            ocheck.Checked = False
                            If oform.Items.Item("Item_1").Specific.String = "" Then
                                oform.Items.Item("Item_1").Specific.active = True
                                Oapplication_SD.StatusBar.SetText("Day from should not be empty ........ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End If

                            If oform.Items.Item("Item_3").Specific.String = "" Then
                                oform.Items.Item("Item_3").Specific.active = True
                                Oapplication_SD.StatusBar.SetText("Day to should not be empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End If

                            If oform.Items.Item("15").Specific.String = "" Then
                                Customer = "%"
                                oform.Items.Item("16").Specific.String = ""
                            Else
                                Customer = oform.Items.Item("15").Specific.String
                            End If

                            Dim Str_query As String = "SELECT T0.[DocNum], T0.[DocEntry], substring(T0.[NumAtCard],3,len(T0.[NumAtCard]) -2) as 'NumAtCard', T0.[CardName], T0.[DocTotal], T0.[DocDate] " & _
                                "FROM OINV T0 WHERE day(T0.[DocDate]) >= '" & CInt(oform.Items.Item("Item_1").Specific.String) & "' and  day(T0.[DocDate]) <= '" & CInt(oform.Items.Item("Item_3").Specific.String) & "' and month( T0.[DocDate] ) = '" & oform.Items.Item("12").Specific.selected.description & "' and " & _
                                "year( T0.[DocDate] )  = '" & Now.Year & "'  and  T0.[CardCode]  like '" & Customer & "' and left(T0.NumAtCard,2) = 'SD'"


                            Try
                                oform.DataSources.DataTables.Add("@AE_SBOOKING")
                            Catch ex As Exception

                            End Try

                            oform.DataSources.DataTables.Item("@AE_SBOOKING").ExecuteQuery(Str_query)

                            oMatrix.Clear()
                            oform.Items.Item("Item_6").Specific.columns.item("V_0mjd").databind.bind("@AE_SBOOKING", "DocEntry")
                            oform.Items.Item("Item_6").Specific.columns.item("BNO").databind.bind("@AE_SBOOKING", "NumAtCard")
                            oform.Items.Item("Item_6").Specific.columns.item("Col_0m").databind.bind("@AE_SBOOKING", "CardName")
                            oform.Items.Item("Item_6").Specific.columns.item("Col_4mjs").databind.bind("@AE_SBOOKING", "DocTotal")
                            oform.Items.Item("Item_6").Specific.columns.item("V_0mj").databind.bind("@AE_SBOOKING", "DocDate")
                            oform.Items.Item("Item_6").Specific.columns.item("V_1INV").databind.bind("@AE_SBOOKING", "DocNum")

                            oform.Items.Item("Item_6").Specific.LoadFromDataSource()
                            oform.Items.Item("Item_6").Specific.AutoResizeColumns()
                            ' oMatrix.Columns.Item("Col_5").Editable = True


                        Catch ex As Exception
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Try
                        End Try
                        Exit Sub
                    End If

                    If pVal.ItemUID = "Item_7" Then
                        Dim oform1 As SAPbouiCOM.Form
                        Try
                            oform1 = Oapplication_SD.Forms.ActiveForm

                            Dim oMAtrix As SAPbouiCOM.Matrix = oform1.Items.Item("Item_6").Specific
                            Dim oCheck1 As SAPbouiCOM.CheckBox = oform1.Items.Item("17").Specific
                            Dim orset As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            orset.DoQuery("SELECT top(1) isnull(T0.[U_printer],'') [Printer] FROM [dbo].[@AE_CRYSTAL]  T0")
                            ' True - Group Invoice , False - Tax Invoice
                            If oCheck1.Checked = True Then
                                Group_Invoice(oform1, oMAtrix)
                            Else
                                TaxInvoice(oform1, oMAtrix, orset.Fields.Item("Printer").Value)
                            End If

                        Catch ex As Exception
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Try
                        End Try
                        Exit Sub
                    End If
                End If

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED And pVal.ItemUID = "Item_6" And pVal.ColUID = "BNO" Then
                    Try
                        Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                        Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_6").Specific
                        Dim docentry As String = omatrix.Columns.Item("BNO").Cells.Item(pVal.Row).Specific.String


                        LoadFromXML("SelfDriving_Booking.srf", Oapplication_SD)
                        oform = Oapplication_SD.Forms.Item("SDB")
                        oform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        oform.Items.Item("Item_14").Enabled = True
                        oform.Items.Item("Item_14").Specific.String = docentry
                        oform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Dim ocombobutton As SAPbouiCOM.ButtonCombo = oform.Items.Item("210").Specific
                        ocombobutton.ValidValues.Add("Copy To", "Copy To A/R Invoice")
                        oform.Items.Item("Item_5").Enabled = True
                        Dim ooption As SAPbouiCOM.OptionBtn = oform.Items.Item("195").Specific
                        Dim ooption1 As SAPbouiCOM.OptionBtn = oform.Items.Item("196").Specific
                        ooption1.GroupWith("195")
                        Dim ooption2 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_104").Specific
                        Dim ooption3 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_105").Specific
                        Dim ooption4 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_106").Specific

                        ooption1.GroupWith("195")

                        ooption3.GroupWith("Item_104")
                        ooption4.GroupWith("Item_104")
                        ' oform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                        oform.Items.Item("Item_14").Enabled = True

                        If ooption2.Selected = True Then
                            oform.Items.Item("Item_108").Specific.caption = "Daily Rates"
                            oform.Items.Item("Item_110").Specific.caption = "Number of Days"
                            oform.Items.Item("Item_114").Specific.caption = "PAI Per Day"
                            oform.Items.Item("Item_122").Specific.caption = "CDW Per Day"
                        ElseIf ooption3.Selected = True Then
                            oform.Items.Item("Item_108").Specific.caption = "Weekly Rates"
                            oform.Items.Item("Item_110").Specific.caption = "Number of Days"
                            oform.Items.Item("Item_114").Specific.caption = "PAI Per Month"
                            oform.Items.Item("Item_122").Specific.caption = "CDW Per Month"
                        ElseIf ooption4.Selected = True Then
                            oform.Items.Item("Item_108").Specific.caption = "Monthly Rates"
                            oform.Items.Item("Item_110").Specific.caption = "Number of Months"
                            oform.Items.Item("Item_114").Specific.caption = "PAI Per Month"
                            oform.Items.Item("Item_122").Specific.caption = "CDW Per Month"
                        End If
                        oform.Items.Item("Item_22").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oform.Visible = True

                    Catch ex As Exception
                        Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        BubbleEvent = False
                        Exit Try
                    End Try
                End If

            End If
        End If

        If pVal.FormUID = "SDBIL" Then
            If pVal.Before_Action = True Then

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED And pVal.ItemUID = "Item_6" And pVal.ColUID = "Col_4mjs" Then
                    Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                    Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_6").Specific
                    Dim sCalDesc As String = String.Empty
                    sCalDesc = omatrix.Columns.Item("Col_Cal").Cells.Item(pVal.Row).Specific.String
                    LoadFromXML("LongtermInvoiceCalculation.srf", Oapplication_SD)
                    oform = Oapplication_SD.Forms.Item("RentCal")
                    oform.Items.Item("1000002").Specific.String = sCalDesc
                End If

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED And pVal.ItemUID = "Item_6" And pVal.ColUID = "BNO" Then
                    Try
                        Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                        Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_6").Specific
                        Dim docentry As String = omatrix.Columns.Item("BNO").Cells.Item(pVal.Row).Specific.String


                        LoadFromXML("SelfDriving_Booking.srf", Oapplication_SD)
                        oform = Oapplication_SD.Forms.Item("SDB")
                        oform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        oform.Items.Item("Item_14").Enabled = True
                        oform.Items.Item("Item_14").Specific.String = docentry
                        oform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Dim ocombobutton As SAPbouiCOM.ButtonCombo = oform.Items.Item("210").Specific
                        ocombobutton.ValidValues.Add("Copy To", "Copy To A/R Invoice")
                        oform.Items.Item("Item_5").Enabled = True
                        Dim ooption As SAPbouiCOM.OptionBtn = oform.Items.Item("195").Specific
                        Dim ooption1 As SAPbouiCOM.OptionBtn = oform.Items.Item("196").Specific
                        ooption1.GroupWith("195")
                        Dim ooption2 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_104").Specific
                        Dim ooption3 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_105").Specific
                        Dim ooption4 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_106").Specific

                        ooption1.GroupWith("195")

                        ooption3.GroupWith("Item_104")
                        ooption4.GroupWith("Item_104")
                        ' '' oform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                        ''oform.Items.Item("Item_14").Enabled = True

                        ''If ooption2.Selected = True Then
                        ''    oform.Items.Item("Item_108").Specific.caption = "Daily Rates"
                        ''    oform.Items.Item("Item_110").Specific.caption = "Number of Days"
                        ''    oform.Items.Item("Item_114").Specific.caption = "PAI Per Day"
                        ''    oform.Items.Item("Item_122").Specific.caption = "CDW Per Day"
                        ''ElseIf ooption3.Selected = True Then
                        ''    oform.Items.Item("Item_108").Specific.caption = "Weekly Rates"
                        ''    oform.Items.Item("Item_110").Specific.caption = "Number of Days"
                        ''    oform.Items.Item("Item_114").Specific.caption = "PAI Per Month"
                        ''    oform.Items.Item("Item_122").Specific.caption = "CDW Per Month"
                        ''ElseIf ooption4.Selected = True Then
                        ''    oform.Items.Item("Item_108").Specific.caption = "Monthly Rates"
                        ''    oform.Items.Item("Item_110").Specific.caption = "Number of Months"
                        ''    oform.Items.Item("Item_114").Specific.caption = "PAI Per Month"
                        ''    oform.Items.Item("Item_122").Specific.caption = "CDW Per Month"
                        ''End If
                        NavigationValidation_SelfDriver(oform, Ocompany_SD, Oapplication_SD)
                        oform.Items.Item("Item_22").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oform.Visible = True

                    Catch ex As Exception
                        Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        BubbleEvent = False
                        Exit Try
                    End Try
                End If
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.ItemUID = "Item_5" Then
                        Dim oform As SAPbouiCOM.Form
                        Try
                            oform = Oapplication_SD.Forms.ActiveForm
                            Dim orset As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_6").Specific
                            Dim oCheck As SAPbouiCOM.CheckBox = oform.Items.Item("13").Specific
                            If oform.Items.Item("Item_1").Specific.String = "" Then
                                oform.Items.Item("Item_1").Specific.active = True
                                Oapplication_SD.StatusBar.SetText("Day from should not be empty ........ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End If

                            If oform.Items.Item("Item_3").Specific.String = "" Then
                                oform.Items.Item("Item_3").Specific.active = True
                                Oapplication_SD.StatusBar.SetText("Day to should not be empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End If

                            If oform.Items.Item("Item_8").Specific.String = "" Then
                                oform.Items.Item("Item_8").Specific.active = True
                                Oapplication_SD.StatusBar.SetText("Posting Date should not be empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End If

                            Dim Str_query As String = ""

                            ''oCheck.Checked = False

                            oform.Freeze(True)

                            ''                            Str_query = "SELECT T0.[DocNum], T0.[U_AE_Bname], T0.[U_AE_Bcode], T0.[U_AE_Vregno], T0.[U_AE_Rno], T0.[U_AE_Contract], " & _
                            ''"case when month(T0.[U_AE_Vdatein]) = '" & oform.Items.Item("12").Specific.selected.description & "' " & _
                            ''" then ( case when T0.U_AE_Sdate = T0.[U_AE_NXDT] then (T0.[U_AE_Rate]/DAY(EOMONTH(DATEFROMPARTS(year(T0.[U_AE_Vdatein]),Month(T0.[U_AE_Vdatein]),1))) )*(DATEDIFF(day,U_AE_Sdate,U_AE_Vdatein)+1) " & _
                            '' " else T0.[U_AE_rate] end ) else T0.[U_AE_rate] end  as 'U_AE_rate', " & _
                            '' "case when month(T0.[U_AE_Vdatein]) in ( '" & oform.Items.Item("12").Specific.selected.description & "','" & oform.Items.Item("12").Specific.selected.description + 1 & "') " & _
                            '' "then ( case when T0.U_AE_Sdate = T0.[U_AE_NXDT] " & _
                            '' "then ('Long Term Return Calculation ' + CHAR(13) + '----------------------------------------' + CHAR(13) + CHAR(13) + 'Monthly Rental          : ' + cast(cast(T0.[U_AE_Rate] as decimal(19,4)) as nvarchar ) + char(13) + 'No. of Days in month : ' +  " & _
                            '' "cast(DAY(EOMONTH(DATEFROMPARTS(year(T0.[U_AE_Vdatein]),Month(T0.[U_AE_Vdatein]),1))) as nvarchar) + char(13) +  " & _
                            '' "'Consumed Days         : ' + cast((DATEDIFF(day,U_AE_Sdate,U_AE_Vdatein)+1) as nvarchar) + char(13) + 'Calculated Amount     : ' + " & _
                            '' "cast(cast((T0.[U_AE_Rate]/DAY(EOMONTH(DATEFROMPARTS(year(T0.[U_AE_Vdatein]),Month(T0.[U_AE_Vdatein]),1))) )*(DATEDIFF(day,U_AE_Sdate,U_AE_Vdatein)+1) as decimal(19,4)) as varchar)) " & _
                            '' "else 'Long Term Recurring Calculation' + CHAR(13) + '----------------------------------------' + CHAR(13) + CHAR(13) + 'Monthly Rental         : ' + cast(cast(T0.[U_AE_rate] as decimal(19,4)) as nvarchar) end ) else " & _
                            '' "'Long Term Recurring Calculation' + CHAR(13) + '----------------------------------------' + CHAR(13) + CHAR(13) + 'Monthly Rental         : ' + cast(cast(T0.[U_AE_rate] as decimal(19,4)) as nvarchar) " & _
                            '' "end  as 'Cal', " & _
                            ''" T0.[U_AE_NXDT], T0.[Docentry] " & _
                            ''"FROM [dbo].[@AE_SBOOKING]  T0  , OINV T1 " & _
                            ''"WHERE  DAY(T0.[U_AE_NXDT]) >= '" & oform.Items.Item("Item_1").Specific.String & "' and DAY(T0.[U_AE_NXDT]) <= '" & oform.Items.Item("Item_3").Specific.String & "' and T0.[U_AE_Status] = 'Open' and " & _
                            ''"T0.DocNum not in (select rtrim(substring(TT.NumAtCard , 3 , len(Tt.NumAtCard) -2)) from OINV tt where month(tt.DocDate) = '" & oform.Items.Item("12").Specific.selected.description & "' and left(TT.NumAtCard,2) = 'SD'  ) " & _
                            ''"group by T0.[DocNum], T0.[U_AE_Bname], " & _
                            ''"T0.[U_AE_Bcode], T0.[U_AE_Vregno], T0.[U_AE_Rno], T0.[U_AE_Contract], T0.[U_AE_NXDT], T0.[U_AE_Vdatein],T0.[U_AE_Rate], T0.[DocEntry],T0.[U_AE_Sdate] "

                            Str_query = "SELECT T0.[DocNum], T0.[U_AE_Bname], T0.[U_AE_Bcode], T0.[U_AE_Vregno], T0.[U_AE_Rno], T0.[U_AE_Contract], " & _
"case when month(T0.[U_AE_Vdatein]) in ( '" & oform.Items.Item("12").Specific.selected.description & "','" & oform.Items.Item("12").Specific.selected.description + 1 & "') " & _
" then ( case when T0.U_AE_Sdate = T0.[U_AE_NXDT] then (T0.[U_AE_Rate]/DATEDIFF(day,U_AE_NXDT,DATEADD(m,1,U_AE_NXDT) ) )*(DATEDIFF(day,U_AE_Sdate,U_AE_Vdatein)+1) " & _
" else T0.[U_AE_rate] end ) else T0.[U_AE_rate] end  as 'U_AE_rate', " & _
"case when month(T0.[U_AE_Vdatein]) in ( '" & oform.Items.Item("12").Specific.selected.description & "','" & oform.Items.Item("12").Specific.selected.description + 1 & "') " & _
"then ( case when T0.U_AE_Sdate = T0.[U_AE_NXDT] " & _
"then ('Long Term Return Calculation ' + CHAR(13) + '----------------------------------------' + CHAR(13) + CHAR(13) + 'Monthly Rental          : ' + cast(cast(T0.[U_AE_Rate] as decimal(19,4)) as nvarchar ) + char(13) + 'Days in Billing Cycle   : ' +  " & _
"cast(DATEDIFF(day,U_AE_NXDT,DATEADD(m,1,U_AE_NXDT) ) as nvarchar) + char(13) +  " & _
"'Consumed Days         : ' + cast((DATEDIFF(day,U_AE_Sdate,U_AE_Vdatein)+1) as nvarchar) + char(13) + 'Calculated Amount     : ' + " & _
"cast(cast((T0.[U_AE_Rate]/DATEDIFF(day,U_AE_NXDT,DATEADD(m,1,U_AE_NXDT) ) )*(DATEDIFF(day,U_AE_Sdate,U_AE_Vdatein)+1) as decimal(19,4)) as varchar)) " & _
"else 'Long Term Recurring Calculation' + CHAR(13) + '----------------------------------------' + CHAR(13) + CHAR(13) + 'Monthly Rental         : ' + cast(cast(T0.[U_AE_rate] as decimal(19,4)) as nvarchar) end ) else " & _
"'Long Term Recurring Calculation' + CHAR(13) + '----------------------------------------' + CHAR(13) + CHAR(13) + 'Monthly Rental         : ' + cast(cast(T0.[U_AE_rate] as decimal(19,4)) as nvarchar) " & _
"end  as 'Cal', " & _
" T0.[U_AE_NXDT], T0.[Docentry] , cast('" & oform.Items.Item("Item_8").Specific.value & "' as datetime)  'U_AE_Sdate' FROM [dbo].[@AE_SBOOKING]  T0  , OINV T1 " & _
"WHERE  DAY(T0.[U_AE_NXDT]) >= '" & oform.Items.Item("Item_1").Specific.String & "' and DAY(T0.[U_AE_NXDT]) <= '" & oform.Items.Item("Item_3").Specific.String & "' and T0.[U_AE_Status] = 'Open' and " & _
"T0.DocNum not in (select rtrim(substring(TT.NumAtCard , 3 , len(Tt.NumAtCard) -2)) from OINV tt where month(tt.DocDate) = '" & oform.Items.Item("12").Specific.selected.description & "' and left(TT.NumAtCard,2) = 'SD'  ) " & _
"group by T0.[DocNum], T0.[U_AE_Bname], " & _
"T0.[U_AE_Bcode], T0.[U_AE_Vregno], T0.[U_AE_Rno], T0.[U_AE_Contract], T0.[U_AE_NXDT], T0.[U_AE_Vdatein],T0.[U_AE_Rate], T0.[DocEntry],T0.[U_AE_Sdate] "
                            'and month(t0.[U_AE_NXDT]) <= '" & oform.Items.Item("12").Specific.selected.description & "' " & _
                            ''" and (month(T0.[U_AE_Vdatein]) = '" & oform.Items.Item("12").Specific.selected.description & "' or T0.[U_AE_Vdatein] is null) " & _

                            Try
                                oform.DataSources.DataTables.Add("@AE_SBOOKING")
                            Catch ex As Exception
                            End Try

                            oform.DataSources.DataTables.Item("@AE_SBOOKING").ExecuteQuery(Str_query)
                            oMatrix.Clear()
                            oform.Items.Item("Item_6").Specific.columns.item("BNO").databind.bind("@AE_SBOOKING", "DocNum")
                            oform.Items.Item("Item_6").Specific.columns.item("V_0mj").databind.bind("@AE_SBOOKING", "U_AE_NXDT")
                            oform.Items.Item("Item_6").Specific.columns.item("Col_0m").databind.bind("@AE_SBOOKING", "U_AE_Bname")
                            oform.Items.Item("Item_6").Specific.columns.item("Col_1m").databind.bind("@AE_SBOOKING", "U_AE_Vregno")
                            '  oform.Items.Item("Item_6").Specific.columns.item("Col_2m").databind.bind("@AE_SBOOKING", "U_AE_Rno")
                            oform.Items.Item("Item_6").Specific.columns.item("Col_3").databind.bind("@AE_SBOOKING", "U_AE_Contract")
                            oform.Items.Item("Item_6").Specific.columns.item("Col_4mjs").databind.bind("@AE_SBOOKING", "U_AE_rate")
                            oform.Items.Item("Item_6").Specific.columns.item("V_0ca").databind.bind("@AE_SBOOKING", "U_AE_rate")
                            oform.Items.Item("Item_6").Specific.columns.item("V_0sjr").databind.bind("@AE_SBOOKING", "U_AE_Bcode")
                            oform.Items.Item("Item_6").Specific.columns.item("V_0DOC").databind.bind("@AE_SBOOKING", "Docentry")
                            oform.Items.Item("Item_6").Specific.columns.item("Col_Cal").databind.bind("@AE_SBOOKING", "Cal")
                            oform.Items.Item("Item_6").Specific.columns.item("Col_0pd").databind.bind("@AE_SBOOKING", "U_AE_Sdate")

                            oform.Items.Item("Item_6").Specific.LoadFromDataSource()
                            oform.Items.Item("Item_6").Specific.AutoResizeColumns()
                            oMatrix.Columns.Item("V_0ca").Editable = False
                            oMatrix.Columns.Item("Col_0pd").Editable = False
                            oform.Freeze(False)
                            Oapplication_SD.StatusBar.SetText("Operation Completed Successfully ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        Catch ex As Exception
                            oform.Freeze(False)
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Try
                        End Try

                    End If

                    If pVal.ItemUID = "Item_7" Then
                        Dim oform As SAPbouiCOM.Form
                        Dim oform1 As SAPbouiCOM.Form
                        Dim sFunName As String = String.Empty

                        Try
                            oform1 = Oapplication_SD.Forms.ActiveForm
                            sFunName = "SD Recurring Invoice Generation Function()"
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFunName)

                            Dim oMAtrix As SAPbouiCOM.Matrix = oform1.Items.Item("Item_6").Specific
                            Dim oMAtrix_IN As SAPbouiCOM.Matrix
                            Dim oCheck As SAPbouiCOM.CheckBox
                            Oapplication_SD.StatusBar.SetText("Please wait Invoice generation is in process ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            For mjs As Integer = 1 To oMAtrix.RowCount
                                oCheck = oMAtrix.Columns.Item("Col_5m").Cells.Item(mjs).Specific
                                If oCheck.Checked = True Then
                                    ' MsgBox(oMAtrix.Columns.Item("BNO").Cells.Item(mjs).Specific.String)
                                    ' MsgBox(System.DateTime.Parse(oMAtrix.Columns.Item("V_0mj").Cells.Item(mjs).Specific.String, format1, System.Globalization.DateTimeStyles.None))
                                    sFunName = "SD_InvoiceGeneration() - For RA " & oMAtrix.Columns.Item("BNO").Cells.Item(mjs).Specific.String
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFunName)

                                    If SD_InvoiceGeneration(oMAtrix.Columns.Item("V_0sjr").Cells.Item(mjs).Specific.String, _
                                        oMAtrix.Columns.Item("BNO").Cells.Item(mjs).Specific.String, "Self Drive Billing for this Booking No : " & oMAtrix.Columns.Item("BNO").Cells.Item(mjs).Specific.String, _
                                       SD_GLACC, "SO", oMAtrix.Columns.Item("V_0ca").Cells.Item(mjs).Specific.String, "Self Drive Invoice Based on Booking No : " & oMAtrix.Columns.Item("BNO").Cells.Item(mjs).Specific.String, _
                                        oMAtrix.Columns.Item("V_0mj").Cells.Item(mjs).Specific.String, oform1.Items.Item("12").Specific.selected.description, oMAtrix.Columns.Item("V_0DOC").Cells.Item(mjs).Specific.String, "", oMAtrix.Columns.Item("Col_0pd").Cells.Item(mjs).Specific.String) = False Then
                                        BubbleEvent = False
                                        Exit Try
                                    End If
                                    sFunName = "SD_InvoiceGeneration() - For RA " & oMAtrix.Columns.Item("BNO").Cells.Item(mjs).Specific.String
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS", sFunName)
                                End If
                            Next mjs
                            oMAtrix.Columns.Item("V_0ca").Editable = False
                            oform1.Items.Item("Item_5").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFunName)

                        Catch ex As Exception
                            WriteToLogFile(ex.Message, sFunName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFunName)
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Try
                        End Try
                        Exit Sub
                    End If
                End If

            ElseIf pVal.Before_Action = False Then

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                    If pVal.ItemUID = "Item_6" And pVal.ColUID = "Col_5m" Then
                        Dim oform As SAPbouiCOM.Form
                        Try
                            oform = Oapplication_SD.Forms.ActiveForm
                            oform.Freeze(True)
                            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_6").Specific
                            Dim ocheck As SAPbouiCOM.CheckBox = Nothing
                            Dim Flag As Boolean = False
                            ocheck = oMatrix.Columns.Item("Col_5m").Cells.Item(pVal.Row).Specific
                            If ocheck.Checked = True Then
                                Flag = True
                            Else
                                Flag = False
                            End If
                            oMatrix.CommonSetting.SetCellEditable(pVal.Row, 6, Flag)
                            oMatrix.CommonSetting.SetCellEditable(pVal.Row, 7, Flag)
                            oform.Freeze(False)

                        Catch ex As Exception
                            oform.Freeze(False)
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Try
                        End Try
                        Exit Sub

                    End If
                    If pVal.ItemUID = "11" Then
                        Dim oform As SAPbouiCOM.Form
                        Try
                            oform = Oapplication_SD.Forms.ActiveForm
                            Dim ocheck As SAPbouiCOM.CheckBox = oform.Items.Item("11").Specific
                            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_6").Specific
                            oform.Freeze(False)
                            If ocheck.Checked = True Then
                                oMatrix.Columns.Item("V_0mjs").Visible = True
                                oMatrix.Columns.Item("Col_5").Visible = False
                                oform.Items.Item("Item_7").Enabled = False
                            Else
                                oMatrix.Columns.Item("V_0mjs").Visible = False
                                oMatrix.Columns.Item("Col_5").Visible = True
                                oform.Items.Item("Item_7").Enabled = True
                            End If
                            oform.Freeze(True)
                        Catch ex As Exception
                            oform.Freeze(False)
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Try
                        End Try
                        Exit Sub
                    End If

                    If pVal.ItemUID = "13" Then
                        Dim oform As SAPbouiCOM.Form
                        Try
                            oform = Oapplication_SD.Forms.ActiveForm
                            oform.Freeze(True)
                            Dim ocheck As SAPbouiCOM.CheckBox = oform.Items.Item("13").Specific
                            Dim ocheck1 As SAPbouiCOM.CheckBox
                            Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_6").Specific
                            Dim Flag As Boolean = False
                            Oapplication_SD.SetStatusBarMessage("Processing ......... !", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                            If ocheck.Checked = True Then
                                Flag = True
                            Else
                                Flag = False
                            End If

                            For mjs As Integer = 1 To omatrix.RowCount
                                ocheck1 = omatrix.Columns.Item("Col_5m").Cells.Item(mjs).Specific
                                ocheck1.Checked = Flag
                                omatrix.CommonSetting.SetCellEditable(mjs, 6, Flag)
                                omatrix.CommonSetting.SetCellEditable(mjs, 7, Flag)
                            Next mjs
                            oform.Freeze(False)
                            Oapplication_SD.SetStatusBarMessage("Completed Successfully ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        Catch ex As Exception
                            oform.Freeze(False)
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Try
                        End Try
                        Exit Sub
                    End If

                End If
            End If
        End If
    End Sub

    Private Sub Oapplication_SD_LayoutKeyEvent(ByRef eventInfo As SAPbouiCOM.LayoutKeyInfo, ByRef BubbleEvent As Boolean) Handles Oapplication_SD.LayoutKeyEvent

    End Sub

    Private Sub Oapplication_SD_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles Oapplication_SD.MenuEvent

        Try


            If pVal.MenuUID = "SDVB" And pVal.BeforeAction = True Then

                LoadFromXML("SelfDriving_Booking.srf", Oapplication_SD)
                Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.Item("SDB")
                oform.Items.Item("Item_5").Enabled = True
                oform.Visible = True
                Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("1000001").Specific
                Try
                    oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                    oform.Freeze(True)
                    Dim Tmp_val As String = oform.BusinessObject.GetNextSerialNumber(NextSerialNo(Ocompany_SD, Oapplication_SD, "AE_Sbooking"))
                    oform.Items.Item("Item_14").Specific.String = Tmp_val
                    oform.Items.Item("Item_16").Specific.String = Format(Now.Date, "yyyyMMdd") 'Format(Now.Date, "dd MMM yyyy, ddd")
                    Dim ocombo As SAPbouiCOM.ComboBox
                    ocombo = oform.Items.Item("226").Specific
                    ocombo.Select("No")
                    ocombo = oform.Items.Item("228").Specific
                    ocombo.Select("No")
                    oform.Items.Item("Item_115").Enabled = False

                    oMatrix.Columns.Item("V_2").Visible = False

                    Dim ocombobutton As SAPbouiCOM.ButtonCombo = oform.Items.Item("210").Specific
                    ocombobutton.ValidValues.Add("Copy To", "Copy To A/R Invoice")
                    oform.Items.Item("210").Enabled = False

                    oform.Items.Item("217").Specific.String = Ocompany_SD.UserName
                    oform.Items.Item("218").Specific.String = Company_Name
                    oform.Items.Item("Item_154").Specific.String = "9999"

                    Dim opt As SAPbouiCOM.OptionBtn = oform.Items.Item("195").Specific
                    Dim opt1 As SAPbouiCOM.OptionBtn = oform.Items.Item("196").Specific
                    Dim opt2 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_104").Specific
                    Dim opt3 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_105").Specific
                    Dim opt4 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_106").Specific

                    opt1.GroupWith("195")
                    opt1.Selected = True

                    opt3.GroupWith("Item_104")
                    opt4.GroupWith("Item_104")
                    oform.PaneLevel = 3
                    opt3.Selected = True
                    opt2.Selected = True
                    ocombo = oform.Items.Item("Item_158").Specific
                    ocombo.ValidValues.Add("CHEQUE", "CHEQUE")
                    ocombo.ValidValues.Add("CASH", "CASH")
                    ocombo.ValidValues.Add("CREDIT CARD", "CREDIT CARD")
                    ocombo.ValidValues.Add("INVOICE", "INVOICE")
                    oform.Items.Item("188").Specific.String = "SO"

                    ocombo = oform.Items.Item("237").Specific
                    For imjs As Integer = 1 To 31
                        ocombo.ValidValues.Add(imjs, imjs)
                    Next
                    ocombo.ValidValues.Add("", "")

                    oform.Items.Item("Item_22").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Dim oCFLs As SAPbouiCOM.ChooseFromList
                    Dim oCons As SAPbouiCOM.Conditions
                    Dim oCon As SAPbouiCOM.Condition
                    Dim empty As New SAPbouiCOM.Conditions

                    oCFLs = oform.ChooseFromLists.Item("CFL_5")
                    oCFLs.SetConditions(empty)
                    oCons = oCFLs.GetConditions()
                    oCon = oCons.Add()
                    oCon.Alias = "ItmsGrpCod"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = ItemGroup
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    oCon = oCons.Add()
                    oCon.Alias = "Validfor"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                    oCFLs.SetConditions(oCons)


                    oCFLs = oform.ChooseFromLists.Item("CFL_2")
                    oCFLs.SetConditions(empty)
                    oCons = oCFLs.GetConditions()
                    oCon = oCons.Add()
                    oCon.Alias = "CardType"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "C"
                    oCFLs.SetConditions(oCons)

                    '
                    oform.DataBrowser.BrowseBy = "Item_14"

                    Try
                        ocombo = oform.Items.Item("Item_18").Specific
                        ocombo.ValidValues.Add("Open", "Open")
                        ocombo.ValidValues.Add("Billing", "Billing")
                        ocombo.ValidValues.Add("Cancel", "Cancel")
                        ocombo.ValidValues.Add("Closed", "Closed")


                    Catch ex As Exception
                    End Try
                    ocombo.Select("Open")

                    ocombo = oform.Items.Item("228").Specific
                    ocombo.Select("-")
                    oform.Freeze(False)
                Catch ex As Exception
                    oform.Freeze(False)
                End Try


            End If


            If pVal.MenuUID = "SDB" And pVal.BeforeAction = True Then

                Try
                    LoadFromXML("SelfDriving_Billing.srf", Oapplication_SD)
                    Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.Item("SDBIL")
                    oform.Visible = True
                    oform.Items.Item("Item_10").Specific.String = Format(Now.Date, "yyyyMMdd")
                    Dim oCombo As SAPbouiCOM.ComboBox = oform.Items.Item("12").Specific
                    ' MsgBox(MonthName(Month(Now) - 1))


                    ' oCombo.ValidValues.Add(MonthName(Month(Now) + 1), Now.Month + 1)
                    If Month(Now) = 12 Then
                        oCombo.ValidValues.Add(MonthName(Month(Now) - 1), Now.Month - 1)
                        oCombo.ValidValues.Add(MonthName(Month(Now)), Now.Month)
                        oCombo.ValidValues.Add(MonthName(1), 1)
                    ElseIf Month(Now) = 1 Then
                        oCombo.ValidValues.Add(MonthName(12), 12)
                        oCombo.ValidValues.Add(MonthName(Month(Now)), Now.Month)
                        oCombo.ValidValues.Add(MonthName(Month(Now) + 1), Now.Month + 1)
                    Else
                        oCombo.ValidValues.Add(MonthName(Month(Now) - 1), Now.Month - 1)
                        oCombo.ValidValues.Add(MonthName(Month(Now)), Now.Month)
                        oCombo.ValidValues.Add(MonthName(Month(Now) + 1), Now.Month + 1)
                    End If
                    oCombo.Select(MonthName(Month(Now)), SAPbouiCOM.BoSearchKey.psk_ByValue)
                    Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_6").Specific
                    oform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
                Catch ex As Exception
                    Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try

            End If
            'SDBR

            If pVal.MenuUID = "SDBR" And pVal.BeforeAction = True Then

                Try
                    LoadFromXML("SelfDrive_BookingReport.srf", Oapplication_SD)
                    Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.Item("SDBR")
                    oform.Visible = True
                    Try
                        oform.DataSources.DataTables.Add("SDBR")
                    Catch ex As Exception

                    End Try

                    Dim ogrid As SAPbouiCOM.Grid = oform.Items.Item("8").Specific

                    ogrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto
                    oform.DataSources.DataTables.Item(0).ExecuteQuery("SELECT T0.[DocNum] as 'Billing No', T0.[U_AE_Vregno] as 'Vehicle Number', T0.[U_AE_Vmodel] as 'Vehicle Model', T0.[U_AE_Bname] as 'Customer', T0.[U_AE_DName] as 'Driver', T0.[U_AE_Sdate] as 'Start Date', T0.[U_AE_Edate] as 'End Date' FROM [dbo].[@AE_SBOOKING]  T0 where docnum = ''")
                    ogrid.DataTable = oform.DataSources.DataTables.Item("SDBR")
                    ogrid.AutoResizeColumns()

                    oform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
                    oform.Visible = True

                    Dim oCFLs As SAPbouiCOM.ChooseFromList
                    Dim oCons As SAPbouiCOM.Conditions
                    Dim oCon As SAPbouiCOM.Condition
                    Dim empty As New SAPbouiCOM.Conditions


                    oCFLs = oform.ChooseFromLists.Item("CFL_2")
                    oCFLs.SetConditions(empty)
                    oCons = oCFLs.GetConditions()
                    oCon = oCons.Add()
                    oCon.Alias = "CardType"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "C"
                    oCFLs.SetConditions(oCons)


                Catch ex As Exception
                    Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try

            End If


            If pVal.MenuUID = "SDIR" And pVal.BeforeAction = True Then

                Try
                    Dim oEdit As SAPbouiCOM.EditText
                    LoadFromXML("SelfDriving_InvoiceReport.srf", Oapplication_SD)
                    Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.Item("SDINV")
                    oform.DataSources.UserDataSources.Add("sDate", SAPbouiCOM.BoDataType.dt_DATE)
                    oEdit = oform.Items.Item("Item_10").Specific
                    oEdit.DataBind.SetBound(True, "", "sDate")
                    oform.Items.Item("Item_10").Specific.String = Format(Now.Date, "yyyyMMdd") 'Now.Date
                    Dim oCombo As SAPbouiCOM.ComboBox = oform.Items.Item("12").Specific
                    oform.Visible = True
                    ' MsgBox(MonthName(Month(Now) - 1))
                    ''oCombo.ValidValues.Add(MonthName(Month(Now) - 1), Now.Month - 1)
                    ''oCombo.ValidValues.Add(MonthName(Month(Now)), Now.Month)
                    ''If Month(Now) = 12 Then
                    ''    oCombo.ValidValues.Add(MonthName(1), 1)
                    ''Else
                    ''    oCombo.ValidValues.Add(MonthName(Month(Now) + 1), Now.Month + 1)
                    ''End If

                    If Month(Now) = 12 Then
                        oCombo.ValidValues.Add(MonthName(Month(Now) - 1), Now.Month - 1)
                        oCombo.ValidValues.Add(MonthName(Month(Now)), Now.Month)
                        oCombo.ValidValues.Add(MonthName(1), 1)
                    ElseIf Month(Now) = 1 Then
                        oCombo.ValidValues.Add(MonthName(12), 12)
                        oCombo.ValidValues.Add(MonthName(Month(Now)), Now.Month)
                        oCombo.ValidValues.Add(MonthName(Month(Now) + 1), Now.Month + 1)
                    Else
                        oCombo.ValidValues.Add(MonthName(Month(Now) - 1), Now.Month - 1)
                        oCombo.ValidValues.Add(MonthName(Month(Now)), Now.Month)
                        oCombo.ValidValues.Add(MonthName(Month(Now) + 1), Now.Month + 1)
                    End If

                    oCombo.Select(MonthName(Month(Now)), SAPbouiCOM.BoSearchKey.psk_ByValue)


                    oform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
                    oform.Visible = True

                    Dim oCFLs As SAPbouiCOM.ChooseFromList
                    Dim oCons As SAPbouiCOM.Conditions
                    Dim oCon As SAPbouiCOM.Condition
                    Dim empty As New SAPbouiCOM.Conditions


                    oCFLs = oform.ChooseFromLists.Item("CFL_2")
                    oCFLs.SetConditions(empty)
                    oCons = oCFLs.GetConditions()
                    oCon = oCons.Add()
                    oCon.Alias = "CardType"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "C"
                    oCFLs.SetConditions(oCons)


                Catch ex As Exception
                    Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try

            End If

            If pVal.MenuUID = "1281" And pVal.BeforeAction = False Then
                Try
                    Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                    If oform.UniqueID = "SDB" Then
                        oform.Items.Item("Item_14").Enabled = True
                        oform.Items.Item("210").Enabled = True
                        oform.Items.Item("233").Enabled = True
                    End If

                Catch ex As Exception

                End Try
            End If

            If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.BeforeAction = False Then

                Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm

                If oform.UniqueID = "SDB" Then
                    NavigationValidation_SelfDriver(oform, Ocompany_SD, Oapplication_SD)
                End If

            End If


        Catch ex As Exception

        End Try
    End Sub





    Public Function showOpenFileDialog() As String

        Dim ShowFolderBrowserThread As Threading.Thread
        Try
            ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowser)
            If ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Unstarted Then
                ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA)
                ShowFolderBrowserThread.Start()
                ShowFolderBrowserThread.Join()
            ElseIf ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Stopped Then
                ShowFolderBrowserThread.Start()
                ShowFolderBrowserThread.Join()

            End If
            While ShowFolderBrowserThread.ThreadState = Threading.ThreadState.Running
                Windows.Forms.Application.DoEvents()
            End While
            If FileName <> "" Then
                Return FileName
            End If
        Catch ex As Exception
            'SBO_Application.MessageBox("FileFile" & ex.Message)
            MessageBox.Show(ex.ToString())
        End Try

        Return ""

    End Function

    Private Function SD_InvoiceGeneration(ByVal CardCode As String, ByVal NumAtCard As String, ByVal Description As String, ByVal AcountCode As String, ByVal Taxcode As String, ByVal LineTotal As Double, _
                                          ByVal comments As String, ByVal ADate As String, ByVal month_ As Integer, ByVal DocEntry_ As String, ByVal sNInvoiceDate As String) As Boolean
        Dim sFunName As String = String.Empty
        Try
            Dim oInvoice As SAPbobsCOM.Documents = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fetching the information for RA ", NumAtCard)

            Dim val As Integer = 0
            Dim InvoiceDocEntry, InvoiceDocNum As String
            Dim dSDate, dEDate As Date
            Dim dPetrol As Double = 0
            Dim orset As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset.DoQuery("SELECT T0.[U_AE_CDW], T0.[U_AE_CDW1], T0.[U_AE_PAI],T0.[U_AE_PAI1], isnull(T0.[U_AE_Ocharges],0) [U_AE_Ocharges], isnull(T0.[U_AE_Rcharg],0) [U_AE_Rcharg], isnull(T0.[U_AE_petrol],0) [U_AE_petrol] FROM [dbo].[@AE_SBOOKING]  T0 WHERE T0.[DocEntry]  = '" & DocEntry_ & "'")
            Dim sCDWFlag As String = orset.Fields.Item("U_AE_CDW1").Value
            Dim dCDW As Double = orset.Fields.Item("U_AE_CDW").Value
            Dim sPAIFlag As String = orset.Fields.Item("U_AE_PAI1").Value
            Dim dPAI As Double = orset.Fields.Item("U_AE_PAI").Value

            Dim dOtherCharges As Double = orset.Fields.Item("U_AE_Ocharges").Value
            Dim dMontlyRecurring As Double = orset.Fields.Item("U_AE_Rcharg").Value
            If Not String.IsNullOrEmpty(orset.Fields.Item("U_AE_petrol").Value) Then
                dPetrol = orset.Fields.Item("U_AE_petrol").Value
            End If


            orset.DoQuery("SELECT T0.[Code] FROM OLCT T0")
            Dim _tmpdate As Date = Convert.ToDateTime(ADate)
            oInvoice.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
            'MsgBox(month_)
            'MsgBox(_tmpdate.Day)
            oInvoice.CardCode = CardCode
            oInvoice.NumAtCard = "SD " & NumAtCard

            oInvoice.DocDate = PostDate(_tmpdate.Day, month_, Now.Year, Oapplication_SD)
            dSDate = PostDate(_tmpdate.Day, month_, Now.Year, Oapplication_SD)
            dEDate = dSDate.AddMonths(1)
            Description = "Rental period from " & Format(dSDate, "dd/MM/yyyy") & " to " & Format(dEDate.AddDays(-1), "dd/MM/yyyy") & " (" & DateDiff(DateInterval.Day, dSDate, dEDate) & " Days)"
            oInvoice.TaxDate = PostDate(_tmpdate.Day, month_, Now.Year, Oapplication_SD) 'Server_Date

            oInvoice.Lines.ItemDescription = Description
            oInvoice.Lines.AccountCode = AcountCode
            oInvoice.Lines.TaxCode = Taxcode
            oInvoice.Lines.LineTotal = LineTotal
            'MsgBox(orset.Fields.Item("Code").Value)
            oInvoice.Lines.LocationCode = "1"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CDW Flag status ", sCDWFlag)

            If sCDWFlag = "Yes" And CInt(dCDW) > 0 Then
                oInvoice.Lines.Add()
                oInvoice.Lines.ItemDescription = "Collission Damage Waiver Fee"
                oInvoice.Lines.TaxCode = Taxcode
                oInvoice.Lines.LineTotal = dCDW
                oInvoice.Lines.AccountCode = CDW_GLACC
                oInvoice.Lines.LocationCode = "1"
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("PAI Flag status ", sPAIFlag)

            If sPAIFlag = "Yes" And CInt(dPAI) > 0 Then
                oInvoice.Lines.Add()
                oInvoice.Lines.ItemDescription = "Personal Accidental Insurance"
                oInvoice.Lines.TaxCode = Taxcode
                oInvoice.Lines.LineTotal = dPAI
                oInvoice.Lines.AccountCode = PAI_GLACC
                oInvoice.Lines.LocationCode = "1"
            End If

            If CInt(dOtherCharges) > 0 Then
                oInvoice.Lines.Add()
                oInvoice.Lines.ItemDescription = "Other Charges"
                oInvoice.Lines.TaxCode = Taxcode
                oInvoice.Lines.LineTotal = dOtherCharges
                oInvoice.Lines.AccountCode = sOC_GLACC
                oInvoice.Lines.LocationCode = "1"
            End If

            If CInt(dMontlyRecurring) > 0 Then
                oInvoice.Lines.Add()
                oInvoice.Lines.ItemDescription = "Monthly Recurring Other Charges"
                oInvoice.Lines.TaxCode = Taxcode
                oInvoice.Lines.LineTotal = dMontlyRecurring
                oInvoice.Lines.AccountCode = sMR_GLACC
                oInvoice.Lines.LocationCode = "1"
            End If

            If CInt(dPetrol) > 0 Then
                oInvoice.Lines.Add()
                oInvoice.Lines.ItemDescription = "Petrol Charges"
                oInvoice.Lines.TaxCode = Taxcode
                oInvoice.Lines.LineTotal = dPetrol
                oInvoice.Lines.AccountCode = sPet_GLACC
                oInvoice.Lines.LocationCode = "1"
            End If

            oInvoice.Comments = Description ''comments

            If Ocompany_SD.InTransaction = False Then
                Ocompany_SD.StartTransaction()
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Generating the Invoice for RA ", NumAtCard)

            val = oInvoice.Add

            If val <> 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR  " & Ocompany_SD.GetLastErrorDescription & " for RA ", NumAtCard)
                WriteToLogFile(Ocompany_SD.GetLastErrorDescription, "ERROR Generating the Invoice for RA" & NumAtCard)
                Oapplication_SD.MessageBox("Error Msg : " & Ocompany_SD.GetLastErrorDescription, 1, "Ok")
                If Ocompany_SD.InTransaction = True Then
                    Ocompany_SD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                End If

                Return False
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with  SUCCESS ", NumAtCard)
                InvoiceDocEntry = Ocompany_SD.GetNewObjectKey

                orset.DoQuery("SELECT DocNum,DocDate FROM OINV T0 WHERE Docentry = '" & InvoiceDocEntry & "'")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Generated Invoice Number is ", orset.Fields.Item("DocNum").Value)

                sFunName = "UDO_Update_SelfDrive"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFunName)

                If UDO_Update_SelfDrive(NumAtCard, ADate, LineTotal, InvoiceDocEntry, orset.Fields.Item("DocNum").Value, DocEntry_, Now.Year & "/" & month_ & "/" & _tmpdate.Day, orset.Fields.Item("DocDate").Value, sNInvoiceDate) = False Then
                    If Ocompany_SD.InTransaction = True Then
                        Ocompany_SD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If
                    Return False
                End If

            End If
            If Ocompany_SD.InTransaction = True Then
                Ocompany_SD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", "SD_InvoiceGeneration()")
            Oapplication_SD.StatusBar.SetText("Successfully Completed the Invoice Generation for the Booking : " & NumAtCard, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


            Return True
        Catch ex As Exception
            WriteToLogFile(ex.Message, "SD_InvoiceGeneration()")
            Oapplication_SD.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
            If Ocompany_SD.InTransaction = True Then
                Ocompany_SD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            Return False
        End Try
    End Function

    Private Function SD_InvoiceGeneration(ByVal CardCode As String, ByVal NumAtCard As String, ByVal Description As String, ByVal AcountCode As String, ByVal Taxcode As String, ByVal LineTotal As Double, _
                                          ByVal comments As String, ByVal ADate As String, ByVal month_ As Integer, ByVal DocEntry_ As String, ByVal sNInvoiceDate As String, ByVal sPDate As String) As Boolean
        Dim sFunName As String = String.Empty
        Try
            Dim oInvoice As SAPbobsCOM.Documents = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fetching the information for RA ", NumAtCard)

            Dim val As Integer = 0
            Dim dPetrol As Double = 0
            Dim InvoiceDocEntry, InvoiceDocNum As String
            Dim dSDate, dEDate As Date

            Dim orset As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset.DoQuery("SELECT isnull(T0.[U_AE_CDW],0) [U_AE_CDW], T0.[U_AE_CDW1], isnull(T0.[U_AE_PAI],0) [U_AE_PAI],T0.[U_AE_PAI1], isnull(T0.[U_AE_Ocharges],0) [U_AE_Ocharges], isnull(T0.[U_AE_Rcharg],0) [U_AE_Rcharg], isnull(T0.[U_AE_petrol],0) [U_AE_petrol] FROM [dbo].[@AE_SBOOKING]  T0 WHERE T0.[DocEntry]  = '" & DocEntry_ & "'")
            Dim sCDWFlag As String = orset.Fields.Item("U_AE_CDW1").Value
            Dim dCDW As Double = orset.Fields.Item("U_AE_CDW").Value
            Dim sPAIFlag As String = orset.Fields.Item("U_AE_PAI1").Value
            Dim dPAI As Double = orset.Fields.Item("U_AE_PAI").Value

            Dim dOtherCharges As Double = orset.Fields.Item("U_AE_Ocharges").Value
            Dim dMontlyRecurring As Double = orset.Fields.Item("U_AE_Rcharg").Value
            If Not String.IsNullOrEmpty(orset.Fields.Item("U_AE_petrol").Value) Then
                dPetrol = orset.Fields.Item("U_AE_petrol").Value
            End If


            orset.DoQuery("SELECT T0.[Code] FROM OLCT T0")
            Dim _tmpdate As Date = Convert.ToDateTime(ADate)
            Dim dPostingdate As Date = Convert.ToDateTime(sPDate)
            oInvoice.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
            'MsgBox(month_)
            'MsgBox(_tmpdate.Day)
            oInvoice.CardCode = CardCode
            oInvoice.NumAtCard = "SD " & NumAtCard

            '' oInvoice.DocDate = PostDate(_tmpdate.Day, month_, Now.Year, Oapplication_SD)
            oInvoice.DocDate = dPostingdate ''PostDate(dPostingdate.Day, month_, Now.Year, Oapplication_SD)
            dSDate = PostDate(_tmpdate.Day, month_, Now.Year, Oapplication_SD)
            dEDate = dSDate.AddMonths(1)
            Description = "Rental period from " & Format(dSDate, "dd/MM/yyyy") & " to " & Format(dEDate.AddDays(-1), "dd/MM/yyyy") & " (" & DateDiff(DateInterval.Day, dSDate, dEDate) & " Days)"
            oInvoice.TaxDate = dPostingdate ''PostDate(_tmpdate.Day, month_, Now.Year, Oapplication_SD) 'Server_Date

            oInvoice.Lines.ItemDescription = Description
            oInvoice.Lines.AccountCode = AcountCode
            oInvoice.Lines.TaxCode = Taxcode
            oInvoice.Lines.LineTotal = LineTotal
            'MsgBox(orset.Fields.Item("Code").Value)
            oInvoice.Lines.LocationCode = "1"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CDW Flag status ", sCDWFlag)

            If sCDWFlag = "Yes" And CInt(dCDW) > 0 Then
                oInvoice.Lines.Add()
                oInvoice.Lines.ItemDescription = "Collission Damage Waiver Fee"
                oInvoice.Lines.TaxCode = Taxcode
                oInvoice.Lines.LineTotal = dCDW
                oInvoice.Lines.AccountCode = CDW_GLACC
                oInvoice.Lines.LocationCode = "1"
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("PAI Flag status ", sPAIFlag)

            If sPAIFlag = "Yes" And CInt(dPAI) > 0 Then
                oInvoice.Lines.Add()
                oInvoice.Lines.ItemDescription = "Personal Accidental Insurance"
                oInvoice.Lines.TaxCode = Taxcode
                oInvoice.Lines.LineTotal = dPAI
                oInvoice.Lines.AccountCode = PAI_GLACC
                oInvoice.Lines.LocationCode = "1"
            End If

            If CInt(dOtherCharges) > 0 Then
                oInvoice.Lines.Add()
                oInvoice.Lines.ItemDescription = "Other Charges"
                oInvoice.Lines.TaxCode = Taxcode
                oInvoice.Lines.LineTotal = dOtherCharges
                oInvoice.Lines.AccountCode = sOC_GLACC
                oInvoice.Lines.LocationCode = "1"
            End If

            If CInt(dMontlyRecurring) > 0 Then
                oInvoice.Lines.Add()
                oInvoice.Lines.ItemDescription = "Monthly Recurring Other Charges"
                oInvoice.Lines.TaxCode = Taxcode
                oInvoice.Lines.LineTotal = dMontlyRecurring
                oInvoice.Lines.AccountCode = sMR_GLACC
                oInvoice.Lines.LocationCode = "1"
            End If

            If CInt(dPetrol) > 0 Then
                oInvoice.Lines.Add()
                oInvoice.Lines.ItemDescription = "Petrol Charges"
                oInvoice.Lines.TaxCode = Taxcode
                oInvoice.Lines.LineTotal = dPetrol
                oInvoice.Lines.AccountCode = sPet_GLACC
                oInvoice.Lines.LocationCode = "1"
            End If

            oInvoice.Comments = Description ''comments

            If Ocompany_SD.InTransaction = False Then
                Ocompany_SD.StartTransaction()
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Generating the Invoice for RA ", NumAtCard)

            val = oInvoice.Add

            If val <> 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR  " & Ocompany_SD.GetLastErrorDescription & " for RA ", NumAtCard)
                WriteToLogFile(Ocompany_SD.GetLastErrorDescription, "ERROR Generating the Invoice for RA" & NumAtCard)
                Oapplication_SD.MessageBox("Error Msg : " & Ocompany_SD.GetLastErrorDescription, 1, "Ok")
                If Ocompany_SD.InTransaction = True Then
                    Ocompany_SD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                End If

                Return False
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with  SUCCESS ", NumAtCard)
                InvoiceDocEntry = Ocompany_SD.GetNewObjectKey

                orset.DoQuery("SELECT DocNum,DocDate FROM OINV T0 WHERE Docentry = '" & InvoiceDocEntry & "'")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Generated Invoice Number is ", orset.Fields.Item("DocNum").Value)

                sFunName = "UDO_Update_SelfDrive"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFunName)

                If UDO_Update_SelfDrive(NumAtCard, ADate, LineTotal, InvoiceDocEntry, orset.Fields.Item("DocNum").Value, DocEntry_, Now.Year & "/" & month_ & "/" & _tmpdate.Day, orset.Fields.Item("DocDate").Value, sNInvoiceDate, _tmpdate) = False Then
                    If Ocompany_SD.InTransaction = True Then
                        Ocompany_SD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If
                    Return False
                End If

            End If
            If Ocompany_SD.InTransaction = True Then
                Ocompany_SD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", "SD_InvoiceGeneration()")
            Oapplication_SD.StatusBar.SetText("Successfully Completed the Invoice Generation for the Booking : " & NumAtCard, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


            Return True
        Catch ex As Exception
            WriteToLogFile(ex.Message, "SD_InvoiceGeneration()")
            Oapplication_SD.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
            If Ocompany_SD.InTransaction = True Then
                Ocompany_SD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            Return False
        End Try
    End Function


    Public Function UDO_Update_SelfDrive(ByVal DocNum As String, ByVal NX_Date As String, ByVal InvoiceAmount As Double, ByVal InvoiceDocEntry As String, ByVal InvoiceDocNum As String, ByVal RADocentry As String, ByVal NInvDateNew As Date, ByVal DocDate As Date, ByVal sNInvoiceDate As String) As Boolean
        Dim sFunName As String = String.Empty

        Try
            Dim oGeneralService As SAPbobsCOM.GeneralService = Nothing
            Dim oGeneralData As SAPbobsCOM.GeneralData = Nothing
            Dim oGeneralDataParam As SAPbobsCOM.GeneralDataParams = Nothing
            Dim oCompanyService As SAPbobsCOM.CompanyService = Ocompany_SD.GetCompanyService
            Dim orset As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim orset1 As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset1.DoQuery("SELECT ISNULL(U_AE_Vdatein,'01.01.1900') AS 'U_AE_Vdatein' FROM [dbo].[@AE_SBOOKING]  T0 WHERE T0.[DocEntry]  = '" & RADocentry & "'")
            Dim DateConvertion As DateTime = Convert.ToDateTime(NX_Date)

            orset.DoQuery("SELECT max(cast(T0.[LineId] as integer)) as 'LineID', T0.[U_AE_Ino] FROM [dbo].[@AE_SBOOKING_R1]  T0 WHERE T0.[DocEntry] = '" & RADocentry & "' GROUP BY T0.[U_AE_Ino]")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting UDO Object ", "AE_Sbooking")
            oGeneralService = oCompanyService.GetGeneralService("AE_Sbooking")
            oGeneralDataParam = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Set the RA Document Entry for update ", RADocentry)
            oGeneralDataParam.SetProperty("DocEntry", RADocentry) ' Trim(DocEntry.ToString.Substring(3, DocEntry.ToString.Length - 3)))
            oGeneralData = oGeneralService.GetByParams(oGeneralDataParam)
            ' Document Fields to update
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("NInvDateNew ", Format(NInvDateNew, "dd/MM/yyyy"))

            If Not String.IsNullOrEmpty(sNInvoiceDate) Then
                oGeneralData.SetProperty("U_AE_NXDT", DateTime.ParseExact(sNInvoiceDate, "yyyyMMdd", Nothing))
            Else
                If orset1.Fields.Item("U_AE_Vdatein").Value > DateConvertion.AddMonths(1) Or CDate(orset1.Fields.Item("U_AE_Vdatein").Value).ToString("yyyyMMdd") = "19000101" Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("NInvDateNew.AddMonths(1) ", Format(NInvDateNew.AddMonths(1), "dd/MM/yyyy"))
                    oGeneralData.SetProperty("U_AE_NXDT", NInvDateNew.AddMonths(1))
                Else
                    oGeneralData.SetProperty("U_AE_Status", "Closed")
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Set the Next service date ", Format(NInvDateNew.AddMonths(1), "dd/MM/yyyy"))
            'oGeneralService.Update(oGeneralData)

            Dim oChildTableRows As SAPbobsCOM.GeneralDataCollection = oGeneralData.Child("AE_SBOOKING_R1")

            '   Dim LineNum As Integer = 0
            ' Document Line Fields to update
            Dim oChildTableRow As SAPbobsCOM.GeneralData

            ' MsgBox(orset.Fields.Item("LineID").Value)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Property check for idenftify the Line Num  ", orset.Fields.Item("U_AE_Ino").Value & " , " & orset.Fields.Item("LineID").Value)

            If orset.Fields.Item("U_AE_Ino").Value = "" Then

                If orset.Fields.Item("LineID").Value = 0 Then
                    oChildTableRow = oChildTableRows.Add
                    ' oChildTableRow.SetProperty("LineID", orset.Fields.Item("LineID").Value)
                    oChildTableRow.SetProperty("U_AE_Adate", DocDate)
                    oChildTableRow.SetProperty("U_AE_Pdate", DocDate)
                    oChildTableRow.SetProperty("U_AE_Idate", Server_Date)
                    oChildTableRow.SetProperty("U_AE_Amoun", InvoiceAmount)
                    oChildTableRow.SetProperty("U_AE_Ino", InvoiceDocEntry)
                    oChildTableRow.SetProperty("U_AE_Idocn", InvoiceDocNum)
                Else
                    oChildTableRow = oChildTableRows.Item(orset.Fields.Item("LineID").Value - 1)
                    If oChildTableRow.GetProperty("LineId") = orset.Fields.Item("LineID").Value Then 'LINENUM_MACTHING
                        oChildTableRow.SetProperty("U_AE_Adate", DocDate)
                        oChildTableRow.SetProperty("U_AE_Pdate", DocDate)
                        oChildTableRow.SetProperty("U_AE_Idate", Server_Date)
                        oChildTableRow.SetProperty("U_AE_Amoun", InvoiceAmount)
                        oChildTableRow.SetProperty("U_AE_Ino", InvoiceDocEntry)
                        oChildTableRow.SetProperty("U_AE_Idocn", InvoiceDocNum)
                        ' oGeneralService.Update(oGeneralData)
                    End If
                End If
            Else
                oChildTableRow = oChildTableRows.Add
                ' oChildTableRow.SetProperty("LineID", orset.Fields.Item("LineID").Value)
                oChildTableRow.SetProperty("U_AE_Adate", DocDate)
                oChildTableRow.SetProperty("U_AE_Pdate", DocDate)
                oChildTableRow.SetProperty("U_AE_Idate", Server_Date)
                oChildTableRow.SetProperty("U_AE_Amoun", InvoiceAmount)
                oChildTableRow.SetProperty("U_AE_Ino", InvoiceDocEntry)
                oChildTableRow.SetProperty("U_AE_Idocn", InvoiceDocNum)
            End If
            'oApplication.StatusBar.SetText("Updating your information in Budget Proposal object ......... " & mj, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Start updating the UDO ", "AE_Sbooking")
            oGeneralService.Update(oGeneralData)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", "UDO_Update_SelfDrive()")
            Return True
        Catch ex As Exception
            WriteToLogFile(ex.Message, "UDO_Update_SelfDrive() for RA " & RADocentry)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR - AE_Sbooking", ex.Message)
            Oapplication_SD.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
            Return False
        End Try

    End Function

    Public Function UDO_Update_SelfDrive(ByVal DocNum As String, ByVal NX_Date As String, ByVal InvoiceAmount As Double, ByVal InvoiceDocEntry As String, ByVal InvoiceDocNum As String, ByVal RADocentry As String, ByVal NInvDateNew As Date, ByVal DocDate As Date, ByVal sNInvoiceDate As String, ByVal dbdate As Date) As Boolean
        Dim sFunName As String = String.Empty

        Try
            Dim oGeneralService As SAPbobsCOM.GeneralService = Nothing
            Dim oGeneralData As SAPbobsCOM.GeneralData = Nothing
            Dim oGeneralDataParam As SAPbobsCOM.GeneralDataParams = Nothing
            Dim oCompanyService As SAPbobsCOM.CompanyService = Ocompany_SD.GetCompanyService
            Dim orset As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim orset1 As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset1.DoQuery("SELECT ISNULL(U_AE_Vdatein,'01.01.1900') AS 'U_AE_Vdatein' FROM [dbo].[@AE_SBOOKING]  T0 WHERE T0.[DocEntry]  = '" & RADocentry & "'")
            Dim DateConvertion As DateTime = Convert.ToDateTime(NX_Date)

            orset.DoQuery("SELECT max(cast(T0.[LineId] as integer)) as 'LineID', T0.[U_AE_Ino] FROM [dbo].[@AE_SBOOKING_R1]  T0 WHERE T0.[DocEntry] = '" & RADocentry & "' GROUP BY T0.[U_AE_Ino]")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting UDO Object ", "AE_Sbooking")
            oGeneralService = oCompanyService.GetGeneralService("AE_Sbooking")
            oGeneralDataParam = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Set the RA Document Entry for update ", RADocentry)
            oGeneralDataParam.SetProperty("DocEntry", RADocentry) ' Trim(DocEntry.ToString.Substring(3, DocEntry.ToString.Length - 3)))
            oGeneralData = oGeneralService.GetByParams(oGeneralDataParam)
            ' Document Fields to update
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("NInvDateNew ", Format(NInvDateNew, "dd/MM/yyyy"))

            If Not String.IsNullOrEmpty(sNInvoiceDate) Then
                oGeneralData.SetProperty("U_AE_NXDT", DateTime.ParseExact(sNInvoiceDate, "yyyyMMdd", Nothing))
            Else
                If orset1.Fields.Item("U_AE_Vdatein").Value > DateConvertion.AddMonths(1) Or CDate(orset1.Fields.Item("U_AE_Vdatein").Value).ToString("yyyyMMdd") = "19000101" Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("NInvDateNew.AddMonths(1) ", Format(NInvDateNew.AddMonths(1), "dd/MM/yyyy"))
                    oGeneralData.SetProperty("U_AE_NXDT", NInvDateNew.AddMonths(1))
                Else
                    oGeneralData.SetProperty("U_AE_Status", "Closed")
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Set the Next service date ", Format(NInvDateNew.AddMonths(1), "dd/MM/yyyy"))
            'oGeneralService.Update(oGeneralData)

            Dim oChildTableRows As SAPbobsCOM.GeneralDataCollection = oGeneralData.Child("AE_SBOOKING_R1")

            '   Dim LineNum As Integer = 0
            ' Document Line Fields to update
            Dim oChildTableRow As SAPbobsCOM.GeneralData

            ' MsgBox(orset.Fields.Item("LineID").Value)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Property check for idenftify the Line Num  ", orset.Fields.Item("U_AE_Ino").Value & " , " & orset.Fields.Item("LineID").Value)

            If orset.Fields.Item("U_AE_Ino").Value = "" Then

                If orset.Fields.Item("LineID").Value = 0 Then
                    oChildTableRow = oChildTableRows.Add
                    ' oChildTableRow.SetProperty("LineID", orset.Fields.Item("LineID").Value)
                    oChildTableRow.SetProperty("U_AE_Adate", dbdate) ''DocDate)
                    oChildTableRow.SetProperty("U_AE_Idate", Server_Date)
                    oChildTableRow.SetProperty("U_AE_Pdate", DocDate)
                    oChildTableRow.SetProperty("U_AE_Amoun", InvoiceAmount)
                    oChildTableRow.SetProperty("U_AE_Ino", InvoiceDocEntry)
                    oChildTableRow.SetProperty("U_AE_Idocn", InvoiceDocNum)
                Else
                    oChildTableRow = oChildTableRows.Item(orset.Fields.Item("LineID").Value - 1)
                    If oChildTableRow.GetProperty("LineId") = orset.Fields.Item("LineID").Value Then 'LINENUM_MACTHING
                        oChildTableRow.SetProperty("U_AE_Adate", dbdate) ''DocDate)
                        oChildTableRow.SetProperty("U_AE_Idate", Server_Date)
                        oChildTableRow.SetProperty("U_AE_Pdate", DocDate)
                        oChildTableRow.SetProperty("U_AE_Amoun", InvoiceAmount)
                        oChildTableRow.SetProperty("U_AE_Ino", InvoiceDocEntry)
                        oChildTableRow.SetProperty("U_AE_Idocn", InvoiceDocNum)
                        ' oGeneralService.Update(oGeneralData)
                    End If
                End If
            Else
                oChildTableRow = oChildTableRows.Add
                ' oChildTableRow.SetProperty("LineID", orset.Fields.Item("LineID").Value)
                oChildTableRow.SetProperty("U_AE_Adate", dbdate) ''DocDate)
                oChildTableRow.SetProperty("U_AE_Idate", Server_Date)
                oChildTableRow.SetProperty("U_AE_Pdate", DocDate)
                oChildTableRow.SetProperty("U_AE_Amoun", InvoiceAmount)
                oChildTableRow.SetProperty("U_AE_Ino", InvoiceDocEntry)
                oChildTableRow.SetProperty("U_AE_Idocn", InvoiceDocNum)
            End If
            'oApplication.StatusBar.SetText("Updating your information in Budget Proposal object ......... " & mj, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Start updating the UDO ", "AE_Sbooking")
            oGeneralService.Update(oGeneralData)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS ", "UDO_Update_SelfDrive()")
            Return True
        Catch ex As Exception
            WriteToLogFile(ex.Message, "UDO_Update_SelfDrive() for RA " & RADocentry)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR - AE_Sbooking", ex.Message)
            MsgBox(ex.Message)
            Return False
        End Try

    End Function

    Public Function UDO_Update_SelfDriveST(ByVal DocEntry As String, ByVal InvoiceAmount As Double, ByVal InvoiceDocEntry As String, ByVal InvoiceDocNum As String) As Boolean
        Try
            Dim oGeneralService As SAPbobsCOM.GeneralService = Nothing
            Dim oGeneralData As SAPbobsCOM.GeneralData = Nothing
            Dim oGeneralDataParam As SAPbobsCOM.GeneralDataParams = Nothing
            Dim oCompanyService As SAPbobsCOM.CompanyService = Ocompany_SD.GetCompanyService
            Dim RANumber As String
            Dim orset As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset.DoQuery("SELECT T0.[DocEntry] FROM [dbo].[@AE_SBOOKING]  T0 WHERE T0.[DocNum] = '" & Trim(DocEntry.ToString.Substring(3, DocEntry.ToString.Length - 3)) & "'")

            RANumber = orset.Fields.Item("DocEntry").Value
            oGeneralService = oCompanyService.GetGeneralService("AE_Sbooking")
            oGeneralDataParam = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oGeneralDataParam.SetProperty("DocEntry", RANumber)
            oGeneralData = oGeneralService.GetByParams(oGeneralDataParam)
            ' Document Fields to update
            'oGeneralService.Update(oGeneralData)
            oGeneralData.SetProperty("U_AE_Status", "Closed")
            Dim oChildTableRows As SAPbobsCOM.GeneralDataCollection = oGeneralData.Child("AE_SBOOKING_R1")
            '   Dim LineNum As Integer = 0
            ' Document Line Fields to update
            Dim oChildTableRow As SAPbobsCOM.GeneralData
            Try
                ' MsgBox(orset.Fields.Item("LineID").Value)

                orset.DoQuery("SELECT max(cast(T0.[LineId] as integer)) as 'LineID', T0.[U_AE_Ino] FROM [dbo].[@AE_SBOOKING_R1]  T0 WHERE T0.[DocEntry] = '" & RANumber & "' GROUP BY T0.[U_AE_Ino]")

                If orset.Fields.Item("U_AE_Ino").Value = "" Then
                    If orset.Fields.Item("LineID").Value = 0 Then
                        oChildTableRow = oChildTableRows.Add
                        ' oChildTableRow.SetProperty("LineID", orset.Fields.Item("LineID").Value)
                        oChildTableRow.SetProperty("U_AE_Idate", Server_Date)
                        oChildTableRow.SetProperty("U_AE_Amoun", InvoiceAmount)
                        oChildTableRow.SetProperty("U_AE_Ino", InvoiceDocEntry)
                        oChildTableRow.SetProperty("U_AE_Idocn", InvoiceDocNum)
                    Else
                        oChildTableRow = oChildTableRows.Item(orset.Fields.Item("LineID").Value - 1)
                        If oChildTableRow.GetProperty("LineId") = orset.Fields.Item("LineID").Value Then 'LINENUM_MACTHING
                            oChildTableRow.SetProperty("U_AE_Idate", Server_Date)
                            oChildTableRow.SetProperty("U_AE_Amoun", InvoiceAmount)
                            oChildTableRow.SetProperty("U_AE_Ino", InvoiceDocEntry)
                            oChildTableRow.SetProperty("U_AE_Idocn", InvoiceDocNum)
                            ' oGeneralService.Update(oGeneralData)
                        End If
                    End If

                Else
                    oChildTableRow = oChildTableRows.Add
                    ' oChildTableRow.SetProperty("LineID", orset.Fields.Item("LineID").Value)
                    oChildTableRow.SetProperty("U_AE_Idate", Server_Date)
                    oChildTableRow.SetProperty("U_AE_Amoun", InvoiceAmount)
                    oChildTableRow.SetProperty("U_AE_Ino", InvoiceDocEntry)
                    oChildTableRow.SetProperty("U_AE_Idocn", InvoiceDocNum)
                End If
                'oApplication.StatusBar.SetText("Updating your information in Budget Proposal object ......... " & mj, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Catch ex As Exception
                MsgBox(ex.Message)
                Return False
            End Try
            oGeneralService.Update(oGeneralData)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try

    End Function


    Public Sub ShowFolderBrowser()

        Dim MyProcs() As System.Diagnostics.Process
        FileName = ""
        Dim OpenFile As New OpenFileDialog
        ' Dim stFilePathAndName As String
        Dim orset As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        orset.DoQuery("SELECT attachpath from oadp")

        Try
            OpenFile.Multiselect = False

            OpenFile.Filter = "All files (*.*)|*.*|All files (*.*)|*.*" '"All files(*.CSV)|*.CSV"
            Dim filterindex As Integer = 0
            Try
                filterindex = 0
            Catch ex As Exception
            End Try

            OpenFile.FilterIndex = filterindex
            OpenFile.InitialDirectory = "C:\" ' orset.Fields.Item("C:\").Value
            OpenFile.RestoreDirectory = False
            MyProcs = System.Diagnostics.Process.GetProcessesByName("SAP Business One")

            ' If MyProcs.Length = 1 Then
            ' For i As Integer = 0 To MyProcs.Length - 1

            Dim MyWindow As New WindowWrapper(MyProcs(0).MainWindowHandle)
            Dim ret As DialogResult = OpenFile.ShowDialog(MyWindow)
            If ret = DialogResult.OK Then
                stFilePathAndName = OpenFile.FileName
                Dim MyFile As IO.FileInfo = New IO.FileInfo(stFilePathAndName)
                FileName = MyFile.Name

                ''FileName = OpenFile.FileName
                OpenFile.Dispose()
            Else
                FileName = ""
                System.Windows.Forms.Application.ExitThread()
            End If
            ' Next
            ' End If
        Catch ex As Exception
            'SBO_Application.StatusBar.SetText(ex.Message)
            MessageBox.Show(ex.ToString())
            FileName = ""
        Finally
            OpenFile.Dispose()
        End Try

    End Sub

    Private Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window
        Private _hwnd As System.IntPtr

        Public Sub New(ByVal handle As System.IntPtr)
            _hwnd = handle
        End Sub

        Private ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property
    End Class

    Private Sub Group_Invoice(ByRef oForm1 As SAPbouiCOM.Form, ByRef oMatrix As SAPbouiCOM.Matrix)

        Dim oCheck As SAPbouiCOM.CheckBox
        Dim DocEntry_ As String = String.Empty
        Dim GroupInvoice_Thread As System.Threading.Thread


        Try

            If oForm1.Items.Item("15").Specific.String = "" Then
                oForm1.Items.Item("15").Specific.active = True
                Oapplication_SD.StatusBar.SetText("Customer should not be empty for the group invoce  ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Exit Sub
            End If

            For mjs As Integer = 1 To oMatrix.RowCount
                oCheck = oMatrix.Columns.Item("Col_5m").Cells.Item(mjs).Specific
                If oCheck.Checked = True Then
                    DocEntry_ = DocEntry_ & oMatrix.Columns.Item("V_0mjd").Cells.Item(mjs).Specific.String & ","
                End If
            Next mjs

            DocEntry_ = DocEntry_.Substring(0, DocEntry_.Length - 1)
            If DocEntry_ = "" Then
                Oapplication_SD.StatusBar.SetText("No invoice has been selected to print ........ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Exit Sub

            Else
                GroupInvoice_Thread = New System.Threading.Thread(AddressOf Class_Report.GroupTaxInvoice_CallingFunction)
                Class_Report.oApplication = Oapplication_SD
                Class_Report.oCompany = Ocompany_SD

                Class_Report.Report_Name = "AE_RP008_SDGroupTaxInvoice.rpt"
                Class_Report.Report_Parameter = "@DocKey"
                Class_Report.Report_Title = "Group Invoice"

                Class_Report.DocKey = DocEntry_

                If GroupInvoice_Thread.IsAlive Then
                    Oapplication_SD.MessageBox("Report is already open....")
                Else
                    GroupInvoice_Thread.TrySetApartmentState(Threading.ApartmentState.STA)
                    Oapplication_SD.StatusBar.SetText("Group Invoice Report Opening in process ......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    GroupInvoice_Thread.Start()
                End If
            End If

        Catch ex As Exception
            GroupInvoice_Thread.Abort()
            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Exit Try
        End Try

    End Sub

    Private Sub TaxInvoice(ByRef oForm1 As SAPbouiCOM.Form, ByRef oMatrix As SAPbouiCOM.Matrix, ByVal sPrinter As String)

        Dim oCheck As SAPbouiCOM.CheckBox
        Dim DocEntry_ As String = String.Empty
        Dim Invoice_Thread As System.Threading.Thread


        Try
            ''Dim printerSettings As New PrinterSettings()
            ''Dim printDialog As New PrintDialog()
            ''printDialog.PrinterSettings = printerSettings
            ''printDialog.AllowPrintToFile = False
            ''printDialog.AllowSomePages = True
            ''printDialog.UseEXDialog = True

            ''Dim result As DialogResult = printDialog.ShowDialog()

            ''If result = DialogResult.Cancel Then
            ''    Oapplication_SD.MessageBox("No Printer was selected ......... !", 1, "Ok")
            ''    Exit Sub
            ''End If

            For mjs As Integer = 1 To oMatrix.RowCount
                oCheck = oMatrix.Columns.Item("Col_5m").Cells.Item(mjs).Specific
                If oCheck.Checked = True Then
                    DocEntry_ = DocEntry_ & oMatrix.Columns.Item("V_0mjd").Cells.Item(mjs).Specific.String & ","
                End If
            Next mjs

            DocEntry_ = DocEntry_.Substring(0, DocEntry_.Length - 1)
            If DocEntry_ = "" Then
                Oapplication_SD.StatusBar.SetText("No invoice has been selected to Display ........ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Exit Sub
            Else
                Invoice_Thread = New System.Threading.Thread(AddressOf Class_Report.TaxInvoice_CallingFunction)
                Class_Report.oApplication = Oapplication_SD
                Class_Report.oCompany = Ocompany_SD
                '  Class_Report.Report_Name = "AE_RP007_SDTaxInvoice_Batch.rpt"
                Class_Report.Report_Name = "AE_FRM01_PrePrint_TaxInvoice_SD.rpt" '"AE_FRM01_PrePrint_TaxInvoice_SD_Original.rpt"
                Class_Report.Report_Parameter = "DocKey@"
                Class_Report.Report_Title = "Tax Invoice"
                Class_Report.DocKey = DocEntry_
                Class_Report.PrinterName = sPrinter  '"Invoice Print" 'printerSettings.PrinterName

                Oapplication_SD.StatusBar.SetText("Tax Invoice Report in process ......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If Invoice_Thread.IsAlive Then
                    Oapplication_SD.MessageBox("Report is already open....")
                Else
                    Invoice_Thread.TrySetApartmentState(Threading.ApartmentState.STA)
                    Invoice_Thread.Start()
                    ''' Invoice_Thread.Join()
                End If

            End If

        Catch ex As Exception
            Invoice_Thread.Abort()
            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Exit Try
        End Try

    End Sub


End Class
