Imports System.IO

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

                End If


                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
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
                                oMatrix.Columns.Item("Col_0").Cells.Item(oMatrix.RowCount).Specific.string = Format(Now.Date, "dd/MM/yyyy")
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
                                oForm.Items.Item("Item_42").Specific.string = Format(oDataTable.GetValue("U_AE_DOB", 0), "dd/MM/yyyy")
                                oForm.Items.Item("Item_44").Specific.string = oDataTable.GetValue("U_AE_LicenseNo", 0)
                                oForm.Items.Item("Item_49").Specific.string = oDataTable.GetValue("U_AE_Lplace", 0)
                                oForm.Items.Item("Item_50").Specific.string = Format(oDataTable.GetValue("U_AE_LEdate", 0), "dd/MM/yyyy")
                                oForm.Items.Item("Item_51").Specific.string = oDataTable.GetValue("U_AE_Pno", 0)
                                oForm.Items.Item("Item_52").Specific.string = oDataTable.GetValue("U_AE_Pplace", 0)
                                oForm.Items.Item("Item_54").Specific.string = Format(oDataTable.GetValue("U_AE_PEdate", 0), "dd/MM/yyyy")
                                oForm.Items.Item("Item_30").Specific.string = oDataTable.GetValue("U_AE_Dcode", 0)

                            ElseIf pVal.ItemUID = "Item_57" Then 'Issued By
                                oForm.Items.Item("Item_58").Specific.string = oDataTable.GetValue("U_AE_Dname", 0)
                                oForm.Items.Item("Item_60").Specific.string = oDataTable.GetValue("U_AE_Ladd", 0)
                                oForm.Items.Item("Item_62").Specific.string = oDataTable.GetValue("U_AE_Hphone", 0)
                                oForm.Items.Item("Item_64").Specific.string = oDataTable.GetValue("U_AE_Occ", 0)
                                oForm.Items.Item("Item_66").Specific.string = oDataTable.GetValue("U_AE_COB", 0)
                                oForm.Items.Item("Item_68").Specific.string = Format(oDataTable.GetValue("U_AE_DOB", 0), "dd/MM/yyyy")
                                oForm.Items.Item("Item_70").Specific.string = oDataTable.GetValue("U_AE_LicenseNo", 0)
                                oForm.Items.Item("Item_75").Specific.string = oDataTable.GetValue("U_AE_Lplace", 0)
                                oForm.Items.Item("Item_76").Specific.string = Format(oDataTable.GetValue("U_AE_LEdate", 0), "dd/MM/yyyy")
                                oForm.Items.Item("Item_77").Specific.string = oDataTable.GetValue("U_AE_Pno", 0)
                                oForm.Items.Item("Item_78").Specific.string = oDataTable.GetValue("U_AE_Pplace", 0)
                                oForm.Items.Item("Item_80").Specific.string = Format(oDataTable.GetValue("U_AE_PEdate", 0), "dd/MM/yyyy")
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
                        Dim iPAI, iCDW As Double
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

                                oMAtrix_IN.Columns.Item("1").Cells.Item(1).Specific.String = "Self Drive Billing for this Booking No : " & oform1.Items.Item("Item_14").Specific.String
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


                                Dim ocombo As SAPbouiCOM.ComboBox
                                If ContactP <> "" Then
                                    ocombo = oform.Items.Item("85").Specific
                                    ocombo.Select(ContactP)
                                End If

                                ' oMAtrix_IN.Columns.Item("95").Cells.Item(1).Specific.String = "SO"
                                'Dim ocombo As SAPbouiCOM.ComboBox = oform.Items.Item("20").Specific
                                'ocombo.Select(salesemp)
                                oform.Items.Item("16").Specific.String = "Self Drive Invoice Based on Booking No : " & oform1.Items.Item("Item_14").Specific.String
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
                                oform.Items.Item("Item_16").Specific.String = Now.Date 'Format(Now.Date, "dd MMM yyyy") 'Now.Date
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

                            Dim opt As SAPbouiCOM.OptionBtn = oform.Items.Item("195").Specific ' Long Term
                            Dim opt1 As SAPbouiCOM.OptionBtn = oform.Items.Item("196").Specific

                            If oform.Items.Item("226").Specific.value.ToString.Trim = "Yes" Then
                                oform.Items.Item("Item_115").Enabled = True
                            Else
                                oform.Items.Item("Item_115").Enabled = True
                            End If
                            oform.Items.Item("201").Enabled = True
                            If opt.Selected = True Then
                                'oform.Items.Item("201").Enabled = False
                            Else
                                If oform.Items.Item("Item_18").Specific.value.ToString.Trim = "Billing" Then
                                    'oform.Items.Item("201").Enabled = True
                                    oform.Items.Item("Item_18").Enabled = True
                                ElseIf oform.Items.Item("Item_18").Specific.value.ToString.Trim = "Closed" Then
                                    ' oform.Items.Item("201").Enabled = False
                                    oform.Items.Item("Item_18").Enabled = False
                                Else
                                    ' oform.Items.Item("201").Enabled = False
                                    oform.Items.Item("Item_18").Enabled = True
                                End If


                                oform.Items.Item("Item_14").Enabled = False
                            End If
                        Catch ex As Exception
                            Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Try
                        End Try

                    End If

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
                            oform.Items.Item("Item_6").Specific.columns.item("V_1INV").databind.bind("@AE_SBOOKING", "DocNum")
                            oform.Items.Item("Item_6").Specific.columns.item("BNO").databind.bind("@AE_SBOOKING", "NumAtCard")
                            oform.Items.Item("Item_6").Specific.columns.item("Col_0m").databind.bind("@AE_SBOOKING", "CardName")
                            oform.Items.Item("Item_6").Specific.columns.item("Col_4mjs").databind.bind("@AE_SBOOKING", "DocTotal")
                            oform.Items.Item("Item_6").Specific.columns.item("V_0mj").databind.bind("@AE_SBOOKING", "DocDate")
                            oform.Items.Item("Item_6").Specific.columns.item("V_0mjd").databind.bind("@AE_SBOOKING", "DocEntry")

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
                        Dim Invoice_Thread As System.Threading.Thread

                        Try
                            oform1 = Oapplication_SD.Forms.ActiveForm

                            Dim oMAtrix As SAPbouiCOM.Matrix = oform1.Items.Item("Item_6").Specific
                            Dim DocEntry_ As String
                            Dim oCheck As SAPbouiCOM.CheckBox
                            Dim oCheck1 As SAPbouiCOM.CheckBox = oform1.Items.Item("17").Specific

                            If oCheck1.Checked = True Then
                                If oform1.Items.Item("15").Specific.String = "" Then
                                    oform1.Items.Item("15").Specific.active = True
                                    Oapplication_SD.StatusBar.SetText("Customer should not be empty for the group invoce  ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                            For mjs As Integer = 1 To oMAtrix.RowCount
                                oCheck = oMAtrix.Columns.Item("Col_5m").Cells.Item(mjs).Specific
                                If oCheck.Checked = True Then
                                    DocEntry_ = DocEntry_ & oMAtrix.Columns.Item("V_0mjd").Cells.Item(mjs).Specific.String & ","
                                End If
                            Next mjs

                            DocEntry_ = DocEntry_.Substring(0, DocEntry_.Length - 1)
                            If DocEntry_ = "" Then
                                Oapplication_SD.StatusBar.SetText("No invoice has been selected to print ........ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Sub

                            Else
                                Invoice_Thread = New System.Threading.Thread(AddressOf Class_Report.TaxInvoice_CallingFunction)
                                Class_Report.oApplication = Oapplication_SD
                                Class_Report.oCompany = Ocompany_SD
                                If oCheck1.Checked = True Then
                                    Class_Report.Report_Name = "AE_RP008_SDGroupTaxInvoice.rpt"
                                    Class_Report.Report_Parameter = "@DocKey"
                                    Class_Report.Report_Title = "Group Invoice"
                                Else
                                    Class_Report.Report_Name = "AE_RP007_SDTaxInvoice_Batch.rpt"
                                    Class_Report.Report_Parameter = "@DocKey"
                                    Class_Report.Report_Title = "Tax Invoice"
                                End If

                                Class_Report.DocKey = DocEntry_

                                If Invoice_Thread.IsAlive Then
                                    Oapplication_SD.MessageBox("Report is already open....")
                                Else
                                    Invoice_Thread.TrySetApartmentState(Threading.ApartmentState.STA)
                                    Oapplication_SD.StatusBar.SetText("Tax Invoice Report Opening in process ......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    Invoice_Thread.Start()
                                End If


                            End If

                        Catch ex As Exception
                            Invoice_Thread.Abort()
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
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.ItemUID = "Item_5" Then
                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                            Dim orset As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim Tmp_Date, Tmp_Date1 As String
                            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_6").Specific

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

                            Dim Str_query As String = ""


                            Str_query = "SELECT T0.[DocNum], T0.[U_AE_Bname], T0.[U_AE_Bcode], T0.[U_AE_Vregno], T0.[U_AE_Rno], T0.[U_AE_Contract], " & _
"case when month(T0.[U_AE_Vdatein]) = '" & oform.Items.Item("12").Specific.selected.description & "' then (T0.[U_AE_Rate]/30) * day(T0.[U_AE_Vdatein] ) " & _
"else T0.[U_AE_rate] end as 'U_AE_rate', T0.[U_AE_NXDT], T0.[Docentry] " & _
"FROM [dbo].[@AE_SBOOKING]  T0  , OINV T1 " & _
"WHERE  DAY(T0.[U_AE_NXDT]) >= '" & oform.Items.Item("Item_1").Specific.String & "' and DAY(T0.[U_AE_NXDT]) <= '" & oform.Items.Item("Item_3").Specific.String & "' and T0.[U_AE_Status] = 'Open' and " & _
"T0.DocNum not in (select rtrim(substring(TT.NumAtCard , 3 , len(Tt.NumAtCard) -2)) from OINV tt where month(tt.DocDate) = '" & oform.Items.Item("12").Specific.selected.description & "' and left(TT.NumAtCard,2) = 'SD'  ) " & _
"group by T0.[DocNum], T0.[U_AE_Bname], " & _
"T0.[U_AE_Bcode], T0.[U_AE_Vregno], T0.[U_AE_Rno], T0.[U_AE_Contract], T0.[U_AE_NXDT], T0.[U_AE_Vdatein],T0.[U_AE_Rate], T0.[DocEntry] "
                            'and month(t0.[U_AE_NXDT]) <= '" & oform.Items.Item("12").Specific.selected.description & "' " & _

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
                            oform.Items.Item("Item_6").Specific.columns.item("V_0sjr").databind.bind("@AE_SBOOKING", "U_AE_Bcode")
                            oform.Items.Item("Item_6").Specific.columns.item("V_0DOC").databind.bind("@AE_SBOOKING", "Docentry")

                            oform.Items.Item("Item_6").Specific.LoadFromDataSource()
                            oform.Items.Item("Item_6").Specific.AutoResizeColumns()
                            ' oMatrix.Columns.Item("Col_5").Editable = True


                        Catch ex As Exception
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
                            For mjs As Integer = 1 To oMAtrix.RowCount
                                oCheck = oMAtrix.Columns.Item("Col_5m").Cells.Item(mjs).Specific
                                If oCheck.Checked = True Then
                                    ' MsgBox(oMAtrix.Columns.Item("BNO").Cells.Item(mjs).Specific.String)
                                    ' MsgBox(System.DateTime.Parse(oMAtrix.Columns.Item("V_0mj").Cells.Item(mjs).Specific.String, format1, System.Globalization.DateTimeStyles.None))
                                    sFunName = "SD_InvoiceGeneration() - For RA " & oMAtrix.Columns.Item("BNO").Cells.Item(mjs).Specific.String
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFunName)

                                    If SD_InvoiceGeneration(oMAtrix.Columns.Item("V_0sjr").Cells.Item(mjs).Specific.String, _
                                        oMAtrix.Columns.Item("BNO").Cells.Item(mjs).Specific.String, "Self Drive Billing for this Booking No : " & oMAtrix.Columns.Item("BNO").Cells.Item(mjs).Specific.String, _
                                       SD_GLACC, "SO", oMAtrix.Columns.Item("Col_4mjs").Cells.Item(mjs).Specific.String, "Self Drive Invoice Based on Booking No : " & oMAtrix.Columns.Item("BNO").Cells.Item(mjs).Specific.String, _
                                        oMAtrix.Columns.Item("V_0mj").Cells.Item(mjs).Specific.String, oform1.Items.Item("12").Specific.selected.description, oMAtrix.Columns.Item("V_0DOC").Cells.Item(mjs).Specific.String) = False Then
                                        BubbleEvent = False
                                        Exit Try
                                    End If
                                    sFunName = "SD_InvoiceGeneration() - For RA " & oMAtrix.Columns.Item("BNO").Cells.Item(mjs).Specific.String
                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed With SUCCESS", sFunName)
                                End If
                            Next mjs
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
                Try
                    oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                    oform.Freeze(True)
                    Dim Tmp_val As String = oform.BusinessObject.GetNextSerialNumber(NextSerialNo(Ocompany_SD, Oapplication_SD, "AE_Sbooking"))
                    oform.Items.Item("Item_14").Specific.String = Tmp_val
                    oform.Items.Item("Item_16").Specific.String = Now.Date ' Format(Now.Date, "dd MMM yyyy") 'Format(Now.Date, "dd MMM yyyy, ddd")
                    Dim ocombo As SAPbouiCOM.ComboBox
                    ocombo = oform.Items.Item("226").Specific
                    ocombo.Select("No")
                    ocombo = oform.Items.Item("228").Specific
                    ocombo.Select("No")
                    oform.Items.Item("Item_115").Enabled = False

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

                LoadFromXML("SelfDriving_Billing.srf", Oapplication_SD)
                Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.Item("SDBIL")
                oform.Visible = True
                oform.Items.Item("Item_10").Specific.String = Now.Date
                Dim oCombo As SAPbouiCOM.ComboBox = oform.Items.Item("12").Specific
                ' MsgBox(MonthName(Month(Now) - 1))

                oCombo.ValidValues.Add(MonthName(Month(Now) - 1), Now.Month - 1)
                oCombo.ValidValues.Add(MonthName(Month(Now)), Now.Month)
                oCombo.ValidValues.Add(MonthName(Month(Now) + 1), Now.Month + 1)
                oCombo.Select(MonthName(Month(Now)), SAPbouiCOM.BoSearchKey.psk_ByValue)
                Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_6").Specific
                oform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
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
                    LoadFromXML("SelfDriving_InvoiceReport.srf", Oapplication_SD)
                    Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.Item("SDINV")
                    oform.Items.Item("Item_10").Specific.String = Now.Date
                    Dim oCombo As SAPbouiCOM.ComboBox = oform.Items.Item("12").Specific
                    ' MsgBox(MonthName(Month(Now) - 1))

                    oCombo.ValidValues.Add(MonthName(Month(Now) - 1), Now.Month - 1)
                    oCombo.ValidValues.Add(MonthName(Month(Now)), Now.Month)
                    oCombo.ValidValues.Add(MonthName(Month(Now) + 1), Now.Month + 1)
                    oCombo.Select(MonthName(Month(Now)), SAPbouiCOM.BoSearchKey.psk_ByValue)
                    oform.Visible = True

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

                    End If

                Catch ex As Exception

                End Try
            End If

            If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.BeforeAction = False Then
                Try
                    Dim oform As SAPbouiCOM.Form = Oapplication_SD.Forms.ActiveForm
                    If oform.UniqueID = "SDB" Then

                        Try

                            Dim opt As SAPbouiCOM.OptionBtn = oform.Items.Item("195").Specific
                            Dim opt1 As SAPbouiCOM.OptionBtn = oform.Items.Item("196").Specific
                            Dim oopt As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_104").Specific
                            Dim oopt1 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_105").Specific
                            Dim oopt3 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_106").Specific
                            Dim orset As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim ocombo As SAPbouiCOM.ComboBox
                            Dim sAttention As String

                            oform.Freeze(True)

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

                            oform.Items.Item("201").Enabled = True
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

                            If opt.Selected = True Then
                                oform.Items.Item("210").Enabled = False
                                'oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                If Trim(oform.Items.Item("Item_18").Specific.value) = "Open" Then
                                    oform.Items.Item("Item_18").Enabled = True
                                Else
                                    oform.Items.Item("Item_18").Enabled = False
                                End If
                            Else
                                If Trim(oform.Items.Item("Item_18").Specific.value) = "Billing" Then
                                    oform.Items.Item("Item_2").Specific.active = True
                                    oform.Items.Item("Item_18").Enabled = False
                                    oform.Items.Item("210").Enabled = True
                                    ''oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                ElseIf Trim(oform.Items.Item("Item_18").Specific.value) <> "Open" Then
                                    oform.Items.Item("Item_18").Enabled = False
                                ElseIf Trim(oform.Items.Item("Item_18").Specific.value) = "Open" Then

                                    oform.Items.Item("Item_2").Specific.active = True
                                    oform.Items.Item("210").Enabled = False
                                    oform.Items.Item("Item_18").Enabled = True
                                    'oform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                End If

                            End If

                            If oopt.Selected = True Then
                                oform.Items.Item("Item_108").Specific.caption = "Daily Rates"
                                oform.Items.Item("Item_110").Specific.caption = "Number of Days"
                                oform.Items.Item("Item_114").Specific.caption = "PAI Per Day"
                                oform.Items.Item("Item_122").Specific.caption = "CDW Per Day"
                            ElseIf oopt1.Selected = True Then
                                oform.Items.Item("Item_108").Specific.caption = "Weekly Rates"
                                oform.Items.Item("Item_110").Specific.caption = "Number of Days"
                                oform.Items.Item("Item_114").Specific.caption = "PAI Per Day"
                                oform.Items.Item("Item_122").Specific.caption = "CDW Per Day"
                            ElseIf oopt3.Selected = True Then
                                oform.Items.Item("Item_108").Specific.caption = "Monthly Rates"
                                oform.Items.Item("Item_110").Specific.caption = "Number of Months"
                                oform.Items.Item("Item_114").Specific.caption = "PAI Per Month"
                                oform.Items.Item("Item_122").Specific.caption = "CDW Per Month"
                            End If

                            ' oform.PaneLevel = 3
                            oform.Items.Item("Item_22").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oform.Items.Item("233").Enabled = True
                            oform.Freeze(False)
                        Catch ex As Exception
                            oform.Freeze(False)
                        End Try
                    End If
                Catch ex As Exception
                    Oapplication_SD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try
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
                                          ByVal comments As String, ByVal ADate As String, ByVal month_ As Integer, ByVal DocEntry_ As String) As Boolean
        Dim sFunName As String = String.Empty
        Try
            Dim oInvoice As SAPbobsCOM.Documents = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fetching the information for RA ", NumAtCard)

            Dim val As Integer = 0
            Dim InvoiceDocEntry, InvoiceDocNum As String
            Dim orset As SAPbobsCOM.Recordset = Ocompany_SD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset.DoQuery("SELECT T0.[U_AE_CDW], T0.[U_AE_CDW1], T0.[U_AE_PAI],T0.[U_AE_PAI1] FROM [dbo].[@AE_SBOOKING]  T0 WHERE T0.[DocEntry]  = '" & DocEntry_ & "'")
            Dim sCDWFlag As String = orset.Fields.Item("U_AE_CDW1").Value
            Dim dCDW As Double = orset.Fields.Item("U_AE_CDW").Value
            Dim sPAIFlag As String = orset.Fields.Item("U_AE_PAI1").Value
            Dim dPAI As Double = orset.Fields.Item("U_AE_PAI").Value

            orset.DoQuery("SELECT T0.[Code] FROM OLCT T0")
            Dim _tmpdate As Date = Convert.ToDateTime(ADate)
            oInvoice.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
            'MsgBox(month_)
            'MsgBox(_tmpdate.Day)
            oInvoice.CardCode = CardCode
            oInvoice.NumAtCard = "SD " & NumAtCard
            
            oInvoice.DocDate = PostDate(_tmpdate.Day, month_, Now.Year, Oapplication_SD)
            oInvoice.TaxDate = PostDate(_tmpdate.Day, month_, Now.Year, Oapplication_SD) 'Server_Date
            oInvoice.Lines.ItemDescription = Description
            oInvoice.Lines.AccountCode = AcountCode
            oInvoice.Lines.TaxCode = Taxcode
            oInvoice.Lines.LineTotal = LineTotal
            'MsgBox(orset.Fields.Item("Code").Value)
            oInvoice.Lines.LocationCode = "1"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice - Cardcode ", CardCode)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invocie - DocDate ", _tmpdate.Day & "/" & month_ & "/" & Now.Year)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice - Description ", Description)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice - Account Code", AcountCode)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice - Taxcode ", Taxcode)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice - Line Total  ", LineTotal)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CDW Flag status ", sCDWFlag)

            If sCDWFlag = "Yes" And CInt(dCDW) > 0 Then
                oInvoice.Lines.Add()
                oInvoice.Lines.ItemDescription = "Collission Damage Waiver Fee"
                oInvoice.Lines.TaxCode = Taxcode
                oInvoice.Lines.LineTotal = dCDW
                oInvoice.Lines.AccountCode = CDW_GLACC
                oInvoice.Lines.LocationCode = "1"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fetching CDW information for RA ", NumAtCard)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice - ItemDescription ", "Collission Damage Waiver Fee")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invocie - TaxCode ", Taxcode)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice - LineTotal ", dCDW)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice - Account Code", CDW_GLACC)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("PAI Flag status ", sPAIFlag)

            If sPAIFlag = "Yes" And CInt(dPAI) > 0 Then
                oInvoice.Lines.Add()
                oInvoice.Lines.ItemDescription = "Personal Accidental Insurance"
                oInvoice.Lines.TaxCode = Taxcode
                oInvoice.Lines.LineTotal = dPAI
                oInvoice.Lines.AccountCode = PAI_GLACC
                oInvoice.Lines.LocationCode = "1"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fetching PAI information for RA ", NumAtCard)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice - ItemDescription ", "Personal Accidental Insurance")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invocie - TaxCode ", Taxcode)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice - LineTotal ", dPAI)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice - Account Code", PAI_GLACC)
            End If

            oInvoice.Comments = comments

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

                If UDO_Update_SelfDrive(NumAtCard, ADate, LineTotal, InvoiceDocEntry, orset.Fields.Item("DocNum").Value, DocEntry_, Now.Year & "/" & month_ & "/" & _tmpdate.Day, orset.Fields.Item("DocDate").Value) = False Then
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
            Ocompany_SD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            MsgBox(ex.Message)
            Return False
        End Try
    End Function


    Public Function UDO_Update_SelfDrive(ByVal DocNum As String, ByVal NX_Date As String, ByVal InvoiceAmount As Double, ByVal InvoiceDocEntry As String, ByVal InvoiceDocNum As String, ByVal RADocentry As String, ByVal NInvDateNew As Date, ByVal DocDate As Date) As Boolean
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

            If orset1.Fields.Item("U_AE_Vdatein").Value > DateConvertion.AddMonths(1) Or CDate(orset1.Fields.Item("U_AE_Vdatein").Value).ToString("yyyyMMdd") = "19000101" Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("NInvDateNew.AddMonths(1) ", Format(NInvDateNew.AddMonths(1), "dd/MM/yyyy"))
                oGeneralData.SetProperty("U_AE_NXDT", NInvDateNew.AddMonths(1))
            Else
                oGeneralData.SetProperty("U_AE_Status", "Closed")
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
                    oChildTableRow.SetProperty("U_AE_Idate", Server_Date)
                    oChildTableRow.SetProperty("U_AE_Amoun", InvoiceAmount)
                    oChildTableRow.SetProperty("U_AE_Ino", InvoiceDocEntry)
                    oChildTableRow.SetProperty("U_AE_Idocn", InvoiceDocNum)
                Else
                    oChildTableRow = oChildTableRows.Item(orset.Fields.Item("LineID").Value - 1)
                    If oChildTableRow.GetProperty("LineId") = orset.Fields.Item("LineID").Value Then 'LINENUM_MACTHING
                        oChildTableRow.SetProperty("U_AE_Adate", DocDate)
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
                oChildTableRow.SetProperty("U_AE_Idate", Server_Date)
                oChildTableRow.SetProperty("U_AE_Amoun", InvoiceAmount)
                oChildTableRow.SetProperty("U_AE_Ino", InvoiceDocEntry)
                oChildTableRow.SetProperty("U_AE_Idocn", InvoiceDocNum)
            End If
            'oApplication.StatusBar.SetText("Updating your information in Budget Proposal object ......... " & mj, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fetching the line items for RA ", RADocentry)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice Date", Format(DocDate, "dd/MM/yyyy"))
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice raised date ", Format(Server_Date, "dd/MM/yyyy"))
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice Amount ", InvoiceAmount)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice Doc Num ", InvoiceDocNum)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice Docentry ", InvoiceDocEntry)


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


End Class
