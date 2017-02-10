
Imports System.IO

Public Class ServiceMaintenance

    Dim WithEvents Oapplication_SM As SAPbouiCOM.Application
    Dim Ocompany_SM As New SAPbobsCOM.Company
    Private FileName As String


    Public Sub New(ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company)

        Oapplication_SM = oApplication
        Ocompany_SM = oCompany

    End Sub

    Private Sub Oapplication_SM_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles Oapplication_SM.FormDataEvent
        If BusinessObjectInfo.FormUID = "SM" Then
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then

                Try
                    Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm
                    Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_13").Specific

                    Dim orset As SAPbobsCOM.Recordset = Ocompany_SM.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim Mileage As String

                    For mjs As Integer = 1 To oMAtrix.RowCount
                        If oMAtrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.value = "General Service" Then
                            Mileage = Trim(oform.DataSources.DBDataSources.Item(0).GetValue("U_AE_Vmileage", 0))
                            orset.DoQuery("update oitm set [U_AE_SDate] = '" & GateDate(oMAtrix.Columns.Item("Col_0").Cells.Item(mjs).Specific.String, Ocompany_SM) & "' , [U_AE_VSKM] = '" & Mileage & "' where [ItemCode] = '" & Trim(oform.DataSources.DBDataSources.Item(0).GetValue("U_AE_Vno", 0)) & "'")
                        End If
                        If oMAtrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.value = "Battery" Then
                            orset.DoQuery("update oitm set [U_AE_BDate] = '" & GateDate(oMAtrix.Columns.Item("Col_0").Cells.Item(mjs).Specific.String, Ocompany_SM) & "' where [ItemCode] = '" & Trim(oform.DataSources.DBDataSources.Item(0).GetValue("U_AE_Vno", 0)) & "'")
                        End If
                        If oMAtrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.value = "Tyre" Then
                            Mileage = Trim(oform.DataSources.DBDataSources.Item(0).GetValue("U_AE_Vmileage", 0))
                            orset.DoQuery("update oitm set [U_AE_TKM] = '" & Mileage & "' where [ItemCode] = '" & Trim(oform.DataSources.DBDataSources.Item(0).GetValue("U_AE_Vno", 0)) & "'")
                        End If
                    Next mjs
                Catch ex As Exception
                    Oapplication_SM.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End Try

            End If
        End If
    End Sub

    Private Sub Oapplication_SM_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles Oapplication_SM.ItemEvent

        If pVal.FormUID = "SM" Then

            If pVal.Before_Action = False Then

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then

                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvento = pVal
                    Dim sCFL_ID As String
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    Dim oForm As SAPbouiCOM.Form = Oapplication_SM.Forms.Item(FormUID)
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                    Try

                        If oCFLEvento.BeforeAction = False Then
                            Dim oDataTable As SAPbouiCOM.DataTable
                            Dim oedit1 As SAPbouiCOM.EditText
                            oDataTable = oCFLEvento.SelectedObjects
                            If pVal.ItemUID = "Item_1" Then 'Billing Code
                                oForm.Items.Item("17").Specific.string = oDataTable.GetValue("ItemName", 0)
                                Try
                                    oForm.Items.Item("Item_5").Specific.string = oDataTable.GetValue("U_AE_RKM", 0)
                                Catch ex As Exception
                                End Try
                                oForm.Items.Item("Item_1").Specific.string = oDataTable.GetValue("ItemCode", 0)
                            End If

                            If pVal.ItemUID = "Item_13" And pVal.ColUID = "Col_4" Then
                                Dim omatrix As SAPbouiCOM.Matrix = oForm.Items.Item("Item_13").Specific
                                omatrix.Columns.Item("Col_5").Cells.Item(pVal.Row).Specific.String = oDataTable.GetValue("CardName", 0)
                                omatrix.Columns.Item("Col_4").Cells.Item(pVal.Row).Specific.String = oDataTable.GetValue("CardCode", 0)
                            End If

                        End If
                    Catch ex As Exception
                        ' Oapplication_SM.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        BubbleEvent = False
                    End Try
                    Exit Sub
                End If

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm
                        Dim ocombo As SAPbouiCOM.ComboBox = oform.Items.Item("23").Specific
                        Dim oButton As SAPbouiCOM.ButtonCombo = oform.Items.Item("21").Specific
                        Dim orset As SAPbobsCOM.Recordset = Ocompany_SM.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        If pVal.Action_Success = True Then
                            Dim Tmp_val As String = oform.BusinessObject.GetNextSerialNumber(NextSerialNo(Ocompany_SM, Oapplication_SM, "AE_SM"))
                            oform.Items.Item("Item_7").Specific.String = Tmp_val
                            oform.Items.Item("Item_9").Specific.String = Format(Now.Date, "yyyyMMdd") ' Now.Date 'Format(Now.Date, "dd MMM yyyy")
                            oform.Items.Item("1000002").Specific.String = Ocompany_SM.UserName
                            oform.Items.Item("24").Specific.String = Company_Name
                            Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_13").Specific
                            omatrix.Columns.Item("Col_0").Editable = True
                            omatrix.AddRow()
                            omatrix.Columns.Item("#").Cells.Item(omatrix.RowCount).Specific.String = omatrix.RowCount
                            ocombo.Select("Open")
                            oform.PaneLevel = 1
                            oform.DataBrowser.BrowseBy = "Item_7"
                        End If
                        Exit Sub
                    End If

                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm
                        Try
                            If bSVH = False Then
                                If SM = True Then
                                    Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_13").Specific
                                    oform.DataSources.DBDataSources.Item(1).Clear()
                                    omatrix.AddRow()
                                    omatrix.Columns.Item("#").Cells.Item(omatrix.RowCount).Specific.String = omatrix.RowCount
                                Else
                                    SM = False
                                End If

                                If Trim(oform.Items.Item("23").Specific.value) = "Open" Then
                                    oform.Items.Item("21").Enabled = True
                                    oform.Items.Item("Item_13").Enabled = True
                                    oform.Items.Item("Item_1").Enabled = True
                                    oform.Items.Item("Item_5").Enabled = True
                                    oform.Items.Item("1").Enabled = True
                                    '.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                Else
                                    oform.Items.Item("21").Enabled = False
                                    oform.Items.Item("Item_13").Enabled = False
                                    oform.Items.Item("Item_1").Enabled = False
                                    oform.Items.Item("Item_5").Enabled = False
                                    bSVH = False
                                End If
                                oform.Items.Item("Item_7").Enabled = False
                            End If
                        Catch ex As Exception
                            Oapplication_SM.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                        End Try
                        Exit Sub
                    End If

                End If

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then
                    If pVal.ItemUID = "Item_13" And pVal.ColUID = "Col_2" Then
                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm
                            Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_13").Specific
                            If pVal.Row = omatrix.RowCount Then
                                If omatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.String <> "" And omatrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific.String <> "" Then
                                    omatrix.AddRow()
                                    omatrix.Columns.Item("#").Cells.Item(omatrix.RowCount).Specific.String = omatrix.RowCount
                                End If

                            End If
                        Catch ex As Exception
                            Oapplication_SM.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                        End Try
                        Exit Sub
                    End If

                    If pVal.ItemUID = "21" Then
                        Dim oform_In As SAPbouiCOM.Form
                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm

                            Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_13").Specific
                            Dim Scode As String = ""
                            Dim orset As SAPbobsCOM.Recordset = Ocompany_SM.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim ocheck As SAPbouiCOM.CheckBox
                            Dim oVendor As String = omatrix.Columns.Item("Col_4").Cells.Item(1).Specific.String
                            Dim oButtton As SAPbouiCOM.ButtonCombo = oform.Items.Item("21").Specific
                            Dim oVno As String = oform.Items.Item("Item_1").Specific.String
                            Dim Docnum As String = oform.Items.Item("Item_7").Specific.String
                            Dim VrefNo As String = oform.Items.Item("26").Specific.String
                            Dim sdocdate As String = oform.Items.Item("Item_9").Specific.String

                            Dim ocombo As SAPbouiCOM.ComboBox
                            If oButtton.Selected.Value = "Copy To" Then
                                If omatrix.Columns.Item("V_0sjr").Cells.Item(1).Specific.String = "" Then
                                    Scode = omatrix.Columns.Item("Col_4").Cells.Item(1).Specific.String
                                    For mjs As Integer = 1 To omatrix.RowCount
                                        If omatrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.String <> "" Then
                                            If Scode <> omatrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.String Then
                                                Oapplication_SM.StatusBar.SetText("Different supplier codes are mapped, kindly do the invoice manually ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                                BubbleEvent = False
                                                Exit Try
                                            End If
                                        End If
                                    Next mjs

                                    Oapplication_SM.ActivateMenuItem("2308")
                                    oform_In = Oapplication_SM.Forms.GetFormByTypeAndCount(141, FormType_SM) 'pVal.FormTypeCount)
                                    oform_In.Freeze(True)
                                    Dim oMatrix_in As SAPbouiCOM.Matrix = oform_In.Items.Item("39").Specific
                                    oform_In.Items.Item("3").Specific.select("S")
                                    oform_In.Items.Item("4").Specific.String = oVendor
                                    oform_In.Items.Item("14").Specific.String = VrefNo
                                    oform_In.Items.Item("D0002").Specific.String = oVno
                                    oform_In.Items.Item("D0003").Specific.String = Docnum
                                    oform_In.Items.Item("10").Specific.String = sdocdate


                                    For mjs As Integer = 1 To omatrix.RowCount
                                        If omatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String <> "" And omatrix.Columns.Item("Col_0").Cells.Item(mjs).Specific.String <> "" Then
                                            oMatrix_in.Columns.Item("1").Cells.Item(mjs).Specific.String = oVno & " - " & omatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.value
                                            oMatrix_in.Columns.Item("12").Cells.Item(mjs).Specific.String = omatrix.Columns.Item("V_0mjs").Cells.Item(mjs).Specific.String
                                            orset.DoQuery("SELECT T0.[U_account] FROM [dbo].[@AE_GLACC]  T0 WHERE T0.[Code] = '" & Trim(omatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.value) & "'")
                                            oMatrix_in.Columns.Item("2").Cells.Item(mjs).Specific.String = orset.Fields.Item("U_account").Value
                                            ocheck = omatrix.Columns.Item("GST").Cells.Item(mjs).Specific
                                            Try
                                                If ocheck.Checked = False Then
                                                    '    oMatrix_in.Columns.Item("95").Cells.Item(mjs).Specific.String = "SI"
                                                    'Else
                                                    ocombo = oMatrix_in.Columns.Item("57").Cells.Item(mjs).Specific
                                                    ocombo.Select("ZI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                                End If

                                            Catch ex As Exception

                                            End Try
                                        End If
                                    Next

                                    oform_In.Items.Item("16").Specific.String = "Service Expenses for this vehicle - " & oVno
                                    oform_In.Freeze(False)
                                Else
                                    Oapplication_SM.StatusBar.SetText("Invoice has been generated for this document ........ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                End If

                            End If

                        Catch ex As Exception
                            oform_In.Freeze(False)
                            Oapplication_SM.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Try
                        End Try
                        Exit Sub
                    End If
                End If

            ElseIf pVal.Before_Action = True Then

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                    Dim oform As SAPbouiCOM.Form
                    Try
                        oform = Oapplication_SM.Forms.ActiveForm
                        Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_13").Specific
                        oform.Freeze(True)
                        If oMatrix.RowCount > 1 Then
                            If oMatrix.Columns.Item("Col_0").Cells.Item(oMatrix.RowCount).Specific.String = "" Then
                                oMatrix.DeleteRow(oMatrix.RowCount)
                                oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            End If
                        End If
                        If oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            oform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If
                        oform.Freeze(False)

                    Catch ex As Exception
                        oform.Freeze(False)
                    End Try
                End If

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then


                    If pVal.ItemUID = "20" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.Item("SM")
                            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("18").Specific
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
                                    oMatrix.Columns.Item("V_-1").Cells.Item(mjs).Specific.string = mjs
                                Next
                                If oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If
                            End If

                        Catch ex As Exception
                            Oapplication_SM.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                        End Try
                        Exit Sub
                    End If

                    If pVal.ItemUID = "19" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.Item("SM")
                            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("18").Specific

                            Dim orset As SAPbobsCOM.Recordset = Ocompany_SM.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            orset.DoQuery("SELECT attachpath from oadp")
                            showOpenFileDialog()
                            If FileName <> "" Then
                                Dim file = New FileInfo(stFilePathAndName)
                                file.CopyTo(Path.Combine(orset.Fields.Item("attachpath").Value, file.Name), True)

                                oMatrix = oform.Items.Item("18").Specific
                                If oMatrix.RowCount = 0 Then
                                    oMatrix.AddRow()
                                Else
                                    If oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific.String <> "" And oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific.String <> "" Then
                                        oMatrix.AddRow()
                                    End If
                                End If

                                oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.string = oMatrix.RowCount
                                oMatrix = oform.Items.Item("18").Specific
                                oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific.string = orset.Fields.Item("attachpath").Value & FileName
                                oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.string = FileName
                                oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific.string = Format(Now.Date, "yyyyMMdd")
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
                            Oapplication_SM.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Try
                        End Try
                        Exit Sub
                    End If

                    If pVal.ItemUID = "1" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE ) Then
                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm

                            If Validate(oform) = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If

                        Catch ex As Exception
                            Oapplication_SM.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Try
                        End Try
                        Exit Sub
                    End If
                End If

            End If
        End If

        If pVal.FormUID = "VSH" Then
            If pVal.Before_Action = False Then
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvento = pVal
                    Dim sCFL_ID As String
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    Dim oForm As SAPbouiCOM.Form = Oapplication_SM.Forms.Item(FormUID)
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                    Try
                        If oCFLEvento.BeforeAction = False Then
                            Dim oDataTable As SAPbouiCOM.DataTable
                            oDataTable = oCFLEvento.SelectedObjects
                            If pVal.ItemUID = "16" Then
                                oForm.Items.Item("3").Specific.string = oDataTable.GetValue("ItemName", 0)
                                oForm.Items.Item("5").Specific.string = oDataTable.GetValue("U_AE_MAKE", 0)
                                oForm.Items.Item("7").Specific.string = oDataTable.GetValue("U_AE_MODEL", 0)
                                ' MsgBox(oDataTable.GetValue("U_AE_REG_DATE", 0) & " " & Format(oDataTable.GetValue("U_AE_REG_DATE", 0), "yyyyMMdd"))
                                oForm.Items.Item("9").Specific.string = Format(oDataTable.GetValue("U_AE_REG_DATE", 0), "yyyyMMdd")
                                ' MsgBox(oForm.Items.Item("9").Specific.string)
                                oForm.Items.Item("16").Specific.string = oDataTable.GetValue("ItemCode", 0)
                            End If
                        End If

                    Catch ex As Exception

                    End Try
                End If

            ElseIf pVal.Before_Action = True Then

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                    If pVal.ItemUID = "10" Then
                        Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm
                        Try

                            Dim ogrid As SAPbouiCOM.Grid = oform.Items.Item("13").Specific
                            Dim oCol As SAPbouiCOM.EditTextColumn
                            Dim sVehicle As String

                            If Not String.IsNullOrEmpty(oform.Items.Item("16").Specific.String) Then
                                sVehicle = oform.Items.Item("16").Specific.String
                            Else
                                sVehicle = "%"
                                oform.Items.Item("3").Specific.String = ""
                                oform.Items.Item("5").Specific.String = ""
                                oform.Items.Item("7").Specific.String = ""
                                oform.Items.Item("9").Specific.String = ""

                            End If

                            oform.Freeze(True)
                            Try
                                oform.DataSources.DataTables.Add("VSM")
                            Catch ex As Exception
                            End Try

                            oform.DataSources.DataTables.Item("VSM").ExecuteQuery("SELECT T0.[DocNum] as 'Document No',  T0.[U_AE_Vmileage] as 'Vehicle Mileage', T1.[U_AE_Idate] as 'IN Date', " & _
                "T1.[U_AE_Odate] as 'OUT Date', T1.[U_AE_Stype] as 'Service Type', T1.[U_AE_Desc] as 'Description', cast(T1.[U_AE_RC] as decimal(19,2)) as 'Repair Cost', T1.[U_AE_Sname] as 'Supplier Name', " & _
                "T1.[U_AE_Inv] as 'Invoice No', T1.[U_AE_Remar] as 'Remark' FROM [dbo].[@AE_SM]  T0 inner join  [dbo].[@AE_SM_R]  T1 on T0.Docentry = T1.Docentry  where T0.U_AE_Vno like '" & sVehicle & "' " & _
                "GROUP BY T0.[DocNum],  T0.[U_AE_Vmileage], T1.[U_AE_Idate], T1.[U_AE_Odate], T1.[U_AE_Stype], T1.[U_AE_Desc], T1.[U_AE_RC], T1.[U_AE_Sname], T1.[U_AE_Inv], T1.[U_AE_Remar] " & _
                "ORDER BY T1.[U_AE_Idate] DESC")

                            ogrid.DataTable = oform.DataSources.DataTables.Item("VSM")
                            ogrid.AutoResizeColumns()

                            oCol = ogrid.Columns.Item("Document No")
                            oCol.LinkedObjectType = "AE_SM"
                            'ogrid.CollapseLevel = 1

                            ogrid.Columns.Item("Document No").TextStyle = 1
                            ogrid.Columns.Item("Vehicle Mileage").TextStyle = 1
                            ogrid.Columns.Item("IN Date").TextStyle = 1
                            ogrid.Columns.Item("OUT Date").TextStyle = 1
                            ogrid.Columns.Item("Service Type").TextStyle = 1
                            ogrid.Columns.Item("Description").TextStyle = 1
                            ogrid.Columns.Item("Repair Cost").TextStyle = 1
                            ogrid.Columns.Item("Supplier Name").TextStyle = 1
                            ogrid.Columns.Item("Invoice No").TextStyle = 1
                            ogrid.Columns.Item("Remark").TextStyle = 1


                            oform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
                            oform.Freeze(False)
                        Catch ex As Exception
                            oform.Freeze(False)
                            Oapplication_SM.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Sub
                        End Try
                        Exit Sub
                    End If


                End If

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED Then
                    If pVal.ItemUID = "13" And pVal.ColUID = "Document No" Then

                        Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm

                        Try
                            Dim ogrid As SAPbouiCOM.Grid = oform.Items.Item("13").Specific
                            '' MsgBox(ogrid.DataTable.Columns.Item(0).Cells.Item(pVal.Row).Value)
                            Dim sDocnum As String = ogrid.DataTable.Columns.Item(0).Cells.Item(pVal.Row).Value 'ogrid.DataTable.GetValue("Document No", pVal.Row - 1)
                            bSVH = True
                            LoadFromXML("ServiceMaintenance.srf", Oapplication_SM)
                            oform = Oapplication_SM.Forms.Item("SM")
                            oform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oform.Items.Item("Item_7").Enabled = True
                            oform.Items.Item("Item_7").Specific.String = sDocnum
                            oform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Dim ocombobutton As SAPbouiCOM.ButtonCombo = oform.Items.Item("21").Specific
                            ocombobutton.ValidValues.Add("Copy To", "Copy To A/P Invoice")
                            ' '' oform.PaneLevel = 1
                            oform.Items.Item("Item_7").Enabled = False
                            oform.Visible = True
                        Catch ex As Exception
                            Oapplication_SM.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Sub
                        End Try
                        Exit Sub
                    End If
                End If

            End If

        End If

        If pVal.FormUID = "SMS" Then
            If pVal.Before_Action = True Then
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.InnerEvent = False Then

                    If pVal.ItemUID = "Item_14" Then
                        Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm
                        Try
                            If oform.Items.Item("Item_14").Specific.String <> "" Then
                                If oform.Items.Item("Item_12").Specific.String <> "" Then
                                    oform.Items.Item("Item_16").Specific.String = CInt(oform.Items.Item("Item_12").Specific.String) + CInt(oform.Items.Item("Item_14").Specific.String)
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                        Exit Sub
                    End If

                    If pVal.ItemUID = "Item_20" Then
                        Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm
                        Try
                            Dim oDate As Date
                            If oform.Items.Item("Item_20").Specific.String <> "" Then
                                If oform.Items.Item("Item_18").Specific.String <> "" Then
                                    oDate = GateDate(oform.Items.Item("Item_18").Specific.String, Ocompany_SM)

                                    oform.Items.Item("Item_22").Specific.String = Format(oDate.AddDays(CInt(oform.Items.Item("Item_20").Specific.String)), "yyyyMMdd")
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                        Exit Sub
                    End If

                    If pVal.ItemUID = "Item_28" Then
                        Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm
                        Try
                            Dim oDate As Date
                            If oform.Items.Item("Item_28").Specific.String <> "" Then
                                If oform.Items.Item("Item_26").Specific.String <> "" Then
                                    oDate = GateDate(oform.Items.Item("Item_26").Specific.String, Ocompany_SM)
                                    oform.Items.Item("Item_30").Specific.String = Format(oDate.AddDays(CInt(oform.Items.Item("Item_28").Specific.String)), "yyyyMMdd")
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                        Exit Sub
                    End If

                    If pVal.ItemUID = "Item_44" Then
                        Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm
                        Try
                            If oform.Items.Item("Item_44").Specific.String <> "" Then
                                If oform.Items.Item("Item_42").Specific.String <> "" Then
                                    oform.Items.Item("Item_46").Specific.String = CInt(oform.Items.Item("Item_42").Specific.String) + CInt(oform.Items.Item("Item_44").Specific.String)
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                        Exit Sub
                    End If


                End If

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.ItemUID = "Item_7" Then
                        Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm
                        oform.PaneLevel = 1
                    End If
                    If pVal.ItemUID = "Item_8" Then
                        Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm
                        oform.PaneLevel = 2
                    End If
                    If pVal.ItemUID = "Item_9" Then
                        Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm
                        oform.PaneLevel = 3
                    End If

                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                        Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm

                        If oform.Items.Item("Item_1").Specific.String = "" Then
                            oform.Items.Item("Item_1").Specific.active = True
                            Oapplication_SM.StatusBar.SetText("Vehicle Number should not be empty ..........!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If

                End If

            ElseIf pVal.Before_Action = False Then

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then

                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvento = pVal
                    Dim sCFL_ID As String
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    Dim oForm As SAPbouiCOM.Form = Oapplication_SM.Forms.Item(FormUID)
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                    Try

                        If oCFLEvento.BeforeAction = False Then
                            Dim oDataTable As SAPbouiCOM.DataTable
                            Dim oedit1 As SAPbouiCOM.EditText
                            oDataTable = oCFLEvento.SelectedObjects
                            If pVal.ItemUID = "Item_1" Then 'Billing Code
                                oForm.Items.Item("Item_2").Specific.string = oDataTable.GetValue("ItemName", 0)
                                oForm.Items.Item("Item_1").Specific.string = oDataTable.GetValue("ItemCode", 0)
                            End If

                            If pVal.ItemUID = "Item_13" And pVal.ColUID = "Col_4" Then

                                Dim omatrix As SAPbouiCOM.Matrix = oForm.Items.Item("Item_13").Specific

                                omatrix.Columns.Item("Col_5").Cells.Item(pVal.Row).Specific.String = oDataTable.GetValue("ItemName", 0)
                                omatrix.Columns.Item("Col_4").Cells.Item(pVal.Row).Specific.String = oDataTable.GetValue("ItemName", 0)


                            End If
                        End If
                    Catch ex As Exception
                    End Try
                End If

                If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm
                    If pVal.Action_Success = True Then
                        oform.Items.Item("Item_4").Specific.String = Format(Now.Date, "yyyyMMdd") 'Now.Date


                    End If
                End If
            End If
        End If



    End Sub




    Private Sub Oapplication_SM_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles Oapplication_SM.MenuEvent
        Try

            If pVal.MenuUID = "SMS" And pVal.BeforeAction = True Then

                LoadFromXML("ServiceMaintenenceSetup.srf", Oapplication_SM)
                Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.Item("SMS")

                oform.Visible = True
                oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                oform.Items.Item("Item_7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                oform.Items.Item("Item_4").Specific.String = Format(Now.Date, "yyyyMMdd") ' Now.Date ' Format(Now.Date, "dd MMM yyyy")
                oform.PaneLevel = 7
                
                Dim oCFLs As SAPbouiCOM.ChooseFromList
                Dim oCons As SAPbouiCOM.Conditions
                Dim oCon As SAPbouiCOM.Condition
                Dim empty As New SAPbouiCOM.Conditions

                oCFLs = oform.ChooseFromLists.Item("CFL_2")
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

                oform.DataBrowser.BrowseBy = "Item_7"

            End If

            If pVal.MenuUID = "VSH" And pVal.BeforeAction = True Then

                LoadFromXML("VehicleServiceHistory.srf", Oapplication_SM)
                Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.Item("VSH")

                oform.Visible = True
                oform.Items.Item("12").Specific.String = Format(Now.Date, "yyyyMMdd") 'Now.Date ' Format(Now.Date, "dd MMM yyyy")

                Try
                    oform.DataSources.DataTables.Add("VSH")
                Catch ex As Exception

                End Try
                Dim ogrid As SAPbouiCOM.Grid = oform.Items.Item("13").Specific
                oform.DataSources.DataTables.Item(0).ExecuteQuery("SELECT T0.[DocNum] as 'Document No.',  T0.[U_AE_Vmileage] as 'Vehicle Mileage', T1.[U_AE_Idate] as 'IN Date', " & _
                "T1.[U_AE_Odate] as 'OUT Date', T1.[U_AE_Stype] as 'Service Type', T1.[U_AE_Desc] as 'Description', cast(T1.[U_AE_RC] as decimal(19,2)) as 'Repair Cost', T1.[U_AE_Sname] as 'Supplier Name', " & _
                "T1.[U_AE_Inv] as 'Invoice No.', T1.[U_AE_Remar] as 'Remark' FROM [dbo].[@AE_SM]  T0 inner join  [dbo].[@AE_SM_R]  T1 on T0.Docentry = T1.Docentry  where T0.Docnum = '' " & _
                "GROUP BY T0.[DocNum],  T0.[U_AE_Vmileage], T1.[U_AE_Idate], T1.[U_AE_Odate], T1.[U_AE_Stype], T1.[U_AE_Desc], T1.[U_AE_RC], T1.[U_AE_Sname], T1.[U_AE_Inv], T1.[U_AE_Remar] " & _
                "ORDER BY T1.[U_AE_Idate] desc")
                ogrid.DataTable = oform.DataSources.DataTables.Item("VSH")
                ogrid.AutoResizeColumns()

                Dim oCFLs As SAPbouiCOM.ChooseFromList
                Dim oCons As SAPbouiCOM.Conditions
                Dim oCon As SAPbouiCOM.Condition
                Dim empty As New SAPbouiCOM.Conditions

                oCFLs = oform.ChooseFromLists.Item("CFL_2")
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
                oform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            End If


            If pVal.MenuUID = "SMSM" And pVal.BeforeAction = True Then

                LoadFromXML("ServiceMaintenance.srf", Oapplication_SM)
                Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.Item("SM")
                Dim orset As SAPbobsCOM.Recordset = Ocompany_SM.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oform.Items.Item("Item_11").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                oform.Visible = True
                orset.DoQuery("SELECT Code, Name  FROM [dbo].[@AE_SERVICETYPE]  T0")
                Dim oButton As SAPbouiCOM.ButtonCombo = oform.Items.Item("21").Specific
                Dim Tmp_val As String = oform.BusinessObject.GetNextSerialNumber(NextSerialNo(Ocompany_SM, Oapplication_SM, "AE_SM"))
                oform.Items.Item("Item_7").Specific.String = Tmp_val
                oform.Items.Item("Item_9").Specific.String = Format(Now.Date, "yyyyMMdd") 'Now.Date 'Format(Now.Date, "dd MMM yyyy")
                oform.PaneLevel = 1
                oButton.ValidValues.Add("Copy To", "Copy To A/P Invoice")
                oform.Items.Item("21").Enabled = False
                Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_13").Specific
                omatrix.Columns.Item("Col_0").Editable = True
                omatrix.AddRow()
                omatrix.Columns.Item("#").Cells.Item(omatrix.RowCount).Specific.String = omatrix.RowCount

                Dim oCombo As SAPbouiCOM.ComboBox = oform.Items.Item("23").Specific
                Try
                    oCombo.ValidValues.Add("Open", "Open")
                    oCombo.ValidValues.Add("Closed", "Closed")
                Catch ex As Exception
                End Try

                oCombo.Select("Open")
                oCombo = omatrix.Columns.Item("Col_2").Cells.Item(1).Specific
                Try
                    For mjs As Integer = 1 To orset.RecordCount
                        oCombo.ValidValues.Add(orset.Fields.Item("Code").Value, orset.Fields.Item("Name").Value)
                        orset.MoveNext()
                    Next mjs
                Catch ex As Exception
                End Try

                oform.Items.Item("1000002").Specific.String = Ocompany_SM.UserName
                oform.Items.Item("24").Specific.String = Company_Name

                Dim oCFLs As SAPbouiCOM.ChooseFromList
                Dim oCons As SAPbouiCOM.Conditions
                Dim oCon As SAPbouiCOM.Condition
                Dim empty As New SAPbouiCOM.Conditions

                oCFLs = oform.ChooseFromLists.Item("CFL_2")
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


                oCFLs = oform.ChooseFromLists.Item("CFL_3")
                oCFLs.SetConditions(empty)
                oCons = oCFLs.GetConditions()
                oCon = oCons.Add()
                oCon.Alias = "CardType"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "S"
                oCFLs.SetConditions(oCons)


                oform.DataBrowser.BrowseBy = "Item_7"


            End If

            If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.BeforeAction = False Then
                Try
                    Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm
                    Dim oEdit As SAPbouiCOM.EditText
                    If oform.UniqueID = "SM" Then
                        Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_13").Specific
                        If omatrix.Columns.Item("Col_0").Cells.Item(omatrix.RowCount).Specific.String <> "" Then
                            oform.DataSources.DBDataSources.Item(1).Clear()
                            omatrix.AddRow(1)
                            ''If oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            ''    oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            ''End If
                            'oEdit = omatrix.Columns.Item("Col_0").Cells.Item(omatrix.RowCount).Specific
                            'oEdit.Active = True
                        End If
                        omatrix.Columns.Item("#").Cells.Item(omatrix.RowCount).Specific.String = omatrix.RowCount
                        If Trim(oform.Items.Item("23").Specific.value) = "Open" Then
                            oform.Items.Item("21").Enabled = True
                            oform.Items.Item("23").Enabled = True
                            oform.Items.Item("Item_13").Enabled = True
                            oform.Items.Item("Item_1").Enabled = True
                            oform.Items.Item("Item_5").Enabled = True
                            oform.Items.Item("1").Enabled = True
                            oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                        Else
                            oform.Items.Item("21").Enabled = False
                            oform.Items.Item("Item_13").Enabled = False
                            oform.Items.Item("Item_1").Enabled = False
                            oform.Items.Item("Item_5").Enabled = False
                            oform.Items.Item("23").Enabled = False
                        End If
                    End If
                   

                Catch ex As Exception
                    Oapplication_SM.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try
            End If

            If pVal.MenuUID = "1281" And pVal.BeforeAction = True Then
                Dim oform As SAPbouiCOM.Form = Oapplication_SM.Forms.ActiveForm
                If oform.UniqueID = "SM" Then
                    SM = True
                    oform.Items.Item("Item_7").Enabled = True
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


    Public Sub ShowFolderBrowser()

        Dim MyProcs() As System.Diagnostics.Process
        FileName = ""
        Dim OpenFile As New OpenFileDialog
        ' Dim stFilePathAndName As String
        Dim orset As SAPbobsCOM.Recordset = Ocompany_SM.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
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
            OpenFile.InitialDirectory = "C:\" 'orset.Fields.Item("attachpath").Value
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


    Private Function Validate(ByRef oform As SAPbouiCOM.Form) As Boolean
        Try



            If oform.Items.Item("Item_1").Specific.String = "" Then
                oform.Items.Item("Item_1").Specific.active = True
                Oapplication_SM.StatusBar.SetText("Vehicle Number should not be empty ........ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False
            End If

            If oform.Items.Item("Item_5").Specific.String = "" Then
                oform.Items.Item("Item_5").Specific.active = True
                Oapplication_SM.StatusBar.SetText("Vehicle Mileage should not be empty ........ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False
            End If

            Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_13").Specific

            If oMAtrix.RowCount > 1 Then 'And oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                If oMAtrix.Columns.Item("Col_0").Cells.Item(oMAtrix.RowCount).Specific.String = "" And oMAtrix.Columns.Item("Col_1").Cells.Item(oMAtrix.RowCount).Specific.String = "" And oMAtrix.Columns.Item("Col_2").Cells.Item(oMAtrix.RowCount).Specific.value = "" Then
                    oMAtrix.DeleteRow(oMAtrix.RowCount)
                End If
            End If

            For mjs As Integer = 1 To oMAtrix.RowCount
                If oMAtrix.Columns.Item("Col_0").Cells.Item(mjs).Specific.String = "" Then
                    oMAtrix.Columns.Item("Col_0").Cells.Item(mjs).Specific.active = True
                    Oapplication_SM.StatusBar.SetText("In Date should not be empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Return False
                End If

                If oMAtrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String = "" Then
                    oMAtrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.active = True
                    Oapplication_SM.StatusBar.SetText("Out Date should not be empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Return False
                End If

                If oMAtrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.value = "" Then
                    oMAtrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.active = True
                    Oapplication_SM.StatusBar.SetText("Service Type should not be empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Return False
                End If

                If oMAtrix.Columns.Item("Col_3").Cells.Item(mjs).Specific.String = "" Then
                    oMAtrix.Columns.Item("Col_3").Cells.Item(mjs).Specific.active = True
                    Oapplication_SM.StatusBar.SetText("Descrition should not be empty ......... ! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Return False
                End If

                If oMAtrix.Columns.Item("V_0mjs").Cells.Item(mjs).Specific.String = "0.00" Then
                    oMAtrix.Columns.Item("Col_3").Cells.Item(mjs).Specific.active = True
                    Oapplication_SM.StatusBar.SetText("Repair Cost should not be empty ......... ! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Return False
                End If

                If oMAtrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.String = "" Then
                    oMAtrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.active = True
                    Oapplication_SM.StatusBar.SetText("Supplier should not be empty ............ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Return False
                End If

            Next mjs

            Return True

        Catch ex As Exception
            Oapplication_SM.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try

    End Function

End Class
