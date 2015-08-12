Imports System.IO
Imports System.Globalization
Public Class TrafficAccident


    Dim WithEvents Oapplication_TA As SAPbouiCOM.Application
    Dim Ocompany_TA As New SAPbobsCOM.Company
    Private FileName As String = ""



    Public Sub New(ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company)

        Oapplication_TA = oApplication
        Ocompany_TA = oCompany

    End Sub

    Private Sub Oapplication_TA_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles Oapplication_TA.FormDataEvent
        If BusinessObjectInfo.FormUID = "VT" Then
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                Dim orset As SAPbobsCOM.Recordset = Ocompany_TA.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim DEntry As String = Trim(oform.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0))
                orset.DoQuery("update [@AE_VTRACK] set U_AE_stat = 'C' where [U_AE_Vno] = '" & oform.Items.Item("Item_3").Specific.String & "' and  [DocEntry] <> " & DEntry & "")
                orset.DoQuery("update [@AE_VTRACK] set U_AE_stat = 'O' where [U_AE_Vno] = '" & oform.Items.Item("Item_3").Specific.String & "' and  [DocEntry] = " & DEntry & "")
            End If
        End If
    End Sub

    Private Sub Oapplication_TA_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles Oapplication_TA.ItemEvent
        Try

            If pVal.FormUID = "VTLR" Then
                If pVal.Before_Action = True Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                            CD_Timer.Stop()
                        Catch ex As Exception

                        End Try


                    End If

                End If
            End If

            If pVal.FormUID = "ACD" Then

                If pVal.Before_Action = False Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then

                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pVal
                        Dim sCFL_ID As String
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        Dim oForm As SAPbouiCOM.Form = Oapplication_TA.Forms.Item(FormUID)
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                        Try

                            If oCFLEvento.BeforeAction = False Then
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim oedit1 As SAPbouiCOM.EditText
                                oDataTable = oCFLEvento.SelectedObjects

                                If pVal.ItemUID = "Item_30" Then 'Order By

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

                                    oForm.Items.Item("Item_58").Specific.string = oDataTable.GetValue("CardName", 0)
                                    oForm.Items.Item("Item_60").Specific.string = oDataTable.GetValue("BillToDef", 0) & ", " & oDataTable.GetValue("Address", 0) & ", " & oDataTable.GetValue("City", 0) & ", " & oDataTable.GetValue("Country", 0) & ", " & oDataTable.GetValue("ZipCode", 0)
                                    oForm.Items.Item("Item_64").Specific.string = oDataTable.GetValue("Cellular", 0)
                                    oForm.Items.Item("Item_62").Specific.string = oDataTable.GetValue("CntctPrsn", 0)
                                    oForm.Items.Item("Item_57").Specific.string = oDataTable.GetValue("CardCode", 0)

                                ElseIf pVal.ItemUID = "129" Then 'Issued By
                                    oForm.Items.Item("181").Specific.string = oDataTable.GetValue("firstName", 0) & " " & oDataTable.GetValue("lastName", 0)
                                    oForm.Items.Item("129").Specific.string = oDataTable.GetValue("empID", 0)

                                ElseIf pVal.ItemUID = "130" Then 'Issued By
                                    oForm.Items.Item("Item_7").Specific.string = oDataTable.GetValue("firstName", 0) & " " & oDataTable.GetValue("lastName", 0)
                                    oForm.Items.Item("130").Specific.string = oDataTable.GetValue("empID", 0)

                                ElseIf pVal.ItemUID = "1000001" Then 'Issued By

                                    Dim orset As SAPbobsCOM.Recordset = Ocompany_TA.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim orset1 As SAPbobsCOM.Recordset = Ocompany_TA.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    ' ------------------ New ----------------------
                                    orset.DoQuery("SELECT T0.[U_AE_DName] FROM [dbo].[@AE_SBOOKING]  T0 WHERE T0.[DocEntry] = '" & oDataTable.GetValue("DocEntry", 0) & "' union all SELECT T0.[U_AE_DName1] as 'U_AE_DName' FROM [dbo].[@AE_SBOOKING]  T0 WHERE T0.[DocEntry] = '" & oDataTable.GetValue("DocEntry", 0) & "'")
                                    orset1.DoQuery("SELECT T0.[U_AE_Bcode], T0.[U_AE_Bname], T0.[U_AE_Address], T0.[U_AE_Atten], T0.[U_AE_Cno] FROM [dbo].[@AE_SBOOKING]  T0 WHERE T0.[DocEntry] = '" & oDataTable.GetValue("DocEntry", 0) & "'")
                                    Dim ocombo As SAPbouiCOM.ComboBox = oForm.Items.Item("1000011").Specific
                                    'MsgBox(ocombo.ValidValues.Count)

                                    Try

                                        oForm.Items.Item("Item_57").Specific.string = orset1.Fields.Item("U_AE_Bcode").Value
                                    Catch ex As Exception
                                    End Try

                                    oForm.Items.Item("Item_58").Specific.string = orset1.Fields.Item("U_AE_Bname").Value
                                    oForm.Items.Item("Item_60").Specific.string = orset1.Fields.Item("U_AE_Address").Value
                                    oForm.Items.Item("Item_62").Specific.string = orset1.Fields.Item("U_AE_Atten").Value
                                    oForm.Items.Item("Item_64").Specific.string = orset1.Fields.Item("U_AE_Cno").Value


                                    For mjs As Integer = ocombo.ValidValues.Count - 1 To 0 Step -1
                                        ocombo.ValidValues.Remove(mjs, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next mjs

                                    For mjs As Integer = 1 To orset.RecordCount
                                        ocombo.ValidValues.Add(orset.Fields.Item("U_AE_DName").Value, "")
                                        orset.MoveNext()
                                    Next mjs
                                    ocombo.ValidValues.Add("-", "")
                                    ocombo.Select("-")
                                    oForm.Items.Item("1000001").Specific.string = oDataTable.GetValue("DocEntry", 0)


                                ElseIf pVal.ItemUID = "131" Then 'Issued By
                                    oForm.Items.Item("Item_12").Specific.string = oDataTable.GetValue("firstName", 0) & " " & oDataTable.GetValue("lastName", 0)
                                    oForm.Items.Item("131").Specific.string = oDataTable.GetValue("empID", 0)

                                ElseIf pVal.ItemUID = "95" Then 'Issued By
                                    oForm.Items.Item("96").Specific.string = oDataTable.GetValue("ItemName", 0)
                                    oForm.Items.Item("1000002").Specific.string = oDataTable.GetValue("U_AE_MODEL", 0)
                                    oForm.Items.Item("1000004").Specific.string = oDataTable.GetValue("U_AE_YEAR_Make", 0)
                                    oForm.Items.Item("104").Specific.string = oDataTable.GetValue("U_AE_CHASSIS_NO", 0)
                                    oForm.Items.Item("95").Specific.string = oDataTable.GetValue("ItemCode", 0)
                                End If

                            End If
                        Catch ex As Exception
                        End Try
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then

                        If pVal.ItemUID = "179" Then

                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                                If Trim(oform.Items.Item("179").Specific.value) = "Staff (CD or Errand)" Then
                                    oform.Items.Item("1000001").Specific.String = ""
                                    oform.Items.Item("Item_16").Specific.active = True
                                    oform.Items.Item("1000001").Enabled = False

                                    oform.Items.Item("Item_30").Visible = True
                                    oform.Items.Item("Item_31").Visible = True
                                    oform.Items.Item("1000011").Visible = False
                                    oform.Items.Item("132").Visible = True

                                    oform.Items.Item("Item_30").Specific.String = ""
                                    oform.Items.Item("132").Specific.String = ""
                                    Dim ocombo As SAPbouiCOM.ComboBox = oform.Items.Item("1000011").Specific
                                    For mjs As Integer = ocombo.ValidValues.Count - 1 To 0 Step -1
                                        ocombo.ValidValues.Remove(mjs, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next mjs
                                    ocombo.ValidValues.Add("", "")
                                    ocombo.Select("")
                                    oform.Items.Item("Item_31").Specific.String = ""

                                Else
                                    oform.Items.Item("132").Specific.String = ""
                                    oform.Items.Item("Item_16").Specific.active = True
                                    oform.Items.Item("132").Visible = False
                                    oform.Items.Item("1000001").Enabled = True

                                    oform.Items.Item("Item_30").Visible = False
                                    oform.Items.Item("Item_31").Visible = False
                                    oform.Items.Item("1000011").Visible = True

                                    oform.Items.Item("Item_30").Specific.String = ""
                                    oform.Items.Item("Item_31").Specific.String = ""
                                    oform.Items.Item("132").Specific.String = ""
                                    Dim ocombo As SAPbouiCOM.ComboBox = oform.Items.Item("1000011").Specific
                                    For mjs As Integer = ocombo.ValidValues.Count - 1 To 0 Step -1
                                        ocombo.ValidValues.Remove(mjs, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next mjs
                                End If

                                oform.Items.Item("Item_34").Specific.String = ""
                                oform.Items.Item("Item_36").Specific.String = ""
                                oform.Items.Item("Item_38").Specific.String = ""
                                oform.Items.Item("Item_40").Specific.String = ""
                                oform.Items.Item("Item_42").Specific.String = ""
                                oform.Items.Item("Item_44").Specific.String = ""
                                oform.Items.Item("Item_49").Specific.String = ""
                                oform.Items.Item("Item_50").Specific.String = ""
                                oform.Items.Item("Item_51").Specific.String = ""

                                oform.Items.Item("Item_52").Specific.String = ""
                                oform.Items.Item("Item_54").Specific.String = ""

                            Catch ex As Exception
                            End Try
                        End If

                        If pVal.ItemUID = "1000011" Then
                            Dim oform As SAPbouiCOM.Form
                            Try
                                oform = Oapplication_TA.Forms.ActiveForm
                                Dim orset As SAPbobsCOM.Recordset = Ocompany_TA.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oform.Freeze(True)
                                If oform.Items.Item("1000011").Specific.selected.value <> "-" Then
                                    orset.DoQuery("SELECT T0.[U_AE_Dadd], T0.[U_AE_Dcno], T0.[U_AE_Occuption], T0.[U_AE_Nation], T0.[U_AE_DOB], T0.[U_AE_License], T0.[U_AE_Pissue], T0.[U_AE_Exdate], T0.[U_AE_Passno], T0.[U_AE_Pissuepno], T0.[U_AE_Pexdate] FROM [dbo].[@AE_SBOOKING]  T0 WHERE T0.[DocEntry] = '" & oform.Items.Item("1000001").Specific.string & "' and T0.[U_AE_DName] = '" & oform.Items.Item("1000011").Specific.selected.value & "'")
                                    oform.Items.Item("Item_34").Specific.string = orset.Fields.Item("U_AE_Dadd").Value
                                    oform.Items.Item("Item_36").Specific.string = orset.Fields.Item("U_AE_Dcno").Value
                                    oform.Items.Item("Item_38").Specific.string = orset.Fields.Item("U_AE_Occuption").Value
                                    oform.Items.Item("Item_40").Specific.string = orset.Fields.Item("U_AE_Nation").Value
                                    oform.Items.Item("Item_42").Specific.string = Format(orset.Fields.Item("U_AE_DOB").Value, "dd/MM/yyyy")
                                    oform.Items.Item("Item_44").Specific.string = orset.Fields.Item("U_AE_License").Value
                                    oform.Items.Item("Item_49").Specific.string = orset.Fields.Item("U_AE_Pissue").Value
                                    oform.Items.Item("Item_50").Specific.string = Format(orset.Fields.Item("U_AE_Exdate").Value, "dd/MM/yyyy")
                                    oform.Items.Item("Item_51").Specific.string = orset.Fields.Item("U_AE_Passno").Value
                                    oform.Items.Item("Item_52").Specific.string = orset.Fields.Item("U_AE_Pissuepno").Value
                                    oform.Items.Item("Item_54").Specific.string = Format(orset.Fields.Item("U_AE_Pexdate").Value, "dd/MM/yyyy")
                                Else
                                    oform.Items.Item("Item_34").Specific.string = ""
                                    oform.Items.Item("Item_36").Specific.string = ""
                                    oform.Items.Item("Item_38").Specific.string = ""
                                    oform.Items.Item("Item_40").Specific.string = ""
                                    oform.Items.Item("Item_42").Specific.string = ""
                                    oform.Items.Item("Item_44").Specific.string = ""
                                    oform.Items.Item("Item_49").Specific.string = ""
                                    oform.Items.Item("Item_50").Specific.string = ""
                                    oform.Items.Item("Item_51").Specific.string = ""
                                    oform.Items.Item("Item_52").Specific.string = ""
                                    oform.Items.Item("Item_54").Specific.string = ""
                                End If

                                oform.Freeze(False)

                            Catch ex As Exception
                                oform.Freeze(False)
                            End Try


                        End If

                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        If pVal.ItemUID = "1" And pVal.Action_Success = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                            Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                            ' oform.Close()

                            ' LoadFromXML("Accident_Claim.srf", Oapplication_TA)
                            oform = Oapplication_TA.Forms.Item("ACD")
                            oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                            Dim Tmp_val As String = oform.BusinessObject.GetNextSerialNumber(NextSerialNo(Ocompany_TA, Oapplication_TA, "AE_Accident"))
                            oform.Items.Item("Item_14").Specific.String = Tmp_val
                            oform.Items.Item("134").Specific.select("Open")
                            oform.Items.Item("Item_20").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oform.Visible = True
                            oform.Items.Item("Item_16").Specific.String = Now.Date ' Format(Now.Date, "dd MMM yyyy")
                            oform.Items.Item("1000001").Enabled = False
                            oform.Items.Item("138").Specific.String = Ocompany_TA.UserName
                            oform.Items.Item("139").Specific.String = Company_Name
                            oform.DataBrowser.BrowseBy = "Item_14"
                            oform.PaneLevel = 1
                        End If
                    End If


                    '--------------- New -------------------------------
                    If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm

                        Try
                            If ACD_Flag = True Then
                                oform.Freeze(True)
                                If Trim(oform.Items.Item("179").Specific.value) = "Staff (CD or Errand)" Then
                                    oform.Items.Item("Item_30").Visible = True
                                    oform.Items.Item("Item_31").Visible = True
                                    oform.Items.Item("1000011").Visible = False
                                Else
                                    oform.Items.Item("Item_30").Visible = False
                                    oform.Items.Item("Item_31").Visible = False
                                    oform.Items.Item("1000011").Visible = True
                                End If
                                oform.Freeze(False)
                                ACD_Flag = False
                            End If
                        Catch ex As Exception
                            oform.Freeze(False)
                        End Try
                    End If
                    ' ---------------------------------------------------------------------------------


                End If
                If pVal.Before_Action = True Then

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.InnerEvent = False Then
                        If pVal.ItemUID = "1000012" Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                                If oform.Items.Item("1000012").Specific.String <> "" Then
                                    If oform.Items.Item("1000012").Specific.String.ToString.Length = "4" Then
                                        If IsDate(oform.Items.Item("1000012").Specific.String.ToString.Substring(0, 2) & ":" & oform.Items.Item("1000012").Specific.String.ToString.Substring(2, 2) & ":00") = False Then
                                            oform.Items.Item("1000012").Specific.active = True
                                            Oapplication_TA.StatusBar.SetText("Invalid Time format ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Try
                                        Else
                                            oform.Items.Item("1000012").Specific.String = oform.Items.Item("1000012").Specific.String.ToString.Substring(0, 2) & ":" & oform.Items.Item("1000012").Specific.String.ToString.Substring(2, 2)
                                        End If
                                    ElseIf oform.Items.Item("1000012").Specific.String.ToString.Length = "5" Then
                                        If IsDate(oform.Items.Item("1000012").Specific.String & ":00") = False Then
                                            oform.Items.Item("1000012").Specific.active = True
                                            Oapplication_TA.StatusBar.SetText("Invalid Time format ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Try
                                        End If
                                    Else

                                        oform.Items.Item("1000012").Specific.active = True
                                        Oapplication_TA.StatusBar.SetText("Invalid Time format ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Try

                                    End If
                                End If
                            Catch ex As Exception

                            End Try
                            Exit Sub
                        End If

                        If pVal.ItemUID = "108" Or pVal.ItemUID = "115" Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                                If oform.Items.Item(pVal.ItemUID).Specific.String <> "" Then
                                    If oform.Items.Item(pVal.ItemUID).Specific.String.ToString.Length = "4" Then
                                        If IsDate(oform.Items.Item(pVal.ItemUID).Specific.String.ToString.Substring(0, 2) & ":" & oform.Items.Item(pVal.ItemUID).Specific.String.ToString.Substring(2, 2) & ":00") = False Then
                                            oform.Items.Item(pVal.ItemUID).Specific.active = True
                                            Oapplication_TA.StatusBar.SetText("Invalid Time format ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Try
                                        Else
                                            oform.Items.Item(pVal.ItemUID).Specific.String = oform.Items.Item(pVal.ItemUID).Specific.String.ToString.Substring(0, 2) & ":" & oform.Items.Item(pVal.ItemUID).Specific.String.ToString.Substring(2, 2)
                                        End If
                                    ElseIf oform.Items.Item(pVal.ItemUID).Specific.String.ToString.Length = "5" Then
                                        If IsDate(oform.Items.Item(pVal.ItemUID).Specific.String & ":00") = False Then
                                            oform.Items.Item(pVal.ItemUID).Specific.active = True
                                            Oapplication_TA.StatusBar.SetText("Invalid Time format ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Try
                                        End If
                                    Else

                                        oform.Items.Item(pVal.ItemUID).Specific.active = True
                                        Oapplication_TA.StatusBar.SetText("Invalid Time format ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Try

                                    End If
                                End If
                            Catch ex As Exception

                            End Try
                            Exit Sub
                        End If

                    End If


                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                        If pVal.ItemUID = "141" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                                Dim docnum = oform.Items.Item("1000001").Specific.String
                                If docnum <> "" Then

                                    LoadFromXML("SelfDriving_Booking.srf", Oapplication_TA)
                                    oform = Oapplication_TA.Forms.Item("SDB")
                                    oform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    oform.Items.Item("Item_14").Enabled = True
                                    oform.Items.Item("Item_14").Specific.String = docnum
                                    oform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    ' oform.Items.Item("1000001").Specific.active = True


                                    Dim ocombobutton As SAPbouiCOM.ButtonCombo = oform.Items.Item("210").Specific
                                    ocombobutton.ValidValues.Add("Copy To", "Copy To A/R Invoice")

                                    Dim ooption As SAPbouiCOM.OptionBtn = oform.Items.Item("195").Specific
                                    Dim ooption1 As SAPbouiCOM.OptionBtn = oform.Items.Item("196").Specific

                                    Dim ooption2 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_104").Specific
                                    Dim ooption3 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_105").Specific
                                    Dim ooption4 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_106").Specific

                                    ooption1.GroupWith("195")

                                    ooption3.GroupWith("Item_104")
                                    ooption4.GroupWith("Item_104")
                                    ' oform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
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
                                    oform.Items.Item("Item_14").Enabled = False
                                    'oform.Items.Item("Item_22").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    'oform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                    oform.Visible = True
                                End If

                            Catch ex As Exception

                            End Try
                            Exit Sub
                        End If

                        If pVal.ItemUID = "Item_20" Or pVal.ItemUID = "93" Or pVal.ItemUID = "100" Or pVal.ItemUID = "Item_25" Or pVal.ItemUID = "135" Then

                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm

                                Select Case pVal.ItemUID

                                    Case "Item_20"
                                        oform.PaneLevel = 1
                                    Case "93"
                                        oform.PaneLevel = 2
                                    Case "100"
                                        oform.PaneLevel = 3
                                    Case "Item_25"
                                        oform.PaneLevel = 6
                                    Case "135"
                                        oform.PaneLevel = 4
                                End Select

                            Catch ex As Exception

                            End Try
                        End If

                        If pVal.ItemUID = "Item_209" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.Item("ACD")
                                Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_207").Specific

                                Dim orset As SAPbobsCOM.Recordset = Ocompany_TA.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
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
                                Oapplication_TA.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub
                        End If


                        If pVal.ItemUID = "Item_208" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.Item("ACD")
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
                                Oapplication_TA.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End Try
                            Exit Sub
                        End If

                        If pVal.ItemUID = "1" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                            Try

                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm

                                If oform.Items.Item("179").Specific.value = "" Then
                                    oform.Items.Item("179").Specific.active = True
                                    Oapplication_TA.StatusBar.SetText("Category should not be Empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                If oform.Items.Item("Item_30").Specific.value = "" Then
                                    oform.Items.Item("Item_30").Specific.active = True
                                    Oapplication_TA.StatusBar.SetText("Driver Name should not be Empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                If oform.Items.Item("Item_82").Specific.value = "" Then
                                    oform.Items.Item("Item_82").Specific.active = True
                                    Oapplication_TA.StatusBar.SetText("Vehicle Number should not be Empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                            Catch ex As Exception
                            End Try

                        End If
                    End If
                End If
            End If


            If pVal.FormUID = "TPODR" Then
                If pVal.Before_Action = True Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        If pVal.ItemUID = "7" Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                                Dim orset As SAPbobsCOM.Recordset = Ocompany_TA.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim sqlstr As String = ""

                                If oform.Items.Item("10").Specific.String = "" And oform.Items.Item("4").Specific.String = "" Then

                                    sqlstr = "SELECT T0.[U_AE_Agency] as 'Agency', T0.[U_AE_Offense] as 'Type of Offense', T0.[U_AE_Edate] as 'Expiry Date', T0.[U_AE_nno] as 'Notice Number',T0.[U_AE_fine] as 'Fine Amount',  T0.[U_AE_Submit] as 'Created By', T0.[U_AE_Status] as 'Status'  FROM [dbo].[@AE_TRAFFICO]  T0 WHERE T0.[U_AE_Status] = '" & Trim(oform.Items.Item("6").Specific.value) & "'"
                                ElseIf oform.Items.Item("10").Specific.String <> "" And oform.Items.Item("4").Specific.String <> "" Then
                                    sqlstr = "SELECT T0.[U_AE_Agency] as 'Agency', T0.[U_AE_Offense] as 'Type of Offense', T0.[U_AE_Edate] as 'Expiry Date', T0.[U_AE_nno] as 'Notice Number',T0.[U_AE_fine] as 'Fine Amount',  T0.[U_AE_Submit] as 'Created By', T0.[U_AE_Status] as 'Status'  FROM [dbo].[@AE_TRAFFICO]  T0 WHERE T0.[U_AE_Odate] >= '" & GateDate(oform.Items.Item("10").Specific.String, Ocompany_TA) & "' and T0.[U_AE_Odate] <= '" & GateDate(oform.Items.Item("4").Specific.String, Ocompany_TA) & "' and T0.[U_AE_Status] = '" & Trim(oform.Items.Item("6").Specific.value) & "'"
                                Else
                                    Oapplication_TA.StatusBar.SetText("Invalid Date parameter ............. !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                Dim ogrid As SAPbouiCOM.Grid = oform.Items.Item("8").Specific
                                oform.Items.Item("8").Enabled = False
                                Try
                                    oform.DataSources.DataTables.Add("Offence")
                                Catch ex As Exception

                                End Try
                                oform.DataSources.DataTables.Item("Offence")
                                ogrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto
                                oform.DataSources.DataTables.Item(0).ExecuteQuery(sqlstr)
                                ogrid.DataTable = oform.DataSources.DataTables.Item("Offence")
                                ogrid.AutoResizeColumns()

                            Catch ex As Exception

                            End Try
                            Exit Sub
                        End If

                    End If
                End If
            End If

            If pVal.FormUID = "TPOD" Then

                If pVal.Before_Action = True Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                        If pVal.ItemUID = "106" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm

                                Dim TPOD_Thread As Threading.Thread

                                TPOD_Thread = New System.Threading.Thread(AddressOf Class_Report.Report_CallingFunction)
                                Class_Report.oApplication = Oapplication_TA
                                Class_Report.oCompany = Ocompany_TA
                                Class_Report.Report_Name = "AE_RP006_SubmissionReport.rpt"
                                Class_Report.Report_Parameter = "@DocNum"
                                Class_Report.Docnum = oform.Items.Item("Item_14").Specific.string
                                Class_Report.Report_Title = "Submission Report"
                                If TPOD_Thread.IsAlive Then
                                    Oapplication_TA.MessageBox("Report is already open....")
                                Else
                                    TPOD_Thread.TrySetApartmentState(Threading.ApartmentState.STA)
                                    Oapplication_TA.StatusBar.SetText("Submission Report Opening in process ......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    TPOD_Thread.Start()
                                End If

                            Catch ex As Exception

                            End Try


                        End If

                        If pVal.ItemUID = "111" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                                Dim docnum = oform.Items.Item("1000001").Specific.String
                                If docnum <> "" Then

                                    LoadFromXML("SelfDriving_Booking.srf", Oapplication_TA)
                                    oform = Oapplication_TA.Forms.Item("SDB")
                                    oform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    oform.Items.Item("Item_14").Enabled = True
                                    oform.Items.Item("Item_14").Specific.String = docnum
                                    oform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    ' oform.Items.Item("1000001").Specific.active = True


                                    Dim ocombobutton As SAPbouiCOM.ButtonCombo = oform.Items.Item("210").Specific
                                    ocombobutton.ValidValues.Add("Copy To", "Copy To A/R Invoice")

                                    Dim ooption As SAPbouiCOM.OptionBtn = oform.Items.Item("195").Specific
                                    Dim ooption1 As SAPbouiCOM.OptionBtn = oform.Items.Item("196").Specific

                                    Dim ooption2 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_104").Specific
                                    Dim ooption3 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_105").Specific
                                    Dim ooption4 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_106").Specific

                                    ooption1.GroupWith("195")

                                    ooption3.GroupWith("Item_104")
                                    ooption4.GroupWith("Item_104")
                                    ' oform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
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
                                    oform.Items.Item("Item_14").Enabled = False
                                    'oform.Items.Item("Item_22").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    'oform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                    oform.Visible = True
                                End If

                            Catch ex As Exception

                            End Try
                            Exit Sub
                        End If

                        If pVal.ItemUID = "1" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            Try

                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm

                                If oform.Items.Item("179").Specific.value = "" Then
                                    oform.Items.Item("179").Specific.active = True
                                    Oapplication_TA.StatusBar.SetText("Category should not be Empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                If oform.Items.Item("Item_30").Specific.value = "" Then
                                    oform.Items.Item("Item_30").Specific.active = True
                                    Oapplication_TA.StatusBar.SetText("Driver Name should not be Empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                If oform.Items.Item("68").Specific.value = "" Then
                                    oform.Items.Item("68").Specific.active = True
                                    Oapplication_TA.StatusBar.SetText("Vehicle Number should not be Empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                            Catch ex As Exception
                            End Try
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                            Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                            Select Case pVal.ItemUID

                                Case "Item_20"
                                    oform.PaneLevel = 1
                                Case "Item_21"
                                    oform.PaneLevel = 2
                                Case "Item_25"
                                    oform.PaneLevel = 6
                            End Select

                        End If

                        If pVal.ItemUID = "Item_209" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.Item("TPOD")
                                Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_207").Specific

                                Dim orset As SAPbobsCOM.Recordset = Ocompany_TA.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                orset.DoQuery("SELECT attachpath from oadp")
                                showOpenFileDialog()
                                If FileName <> "" Then
                                    ' For Each file__1 As String In Directory.GetFiles(stFilePathAndName)
                                    ' Dim dest As String = Path.Combine(orset.Fields.Item("attachpath").Value, Path.GetFileName(file__1))
                                    Dim file = New FileInfo(stFilePathAndName)
                                    file.CopyTo(Path.Combine(orset.Fields.Item("attachpath").Value, file.Name), True)
                                    ' Next
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
                                    stFilePathAndName = ""
                                    'oMatrix.AutoResizeColumns()
                                End If

                                If oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If

                                oform.Freeze(True)
                                oMatrix.AutoResizeColumns()
                                oform.Freeze(False)

                            Catch ex As Exception
                                Oapplication_TA.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub
                        End If


                        If pVal.ItemUID = "Item_208" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.Item("TPOD")
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
                                Oapplication_TA.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End Try
                            Exit Sub
                        End If



                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.InnerEvent = False Then
                        If pVal.ItemUID = "85" Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                                If oform.Items.Item("85").Specific.String <> "" Then
                                    If oform.Items.Item("85").Specific.String.ToString.Length = "4" Then
                                        If IsDate(oform.Items.Item("85").Specific.String.ToString.Substring(0, 2) & ":" & oform.Items.Item("85").Specific.String.ToString.Substring(2, 2) & ":00") = False Then
                                            oform.Items.Item("85").Specific.active = True
                                            Oapplication_TA.StatusBar.SetText("Invalid Time format ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Try
                                        Else
                                            oform.Items.Item("85").Specific.String = oform.Items.Item("85").Specific.String.ToString.Substring(0, 2) & ":" & oform.Items.Item("85").Specific.String.ToString.Substring(2, 2)
                                        End If
                                    ElseIf oform.Items.Item("85").Specific.String.ToString.Length = "5" Then
                                        If IsDate(oform.Items.Item("85").Specific.String & ":00") = False Then
                                            oform.Items.Item("85").Specific.active = True
                                            Oapplication_TA.StatusBar.SetText("Invalid Time format ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Try
                                        End If
                                    Else

                                        oform.Items.Item("85").Specific.active = True
                                        Oapplication_TA.StatusBar.SetText("Invalid Time format ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Try

                                    End If
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
                        Dim oForm As SAPbouiCOM.Form = Oapplication_TA.Forms.Item(FormUID)
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                        Try
                            If oCFLEvento.BeforeAction = False Then
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim oedit1 As SAPbouiCOM.EditText
                                oDataTable = oCFLEvento.SelectedObjects
                                If pVal.ItemUID = "Item_30" Then 'Order By
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

                                    Dim oAdress As SAPbobsCOM.Recordset = Ocompany_TA.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oAdress.DoQuery("select T1.cardcode , (ISNULL(T2.Street,'') +  ', ' + ISNULL(T2.Block,'') + char(13) + ISNULL(T2.City,'') + ', ' + ISNULL(T2.StreetNo,'') + CHAR(13) + isnull(T3.[Name], '') + ' ' + isnull(T2.ZipCode,'')) as 'Address' " & _
"from OCRD T1 " & _
"LEFT JOIN CRD1 T2 ON T1.CardCode=T2.CardCode and [AdresType]='B' " & _
"LEFT JOIN OCRY T3 ON T2.Country=T3.Code where T1.cardcode = '" & oDataTable.GetValue("CardCode", 0) & "' " & _
"group by T1.cardcode , ISNULL(T2.Street,'') +  ', ' + ISNULL(T2.Block,'') + char(13) + ISNULL(T2.City,'') + ', ' + ISNULL(T2.StreetNo,'') + CHAR(13) + isnull(T3.[Name], '') + ' ' + isnull(T2.ZipCode,'') ")

                                    oForm.Items.Item("Item_58").Specific.string = oDataTable.GetValue("CardName", 0)
                                    oForm.Items.Item("Item_60").Specific.string = oAdress.Fields.Item("Address").Value  'oDataTable.GetValue("BillToDef", 0) & ", " & oDataTable.GetValue("Address", 0) & ", " & oDataTable.GetValue("City", 0) & ", " & oDataTable.GetValue("Country", 0) & ", " & oDataTable.GetValue("ZipCode", 0)
                                    oForm.Items.Item("Item_64").Specific.string = oDataTable.GetValue("Cellular", 0)
                                    oForm.Items.Item("Item_62").Specific.string = oDataTable.GetValue("CntctPrsn", 0)
                                    oForm.Items.Item("Item_57").Specific.string = oDataTable.GetValue("CardCode", 0)

                                ElseIf pVal.ItemUID = "100" Then 'Issued By
                                    oForm.Items.Item("181").Specific.string = oDataTable.GetValue("firstName", 0) & " " & oDataTable.GetValue("lastName", 0)
                                    oForm.Items.Item("100").Specific.string = oDataTable.GetValue("empID", 0)

                                ElseIf pVal.ItemUID = "101" Then 'Issued By
                                    oForm.Items.Item("Item_7").Specific.string = oDataTable.GetValue("firstName", 0) & " " & oDataTable.GetValue("lastName", 0)
                                    oForm.Items.Item("101").Specific.string = oDataTable.GetValue("empID", 0)

                                ElseIf pVal.ItemUID = "1000001" Then 'Issued By

                                    Dim orset As SAPbobsCOM.Recordset = Ocompany_TA.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim orset1 As SAPbobsCOM.Recordset = Ocompany_TA.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    ' ------------------ New ----------------------
                                    orset.DoQuery("SELECT T0.[U_AE_DName] FROM [dbo].[@AE_SBOOKING]  T0 WHERE T0.[DocEntry] = '" & oDataTable.GetValue("DocEntry", 0) & "' union all SELECT T0.[U_AE_DName1] as 'U_AE_DName' FROM [dbo].[@AE_SBOOKING]  T0 WHERE T0.[DocEntry] = '" & oDataTable.GetValue("DocEntry", 0) & "'")
                                    orset1.DoQuery("SELECT T0.[U_AE_Bcode], T0.[U_AE_Bname], T0.[U_AE_Address], T0.[U_AE_Atten], T0.[U_AE_Cno] FROM [dbo].[@AE_SBOOKING]  T0 WHERE T0.[DocEntry] = '" & oDataTable.GetValue("DocEntry", 0) & "'")
                                    Dim ocombo As SAPbouiCOM.ComboBox = oForm.Items.Item("1000011").Specific
                                   
                                    Try

                                        oForm.Items.Item("Item_57").Specific.string = orset1.Fields.Item("U_AE_Bcode").Value
                                    Catch ex As Exception
                                    End Try

                                    oForm.Items.Item("Item_58").Specific.string = orset1.Fields.Item("U_AE_Bname").Value
                                    oForm.Items.Item("Item_60").Specific.string = orset1.Fields.Item("U_AE_Address").Value
                                    oForm.Items.Item("Item_62").Specific.string = orset1.Fields.Item("U_AE_Atten").Value
                                    oForm.Items.Item("Item_64").Specific.string = orset1.Fields.Item("U_AE_Cno").Value


                                    For mjs As Integer = ocombo.ValidValues.Count - 1 To 0 Step -1
                                        ocombo.ValidValues.Remove(mjs, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next mjs

                                    For mjs As Integer = 1 To orset.RecordCount
                                        ocombo.ValidValues.Add(orset.Fields.Item("U_AE_DName").Value, "")
                                        orset.MoveNext()
                                    Next mjs
                                    ocombo.ValidValues.Add("-", "")
                                    ocombo.Select("-")
                                    oForm.Items.Item("1000001").Specific.string = oDataTable.GetValue("DocEntry", 0)

                                    '----------------------------------------------------


                                ElseIf pVal.ItemUID = "102" Then 'Issued By
                                    oForm.Items.Item("Item_12").Specific.string = oDataTable.GetValue("firstName", 0) & " " & oDataTable.GetValue("lastName", 0)
                                    oForm.Items.Item("102").Specific.string = oDataTable.GetValue("empID", 0)

                                ElseIf pVal.ItemUID = "68" Then 'Issued By
                                    oForm.Items.Item("1000010").Specific.string = oDataTable.GetValue("ItemName", 0)
                                    oForm.Items.Item("73").Specific.string = oDataTable.GetValue("U_AE_MODEL", 0)
                                    oForm.Items.Item("75").Specific.string = oDataTable.GetValue("U_AE_YEAR_Make", 0)
                                    oForm.Items.Item("77").Specific.string = oDataTable.GetValue("U_AE_CHASSIS_NO", 0)
                                    oForm.Items.Item("68").Specific.string = oDataTable.GetValue("ItemCode", 0)
                                End If

                            End If
                        Catch ex As Exception
                        End Try
                    End If
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then

                        If pVal.ItemUID = "179" Then
                            Dim oform As SAPbouiCOM.Form
                            Try
                                oform = Oapplication_TA.Forms.ActiveForm
                                oform.Freeze(True)
                                If Trim(oform.Items.Item("179").Specific.value) = "Staff (CD or Errand)" Then
                                    oform.Items.Item("1000001").Specific.String = ""
                                    oform.Items.Item("Item_16").Specific.active = True
                                    oform.Items.Item("1000001").Enabled = False
                                    ' ------------------ New ----------------------
                                    oform.Items.Item("Item_30").Visible = True
                                    oform.Items.Item("Item_31").Visible = True
                                    oform.Items.Item("1000011").Visible = False
                                    '-----------------------------------------------
                                    oform.Items.Item("103").Visible = True
                                    oform.Items.Item("Item_30").Specific.String = ""
                                    oform.Items.Item("103").Specific.String = ""
                                    Dim ocombo As SAPbouiCOM.ComboBox = oform.Items.Item("1000011").Specific
                                    For mjs As Integer = ocombo.ValidValues.Count - 1 To 0 Step -1
                                        ocombo.ValidValues.Remove(mjs, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next mjs
                                    ocombo.ValidValues.Add("", "")
                                    ocombo.Select("")
                                    oform.Items.Item("Item_31").Specific.String = ""

                                Else
                                    oform.Items.Item("103").Specific.String = ""
                                    oform.Items.Item("Item_16").Specific.active = True
                                    oform.Items.Item("103").Visible = False
                                    ' ------------------ New ----------------------
                                    oform.Items.Item("Item_30").Visible = False
                                    oform.Items.Item("Item_31").Visible = False
                                    oform.Items.Item("1000011").Visible = True
                                    '--------------------------------------------
                                    oform.Items.Item("1000001").Enabled = True

                                    oform.Items.Item("Item_30").Specific.String = ""
                                    oform.Items.Item("Item_31").Specific.String = ""
                                    oform.Items.Item("103").Specific.String = ""
                                    Dim ocombo As SAPbouiCOM.ComboBox = oform.Items.Item("1000011").Specific
                                    For mjs As Integer = ocombo.ValidValues.Count - 1 To 0 Step -1
                                        ocombo.ValidValues.Remove(mjs, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next mjs


                                End If
                                oform.Items.Item("Item_34").Specific.String = ""
                                oform.Items.Item("Item_36").Specific.String = ""
                                oform.Items.Item("Item_38").Specific.String = ""
                                oform.Items.Item("Item_40").Specific.String = ""
                                oform.Items.Item("Item_42").Specific.String = ""
                                oform.Items.Item("Item_44").Specific.String = ""
                                oform.Items.Item("Item_49").Specific.String = ""
                                oform.Items.Item("Item_50").Specific.String = ""
                                oform.Items.Item("Item_51").Specific.String = ""

                                oform.Items.Item("Item_52").Specific.String = ""
                                oform.Items.Item("Item_54").Specific.String = ""
                                oform.Freeze(False)
                            Catch ex As Exception
                                oform.Freeze(False)
                            End Try
                        End If

                        ' ------------------ New ----------------------

                        If pVal.ItemUID = "1000011" Then
                            Dim oform As SAPbouiCOM.Form
                            Try
                                oform = Oapplication_TA.Forms.ActiveForm
                                Dim orset As SAPbobsCOM.Recordset = Ocompany_TA.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oform.Freeze(True)
                                If oform.Items.Item("1000011").Specific.selected.value <> "-" Then
                                    orset.DoQuery("SELECT T0.[U_AE_Dadd], T0.[U_AE_Dcno], T0.[U_AE_Occuption], T0.[U_AE_Nation], T0.[U_AE_DOB], T0.[U_AE_License], T0.[U_AE_Pissue], T0.[U_AE_Exdate], T0.[U_AE_Passno], T0.[U_AE_Pissuepno], T0.[U_AE_Pexdate] FROM [dbo].[@AE_SBOOKING]  T0 WHERE T0.[DocEntry] = '" & oform.Items.Item("1000001").Specific.string & "' and T0.[U_AE_DName] = '" & oform.Items.Item("1000011").Specific.selected.value & "'")
                                    oform.Items.Item("Item_34").Specific.string = orset.Fields.Item("U_AE_Dadd").Value
                                    oform.Items.Item("Item_36").Specific.string = orset.Fields.Item("U_AE_Dcno").Value
                                    oform.Items.Item("Item_38").Specific.string = orset.Fields.Item("U_AE_Occuption").Value
                                    oform.Items.Item("Item_40").Specific.string = orset.Fields.Item("U_AE_Nation").Value
                                    oform.Items.Item("Item_42").Specific.string = Format(orset.Fields.Item("U_AE_DOB").Value, "dd/MM/yyyy")
                                    oform.Items.Item("Item_44").Specific.string = orset.Fields.Item("U_AE_License").Value
                                    oform.Items.Item("Item_49").Specific.string = orset.Fields.Item("U_AE_Pissue").Value
                                    oform.Items.Item("Item_50").Specific.string = Format(orset.Fields.Item("U_AE_Exdate").Value, "dd/MM/yyyy")
                                    oform.Items.Item("Item_51").Specific.string = orset.Fields.Item("U_AE_Passno").Value
                                    oform.Items.Item("Item_52").Specific.string = orset.Fields.Item("U_AE_Pissuepno").Value
                                    oform.Items.Item("Item_54").Specific.string = Format(orset.Fields.Item("U_AE_Pexdate").Value, "dd/MM/yyyy")
                                Else
                                    oform.Items.Item("Item_34").Specific.string = ""
                                    oform.Items.Item("Item_36").Specific.string = ""
                                    oform.Items.Item("Item_38").Specific.string = ""
                                    oform.Items.Item("Item_40").Specific.string = ""
                                    oform.Items.Item("Item_42").Specific.string = ""
                                    oform.Items.Item("Item_44").Specific.string = ""
                                    oform.Items.Item("Item_49").Specific.string = ""
                                    oform.Items.Item("Item_50").Specific.string = ""
                                    oform.Items.Item("Item_51").Specific.string = ""
                                    oform.Items.Item("Item_52").Specific.string = ""
                                    oform.Items.Item("Item_54").Specific.string = ""
                                End If

                                oform.Freeze(False)

                            Catch ex As Exception
                                oform.Freeze(False)
                            End Try


                        End If

                        '----------------------------------------------

                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        If pVal.ItemUID = "1" And pVal.Action_Success = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                            Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                            ''oform.Close()
                            ''LoadFromXML("TrafficParkingOffense.srf", Oapplication_TA)
                            oform = Oapplication_TA.Forms.Item("TPOD")
                            ''oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                            Dim Tmp_val As String = oform.BusinessObject.GetNextSerialNumber(NextSerialNo(Ocompany_TA, Oapplication_TA, "AE_TrafficO"))
                            oform.Items.Item("Item_14").Specific.String = Tmp_val
                            oform.Items.Item("105").Specific.select("Open")
                            oform.Items.Item("Item_20").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            'oform.Visible = True
                            oform.Items.Item("Item_16").Specific.String = Now.Date 'Format(Now.Date, "dd MMM yyyy")
                            oform.Items.Item("1000001").Enabled = False
                            oform.DataBrowser.BrowseBy = "Item_14"
                            oform.Items.Item("108").Specific.String = Ocompany_TA.UserName

                            oform.Items.Item("109").Specific.String = Company_Name
                            oform.PaneLevel = 1

                        End If



                        '--------------- New -------------------------------
                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm

                            Try
                                If TAO_Flag = True Then
                                    oform.Freeze(True)
                                    If Trim(oform.Items.Item("179").Specific.value) = "Staff (CD or Errand)" Then
                                        oform.Items.Item("Item_30").Visible = True
                                        oform.Items.Item("Item_31").Visible = True
                                        oform.Items.Item("1000011").Visible = False
                                    Else
                                        oform.Items.Item("Item_30").Visible = False
                                        oform.Items.Item("Item_31").Visible = False
                                        oform.Items.Item("1000011").Visible = True
                                    End If
                                    oform.Freeze(False)
                                    TAO_Flag = False
                                End If
                            Catch ex As Exception
                                oform.Freeze(False)
                            End Try
                        End If
                        ' ---------------------------------------------------------------------------------


                    End If
                End If
            End If


            If pVal.FormUID = "VT" Then

                If pVal.Before_Action = False Then

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then

                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pVal
                        Dim sCFL_ID As String
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        Dim oForm As SAPbouiCOM.Form = Oapplication_TA.Forms.Item(FormUID)
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                        If oCFLEvento.BeforeAction = False Then
                            Dim oDataTable As SAPbouiCOM.DataTable
                            oDataTable = oCFLEvento.SelectedObjects

                            If pVal.ItemUID = "Item_22" Then 'Billing Code
                                oForm.Items.Item("Item_24").Specific.string = oDataTable.GetValue("firstName", 0) & " " & oDataTable.GetValue("lastName", 0)
                                oForm.Items.Item("Item_22").Specific.string = oDataTable.GetValue("empID", 0)
                            End If
                            If pVal.ItemUID = "1000002" Then
                                oForm.Items.Item("37").Specific.string = oDataTable.GetValue("DocEntry", 0)
                                oForm.Items.Item("1000002").Specific.string = oDataTable.GetValue("DocNum", 0)
                            End If

                        End If

                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Action_Success = True Then
                            Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                            oform.Close()
                            Try
                                oform = Oapplication_TA.Forms.Item("VTF")
                                oform.Items.Item("Item_1").Specific.String = ""
                                oform.Items.Item("Item_2").Specific.String = ""
                            Catch ex As Exception
                            End Try
                        End If
                    End If
                End If
                'tt
                If pVal.Before_Action = True Then

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.InnerEvent = False Then
                        If pVal.ItemUID = "Item_20" Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                                Dim Time_check As String = ""
                                If oform.Items.Item("Item_20").Specific.String <> "" Then
                                    If oform.Items.Item("Item_20").Specific.String.ToString.Length = "4" Then
                                        Time_check = oform.Items.Item("Item_20").Specific.String.ToString.Substring(0, 2) & ":" & oform.Items.Item("Item_20").Specific.String.ToString.Substring(2, 2)
                                    ElseIf oform.Items.Item("Item_20").Specific.String.ToString.Length = "5" Then
                                        Time_check = oform.Items.Item("Item_20").Specific.String
                                    End If
                                    If IsDate(Time_check & ":00") = False Then
                                        oform.Items.Item("Item_20").Specific.active = True
                                        Oapplication_TA.StatusBar.SetText("Invalid Time Format ............ ,", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Try
                                    Else
                                        oform.Items.Item("Item_20").Specific.String = Time_check
                                    End If
                                End If
                            Catch ex As Exception

                            End Try
                            Exit Sub
                        End If

                    End If
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then




                        If pVal.ItemUID = "41" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            Try
                                Dim oForm As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                                Dim docnum = oForm.Items.Item("1000002").Specific.String
                                If docnum <> "" Then

                                    LoadFromXML("SelfDriving_Booking.srf", Oapplication_TA)
                                    oForm = Oapplication_TA.Forms.Item("SDB")
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    oForm.Items.Item("Item_14").Enabled = True
                                    oForm.Items.Item("Item_14").Specific.String = docnum
                                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    'oForm.Items.Item("1000001").Specific.active = True
                                    Dim ocombobutton As SAPbouiCOM.ButtonCombo = oForm.Items.Item("210").Specific
                                    ocombobutton.ValidValues.Add("Copy To", "Copy To A/R Invoice")

                                    Dim ooption As SAPbouiCOM.OptionBtn = oForm.Items.Item("195").Specific
                                    Dim ooption1 As SAPbouiCOM.OptionBtn = oForm.Items.Item("196").Specific

                                    Dim ooption2 As SAPbouiCOM.OptionBtn = oForm.Items.Item("Item_104").Specific
                                    Dim ooption3 As SAPbouiCOM.OptionBtn = oForm.Items.Item("Item_105").Specific
                                    Dim ooption4 As SAPbouiCOM.OptionBtn = oForm.Items.Item("Item_106").Specific

                                    ooption1.GroupWith("195")

                                    ooption3.GroupWith("Item_104")
                                    ooption4.GroupWith("Item_104")
                                    ' oform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                    If ooption2.Selected = True Then
                                        oForm.Items.Item("Item_108").Specific.caption = "Daily Rates"
                                        oForm.Items.Item("Item_110").Specific.caption = "Number of Days"
                                        oForm.Items.Item("Item_114").Specific.caption = "PAI Per Day"
                                        oForm.Items.Item("Item_122").Specific.caption = "CDW Per Day"
                                    ElseIf ooption3.Selected = True Then
                                        oForm.Items.Item("Item_108").Specific.caption = "Weekly Rates"
                                        oForm.Items.Item("Item_110").Specific.caption = "Number of Days"
                                        oForm.Items.Item("Item_114").Specific.caption = "PAI Per Month"
                                        oForm.Items.Item("Item_122").Specific.caption = "CDW Per Month"
                                    ElseIf ooption4.Selected = True Then
                                        oForm.Items.Item("Item_108").Specific.caption = "Monthly Rates"
                                        oForm.Items.Item("Item_110").Specific.caption = "Number of Months"
                                        oForm.Items.Item("Item_114").Specific.caption = "PAI Per Month"
                                        oForm.Items.Item("Item_122").Specific.caption = "CDW Per Month"
                                    End If
                                    oForm.Items.Item("Item_14").Enabled = False

                                    oForm.Items.Item("Item_22").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oForm.Visible = True
                                End If

                            Catch ex As Exception

                            End Try

                        End If

                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                                Dim orset As SAPbobsCOM.Recordset = Ocompany_TA.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim loc As String = ""
                                Dim opt1, opt2, opt3, opt4, opt5, opt6, opt7, opt8, opt9 As SAPbouiCOM.OptionBtn

                                opt1 = oform.Items.Item("Item_7").Specific
                                opt2 = oform.Items.Item("Item_8").Specific
                                opt3 = oform.Items.Item("Item_9").Specific
                                opt4 = oform.Items.Item("Item_10").Specific
                                opt5 = oform.Items.Item("Item_11").Specific
                                opt6 = oform.Items.Item("Item_12").Specific
                                opt7 = oform.Items.Item("Item_13").Specific
                                opt8 = oform.Items.Item("Item_14").Specific
                                opt9 = oform.Items.Item("Item_15").Specific


                                If oform.Items.Item("39").Specific.String = "Loc R" Then
                                    If oform.Items.Item("1000002").Specific.String = "" Then
                                        oform.Items.Item("1000002").Specific.active = True
                                        Oapplication_TA.StatusBar.SetText("Rental Agreement should not be empty ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If


                                If oform.Items.Item("Item_5").Specific.String = "" Then
                                    oform.Items.Item("Item_5").Specific.active = True
                                    Oapplication_TA.StatusBar.SetText("Mileage should not be empty ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                End If


                                If opt1.Selected = False And opt2.Selected = False And opt3.Selected = False And opt4.Selected = False And opt5.Selected = False And opt6.Selected = False And opt7.Selected = False And opt8.Selected = False And opt9.Selected = False Then
                                    Oapplication_TA.StatusBar.SetText("Petrol should not be empty ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                End If


                                If oform.Items.Item("Item_18").Specific.String = "" Then
                                    oform.Items.Item("Item_18").Specific.active = True
                                    Oapplication_TA.StatusBar.SetText("Date should not be empty ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                If oform.Items.Item("Item_20").Specific.String = "" Then
                                    oform.Items.Item("Item_20").Specific.active = True
                                    Oapplication_TA.StatusBar.SetText("Time should not be empty ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                If oform.Items.Item("Item_22").Specific.String = "" Then
                                    oform.Items.Item("Item_22").Specific.active = True
                                    Oapplication_TA.StatusBar.SetText("Employee No should not be empty ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                If oform.Items.Item("Item_24").Specific.String = "" Then
                                    oform.Items.Item("Item_24").Specific.active = True
                                    Oapplication_TA.StatusBar.SetText("Name should not be empty ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                                If oform.Items.Item("39").Specific.String = "Loc D" Or oform.Items.Item("39").Specific.String = "Loc C" Then
                                    If oform.Items.Item("Item_26").Specific.String = "" Then
                                        oform.Items.Item("Item_26").Specific.active = True
                                        Oapplication_TA.StatusBar.SetText("Remarks should not be empty ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If

                                'If oform.Items.Item("39").Specific.String = "Loc 1" Or oform.Items.Item("39").Specific.String = "Loc 2" Or oform.Items.Item("39").Specific.String = "Loc 3" Or oform.Items.Item("39").Specific.String = "Loc 4" Or oform.Items.Item("39").Specific.String = "Loc 5" Then
                                '    loc = "Y"
                                'ElseIf oform.Items.Item("39").Specific.String = "Loc R" Or oform.Items.Item("39").Specific.String = "Loc C" Or oform.Items.Item("39").Specific.String = "Loc D" Then
                                '    loc = "N"
                                'End If
                                'orset.DoQuery("update OITM set U_AE_IN = '" & loc & "' where [ItemCode] = '" & oform.Items.Item("Item_3").Specific.String & "'")
                            Catch ex As Exception

                            End Try
                        End If
                    End If
                End If
            End If

            If pVal.FormUID = "VTR1" Then

                If pVal.Before_Action = False Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then

                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pVal
                        Dim sCFL_ID As String
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        Dim oForm As SAPbouiCOM.Form = Oapplication_TA.Forms.Item(FormUID)
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)

                        If oCFLEvento.BeforeAction = False Then
                            Dim oDataTable As SAPbouiCOM.DataTable
                            oDataTable = oCFLEvento.SelectedObjects
                            Try

                                If pVal.ItemUID = "9" Then
                                    oForm.Items.Item("9").Specific.string = oDataTable.GetValue("ItemCode", 0)
                                End If

                                If pVal.ItemUID = "11" Then
                                    oForm.Items.Item("12").Specific.string = oDataTable.GetValue("firstName", 0) & " " & oDataTable.GetValue("lastName", 0)
                                    oForm.Items.Item("11").Specific.string = oDataTable.GetValue("empID", 0)
                                End If

                                If pVal.ItemUID = "13" Then
                                    oForm.Items.Item("13").Specific.string = oDataTable.GetValue("DocNum", 0)
                                End If
                            Catch ex As Exception

                            End Try

                        End If
                    End If


                ElseIf pVal.Before_Action = True Then


                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED Then
                        If pVal.ItemUID = "14" And pVal.ColUID = "Rental Agreement" Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                                Dim ogrid As SAPbouiCOM.Grid = oform.Items.Item("14").Specific
                                Dim docNum As String = ogrid.DataTable.GetValue("Rental Agreement", pVal.Row)

                                LoadFromXML("SelfDriving_Booking.srf", Oapplication_TA)
                                oform = Oapplication_TA.Forms.Item("SDB")
                                oform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                oform.Items.Item("Item_14").Enabled = True
                                oform.Items.Item("Item_14").Specific.String = docNum
                                oform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Dim ocombobutton As SAPbouiCOM.ButtonCombo = oform.Items.Item("210").Specific
                                ocombobutton.ValidValues.Add("Copy To", "Copy To A/R Invoice")
                                oform.Visible = True

                                Dim ooption As SAPbouiCOM.OptionBtn = oform.Items.Item("195").Specific
                                Dim ooption1 As SAPbouiCOM.OptionBtn = oform.Items.Item("196").Specific

                                Dim ooption2 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_104").Specific
                                Dim ooption3 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_105").Specific
                                Dim ooption4 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_106").Specific

                                ooption1.GroupWith("195")

                                ooption3.GroupWith("Item_104")
                                ooption4.GroupWith("Item_104")
                                ' oform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE



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
                                oform.Items.Item("Item_14").Enabled = False
                                oform.Items.Item("Item_22").Click(SAPbouiCOM.BoCellClickType.ct_Regular)


                            Catch ex As Exception

                            End Try
                        End If
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        If pVal.ItemUID = "8" And pVal.Before_Action = True Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                                Dim sqlstr, vehicleno, employee, RA As String

                                If oform.Items.Item("9").Specific.String <> "" Then
                                    vehicleno = oform.Items.Item("9").Specific.String
                                Else
                                    vehicleno = "%"
                                End If

                                If oform.Items.Item("11").Specific.String <> "" Then
                                    employee = oform.Items.Item("11").Specific.String
                                Else
                                    employee = "%"
                                End If

                                If oform.Items.Item("13").Specific.String <> "" Then
                                    RA = oform.Items.Item("13").Specific.String
                                Else
                                    RA = "%"
                                End If

                                Dim sFromDate As String = oform.Items.Item("10").Specific.String
                                Dim sToDate As String = oform.Items.Item("4").Specific.String


                                If oform.Items.Item("10").Specific.String = "" And oform.Items.Item("4").Specific.String = "" Then
                                    sqlstr = "SELECT T0.[U_AE_Date] as 'Date', T0.[U_AE_Time] as 'Time', T0.[U_AE_Vno] as 'Vehicle No', case when [U_AE_Petrol] = '1' then 'Empty' when [U_AE_Petrol] = '2' then '1/8' when [U_AE_Petrol] = '3' then '1/4' when [U_AE_Petrol] = '4' then '3/8' when [U_AE_Petrol] = '5' then '1/2' when [U_AE_Petrol] = '6' then '5/8' when [U_AE_Petrol] = '7' then '3/4' when [U_AE_Petrol] = '8' then '7/8' when [U_AE_Petrol] = '9' then 'Full' end as 'Petrol', replace(convert(varchar,convert(Money, T0.[U_AE_Mileage]),1),'.00','') as 'Mileage', T0.[U_AE_Loc] as 'Location', T0.[U_AE_RA] as 'Rental Agreement', T0.[U_AE_Remark] as 'Remarks', T0.[U_AE_Name] as 'Employee Name' FROM [dbo].[@AE_VTRACK]  T0 where T0.[U_AE_Vno] like '" & vehicleno & "' and  T0.[U_AE_NRIC] like '" & employee & "' and  isnull(t0.[U_AE_RA],'') like '" & RA & "'"
                                Else
                                    sqlstr = "SELECT T0.[U_AE_Date] as 'Date', T0.[U_AE_Time] as 'Time', T0.[U_AE_Vno] as 'Vehicle No', case when [U_AE_Petrol] = '1' then 'Empty' when [U_AE_Petrol] = '2' then '1/8' when [U_AE_Petrol] = '3' then '1/4' when [U_AE_Petrol] = '4' then '3/8' when [U_AE_Petrol] = '5' then '1/2' when [U_AE_Petrol] = '6' then '5/8' when [U_AE_Petrol] = '7' then '3/4' when [U_AE_Petrol] = '8' then '7/8' when [U_AE_Petrol] = '9' then 'Full' end as 'Petrol', replace(convert(varchar,convert(Money, T0.[U_AE_Mileage]),1),'.00','') as 'Mileage', T0.[U_AE_Loc] as 'Location', T0.[U_AE_RA] as 'Rental Agreement', T0.[U_AE_Remark] as 'Remarks', T0.[U_AE_Name] as 'Employee Name' FROM [dbo].[@AE_VTRACK]  T0 WHERE T0.[U_AE_Date] >= '" & GateDate(oform.Items.Item("10").Specific.String, Ocompany_TA) & "' and  T0.[U_AE_Date] <= '" & GateDate(oform.Items.Item("4").Specific.String, Ocompany_TA) & "' and  T0.[U_AE_Vno] like '" & vehicleno & "' and  T0.[U_AE_NRIC] like '" & employee & "' and  isnull(t0.[U_AE_RA],'') like '" & RA & "'"
                                End If

                                Dim ogrid As SAPbouiCOM.Grid = oform.Items.Item("14").Specific
                                oform.Items.Item("14").Enabled = False
                                Try
                                    oform.DataSources.DataTables.Add("VT")
                                Catch ex As Exception

                                End Try
                                ogrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto
                                oform.DataSources.DataTables.Item(0).ExecuteQuery(sqlstr)
                                ogrid.DataTable = oform.DataSources.DataTables.Item("VT")
                                Dim ocolumns As SAPbouiCOM.EditTextColumn = ogrid.Columns.Item("Rental Agreement")
                                ocolumns.LinkedObjectType = "AE_Sbooking"



                                ogrid.AutoResizeColumns()

                            Catch ex As Exception

                            End Try
                            Exit Sub
                        End If
                    End If
                End If



            End If

            If pVal.FormUID = "VTR2" Then
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    If pVal.ItemUID = "8" And pVal.Before_Action = True Then
                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                            Dim IN1 As String
                            Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("7").Specific
                            If oform.Items.Item("10").Specific.String = "" Then
                                oform.Items.Item("10").Specific.active = True
                                Oapplication_TA.StatusBar.SetText("Date From should not be empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End If

                            If oform.Items.Item("4").Specific.String = "" Then
                                oform.Items.Item("4").Specific.active = True
                                Oapplication_TA.StatusBar.SetText("Date To Should not be empty ........... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End If

                            Dim Str_query As String = "SELECT T0.[U_AE_Vno] + ' - ' + left(convert(varchar, T0.[U_AE_Date],103),10) + ' - ' + T0.[U_AE_Time] as 'U_AE_Loc','' as 'U_AE_Name', U_AE_Remark as 'U_AE_Remark' " & _
"FROM [dbo].[@AE_VTRACK]  T0 inner join OITM T1 on T1.ItemCode = T0.U_AE_Vno " & _
"WHERE T0.[U_AE_Date]  >= '" & GateDate(oform.Items.Item("10").Specific.String, Ocompany_TA) & "' and T0.[U_AE_Date]  <= '" & GateDate(oform.Items.Item("10").Specific.String, Ocompany_TA) & "' and T0.[U_AE_Loc1] = 'Loc 5' " & _
"and T1.[U_AE_IN] = 'Y'  "

                            Try
                                oform.DataSources.DataTables.Add("@AE_VTRACK")
                            Catch ex As Exception

                            End Try

                            oform.DataSources.DataTables.Item("@AE_VTRACK").ExecuteQuery(Str_query)

                            omatrix.Clear()
                            oform.Items.Item("7").Specific.columns.item("V_0").databind.bind("@AE_VTRACK", "U_AE_Loc")
                            oform.Items.Item("7").Specific.columns.item("V_3").databind.bind("@AE_VTRACK", "U_AE_Remark")

                            oform.Items.Item("7").Specific.LoadFromDataSource()
                            oform.Items.Item("7").Specific.AutoResizeColumns()

                        Catch ex As Exception

                        End Try
                        Exit Sub
                    End If
                End If
            End If

            If pVal.FormUID = "VTF" Then


                If pVal.Before_Action = False Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then

                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pVal
                        Dim sCFL_ID As String
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        Dim oForm As SAPbouiCOM.Form = Oapplication_TA.Forms.Item(FormUID)
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                        Try

                            If oCFLEvento.BeforeAction = False Then
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim oedit1 As SAPbouiCOM.EditText
                                oDataTable = oCFLEvento.SelectedObjects
                                If pVal.ItemUID = "Item_1" Then 'Order By
                                    oForm.Items.Item("Item_2").Specific.string = oDataTable.GetValue("ItemName", 0)

                                    Dim SQL_String As String = "SELECT top(5) T0.[Docentry], T0.[U_AE_Date], T0.[U_AE_Time], " & _
"T0.[U_AE_Vno] , " & _
"case when [U_AE_Petrol] = '1' then 'Empty' " & _
"when [U_AE_Petrol] = '2' then '1/8' when [U_AE_Petrol] = '3' " & _
"then '1/4' when [U_AE_Petrol] = '4' then '3/8' " & _
"when [U_AE_Petrol] = '5' then '1/2' when [U_AE_Petrol] = '6' " & _
"then '5/8' when [U_AE_Petrol] = '7' then '3/4' " & _
"when [U_AE_Petrol] = '8' then '7/8' when [U_AE_Petrol] = '9' " & _
"then 'Full' end as 'U_AE_Petrol', " & _
"replace(convert(varchar,convert(Money, T0.[U_AE_Mileage]),1),'.00','') " & _
"as 'U_AE_Mileage', T0.[U_AE_Loc] , T0.[U_AE_Remark], T0.[U_AE_RA], " & _
"T0.[U_AE_Name] FROM [dbo].[@AE_VTRACK]  T0 " & _
"where T0.[U_AE_Vno] = '" & oDataTable.GetValue("ItemCode", 0) & "' order by T0.[U_AE_Date] desc,replace(U_AE_Time,':','') desc "

                                    Try
                                        oForm.DataSources.DataTables.Add("@AE_VTRACK")
                                    Catch ex As Exception

                                    End Try

                                    oForm.DataSources.DataTables.Item("@AE_VTRACK").ExecuteQuery(SQL_String)
                                    Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("20").Specific
                                    oMatrix.Clear()
                                    oForm.Items.Item("20").Specific.columns.item("V_0").databind.bind("@AE_VTRACK", "U_AE_Date")
                                    oForm.Items.Item("20").Specific.columns.item("V_7").databind.bind("@AE_VTRACK", "U_AE_Time")
                                    oForm.Items.Item("20").Specific.columns.item("V_6").databind.bind("@AE_VTRACK", "U_AE_Vno")
                                    oForm.Items.Item("20").Specific.columns.item("V_5").databind.bind("@AE_VTRACK", "U_AE_Petrol")
                                    oForm.Items.Item("20").Specific.columns.item("V_4").databind.bind("@AE_VTRACK", "U_AE_Mileage")
                                    oForm.Items.Item("20").Specific.columns.item("V_3").databind.bind("@AE_VTRACK", "U_AE_Loc")
                                    oForm.Items.Item("20").Specific.columns.item("V_2").databind.bind("@AE_VTRACK", "U_AE_Remark")
                                    oForm.Items.Item("20").Specific.columns.item("V_1").databind.bind("@AE_VTRACK", "U_AE_Name")
                                    oForm.Items.Item("20").Specific.columns.item("V_8").databind.bind("@AE_VTRACK", "Docentry")
                                    oForm.Items.Item("20").Specific.columns.item("V_9").databind.bind("@AE_VTRACK", "U_AE_RA")

                                    oForm.Items.Item("20").Specific.LoadFromDataSource()
                                    oForm.Items.Item("20").Specific.AutoResizeColumns()
                                    oForm.Items.Item("Item_1").Specific.string = oDataTable.GetValue("ItemCode", 0)
                                End If

                            End If
                        Catch ex As Exception
                            'Oapplication_TA.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            'BubbleEvent = False
                            'Exit Try
                        End Try
                    End If

                End If

                If pVal.Before_Action = True Then

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then


                        If pVal.ItemUID = "21" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then

                            Try

                                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.Item("VTF")
                                Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("20").Specific
                                Dim rowcount As Integer = oMatrix.RowCount

                                For mjs As Integer = 1 To rowcount
                                    If mjs <= oMatrix.RowCount Then
                                        If oMatrix.IsRowSelected(mjs) = True Then
                                            If Oapplication_TA.MessageBox("Do you want to delete ? ", 1, "Yes", "No") = 1 Then
                                                If VehicleTrackingDeletion(oMatrix.Columns.Item("V_8").Cells.Item(mjs).Specific.String) = False Then
                                                    BubbleEvent = False
                                                    Exit For
                                                Else
                                                    oMatrix.DeleteRow(mjs)
                                                    mjs = mjs - 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        Exit For
                                    End If
                                Next mjs

                              

                            Catch ex As Exception
                                Oapplication_TA.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub

                        End If

                        If pVal.ItemUID = "Item_9" Or pVal.ItemUID = "Item_10" Or pVal.ItemUID = "Item_11" Or pVal.ItemUID = "Item_12" Or _
                            pVal.ItemUID = "Item_13" Or pVal.ItemUID = "Item_14" Or pVal.ItemUID = "Item_15" Or pVal.ItemUID = "Item_16" Then

                            Try
                                Dim oForm As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                                Dim oButton As SAPbouiCOM.Button
                                Dim location, vcode, vname, LocID As String
                                If oForm.Items.Item("Item_1").Specific.String = "" Then
                                    oForm.Items.Item("Item_1").Specific.active = True
                                    Oapplication_TA.StatusBar.SetText("Vehicle Number should not be empty ........ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                oButton = oForm.Items.Item(pVal.ItemUID).Specific

                                Select Case pVal.ItemUID

                                    Case "Item_9"
                                        location = oButton.Caption
                                        LocID = "Loc 1"
                                    Case "Item_10"
                                        location = oButton.Caption
                                        LocID = "Loc 2"
                                    Case "Item_11"
                                        location = oButton.Caption
                                        LocID = "Loc 3"
                                    Case "Item_12"
                                        location = oButton.Caption
                                        LocID = "Loc 4"
                                    Case "Item_13"
                                        location = oButton.Caption
                                        LocID = "Loc 5"
                                    Case "Item_14"
                                        location = "R - [Rental]"
                                        LocID = "Loc R"
                                    Case "Item_15"
                                        location = "C - [Chauffer]"
                                        LocID = "Loc C"
                                    Case "Item_16"
                                        location = "D - [Errand or Operation]"
                                        LocID = "Loc D"
                                End Select

                                vcode = oForm.Items.Item("Item_1").Specific.String
                                vname = oForm.Items.Item("Item_2").Specific.String

                                'If opt.Selected = True Then
                                '    location = "IN"
                                'Else
                                '    location = "OUT"
                                'End If


                                LoadFromXML("Vehicle_TrackingDetails.srf", Oapplication_TA)
                                oForm = Oapplication_TA.Forms.Item("VT")
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                                oForm.Visible = True

                                If pVal.ItemUID <> "Item_14" Then
                                    oForm.Items.Item("1000001").Visible = False
                                    oForm.Items.Item("1000002").Visible = False
                                    oForm.Items.Item("41").Visible = False
                                End If
                                oForm.Items.Item("Item_1").Specific.String = location
                                oForm.Items.Item("Item_3").Specific.String = vcode
                                oForm.Items.Item("38").Specific.String = vname
                                oForm.Items.Item("39").Specific.String = LocID
                                oForm.Items.Item("35").Specific.String = Format(Now.Date, "dd/MM/yyyy")
                                oForm.Items.Item("Item_18").Specific.String = Format(Now.Date, "dd/MM/yyyy")
                                oForm.Items.Item("Item_20").Specific.String = Format(DateTime.Now, "HH:mm")
                                oForm.Items.Item("Item_5").Specific.active = True

                                Dim SQL_String As String = "SELECT top(5) T0.[U_AE_Date], T0.[U_AE_Time], " & _
"T0.[U_AE_Vno] , " & _
"case when [U_AE_Petrol] = '1' then 'Empty' " & _
"when [U_AE_Petrol] = '2' then '1/8' when [U_AE_Petrol] = '3' " & _
"then '1/4' when [U_AE_Petrol] = '4' then '3/8' " & _
"when [U_AE_Petrol] = '5' then '1/2' when [U_AE_Petrol] = '6' " & _
"then '5/8' when [U_AE_Petrol] = '7' then '3/4' " & _
"when [U_AE_Petrol] = '8' then '7/8' when [U_AE_Petrol] = '9' " & _
"then 'Full' end as 'U_AE_Petrol', " & _
"replace(convert(varchar,convert(Money, T0.[U_AE_Mileage]),1),'.00','') " & _
"as 'U_AE_Mileage', T0.[U_AE_Loc] , T0.[U_AE_Remark], T0.[U_AE_RA], " & _
"T0.[U_AE_Name] FROM [dbo].[@AE_VTRACK]  T0 " & _
"where T0.[U_AE_Vno] = '" & vcode & "' order by T0.[U_AE_Date] desc,replace(U_AE_Time,':','') desc "

                                Try
                                    oForm.DataSources.DataTables.Add("@AE_VTRACK")
                                Catch ex As Exception

                                End Try

                                oForm.DataSources.DataTables.Item("@AE_VTRACK").ExecuteQuery(SQL_String)
                                Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("40").Specific
                                oMatrix.Clear()
                                oForm.Items.Item("40").Specific.columns.item("V_0").databind.bind("@AE_VTRACK", "U_AE_Date")
                                oForm.Items.Item("40").Specific.columns.item("V_7").databind.bind("@AE_VTRACK", "U_AE_Time")
                                oForm.Items.Item("40").Specific.columns.item("V_6").databind.bind("@AE_VTRACK", "U_AE_Vno")
                                oForm.Items.Item("40").Specific.columns.item("V_5").databind.bind("@AE_VTRACK", "U_AE_Petrol")
                                oForm.Items.Item("40").Specific.columns.item("V_4").databind.bind("@AE_VTRACK", "U_AE_Mileage")
                                oForm.Items.Item("40").Specific.columns.item("V_3").databind.bind("@AE_VTRACK", "U_AE_Loc")
                                oForm.Items.Item("40").Specific.columns.item("V_2").databind.bind("@AE_VTRACK", "U_AE_Remark")
                                oForm.Items.Item("40").Specific.columns.item("V_1").databind.bind("@AE_VTRACK", "U_AE_Name")
                                oForm.Items.Item("40").Specific.columns.item("V_8").databind.bind("@AE_VTRACK", "U_AE_RA")

                                oForm.Items.Item("40").Specific.LoadFromDataSource()
                                oForm.Items.Item("40").Specific.AutoResizeColumns()

                                Dim opt1 As SAPbouiCOM.OptionBtn = oForm.Items.Item("Item_7").Specific
                                Dim opt2 As SAPbouiCOM.OptionBtn = oForm.Items.Item("Item_8").Specific
                                Dim opt3 As SAPbouiCOM.OptionBtn = oForm.Items.Item("Item_9").Specific
                                Dim opt4 As SAPbouiCOM.OptionBtn = oForm.Items.Item("Item_10").Specific
                                Dim opt5 As SAPbouiCOM.OptionBtn = oForm.Items.Item("Item_11").Specific
                                Dim opt6 As SAPbouiCOM.OptionBtn = oForm.Items.Item("Item_12").Specific
                                Dim opt7 As SAPbouiCOM.OptionBtn = oForm.Items.Item("Item_13").Specific
                                Dim opt8 As SAPbouiCOM.OptionBtn = oForm.Items.Item("Item_14").Specific
                                Dim opt9 As SAPbouiCOM.OptionBtn = oForm.Items.Item("Item_15").Specific

                                opt2.GroupWith("Item_7")
                                opt3.GroupWith("Item_7")
                                opt4.GroupWith("Item_7")
                                opt5.GroupWith("Item_7")
                                opt6.GroupWith("Item_7")
                                opt7.GroupWith("Item_7")
                                opt8.GroupWith("Item_7")
                                opt9.GroupWith("Item_7")
                                'opt2.Selected = True

                                oForm.Items.Item("Item_1").TextStyle = 1
                                oForm.Items.Item("Item_3").TextStyle = 1
                                oForm.Items.Item("39").TextStyle = 1
                                oForm.Items.Item("38").TextStyle = 1
                                oForm.Items.Item("35").TextStyle = 1
                                oForm.Items.Item("Item_18").TextStyle = 1
                                oForm.Items.Item("Item_20").TextStyle = 1
                                oForm.Items.Item("Item_22").TextStyle = 1
                                oForm.Items.Item("Item_24").TextStyle = 1


                            Catch ex As Exception
                                Oapplication_TA.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub
                        End If
                    End If
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub




    Private Sub Oapplication_TA_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles Oapplication_TA.MenuEvent
        Try

            If pVal.MenuUID = "TPO" And pVal.BeforeAction = True Then

                LoadFromXML("TrafficParkingOffense.srf", Oapplication_TA)
                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.Item("TPOD")
                oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                Dim Tmp_val As String = oform.BusinessObject.GetNextSerialNumber(NextSerialNo(Ocompany_TA, Oapplication_TA, "AE_TrafficO"))
                oform.Items.Item("Item_14").Specific.String = Tmp_val

                oform.Items.Item("105").Specific.select("Open")
                oform.Items.Item("Item_20").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                oform.Visible = True
                oform.Items.Item("Item_16").Specific.String = Now.Date 'Format(Now.Date, "dd MMM yyyy")
                oform.Items.Item("1000001").Enabled = False

                oform.Items.Item("108").Specific.String = Ocompany_TA.UserName
                oform.Items.Item("109").Specific.String = Company_Name

                Dim oCFLs As SAPbouiCOM.ChooseFromList
                Dim oCons As SAPbouiCOM.Conditions
                Dim oCon As SAPbouiCOM.Condition
                Dim empty As New SAPbouiCOM.Conditions

                oCFLs = oform.ChooseFromLists.Item("CFL_4")
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
                oCon.CondVal = "C"
                oCFLs.SetConditions(oCons)

                oform.DataBrowser.BrowseBy = "Item_14"
                oform.PaneLevel = 1

            End If

            If pVal.MenuUID = "TPOR" And pVal.BeforeAction = True Then

                LoadFromXML("TrafficParkingReports.srf", Oapplication_TA)
                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.Item("TPODR")
                oform.Visible = True

                oform.Items.Item("10").Specific.String = Format(Now.Date, "dd/MM/yyyy")
                oform.Items.Item("4").Specific.String = Format(Now.Date, "dd/MM/yyyy")
                oform.Items.Item("6").Specific.select("Open")
                Dim ogrid As SAPbouiCOM.Grid = oform.Items.Item("8").Specific
                oform.Items.Item("8").Enabled = False
                Try
                    oform.DataSources.DataTables.Add("Offence")
                Catch ex As Exception

                End Try
                ogrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto
                oform.DataSources.DataTables.Item(0).ExecuteQuery("SELECT T0.[U_AE_Agency] as 'Agency', T0.[U_AE_Offense] as 'Type of Offense', T0.[U_AE_Edate] as 'Expiry Date', T0.[U_AE_nno] as 'Notice Number',T0.[U_AE_fine] as 'Fine Amount',  T0.[U_AE_Submit] as 'Created By', T0.[U_AE_Status] as 'Status'  FROM [dbo].[@AE_TRAFFICO]  T0 WHERE T0.[DocEntry] = ''")
                ogrid.DataTable = oform.DataSources.DataTables.Item("Offence")
                ogrid.AutoResizeColumns()


                oform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            End If


            If pVal.MenuUID = "VAC" And pVal.BeforeAction = True Then
                Try
                    LoadFromXML("Accident_Claim.srf", Oapplication_TA)
                    Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.Item("ACD")
                    oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                    Dim Tmp_val As String = oform.BusinessObject.GetNextSerialNumber(NextSerialNo(Ocompany_TA, Oapplication_TA, "AE_Accident"))
                    oform.Items.Item("Item_14").Specific.String = Tmp_val

                    oform.Items.Item("Item_20").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    oform.Visible = True
                    oform.Items.Item("Item_16").Specific.String = Now.Date '
                    oform.Items.Item("134").Specific.select("Open")
                    oform.Items.Item("1000001").Enabled = False
                    oform.Items.Item("138").Specific.String = Ocompany_TA.UserName
                    oform.Items.Item("139").Specific.String = Company_Name


                    Dim oCFLs As SAPbouiCOM.ChooseFromList
                    Dim oCons As SAPbouiCOM.Conditions
                    Dim oCon As SAPbouiCOM.Condition
                    Dim empty As New SAPbouiCOM.Conditions
                    oCFLs = oform.ChooseFromLists.Item("CFL_4")
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
                    oCon.CondVal = "C"
                    oCFLs.SetConditions(oCons)

                    oform.DataBrowser.BrowseBy = "Item_14"
                    oform.PaneLevel = 1
                Catch ex As Exception
                    Oapplication_TA.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                End Try

            End If



            If pVal.MenuUID = "VT" And pVal.BeforeAction = True Then

                Try
                    LoadFromXML("Vehicle_Tracking.srf", Oapplication_TA)
                    Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.Item("VTF")
                    oform.Visible = True
                    Dim orset As SAPbobsCOM.Recordset = Ocompany_TA.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim oButton As SAPbouiCOM.Button
                    orset.DoQuery("SELECT Top (8)  name FROM [dbo].[@AE_LOCATION]  T0 order by code")

                    For mjs As Integer = 9 To 16
                        oButton = oform.Items.Item("Item_" & mjs).Specific
                        oButton.Caption = orset.Fields.Item(0).Value
                        orset.MoveNext()
                    Next mjs

                    oform.Items.Item("Item_4").Specific.String = Now.Date
                    oform.Items.Item("Item_6").Specific.String = Format(DateTime.Now, "HH:mm")
                    oform.Items.Item("Item_4").TextStyle = 1
                    oform.Items.Item("Item_6").TextStyle = 1
                    oform.Items.Item("Item_1").TextStyle = 1
                    oform.Items.Item("Item_2").TextStyle = 1
                    oform.Items.Item("Item_1").Specific.active = True

                    If p_bSuperUser = "Y" Then
                        oform.Items.Item("21").Visible = True
                    Else
                        oform.Items.Item("21").Visible = False
                    End If


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



                Catch ex As Exception
                    Oapplication_TA.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End Try
            End If

            If pVal.MenuUID = "VTR1" And pVal.BeforeAction = True Then

                LoadFromXML("VehicleTrackingR1.srf", Oapplication_TA)
                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.Item("VTR1")
                oform.Visible = True
                Dim orset As SAPbobsCOM.Recordset = Ocompany_TA.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim ogrid As SAPbouiCOM.Grid = oform.Items.Item("14").Specific
                oform.Items.Item("14").Enabled = False
                Try
                    oform.DataSources.DataTables.Add("VT")
                Catch ex As Exception

                End Try
                ogrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto
                oform.DataSources.DataTables.Item(0).ExecuteQuery("SELECT T0.[U_AE_Date] as 'Date', T0.[U_AE_Time] as 'Time', T0.[U_AE_Vno] as 'Vehicle No', T0.[U_AE_Petrol] as 'Petrol', T0.[U_AE_Mileage] as 'Mileage', T0.[U_AE_Loc] as 'Location', T0.[U_AE_Remark] as 'Remarks', T0.[U_AE_Name] as 'Employee Name' FROM [dbo].[@AE_VTRACK]  T0 where T0.[U_AE_Date] = ''")
                ogrid.DataTable = oform.DataSources.DataTables.Item("VT")
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

            If pVal.MenuUID = "VTR2" And pVal.BeforeAction = True Then

                LoadFromXML("VehicleTrackingR2.srf", Oapplication_TA)
                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.Item("VTR2")
                Dim orset As SAPbobsCOM.Recordset = Ocompany_TA.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                orset.DoQuery("SELECT Top (5)  name FROM [dbo].[@AE_LOCATION]  T0 order by code")
                Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("7").Specific

                orset.MoveLast()
                omatrix.Columns.Item("V_0").TitleObject.Caption = orset.Fields.Item("name").Value

                oform.Items.Item("10").Specific.String = Format(Now.Date, "dd/MM/yyyy")
                oform.Items.Item("4").Specific.String = Format(Now.Date, "dd/MM/yyyy")

                oform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
                oform.Visible = True
                oform.Visible = True

            End If

            If pVal.MenuUID = "VTLR" And pVal.BeforeAction = True Then

                Try

                    LoadFromXML("VehicleTrackingLive.srf", Oapplication_TA)
                    Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.Item("VTLR")
                    oform.Visible = True
                    Dim orset As SAPbobsCOM.Recordset = Ocompany_TA.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    orset.DoQuery("SELECT Top (8)  name FROM [dbo].[@AE_LOCATION]  T0 order by code")
                    Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("7").Specific

                    omatrix.Columns.Item("V_0mjs").TitleObject.Caption = orset.Fields.Item("name").Value
                    orset.MoveNext()
                    omatrix.Columns.Item("V_3mjs").TitleObject.Caption = orset.Fields.Item("name").Value
                    orset.MoveNext()
                    omatrix.Columns.Item("V_2mjs").TitleObject.Caption = orset.Fields.Item("name").Value
                    orset.MoveNext()
                    omatrix.Columns.Item("V_1mjs").TitleObject.Caption = orset.Fields.Item("name").Value
                    orset.MoveNext()
                    omatrix.Columns.Item("V_7mjs").TitleObject.Caption = orset.Fields.Item("name").Value
                    orset.MoveNext()

                    omatrix.Columns.Item("V_6mjs").TitleObject.Caption = orset.Fields.Item("name").Value
                    orset.MoveNext()
                    omatrix.Columns.Item("V_5mjs").TitleObject.Caption = orset.Fields.Item("name").Value
                    orset.MoveNext()
                    omatrix.Columns.Item("V_4mjs").TitleObject.Caption = orset.Fields.Item("name").Value
                    orset.MoveNext()

                    Dim str_query = "VehicleTrackingLiveReport"


                    Try
                        oform.DataSources.DataTables.Add("@AE_VTRACK")
                    Catch ex As Exception

                    End Try

                    oform.DataSources.DataTables.Item("@AE_VTRACK").ExecuteQuery(Str_query)

                    omatrix.Clear()
                    oform.Items.Item("7").Specific.columns.item("V_0mjs").databind.bind("@AE_VTRACK", "U_AE_Loc")
                    oform.Items.Item("7").Specific.columns.item("V_3mjs").databind.bind("@AE_VTRACK", "U_AE_Vno")
                    oform.Items.Item("7").Specific.columns.item("V_2mjs").databind.bind("@AE_VTRACK", "U_AE_Vdesc")
                    oform.Items.Item("7").Specific.columns.item("V_1mjs").databind.bind("@AE_VTRACK", "U_AE_Mileage")
                    oform.Items.Item("7").Specific.columns.item("V_7mjs").databind.bind("@AE_VTRACK", "U_AE_NRIC")
                    oform.Items.Item("7").Specific.columns.item("V_6mjs").databind.bind("@AE_VTRACK", "U_AE_Name")
                    oform.Items.Item("7").Specific.columns.item("V_5mjs").databind.bind("@AE_VTRACK", "U_AE_Remark")
                    oform.Items.Item("7").Specific.columns.item("V_4mjs").databind.bind("@AE_VTRACK", "U_AE_Loc1")

                    oform.Items.Item("7").Specific.LoadFromDataSource()
                    oform.Items.Item("7").Specific.AutoResizeColumns()
                    oform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
                    Class_MainClass.VT_Live_Timer()



                Catch ex As Exception
                    Oapplication_TA.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try

                Exit Sub

            End If


            If pVal.MenuUID = "VLOC" And pVal.BeforeAction = True Then
                Try
                    Oapplication_TA.ActivateMenuItem(UDTform("AE_LOCATION - AE_Location Master", Oapplication_TA))
                Catch ex As Exception
                    Oapplication_TA.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try
            End If


            '---------- New ---------------------------------------

            If pVal.MenuUID = "1281" And pVal.BeforeAction = True Then
                Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm
                If oform.UniqueID = "TPOD" Then
                    TAO_Flag = True
                    oform.Items.Item("Item_14").Enabled = True

                End If

                If oform.UniqueID = "ACD" Then
                    ACD_Flag = True
                    oform.Items.Item("Item_14").Enabled = True

                End If


            End If


            If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.BeforeAction = False Then
                Try
                    Dim oform As SAPbouiCOM.Form = Oapplication_TA.Forms.ActiveForm

                    If oform.UniqueID = "TPOD" Or oform.UniqueID = "ACD" Then
                        Try
                            oform.Freeze(True)
                            If Trim(oform.Items.Item("179").Specific.value) = "Staff (CD or Errand)" Then
                                oform.Items.Item("Item_30").Visible = True
                                oform.Items.Item("Item_31").Visible = True
                                oform.Items.Item("1000011").Visible = False
                            Else
                                oform.Items.Item("Item_30").Visible = False
                                oform.Items.Item("Item_31").Visible = False
                                oform.Items.Item("1000011").Visible = True
                            End If
                            oform.Freeze(False)
                        Catch ex As Exception
                            oform.Freeze(False)
                            Oapplication_TA.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Try
                        End Try
                    End If

                Catch ex As Exception
                    Oapplication_TA.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try
            End If

            '------------------------------------------------------------------------------



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

    Private Function VehicleTrackingDeletion(ByVal sDocEntry As String) As Boolean

        Try

            Dim orset As SAPbobsCOM.Recordset = Ocompany_TA.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset.DoQuery("delete [dbo].[@AE_VTRACK]  WHERE [DocEntry]  = '" & sDocEntry & "'")

            ''Dim oGeneralService As SAPbobsCOM.GeneralService = Nothing
            ''Dim oGeneralData As SAPbobsCOM.GeneralData = Nothing
            ''Dim oGeneralDataParam As SAPbobsCOM.GeneralDataParams = Nothing
            ''Dim oCompanyService As SAPbobsCOM.CompanyService = Ocompany_TA.GetCompanyService

            ''oGeneralService = oCompanyService.GetGeneralService("AE_Vtrack")
            ''oGeneralDataParam = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            ''oGeneralDataParam.SetProperty("DocEntry", sDocEntry)
            ' '' oGeneralData = oGeneralService.GetByParams(oGeneralDataParam)

            ''oGeneralService.Delete(oGeneralData)
            Return True

        Catch ex As Exception
            Return False
            MsgBox(ex.Message)
        End Try

    End Function



    Public Sub ShowFolderBrowser()

        Dim MyProcs() As System.Diagnostics.Process
        FileName = ""
        Dim OpenFile As New OpenFileDialog

        Dim orset As SAPbobsCOM.Recordset = Ocompany_TA.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
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
            OpenFile.InitialDirectory = "C:\" 'orset.Fields.Item("C:\").Value
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
