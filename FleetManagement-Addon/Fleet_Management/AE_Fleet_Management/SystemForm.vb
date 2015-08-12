Public Class SystemForm

    Dim WithEvents Oapplication_SF As SAPbouiCOM.Application
    Dim Ocompany_SF As New SAPbobsCOM.Company
    Dim oItem As SAPbouiCOM.Item
    Dim oFolItem As SAPbouiCOM.Item
    Dim oFol As SAPbouiCOM.Folder
    Dim oStatic As SAPbouiCOM.StaticText
    Dim oEdit As SAPbouiCOM.EditText
    Dim oCombo As SAPbouiCOM.ComboBox



    Public Sub New(ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company)

        Oapplication_SF = oApplication
        Ocompany_SF = oCompany

    End Sub

    Private Sub Oapplication_SF_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles Oapplication_SF.FormDataEvent

        If BusinessObjectInfo.FormTypeEx = "134" Then
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                Dim oform As SAPbouiCOM.Form = Oapplication_SF.Forms.GetFormByTypeAndCount(134, FormType_BP)
                Dim orset As SAPbobsCOM.Recordset = Ocompany_SF.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim PR As String = Trim(oform.DataSources.DBDataSources.Item(0).GetValue("U_AE_Plcode", 0))
                If PR = "" Then
                    orset.DoQuery("update  [dbo].[@AE_PLIST] set U_AE_ass = '' where U_AE_Pcode = '" & BP_PriceList & "'")
                Else
                    orset.DoQuery("update  [dbo].[@AE_PLIST] set U_AE_ass = 'Y' where U_AE_Pcode = '" & PR & "'")
                End If
            End If
        End If

        ''If BusinessObjectInfo.FormTypeEx = "141" Then
        ''    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
        ''        Try

        ''            Dim oform As SAPbouiCOM.Form = Oapplication_SF.Forms.GetFormByTypeAndCount(141, FormType_SM)
        ''            Dim Docentry As String = oform.Items.Item("D0003").Specific.String
        ''            Dim orset As SAPbobsCOM.Recordset = Ocompany_SF.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ''            orset.DoQuery("update  [dbo].[@AE_SM] set  [U_AE_VID] = 'Closed'  WHERE [DocEntry]  = '" & Docentry & "'")

        ''            oform = Oapplication_SF.Forms.Item("SM")
        ''            oform.Close()
        ''        Catch ex As Exception

        ''        End Try
        ''    End If
        ''End If

        If BusinessObjectInfo.FormTypeEx = "65300" Then
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                Dim oform As SAPbouiCOM.Form = Oapplication_SF.Forms.GetFormByTypeAndCount(65300, FormType_Emp)
                Dim orset As SAPbobsCOM.Recordset = Ocompany_SF.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If oform.Items.Item("14").Specific.String <> "" Then
                    orset.DoQuery("update [@AE_SBOOKING] set [U_AE_Deposit] = '" & CDbl(oform.Items.Item("29").Specific.String.ToString.Substring(3, oform.Items.Item("29").Specific.String.ToString.Length - 3)) & "' where [DocEntry] = '" & oform.Items.Item("14").Specific.String & "'")
                End If

            End If
        End If


        If BusinessObjectInfo.FormTypeEx = "133" Then
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                Try
                    Dim oform As SAPbouiCOM.Form = Oapplication_SF.Forms.GetFormByTypeAndCount(133, FormType_Invoice)
                    Dim Docnum As String = Trim(oform.DataSources.DBDataSources.Item(0).GetValue("DocNum", 0))
                    Dim Docentry As String = Trim(oform.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0))
                    Dim orset As SAPbobsCOM.Recordset = Ocompany_SF.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


                    If Invoice_Type = "CD" Then
                        ' MsgBox(Trim(oform.Items.Item("14").Specific.String.ToString.Substring(3, oform.Items.Item("14").Specific.String.ToString.Length - 3)))
                        orset.DoQuery("update [dbo].[@AE_CDRIVER] set  [U_AE_Statu] = 'Closed' WHERE [DocEntry] = '" & Trim(oform.Items.Item("14").Specific.String.ToString.Substring(3, oform.Items.Item("14").Specific.String.ToString.Length - 3)) & "'")
                        orset.DoQuery("update [dbo].[@AE_CDRIVER_R] set [U_AE_Invno] = '" & Docentry & "', [U_AE_Tdriv] = '" & Docnum & "' WHERE [DocEntry] = '" & Trim(oform.Items.Item("14").Specific.String.ToString.Substring(3, oform.Items.Item("14").Specific.String.ToString.Length - 3)) & "'")

                        oform = Oapplication_SF.Forms.Item("CDB")
                        oform.Close()
                    ElseIf Invoice_Type = "SD" Then
                        ' MsgBox(oform.Items.Item("22").Specific.String.ToString.Substring(3, oform.Items.Item("22").Specific.String.ToString.Length - 3))
                        Class_SelfDrive.UDO_Update_SelfDriveST(oform.Items.Item("14").Specific.String, CDbl(oform.Items.Item("22").Specific.String.ToString.Substring(3, oform.Items.Item("22").Specific.String.ToString.Length - 3)), Docentry, Docnum)
                        'orset.DoQuery("update [dbo].[@AE_SBOOKING] set  [U_AE_Status] = 'Closed' WHERE [DocEntry] = '" & oform.Items.Item("14").Specific.String & "'")
                        oform = Oapplication_SF.Forms.Item("SDB")
                        oform.Close()
                    End If
                Catch ex As Exception
                    Oapplication_SF.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End Try

            End If
        End If

        If BusinessObjectInfo.FormTypeEx = "141" Then
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then

                Try
                    Dim oform As SAPbouiCOM.Form = Oapplication_SF.Forms.GetFormByTypeAndCount(141, FormType_SM) 'FormType_Invoice)
                    Dim Docnum As String = Trim(oform.DataSources.DBDataSources.Item(0).GetValue("U_AE_Docnum", 0))
                    Dim DocEntry As String = Trim(oform.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0))
                    Dim CardCode As String = Trim(oform.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0))
                    Dim orset As SAPbobsCOM.Recordset = Ocompany_SF.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("39").Specific
                    Dim oPosition As Integer = 0
                    Dim SM As String
                    For mjs As Integer = 1 To omatrix.RowCount
                        oPosition = InStr(omatrix.Columns.Item("1").Cells.Item(mjs).Specific.String, "-")
                        If oPosition = 0 Then
                            orset.DoQuery("update [dbo].[@AE_SM_R] set [U_AE_Inv] ='" & DocEntry & "' WHERE [DocEntry] = '" & Docnum & "' and  [U_AE_Stype] = '" & Trim(omatrix.Columns.Item("1").Cells.Item(mjs).Specific.String) & "' and [U_AE_Scode] = '" & CardCode & "'")
                        Else
                            SM = Trim(omatrix.Columns.Item("1").Cells.Item(mjs).Specific.String.ToString.Substring(oPosition, omatrix.Columns.Item("1").Cells.Item(mjs).Specific.String.ToString.Length - oPosition))
                            orset.DoQuery("update [dbo].[@AE_SM_R] set [U_AE_Inv] ='" & DocEntry & "' WHERE [DocEntry] = '" & Docnum & "' and  [U_AE_Stype] = '" & SM & "' and [U_AE_Scode] = '" & CardCode & "'")
                        End If
                    Next mjs
                Catch ex As Exception
                    Oapplication_SF.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End Try

            End If
        End If

    End Sub




    Private Sub Oapplication_SF_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles Oapplication_SF.ItemEvent
        Try

            If pVal.FormTypeEx = "65300" Then
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                    If pVal.ItemUID = "1" Then
                        Dim oform As SAPbouiCOM.Form = Oapplication_SF.Forms.GetFormByTypeAndCount(65300, pVal.FormTypeCount)
                        FormType_Emp = pVal.FormTypeCount
                    End If
                End If
            End If


            If pVal.FormTypeEx = "141" Then
                Try
                    If pVal.Before_Action = True Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            Try
                                FormType_SM = pVal.FormTypeCount

                            Catch ex As Exception

                            End Try
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Then
                            Dim OFORM As SAPbouiCOM.Form = Oapplication_SF.Forms.GetFormByTypeAndCount(141, pVal.FormTypeCount)
                            Try

                                FormType_SM = 0
                                FormType_SM = pVal.FormTypeCount
                                Dim oMatrix_in As SAPbouiCOM.Matrix = OFORM.Items.Item("39").Specific
                                Dim oColumn As SAPbouiCOM.Column
                                oColumn = oMatrix_in.Columns.Item("U_AE_OCNO")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_Type")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_TIN")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_TOUT")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_NH")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_SH")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_NP")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_SP")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_REM")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_IND")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_INT")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_OTD")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_OTT")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_SurC")
                                oColumn.Visible = False

                                oItem = OFORM.Items.Add("D0001", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                                oItem.Left = OFORM.Items.Item("70").Left
                                oItem.Width = OFORM.Items.Item("70").Width
                                oItem.Top = OFORM.Items.Item("70").Top + OFORM.Items.Item("70").Height + 2
                                oItem.Height = OFORM.Items.Item("70").Height
                                oItem.FromPane = 0
                                oItem.ToPane = 0
                                oStatic = oItem.Specific
                                oStatic.Caption = "Vehicle No"

                                oItem = OFORM.Items.Add("D0002", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                oItem.Left = OFORM.Items.Item("63").Left
                                oItem.Width = OFORM.Items.Item("14").Width
                                oItem.Top = OFORM.Items.Item("63").Top + OFORM.Items.Item("63").Height + 2
                                oItem.Height = OFORM.Items.Item("63").Height
                                oItem.FromPane = 0
                                oItem.ToPane = 0
                                oEdit = oItem.Specific
                                oEdit.DataBind.SetBound(True, "OPCH", "U_AE_VNO")
                                OFORM.Items.Item("D0001").LinkTo = "D0002"

                                oItem = OFORM.Items.Add("D0003", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                oItem.Left = OFORM.Items.Item("D0002").Left
                                oItem.Width = OFORM.Items.Item("D0002").Width
                                oItem.Top = OFORM.Items.Item("D0002").Top + OFORM.Items.Item("D0002").Height + 2
                                oItem.Height = OFORM.Items.Item("D0002").Height
                                oItem.FromPane = 0
                                oItem.ToPane = 0
                                oEdit = oItem.Specific
                                oEdit.DataBind.SetBound(True, "OPCH", "U_AE_Docnum")
                                oItem.Visible = False

                            Catch ex As Exception

                            End Try


                        End If

                    End If

                Catch ex As Exception

                End Try
            End If

            If pVal.FormTypeEx = "134" Then

                Try
                    If pVal.Before_Action = True Then

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            Try
                                FormType_BP = pVal.FormTypeCount

                            Catch ex As Exception

                            End Try
                        End If


                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_SF.Forms.GetFormByTypeAndCount(134, pVal.FormTypeCount)
                                Dim orset As SAPbobsCOM.Recordset = Ocompany_SF.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                orset.DoQuery("SELECT T0.[U_AE_Pcode], T0.[U_AE_Pname] FROM [dbo].[@AE_PLIST]  T0 where T0.[U_AE_Pcode] <> ''")

                                oItem = oform.Items.Add("B0001", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                                oItem.Left = oform.Items.Item("12").Left
                                oItem.Width = oform.Items.Item("12").Width
                                oItem.Top = oform.Items.Item("12").Top + oform.Items.Item("12").Height + 2
                                oItem.Height = oform.Items.Item("12").Height
                                oItem.FromPane = 0
                                oItem.ToPane = 0
                                oStatic = oItem.Specific
                                oStatic.Caption = "Price List"

                                oItem = oform.Items.Add("B0002", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                                oItem.Left = oform.Items.Item("11").Left
                                oItem.Width = oform.Items.Item("11").Width
                                oItem.Top = oform.Items.Item("11").Top + oform.Items.Item("11").Height + 2
                                oItem.Height = oform.Items.Item("11").Height
                                oItem.FromPane = 0
                                oItem.ToPane = 0
                                oCombo = oItem.Specific
                                oCombo.DataBind.SetBound(True, "OCRD", "U_AE_Plist")
                                oform.Items.Item("B0001").LinkTo = "B0002"
                                For mjs As Integer = 1 To orset.RecordCount
                                    oCombo.ValidValues.Add(orset.Fields.Item("U_AE_Pname").Value, orset.Fields.Item("U_AE_Pcode").Value)
                                    orset.MoveNext()
                                Next mjs
                                oCombo.ValidValues.Add("", "")


                                oItem = oform.Items.Add("B0003", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                oItem.Left = oform.Items.Item("B0002").Left + oform.Items.Item("B0002").Width + 1
                                oItem.Width = oform.Items.Item("B0002").Width
                                oItem.Top = oform.Items.Item("B0002").Top
                                oItem.Height = oform.Items.Item("B0002").Height
                                oItem.FromPane = 0
                                oItem.ToPane = 0
                                oItem.Enabled = False
                                oEdit = oItem.Specific
                                oEdit.DataBind.SetBound(True, "OCRD", "U_AE_Plcode")
                                oform.Items.Item("B0002").LinkTo = "B0003"

                                FormType_BP = pVal.FormTypeCount
                            Catch ex As Exception
                                MsgBox(ex.Message)
                            End Try
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.ItemUID = "B0002" Then
                            Try

                                Dim oform As SAPbouiCOM.Form = Oapplication_SF.Forms.GetFormByTypeAndCount(134, pVal.FormTypeCount)
                                oCombo = oform.Items.Item("B0002").Specific
                                If Trim(oCombo.Selected.Value) <> "" Then
                                    BP_PriceList = ""
                                    BP_PriceList = oform.Items.Item("B0003").Specific.String
                                End If
                               

                            Catch ex As Exception

                            End Try
                        End If

                        ''If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        ''    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.ItemUID = "1" Then
                        ''        Dim oform As SAPbouiCOM.Form = Oapplication_SF.Forms.GetFormByTypeAndCount(134, pVal.FormTypeCount)
                        ''        oform.Items.Item("B0003").Enabled = False
                        ''    End If
                        ''End If


                    ElseIf pVal.Before_Action = False Then

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.ItemUID = "B0002" Then
                            Try

                                Dim oform As SAPbouiCOM.Form = Oapplication_SF.Forms.GetFormByTypeAndCount(134, pVal.FormTypeCount)
                                oCombo = oform.Items.Item("B0002").Specific
                                oform.Items.Item("B0003").Specific.String = oCombo.Selected.Description
                                oform.Items.Item("B0002").Specific.active = True
                                oform.Items.Item("B0003").Enabled = False

                            Catch ex As Exception

                            End Try

                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE And pVal.ItemUID = "1" Then
                                Dim oform As SAPbouiCOM.Form = Oapplication_SF.Forms.GetFormByTypeAndCount(134, pVal.FormTypeCount)
                                oform.Items.Item("B0003").Enabled = False
                            End If
                        End If
                    End If
                Catch ex As Exception

                End Try

            End If

            If pVal.FormTypeEx = "133" Then

                If pVal.Before_Action = True Then

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Then
                        Dim OFORM As SAPbouiCOM.Form = Oapplication_SF.Forms.GetFormByTypeAndCount(133, pVal.FormTypeCount)
                        Try
                            If Invoice_UDF = True Then
                                FormType_Invoice = 0
                                FormType_Invoice = pVal.FormTypeCount
                                Dim oMatrix_in As SAPbouiCOM.Matrix = OFORM.Items.Item("39").Specific
                                Dim oColumn As SAPbouiCOM.Column
                                oColumn = oMatrix_in.Columns.Item("U_AE_OCNO")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_Type")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_TIN")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_TOUT")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_NH")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_SH")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_NP")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_SP")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_REM")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_IND")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_INT")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_OTD")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_OTT")
                                oColumn.Visible = False

                                oColumn = oMatrix_in.Columns.Item("U_AE_SurC")
                                oColumn.Visible = False
                                Invoice_UDF = False
                            End If


                        Catch ex As Exception

                        End Try


                    End If


                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            FormType_Invoice = pVal.FormTypeCount
                        End If
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.InnerEvent = False Then
                        If pVal.ItemUID = "39" Then
                            If pVal.ColUID = "EMH" Or pVal.ColUID = "HR" Or pVal.ColUID = "EMR" Then
                                Try
                                    Dim oform As SAPbouiCOM.Form = Oapplication_SF.Forms.GetFormByTypeAndCount(133, pVal.FormTypeCount)
                                    Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("39").Specific
                                    If oMAtrix.Columns.Item("Type").Cells.Item(pVal.Row).Specific.String <> "One Way" Then
                                        oMAtrix.Columns.Item("12").Cells.Item(pVal.Row).Specific.String = (CDbl(oMAtrix.Columns.Item("THOW").Cells.Item(pVal.Row).Specific.String) * CDbl(oMAtrix.Columns.Item("HR").Cells.Item(pVal.Row).Specific.String)) + _
                                                                                (CDbl(oMAtrix.Columns.Item("EMH").Cells.Item(pVal.Row).Specific.String) * CDbl(oMAtrix.Columns.Item("EMR").Cells.Item(pVal.Row).Specific.String))
                                    End If

                                Catch ex As Exception

                                End Try
                            End If

                            If pVal.ColUID = "THOW" Then
                                Try
                                    Dim oform As SAPbouiCOM.Form = Oapplication_SF.Forms.GetFormByTypeAndCount(133, pVal.FormTypeCount)
                                    Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("39").Specific
                                    oMAtrix.Columns.Item("12").Cells.Item(pVal.Row).Specific.String = (CDbl(oMAtrix.Columns.Item("THOW").Cells.Item(pVal.Row).Specific.String) * CDbl(oMAtrix.Columns.Item("HR").Cells.Item(pVal.Row).Specific.String)) + _
                                                                            (CDbl(oMAtrix.Columns.Item("EMH").Cells.Item(pVal.Row).Specific.String) * CDbl(oMAtrix.Columns.Item("EMR").Cells.Item(pVal.Row).Specific.String))

                                Catch ex As Exception

                                End Try

                            End If

                            If pVal.ColUID = "TimeOUT" Then
                                Try
                                    Dim oform As SAPbouiCOM.Form = Oapplication_SF.Forms.GetFormByTypeAndCount(133, pVal.FormTypeCount)
                                    Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("39").Specific
                                    If oMAtrix.Columns.Item("Type").Cells.Item(pVal.Row).Specific.String <> "One Way" Then
                                        If oMAtrix.Columns.Item("TimeOUT").Cells.Item(pVal.Row).Specific.String <> "" Then
                                            AR_InvoiceTimeCalculation(pVal.Row, oform, Oapplication_SF)
                                        End If
                                    End If

                                Catch ex As Exception

                                End Try

                            End If

                        End If

                    End If

                End If
            End If

            If pVal.FormTypeEx = "150" Then

                If pVal.Before_Action = True Then

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Then

                        Try
                            Dim orset As SAPbobsCOM.Recordset = Ocompany_SF.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            orset.DoQuery("SELECT *  FROM [dbo].[@AE_VTYPE]  T0")
                            Dim oForm As SAPbouiCOM.Form = Oapplication_SF.Forms.GetFormByTypeAndCount(150, pVal.FormTypeCount)

                            oFolItem = oForm.Items.Add("100045V", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
                            oItem = oForm.Items.Item("163")

                            oFolItem.Top = oItem.Top
                            oFolItem.Height = oItem.Height
                            oFolItem.Width = oItem.Width
                            oFolItem.Left = oItem.Left

                            oFol = oFolItem.Specific
                            oFol.Caption = "Vehicle Details"
                            oFol.GroupWith("163")
                            oForm.PaneLevel = 1

                            oItem = oForm.Items.Add("M0001", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("17").Top
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Vehicle Model"
                          
                            oItem = oForm.Items.Add("M0002", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("17").Top
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_MODEL")
                            oForm.Items.Item("M0001").LinkTo = "M0002"

                            oItem = oForm.Items.Add("S0001", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("M0002").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Vehicle Type"

                            oItem = oForm.Items.Add("S0002", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("M0002").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oCombo = oItem.Specific
                            oCombo.DataBind.SetBound(True, "OITM", "U_AE_VTYPE")

                            For mjs As Integer = 1 To orset.RecordCount
                                oCombo.ValidValues.Add(orset.Fields.Item("Code").Value, orset.Fields.Item("Name").Value)
                                orset.MoveNext()
                            Next mjs
                            oForm.Items.Item("S0001").LinkTo = "S0002"

                            oItem = oForm.Items.Add("S0003", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0001").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Transmission Type"

                            oItem = oForm.Items.Add("S0004", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0001").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oCombo = oItem.Specific
                            oCombo.DataBind.SetBound(True, "OITM", "U_AE_TRANS")
                            oForm.Items.Item("S0003").LinkTo = "S0004"

                            oItem = oForm.Items.Add("S0005", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0003").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Year of Make"

                            oItem = oForm.Items.Add("S0006", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0003").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_YEAR_Make")
                            oForm.Items.Item("S0005").LinkTo = "S0006"

                            oItem = oForm.Items.Add("S0007", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0005").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Vehicle Color"

                            oItem = oForm.Items.Add("S0008", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0005").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_COLOR")
                            oForm.Items.Item("S0007").LinkTo = "S0008"

                            oItem = oForm.Items.Add("S0009", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0007").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Engine Capacity"

                            oItem = oForm.Items.Add("S0010", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0007").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_ENG_CAP")
                            oForm.Items.Item("S0009").LinkTo = "S0010"

                            oItem = oForm.Items.Add("S0011", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0009").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Chassis No"

                            oItem = oForm.Items.Add("S0012", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0009").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_CHASSIS_NO")
                            oForm.Items.Item("S0011").LinkTo = "S0012"

                            oItem = oForm.Items.Add("S0013", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0011").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Engine No"

                            oItem = oForm.Items.Add("S0014", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0011").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_ENGINE_NO")
                            oForm.Items.Item("S0013").LinkTo = "S0014"

                            oItem = oForm.Items.Add("S0015", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0013").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Registration Date"

                            oItem = oForm.Items.Add("S0016", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0013").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific

                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_REG_DATE")
                            oForm.Items.Item("S0015").LinkTo = "S0016"

                            oItem = oForm.Items.Add("S0017", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0015").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Ownership Transfer Date"

                            oItem = oForm.Items.Add("S0018", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0015").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_TRSF_DATE")
                            oForm.Items.Item("S0017").LinkTo = "S0018"

                            oItem = oForm.Items.Add("S0019", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0017").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Net Cost of Purchase"

                            oItem = oForm.Items.Add("S0020", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0017").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_COST")
                            oForm.Items.Item("S0019").LinkTo = "S0020"

                            oItem = oForm.Items.Add("S0021", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0019").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Discount"

                            oItem = oForm.Items.Add("S0022", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0019").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_DISCOUNT")
                            oForm.Items.Item("S0021").LinkTo = "S0022"


                            oItem = oForm.Items.Add("S0023", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0021").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Quotation Premium Paid"

                            oItem = oForm.Items.Add("S0024", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0021").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_COE_QP")
                            oForm.Items.Item("S0023").LinkTo = "S0024"

                            oItem = oForm.Items.Add("S0025", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0023").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "OMV Amount"

                            oItem = oForm.Items.Add("S0026", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0023").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_OMV")
                            oForm.Items.Item("S0025").LinkTo = "S0026"

                            oItem = oForm.Items.Add("S0027", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0025").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Parf Value"

                            oItem = oForm.Items.Add("S0028", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0025").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_PARF")
                            oForm.Items.Item("S0027").LinkTo = "S0028"

                            oItem = oForm.Items.Add("S0029", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0027").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Road Tax (12 Months)"

                            oItem = oForm.Items.Add("S0030", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0027").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_ANNL_RD_TAX")
                            oForm.Items.Item("S0029").LinkTo = "S0030"

                            oItem = oForm.Items.Add("S0031", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0029").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "In Vehicle Unit No."

                            oItem = oForm.Items.Add("S0032", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0029").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_IUNO")
                            oForm.Items.Item("S0031").LinkTo = "S0032"

                            oItem = oForm.Items.Add("S0033", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0031").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Battery Capacity"

                            oItem = oForm.Items.Add("S0034", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0031").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_BATTERY")
                            oForm.Items.Item("S0033").LinkTo = "S0034"


                            oItem = oForm.Items.Add("S0035", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 8
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0033").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Tyre Model"

                            oItem = oForm.Items.Add("S0036", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 8 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0033").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_TYRE")
                            oForm.Items.Item("S0035").LinkTo = "S0036"

                            '------------------------------------------------------------------------
                            oItem = oForm.Items.Add("S0037", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 330
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("17").Top
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Warranty Period"

                            oItem = oForm.Items.Add("S0038", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 330 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("17").Top
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_WARRANTY")
                            oForm.Items.Item("S0037").LinkTo = "S0038"


                            oItem = oForm.Items.Add("S0039", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 330
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0037").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Driver Side Wiper Size"

                            oItem = oForm.Items.Add("S0040", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 330 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0037").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_RHS_WIPER")
                            oForm.Items.Item("S0039").LinkTo = "S0040"

                            oItem = oForm.Items.Add("S0041", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 330
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0039").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Passenger Side Wiper Size"

                            oItem = oForm.Items.Add("S0042", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 330 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0039").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_LHS_WIPER")
                            oForm.Items.Item("S0041").LinkTo = "S0042"

                            '-------------------------------------------------------------------------
                            oItem = oForm.Items.Add("V0001", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 330
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("S0041").Top + 16 + 30
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Vehicle Service Mileage"

                            oItem = oForm.Items.Add("V0002", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 330 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("S0041").Top + 16 + 30
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_VSKM")
                            oForm.Items.Item("V0001").LinkTo = "V0002"

                            oItem = oForm.Items.Add("V0003", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 330
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("V0001").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Running Mileage"

                            oItem = oForm.Items.Add("V0004", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 330 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("V0001").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_RKM")
                            oForm.Items.Item("V0003").LinkTo = "V0004"


                            oItem = oForm.Items.Add("V0005", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 330
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("V0003").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Vehicle Service Date"

                            oItem = oForm.Items.Add("V0006", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 330 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("V0003").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_SDate")
                            oForm.Items.Item("V0005").LinkTo = "V0006"

                            oItem = oForm.Items.Add("V0007", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 330
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("V0005").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Battery Service Date"

                            oItem = oForm.Items.Add("V0008", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 330 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("V0005").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_BDate")
                            oForm.Items.Item("V0007").LinkTo = "V0008"

                            oItem = oForm.Items.Add("V0009", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 330
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("V0007").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Tire Service Mileage"

                            oItem = oForm.Items.Add("V0010", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 330 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("V0007").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_TKM")
                            oForm.Items.Item("V0009").LinkTo = "V0010"

                            '--------------------------------------------------------------------
                            oItem = oForm.Items.Add("A0001", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 330
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("V0009").Top + 16 + 30
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Vehicle Make"

                            oItem = oForm.Items.Add("A0002", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 330 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("V0009").Top + 16 + 30
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_MAKE")
                            oForm.Items.Item("A0001").LinkTo = "A0002"


                            oItem = oForm.Items.Add("A0003", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 330
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("A0001").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Registration Preffix"

                            oItem = oForm.Items.Add("A0004", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 330 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("A0001").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_REGPREFIX")
                            oForm.Items.Item("A0003").LinkTo = "A0004"

                            oItem = oForm.Items.Add("A0005", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 330
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("A0003").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "Registration Suffix"

                            oItem = oForm.Items.Add("A0006", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 330 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("A0003").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_REGSUFFIX")
                            oForm.Items.Item("A0005").LinkTo = "A0006"

                            oItem = oForm.Items.Add("A0007", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 330
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("A0005").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "ACT TAG"

                            oItem = oForm.Items.Add("A0008", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                            oItem.Left = 330 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("A0005").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oCombo = oItem.Specific
                            oCombo.DataBind.SetBound(True, "OITM", "U_AE_ACTTAG")
                            oForm.Items.Item("A0007").LinkTo = "A0008"

                            oItem = oForm.Items.Add("A0009", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                            oItem.Left = 330
                            oItem.Width = 120
                            oItem.Top = oForm.Items.Item("A0007").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oStatic = oItem.Specific
                            oStatic.Caption = "C.C"

                            oItem = oForm.Items.Add("A0010", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            oItem.Left = 330 + 124
                            oItem.Width = 100
                            oItem.Top = oForm.Items.Item("A0007").Top + 16
                            oItem.Height = 14
                            oItem.FromPane = 110
                            oItem.ToPane = 110
                            oEdit = oItem.Specific
                            oEdit.DataBind.SetBound(True, "OITM", "U_AE_CC")
                            oForm.Items.Item("A0009").LinkTo = "A0010"



                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                    End If



                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "100045V" Then
                        Dim oform As SAPbouiCOM.Form = Oapplication_SF.Forms.GetFormByTypeAndCount(150, pVal.FormTypeCount)
                        oform.PaneLevel = 110


                    End If

                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Oapplication_SF_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles Oapplication_SF.MenuEvent
        If pVal.MenuUID = "1282" Or pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291" Then
            If pVal.BeforeAction = False Then
                Dim oform As SAPbouiCOM.Form = Oapplication_SF.Forms.ActiveForm
                If oform.TypeEx = "134" Then
                    oform = Oapplication_SF.Forms.GetFormByTypeAndCount("134", FormType_BP)

                    oform.Items.Item("B0003").Enabled = False
                End If
            End If

        End If
    End Sub
End Class
