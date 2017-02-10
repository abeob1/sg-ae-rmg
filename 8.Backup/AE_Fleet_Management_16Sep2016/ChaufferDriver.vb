Imports System.Data
Imports System.Data.OleDb
Imports System.IO

Public Class ChaufferDriver

    Dim WithEvents Oapplication_CD As SAPbouiCOM.Application
    Dim oCompany_TB_CD As New SAPbobsCOM.Company
    Public Docnum As Integer
    Private FileName As String


    Public Sub New(ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany_TB As SAPbobsCOM.Company)

        Oapplication_CD = oApplication
        oCompany_TB_CD = oCompany_TB

    End Sub

    Private Sub Oapplication_CD_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles Oapplication_CD.ItemEvent
        Try

            If pVal.FormUID = "CDB" Then
                If pVal.Before_Action = False Then

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then

                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pVal
                        Dim sCFL_ID As String
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        Dim oForm As SAPbouiCOM.Form = Oapplication_CD.Forms.Item(FormUID)
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                        Try

                            If oCFLEvento.BeforeAction = False Then
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim oedit1 As SAPbouiCOM.EditText
                                Dim oCombo As SAPbouiCOM.ComboBox
                                oDataTable = oCFLEvento.SelectedObjects
                                If pVal.ItemUID = "Item_4" Then 'Billing Code
                                    Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    orset.DoQuery("SELECT T0.[Cellolar] FROM OCPR T0 WHERE T0.[CardCode]  = '" & oDataTable.GetValue("CardCode", 0) & "'")
                                    oCombo = oForm.Items.Item("48").Specific

                                    For mjs As Integer = oCombo.ValidValues.Count - 1 To 0 Step -1
                                        oCombo.ValidValues.Remove(mjs, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next mjs

                                    oForm.Items.Item("Item_5").Specific.string = oDataTable.GetValue("CardName", 0)
                                    oForm.Items.Item("35").Specific.string = oDataTable.GetValue("U_AE_Plcode", 0)
                                    oForm.Items.Item("Item_11").Specific.string = orset.Fields.Item("Cellolar").Value
                                    orset.DoQuery("SELECT isnull(T0.[FirstName] ,'')+ ' ' + isnull(T0.[LastName],'') as 'Name' FROM OCPR T0 WHERE T0.[CardCode]  = '" & oDataTable.GetValue("CardCode", 0) & "'")


                                    Try
                                        For mjs As Integer = 1 To orset.RecordCount
                                            oCombo.ValidValues.Add(orset.Fields.Item("Name").Value, "")
                                            orset.MoveNext()
                                        Next mjs

                                        If oCombo.ValidValues.Count = 0 Then
                                            oCombo.ValidValues.Add("", "")
                                            oCombo.Select("")
                                        End If

                                    Catch ex As Exception
                                    End Try

                                    oCombo.Select(0)
                                    oForm.Items.Item("Item_4").Specific.string = oDataTable.GetValue("CardCode", 0)

                                ElseIf pVal.ItemUID = "Item_7" Then 'Order By
                                    'oForm.Items.Item("30").Specific.string = oDataTable.GetValue("CardName", 0)
                                    oForm.Items.Item("Item_11").Specific.string = oDataTable.GetValue("Cellular", 0)
                                    oForm.Items.Item("Item_7").Specific.string = oDataTable.GetValue("CardCode", 0)
                                ElseIf pVal.ItemUID = "Item_1" Then 'Issued By
                                    oForm.Items.Item("31").Specific.string = oDataTable.GetValue("firstName", 0) & " " & oDataTable.GetValue("lastName", 0)
                                    oForm.Items.Item("Item_1").Specific.string = oDataTable.GetValue("empID", 0)
                                ElseIf pVal.ItemUID = "33" Then 'Tax
                                    oForm.Items.Item("33").Specific.string = oDataTable.GetValue("Code", 0)
                                ElseIf pVal.ItemUID = "Item_15" Then ' Sale Person
                                    oForm.Items.Item("Item_15").Specific.string = oDataTable.GetValue("SlpName", 0)
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then

                        If pVal.ItemUID = "CT" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            Dim oform As SAPbouiCOM.Form
                            Dim oform_In As SAPbouiCOM.Form
                            Try
                                oform = Oapplication_CD.Forms.Item("CDB")
                                Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific
                                Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim customer, Item, docnum, salesemp, contactperson As String
                                Dim amount As Double = 0
                                Dim NH, SH As Double
                                customer = oform.Items.Item("Item_4").Specific.String
                                docnum = oform.Items.Item("Item_13").Specific.String
                                salesemp = oform.Items.Item("Item_15").Specific.String

                                If Trim(oform.Items.Item("48").Specific.value) <> "" Then
                                    orset.DoQuery("SELECT T0.[Name] FROM OCPR T0 WHERE isnull(T0.[FirstName],'') + ' ' +  isnull(T0.[LastName],'')  = '" & Trim(oform.Items.Item("48").Specific.value) & "' and T0.CardCode = '" & oform.Items.Item("Item_4").Specific.String & "'")
                                    contactperson = Trim(orset.Fields.Item("Name").Value)
                                End If

                                If oform.Items.Item("35").Specific.String = "" Then
                                    For mjs As Integer = 1 To oMatrix.RowCount
                                        Dim jj As String = "SELECT T1.[U_AE_Hrate], T1.[U_AE_Surdi], T1.[U_AEowrate] , T1.[U_AE_Surch] FROM [dbo].[@AE_PLIST]  T0 inner join  [dbo].[@AE_PLIST_R]  T1 on T0.docentry = T1.docentry WHERE  T0.[U_AE_Default] = 'Y' and  T1.[U_AE_Vcode] = '" & oMatrix.Columns.Item("V_0").Cells.Item(mjs).Specific.String & "'"
                                        orset.DoQuery("SELECT T1.[U_AE_Hrate], T1.[U_AE_Surdi], T1.[U_AEowrate] , T1.[U_AE_Surch] FROM [dbo].[@AE_PLIST]  T0 inner join  [dbo].[@AE_PLIST_R]  T1 on T0.docentry = T1.docentry WHERE  T0.[U_AE_Default] = 'Y'  and  T1.[U_AE_Vcode] = '" & oMatrix.Columns.Item("V_0").Cells.Item(mjs).Specific.String & "'")
                                        '  MsgBox(oMatrix.Columns.Item("V_1").Cells.Item(mjs).Specific.String & "   " & Math.Truncate(CDbl(oMatrix.Columns.Item("V_1").Cells.Item(mjs).Specific.String)))
                                        ' MsgBox(Right(oMatrix.Columns.Item("V_1").Cells.Item(mjs).Specific.String, 2))
                                        Select Case Right(oMatrix.Columns.Item("V_1").Cells.Item(mjs).Specific.String, 2)
                                            Case 1 To 60
                                                SH = Math.Ceiling(CDbl(oMatrix.Columns.Item("V_1").Cells.Item(mjs).Specific.String))
                                            Case Else
                                                SH = CDbl(oMatrix.Columns.Item("V_1").Cells.Item(mjs).Specific.String)
                                        End Select
                                        ' SH = Math.Ceiling(CDbl(oMatrix.Columns.Item("V_1").Cells.Item(mjs).Specific.String))
                                        ' MsgBox(Right(oMatrix.Columns.Item("V_2").Cells.Item(mjs).Specific.String, 2))
                                        If oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.value = "Disposal" Then
                                            Select Case Right(oMatrix.Columns.Item("V_2").Cells.Item(mjs).Specific.String, 2)
                                                Case 1 To 39
                                                    NH = Math.Truncate(CDbl(oMatrix.Columns.Item("V_2").Cells.Item(mjs).Specific.String)) + 0.5
                                                Case 40 To 60
                                                    NH = Math.Ceiling(CDbl(oMatrix.Columns.Item("V_2").Cells.Item(mjs).Specific.String))
                                                Case Else
                                                    NH = CDbl(oMatrix.Columns.Item("V_2").Cells.Item(mjs).Specific.String)
                                            End Select
                                            oMatrix.Columns.Item("V_6").Cells.Item(mjs).Specific.String = orset.Fields.Item("U_AE_Hrate").Value ' Normal Price
                                            oMatrix.Columns.Item("V_5").Cells.Item(mjs).Specific.String = orset.Fields.Item("U_AE_Surdi").Value ' Special Price
                                            oMatrix.Columns.Item("V_4").Cells.Item(mjs).Specific.String = NH * orset.Fields.Item("U_AE_Hrate").Value ' Normal Hour * Normal Price
                                            oMatrix.Columns.Item("V_3").Cells.Item(mjs).Specific.String = SH * orset.Fields.Item("U_AE_Surdi").Value  ' Special Hour * Special Price
                                            amount += (NH * orset.Fields.Item("U_AE_Hrate").Value) + (SH * orset.Fields.Item("U_AE_Surdi").Value)
                                        ElseIf oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.value = "One Way" Then
                                            NH = 1
                                            oMatrix.Columns.Item("V_6").Cells.Item(mjs).Specific.String = orset.Fields.Item("U_AEowrate").Value ' Normal Price
                                            oMatrix.Columns.Item("V_5").Cells.Item(mjs).Specific.String = orset.Fields.Item("U_AE_Surch").Value ' Special Price
                                            oMatrix.Columns.Item("V_4").Cells.Item(mjs).Specific.String = NH * orset.Fields.Item("U_AEowrate").Value ' Normal Hour * Normal Price
                                            oMatrix.Columns.Item("V_3").Cells.Item(mjs).Specific.String = SH * orset.Fields.Item("U_AE_Surch").Value  ' Special Hour * Special Price
                                            amount += (NH * orset.Fields.Item("U_AEowrate").Value) + (SH * orset.Fields.Item("U_AE_Surch").Value)
                                        End If
                                    Next mjs

                                Else
                                    For mjs As Integer = 1 To oMatrix.RowCount
                                        Dim jj As String = "SELECT T1.[U_AE_Hrate], T1.[U_AE_Surdi], T1.[U_AEowrate] , T1.[U_AE_Surch] FROM [dbo].[@AE_PLIST]  T0 inner join  [dbo].[@AE_PLIST_R]  T1 on T0.docentry = T1.docentry WHERE T0.[U_AE_Pcode] = '" & oform.Items.Item("35").Specific.String & "' and  T1.[U_AE_Vcode] = '" & oMatrix.Columns.Item("V_0").Cells.Item(mjs).Specific.String & "'"
                                        orset.DoQuery("SELECT T1.[U_AE_Hrate], T1.[U_AE_Surdi], T1.[U_AEowrate] , T1.[U_AE_Surch] FROM [dbo].[@AE_PLIST]  T0 inner join  [dbo].[@AE_PLIST_R]  T1 on T0.docentry = T1.docentry WHERE T0.[U_AE_Pcode] = '" & oform.Items.Item("35").Specific.String & "' and  T1.[U_AE_Vcode] = '" & oMatrix.Columns.Item("V_0").Cells.Item(mjs).Specific.String & "'")
                                        ' MsgBox(Right(oMatrix.Columns.Item("V_1").Cells.Item(mjs).Specific.String, 2))
                                        Select Case Right(oMatrix.Columns.Item("V_1").Cells.Item(mjs).Specific.String, 2)
                                            Case 1 To 60
                                                ''    SH = Math.Truncate(CDbl(oMatrix.Columns.Item("V_1").Cells.Item(mjs).Specific.String)) & "." & 5
                                                ''Case 40 To 60
                                                SH = Math.Ceiling(CDbl(oMatrix.Columns.Item("V_1").Cells.Item(mjs).Specific.String))
                                            Case Else
                                                SH = CDbl(oMatrix.Columns.Item("V_1").Cells.Item(mjs).Specific.String)
                                        End Select
                                        ' SH = Math.Ceiling(CDbl(oMatrix.Columns.Item("V_1").Cells.Item(mjs).Specific.String))
                                        ' MsgBox(Right(oMatrix.Columns.Item("V_2").Cells.Item(mjs).Specific.String, 2))
                                        If oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.value = "Disposal" Then
                                            Select Case Right(oMatrix.Columns.Item("V_2").Cells.Item(mjs).Specific.String, 2)
                                                Case 1 To 39
                                                    NH = Math.Truncate(CDbl(oMatrix.Columns.Item("V_2").Cells.Item(mjs).Specific.String)) + 0.5
                                                Case 40 To 60
                                                    NH = Math.Ceiling(CDbl(oMatrix.Columns.Item("V_2").Cells.Item(mjs).Specific.String))
                                                Case Else
                                                    NH = CDbl(oMatrix.Columns.Item("V_2").Cells.Item(mjs).Specific.String)
                                            End Select
                                            oMatrix.Columns.Item("V_6").Cells.Item(mjs).Specific.String = orset.Fields.Item("U_AE_Hrate").Value ' Normal Price
                                            oMatrix.Columns.Item("V_5").Cells.Item(mjs).Specific.String = orset.Fields.Item("U_AE_Surdi").Value ' Special Price
                                            oMatrix.Columns.Item("V_4").Cells.Item(mjs).Specific.String = NH * orset.Fields.Item("U_AE_Hrate").Value ' Normal Hour * Normal Price
                                            oMatrix.Columns.Item("V_3").Cells.Item(mjs).Specific.String = SH * orset.Fields.Item("U_AE_Surdi").Value ' Special Hour * Special Price
                                            ' MsgBox(NH & " * " & orset.Fields.Item("U_AE_Hrate").Value)
                                            amount += (NH * orset.Fields.Item("U_AE_Hrate").Value) + (SH * orset.Fields.Item("U_AE_Surdi").Value)
                                        ElseIf oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.value = "One Way" Then
                                            NH = 1
                                            oMatrix.Columns.Item("V_6").Cells.Item(mjs).Specific.String = orset.Fields.Item("U_AEowrate").Value ' Normal Price
                                            oMatrix.Columns.Item("V_5").Cells.Item(mjs).Specific.String = orset.Fields.Item("U_AE_Surch").Value ' Special Price
                                            oMatrix.Columns.Item("V_4").Cells.Item(mjs).Specific.String = NH * orset.Fields.Item("U_AEowrate").Value ' Normal Hour * Normal Price
                                            oMatrix.Columns.Item("V_3").Cells.Item(mjs).Specific.String = SH * orset.Fields.Item("U_AE_Surch").Value  ' Special Hour * Special Price
                                            amount += (NH * orset.Fields.Item("U_AEowrate").Value) + (SH * orset.Fields.Item("U_AE_Surch").Value)
                                        End If
                                    Next mjs
                                End If


                                'Item = oMatrix.Columns.Item("Col_13").Cells.Item(1).Specific.String
                                'amount = CDbl(oform.Items.Item("Item_21").Specific.String) + CDbl(oform.Items.Item("Item_23").Specific.String)
                                Invoice_Type = "CD"
                                Dim ocombobuttom As SAPbouiCOM.ButtonCombo = oform.Items.Item("CT").Specific
                                If ocombobuttom.Selected.Description = "Copy To A/R Invoice" Then
                                    Oapplication_CD.StatusBar.SetText("Please Wait A/R Invoice is in process ......... !", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    Invoice_UDF = True
                                    Oapplication_CD.ActivateMenuItem("2053")
                                    oform_In = Oapplication_CD.Forms.GetFormByTypeAndCount(133, FormType_Invoice)
                                    Dim oMatrix_in As SAPbouiCOM.Matrix = oform_In.Items.Item("39").Specific
                                    Dim oColumn As SAPbouiCOM.Column
                                    oMatrix_in.Clear()
                                    oform_In.Freeze(True)

                                    ''                                    item2.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable,
                                    ''SAPbouiCOM.BoAutoFormMode.afm_Ok, BoModeVisualBehavior.mvb_False)

                                    oColumn = oMatrix_in.Columns.Add("OCNO", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                    oColumn.TitleObject.Caption = "Order Chit No."
                                    oColumn.DataBind.SetBound(True, "INV1", "U_AE_OCNO")
                                    oColumn.Editable = False

                                    oColumn = oMatrix_in.Columns.Add("TimeIN", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                    oColumn.TitleObject.Caption = "Time IN"
                                    oColumn.DataBind.SetBound(True, "INV1", "U_AE_TIN")
                                    oColumn.Editable = False

                                    oColumn = oMatrix_in.Columns.Add("TimeOUT", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                    oColumn.TitleObject.Caption = "Time OUT"
                                    oColumn.DataBind.SetBound(True, "INV1", "U_AE_TOUT")

                                    oColumn = oMatrix_in.Columns.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                    oColumn.TitleObject.Caption = "Type"
                                    oColumn.DataBind.SetBound(True, "INV1", "U_AE_Type")
                                    oColumn.Editable = False

                                    oColumn = oMatrix_in.Columns.Add("THOW", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                    oColumn.TitleObject.Caption = "Total Hour / One Way"
                                    oColumn.DataBind.SetBound(True, "INV1", "U_AE_NH")

                                    oColumn = oMatrix_in.Columns.Add("EMH", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                    oColumn.TitleObject.Caption = "Early / Midnight Hour"
                                    oColumn.DataBind.SetBound(True, "INV1", "U_AE_SH")

                                    oColumn = oMatrix_in.Columns.Add("HR", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                    oColumn.TitleObject.Caption = "Hourly Rate"
                                    oColumn.DataBind.SetBound(True, "INV1", "U_AE_NP")

                                    oColumn = oMatrix_in.Columns.Add("EMR", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                    oColumn.TitleObject.Caption = "Early / Midnight Rate"
                                    oColumn.DataBind.SetBound(True, "INV1", "U_AE_SP")

                                    oColumn = oMatrix_in.Columns.Add("SCA", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                    oColumn.TitleObject.Caption = "Surcharge Amount"
                                    oColumn.DataBind.SetBound(True, "INV1", "U_AE_SurC")
                                    oColumn.Editable = False

                                    oColumn = oMatrix_in.Columns.Add("REM", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                                    oColumn.TitleObject.Caption = "Remarks To Billing"
                                    oColumn.DataBind.SetBound(True, "INV1", "U_AE_REM")


                                    oform_In.Items.Item("3").Specific.select("S")
                                    oMatrix_in.AddRow()

                                    oform_In.Items.Item("4").Specific.String = customer
                                    oform_In.Items.Item("14").Specific.String = "CD " & docnum


                                    Dim ocombo As SAPbouiCOM.ComboBox
                                    If contactperson <> "" Then
                                        ocombo = oform_In.Items.Item("85").Specific
                                        ocombo.Select(contactperson)
                                    End If

                                    For mjs As Integer = 1 To oMatrix.RowCount
                                        oMatrix_in.Columns.Item("1").Cells.Item(mjs).Specific.String = "Chauffer Drive Billing for this Vehicle Type : " & oMatrix.Columns.Item("Col_0").Cells.Item(mjs).Specific.value
                                        oMatrix_in.Columns.Item("12").Cells.Item(mjs).Specific.String = CDbl(oMatrix.Columns.Item("V_4").Cells.Item(mjs).Specific.String) + CDbl(oMatrix.Columns.Item("V_3").Cells.Item(mjs).Specific.String)
                                        oMatrix_in.Columns.Item("2").Cells.Item(mjs).Specific.String = CD_GLAcc
                                        oMatrix_in.Columns.Item("OCNO").Cells.Item(mjs).Specific.String = oMatrix.Columns.Item("Col_10").Cells.Item(mjs).Specific.String
                                        oMatrix_in.Columns.Item("TimeIN").Cells.Item(mjs).Specific.String = oMatrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String
                                        oMatrix_in.Columns.Item("TimeOUT").Cells.Item(mjs).Specific.String = oMatrix.Columns.Item("Col_9").Cells.Item(mjs).Specific.String
                                        oMatrix_in.Columns.Item("Type").Cells.Item(mjs).Specific.String = oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.value

                                        If oMatrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.value = "One Way" Then
                                            oMatrix_in.Columns.Item("THOW").Cells.Item(mjs).Specific.String = "1"
                                        Else
                                            Select Case Right(oMatrix.Columns.Item("V_2").Cells.Item(mjs).Specific.String, 2)
                                                Case 1 To 39
                                                    oMatrix_in.Columns.Item("THOW").Cells.Item(mjs).Specific.String = Math.Truncate(CDbl(oMatrix.Columns.Item("V_2").Cells.Item(mjs).Specific.String)) + 0.5
                                                Case 40 To 60
                                                    oMatrix_in.Columns.Item("THOW").Cells.Item(mjs).Specific.String = Math.Ceiling(CDbl(oMatrix.Columns.Item("V_2").Cells.Item(mjs).Specific.String))
                                                Case Else
                                                    oMatrix_in.Columns.Item("THOW").Cells.Item(mjs).Specific.String = CDbl(oMatrix.Columns.Item("V_2").Cells.Item(mjs).Specific.String)
                                            End Select
                                        End If

                                        Select Case Right(oMatrix.Columns.Item("V_1").Cells.Item(mjs).Specific.String, 2)
                                            Case 1 To 60
                                                '    oMatrix_in.Columns.Item("EMH").Cells.Item(mjs).Specific.String = Math.Truncate(CDbl(oMatrix.Columns.Item("V_1").Cells.Item(mjs).Specific.String)) + 0.5
                                                'Case 40 To 60
                                                oMatrix_in.Columns.Item("EMH").Cells.Item(mjs).Specific.String = Math.Ceiling(CDbl(oMatrix.Columns.Item("V_1").Cells.Item(mjs).Specific.String))
                                            Case Else
                                                oMatrix_in.Columns.Item("EMH").Cells.Item(mjs).Specific.String = CDbl(oMatrix.Columns.Item("V_1").Cells.Item(mjs).Specific.String)
                                        End Select
                                        oMatrix_in.Columns.Item("HR").Cells.Item(mjs).Specific.String = oMatrix.Columns.Item("V_6").Cells.Item(mjs).Specific.String
                                        oMatrix_in.Columns.Item("EMR").Cells.Item(mjs).Specific.String = oMatrix.Columns.Item("V_5").Cells.Item(mjs).Specific.String
                                        oMatrix_in.Columns.Item("SCA").Cells.Item(mjs).Specific.String = CDbl(oMatrix_in.Columns.Item("EMR").Cells.Item(mjs).Specific.String) * CDbl(oMatrix_in.Columns.Item("EMH").Cells.Item(mjs).Specific.String)
                                        oMatrix_in.Columns.Item("REM").Cells.Item(mjs).Specific.String = oMatrix.Columns.Item("Col_12").Cells.Item(mjs).Specific.String
                                    Next mjs

                                    oMatrix_in.AutoResizeColumns()

                                    ocombo = oform_In.Items.Item("20").Specific
                                    If salesemp <> "" Then
                                        ocombo.Select(salesemp)
                                    End If

                                    oform_In.Items.Item("16").Specific.String = "Chauffer Drive Invoice Based on Booking No : " & docnum
                                    oform_In.Visible = True
                                    oform_In.Freeze(False)
                                End If

                            Catch ex As Exception
                                oform_In.Freeze(False)
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub
                        End If

                        If pVal.ItemUID = "Item_16" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then

                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.Item("CDB")
                                Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific
                                Dim Tmp_date, Tmp_date1, Tmp_date2, Tmp_date3 As Date
                                Dim Tmp_sub As TimeSpan
                                ' Dim Hour, Min, Hour1, Min1 As Integer
                                ' Dim F_Normal, F_Special As Date

                                ''Special_Hours = 0
                                ''Special_Mins = 0
                                ''Normal_Hours = 0
                                ''Normal_Mins = 0

                                If Trim(oform.Items.Item("Item_16").Specific.selected.value) = "Billing" Then

                                    For mjs As Integer = 1 To oMAtrix.RowCount - 1
                                        If oMAtrix.Columns.Item("Col_13").Cells.Item(mjs).Specific.String = "" Then

                                            oform.Items.Item("Item_16").Specific.select("Open")
                                            Oapplication_CD.StatusBar.SetText("Kindly assign the vehicle and driver in the booking line " & mjs, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Try
                                        End If
                                    Next mjs

                                    If oMAtrix.RowCount > 1 Then
                                        If oMAtrix.Columns.Item("Date").Cells.Item(oMAtrix.RowCount).Specific.String = "" Then
                                            oMAtrix.DeleteRow(oMAtrix.RowCount)
                                        End If
                                    End If

                                    For mjs As Integer = 1 To oMAtrix.RowCount
                                        Special_Hours = 0
                                        Special_Mins = 0
                                        Normal_Hours = 0
                                        Normal_Mins = 0
                                        If oMAtrix.Columns.Item("Col_9").Cells.Item(mjs).Specific.String = "" Then
                                            'oMAtrix.Columns.Item("Col_9").Cells.Item(mjs).Specific.active = True
                                            Oapplication_CD.StatusBar.SetText("Drop Time Should not be empty ........ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Try
                                        Else
                                            oMAtrix.CommonSetting.SetCellEditable(mjs, 11, False)
                                            Tmp_date = oMAtrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String & ":00"
                                            Tmp_date1 = oMAtrix.Columns.Item("Col_9").Cells.Item(mjs).Specific.String & ":00"

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
                                        End If
                                        Tmp_sub = New TimeSpan(0, (Normal_Hours * 60) + Normal_Mins, 0)
                                        If Tmp_sub.Hours < 3 Then
                                            oMAtrix.Columns.Item("V_2").Cells.Item(mjs).Specific.String = "03.00"
                                        Else
                                            oMAtrix.Columns.Item("V_2").Cells.Item(mjs).Specific.String = Format(Tmp_sub.Hours, "00") & "." & Format(Tmp_sub.Minutes, "00")
                                        End If

                                        Tmp_sub = New TimeSpan(0, (Special_Hours * 60) + Special_Mins, 0)
                                        oMAtrix.Columns.Item("V_1").Cells.Item(mjs).Specific.String = Format(Tmp_sub.Hours, "00") & "." & Format(Tmp_sub.Minutes, "00")

                                        'F_Normal = Normal_Hours & ":" & Normal_Mins & ":00"
                                        'F_Special = Special_Hours & ":" & Special_Mins & ":00"
                                        'MsgBox("Normal Hour : " & Normal_Hours & " Normal Min : " & Normal_Mins & " Special Hour : " & Special_Hours & " Special Min " & Special_Mins)
                                        'MsgBox("Normal " & F_Normal & " Special " & F_Special)
                                    Next mjs
                                    oMAtrix.Columns.Item("Col_9").Editable = False
                                    ' oMAtrix.AutoResizeColumns()
                                    oform.Items.Item("1").Enabled = True

                                    If oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    ' oform.Items.Item("CT").Enabled = True
                                End If



                            Catch ex As Exception
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub
                        End If



                        If pVal.ItemUID = "Item_17" And pVal.ColUID = "Col_1" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            Dim oform As SAPbouiCOM.Form
                            Try
                                oform = Oapplication_CD.Forms.Item("CDB")
                                Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific
                                oform.Freeze(True)
                                If oMAtrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific.Selected.value = "One Way" Then
                                    oMAtrix.CommonSetting.SetCellEditable(pVal.Row, 11, False)
                                Else
                                    oMAtrix.CommonSetting.SetCellEditable(pVal.Row, 11, True)
                                End If
                                If oMAtrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific.Selected.value <> "-" Then
                                    If oMAtrix.RowCount = pVal.Row Then
                                        oMAtrix.AddRow()
                                        oMAtrix.Columns.Item("#").Cells.Item(oMAtrix.RowCount).Specific.String = oMAtrix.RowCount
                                        oMAtrix.Columns.Item("Col_3").Cells.Item(oMAtrix.RowCount).Specific.String = "."
                                        oMAtrix.Columns.Item("Col_7").Cells.Item(oMAtrix.RowCount).Specific.String = "."
                                        oMAtrix.Columns.Item("Col_8").Cells.Item(oMAtrix.RowCount).Specific.String = "."
                                        oMAtrix.Columns.Item("Col_11").Cells.Item(oMAtrix.RowCount).Specific.String = "."
                                        oMAtrix.Columns.Item("Col_12").Cells.Item(oMAtrix.RowCount).Specific.String = "."
                                    End If
                                End If

                                'oMAtrix.AutoResizeColumns()
                                oform.Freeze(False)
                            Catch ex As Exception
                                oform.Freeze(False)
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                        End If

                        If pVal.ItemUID = "Item_17" And pVal.ColUID = "Col_0" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                            Try

                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.Item("CDB")
                                Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific
                                Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oMAtrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.String = oMAtrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.selected.description

                                If oMAtrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.value <> "-" Then
                                    orset.DoQuery("SELECT T1.[U_AE_Hrate], T1.[U_AE_Surch], T1.[U_AE_Surdi] FROM [dbo].[@AE_PLIST]  T0 inner join  [dbo].[@AE_PLIST_R]  T1 on T0.docentry = T1.docentry WHERE T0.[U_AE_Pcode] = '" & oform.Items.Item("35").Specific.String & "' and   T1.[U_AE_Vcode] = '" & oMAtrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.selected.description & "'")
                                    oform.Items.Item("36").Specific.String = orset.Fields.Item("U_AE_Hrate").Value
                                    oform.Items.Item("37").Specific.String = orset.Fields.Item("U_AE_Surch").Value
                                    oform.Items.Item("38").Specific.String = orset.Fields.Item("U_AE_Surdi").Value
                                    If oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                End If

                            Catch ex As Exception
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub
                        End If

                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                        If pVal.ItemUID = "1" And pVal.Action_Success = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                'oform.Close()
                                'LoadFromXML("ChaufferDriver_Booking.srf", Oapplication_CD)
                                'oform = Oapplication_CD.Forms.Item("CDB")
                                If Chauffer_Driver_binding_AddMode(oform) = False Then
                                    BubbleEvent = False
                                    Exit Try
                                End If
                                oform.Visible = True
                            Catch ex As Exception
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub
                        End If

                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            Dim oform As SAPbouiCOM.Form
                            Try
                                oform = Oapplication_CD.Forms.Item("CDB")
                                'oform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                ' oform.Items.Item("Item_16").Enabled = True
                                oform.Items.Item("1000003").Enabled = False
                                If Event_CD = False Then
                                    oform.Freeze(True)
                                    ' oform.Items.Item("Item_19").Specific.active = True
                                    oform.Items.Item("Item_13").Enabled = False
                                    CB_Navigation(oform)
                                    oform.Freeze(False)
                                Else
                                    Event_CD = False
                                End If

                            Catch ex As Exception
                                oform.Freeze(False)
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                        End If


                        If pVal.ItemUID = "1000003" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.Item("CDB")
                                Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific
                                showOpenFileDialog()
                                If FileName <> "" Then
                                    Dim file = New FileInfo(stFilePathAndName)
                                    ' file.CopyTo(Path.Combine(orset.Fields.Item("attachpath").Value, file.Name), True)
                                    If CDBooking_ExcelUpload(stFilePathAndName, oform) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                           
                                End If

                                If oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If
                                oMatrix.AutoResizeColumns()
                            Catch ex As Exception
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                            oform = Oapplication_CD.Forms.ActiveForm
                            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific
                            oform.Freeze(True)
                            If oMatrix.RowCount > 1 Then
                                If oMatrix.Columns.Item("Date").Cells.Item(oMatrix.RowCount).Specific.String = "" Then
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

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED Then
                        If pVal.ItemUID = "Item_17" And (pVal.ColUID = "Col_3" Or pVal.ColUID = "Col_7" Or pVal.ColUID = "Col_8" Or pVal.ColUID = "Col_11" Or pVal.ColUID = "Col_12") Then
                            Try

                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm

                                If Trim(oform.Items.Item("Item_16").Specific.value) <> "Closed" Then
                                    Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific
                                    If omatrix.Columns.Item("Date").Cells.Item(pVal.Row).Specific.String <> "" Then
                                        Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        Dim Doc As String = oform.Items.Item("Item_13").Specific.string
                                        Dim Line As String = omatrix.Columns.Item("#").Cells.Item(pVal.Row).Specific.String
                                        Dim ColID As String = pVal.ColUID
                                        Dim Title As String = ""

                                        Select Case pVal.ColUID

                                            Case "Col_3"
                                                Title = "Guest Name"
                                            Case "Col_7"
                                                Title = "Pickup Location"
                                            Case "Col_8"
                                                Title = "Drop Location"
                                            Case "Col_11"
                                                Title = "Remark to Driver"
                                            Case "Col_12"
                                                Title = "Remark for Billing"
                                        End Select

                                        If Extended_Text(oform, Oapplication_CD, oCompany_TB_CD, oform.Items.Item("Item_13").Specific.string, omatrix.Columns.Item("#").Cells.Item(pVal.Row).Specific.String _
                                                         , pVal.ColUID, "CD", Title) = False Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If

                                End If

                            Catch ex As Exception
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub
                        End If
                    End If


                    ''If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                    ''    If pVal.ItemUID = "Item_17" And pVal.ColUID <> "Col_9" Then
                    ''        Try

                    ''            Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                    ''            Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific

                    ''            If oMAtrix.Columns.Item("Col_13").Cells.Item(pVal.Row).Specific.String <> "" Then
                    ''                Oapplication_CD.StatusBar.SetText("Modification not allowed after vehicle assign ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    ''                BubbleEvent = False
                    ''                Exit Try
                    ''            End If

                    ''        Catch ex As Exception

                    ''        End Try
                    ''        Exit Sub


                    ''    End If


                    ''End If


                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.InnerEvent = False Then

                        ''If pVal.ItemUID = "30" Then
                        ''    Oapplication_CD.SendKeys("+{F2}")
                        ''End If

                        ''If pVal.ItemUID = "Item_17" And pVal.ColUID = "Col_8" Then
                        ''    Try
                        ''        Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                        ''        Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific
                        ''        oMAtrix.Columns.Item("Col_11").Cells.Item(pVal.Row).Specific.active = True
                        ''    Catch ex As Exception

                        ''    End Try
                        ''End If
                        If pVal.ItemUID = "Item_17" And pVal.ColUID = "Col_2" Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                Dim oDate As Date
                                Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific
                                If oMAtrix.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.String <> "" Then
                                    If oMAtrix.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.String.ToString.Length = 4 Then
                                        oMAtrix.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.String = oMAtrix.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.String.ToString.Substring(0, 2) & ":" & oMAtrix.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.String.ToString.Substring(2, 2)
                                    ElseIf oMAtrix.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.String.ToString.Length = 5 Then
                                        If IsDate(oMAtrix.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.String & ":00") = False Then
                                            oMAtrix.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.active = True
                                            Oapplication_CD.StatusBar.SetText("Its not a valid time format ............ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Try
                                        End If
                                    Else
                                        oMAtrix.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.active = True
                                        Oapplication_CD.StatusBar.SetText("Its not a valid time format ............ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Try
                                    End If
                                    If oMAtrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific.value = "-" Then
                                        oMAtrix.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.String = ""
                                        oMAtrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific.active = True
                                        Oapplication_CD.StatusBar.SetText("Service type should not be empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        Exit Try
                                    ElseIf oMAtrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific.value = "One Way" Then
                                        oDate = oMAtrix.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific.String
                                        oMAtrix.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific.String = Format(oDate.AddHours(1.5), "HH:mm")
                                    End If

                                End If



                            Catch ex As Exception
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub
                        End If

                        If pVal.ItemUID = "Item_17" And pVal.ColUID = "Col_9" Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific
                                Dim Hour_Rate, Minutes_Rate As Double
                                If oMAtrix.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific.String <> "" Then
                                    If oMAtrix.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific.String.ToString.Length = 4 Then
                                        oMAtrix.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific.String = oMAtrix.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific.String.ToString.Substring(0, 2) & ":" & oMAtrix.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific.String.ToString.Substring(2, 2)
                                    ElseIf oMAtrix.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific.String.ToString.Length = 5 Then
                                        If IsDate(oMAtrix.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific.String & ":00") = False Then
                                            oMAtrix.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific.active = True
                                            Oapplication_CD.StatusBar.SetText("Its not a valid time format ............ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Try
                                        End If
                                    Else
                                        oMAtrix.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific.active = True
                                        Oapplication_CD.StatusBar.SetText("Its not a valid time format ............ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Try
                                    End If

                                    Dim Date1 As Date = oMAtrix.Columns.Item("Col_2").Cells.Item(1).Specific.String & ":00"
                                    Dim Date2 As Date = oMAtrix.Columns.Item("Col_9").Cells.Item(1).Specific.String & ":00"
                                    If Date1 > Date2 Then

                                    Else

                                        Dim TS As TimeSpan = Date2.Subtract(Date1)
                                        Hour_Rate = TS.Hours * CDbl(oform.Items.Item("36").Specific.String)
                                        Minutes_Rate = TS.Minutes * (CDbl(oform.Items.Item("36").Specific.String) / 60)
                                        oform.Items.Item("Item_21").Specific.String = Hour_Rate + Minutes_Rate

                                    End If
                                End If

                            Catch ex As Exception
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub
                        End If

                        If pVal.ItemUID = "Item_17" And pVal.ColUID = "Col_6" Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific
                                If oMAtrix.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific.String <> "" Then
                                    If oMAtrix.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific.String.ToString.Length = 4 Then
                                        oMAtrix.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific.String = oMAtrix.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific.String.ToString.Substring(0, 2) & ":" & oMAtrix.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific.String.ToString.Substring(2, 2)
                                    ElseIf oMAtrix.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific.String.ToString.Length = 5 Then
                                        If IsDate(oMAtrix.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific.String & ":00") = False Then
                                            oMAtrix.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific.active = True
                                            Oapplication_CD.StatusBar.SetText("Its not a valid time format ............ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Try
                                        End If
                                    Else
                                        oMAtrix.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific.active = True
                                        Oapplication_CD.StatusBar.SetText("Its not a valid time format ............ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Try
                                    End If
                                End If

                            Catch ex As Exception
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                        End If

                        ' ''If pVal.ItemUID = "Item_17" And pVal.ColUID = "Date" Then
                        ' ''    Try

                        ' ''        Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                        ' ''        Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific
                        ' ''        If oMAtrix.Columns.Item("Date").Cells.Item(pVal.Row).Specific.String <> "" Then
                        ' ''            Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' ''            orset.DoQuery("SELECT autokey from onnm where ObjectCode = 'AE_Cdriver'")
                        ' ''            oMAtrix.Columns.Item("Col_10").Cells.Item(pVal.Row).Specific.String = orset.Fields.Item("autokey").Value.ToString.PadLeft(7, "0"c) & oMAtrix.Columns.Item("#").Cells.Item(pVal.Row).Specific.String

                        ' ''            If pVal.Row = oMAtrix.RowCount Then
                        ' ''                oMAtrix.AddRow()
                        ' ''                oMAtrix.Columns.Item("#").Cells.Item(oMAtrix.RowCount).Specific.String = oMAtrix.RowCount
                        ' ''            End If
                        ' ''        End If

                        ' ''    Catch ex As Exception
                        ' ''        Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        ' ''        BubbleEvent = False
                        ' ''        Exit Try
                        ' ''    End Try
                        ' ''End If
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        Try
                            If pVal.ItemUID = "1" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific

                                If Chauffer_Driver_Validation(oform) = False Then
                                    BubbleEvent = False
                                    Exit Try
                                End If
                                Event_CD = True

                                If oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                    'oform.Items.Item("CT").Enabled = True
                                    oform.Items.Item("Item_23").Enabled = True
                                    oform.Items.Item("33").Enabled = True
                                End If

                            End If



                            If pVal.ItemUID = "2" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                                Try
                                    Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm

                                    Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    orset.DoQuery("delete  [dbo].[@AE_EXTENDED]  where U_AE_Dno = '" & oform.Items.Item("Item_13").Specific.String & "'")

                                Catch ex As Exception

                                End Try
                            End If

                            If pVal.ItemUID = "46" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                Dim DA1_Thread As System.Threading.Thread
                                Try

                                    Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                    Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific
                                    Dim OCNO As String
                                    Dim Guestname As String
                                    For mjs As Integer = 1 To oMatrix.RowCount
                                        If oMatrix.IsRowSelected(mjs) = True Then
                                            Guestname = oMatrix.Columns.Item("Col_3").Cells.Item(mjs).Specific.String
                                        End If
                                    Next mjs

                                    DA1_Thread = New System.Threading.Thread(AddressOf Class_Report.OpenPagingSign)
                                    Class_Report.oApplication = Oapplication_CD
                                    Class_Report.oCompany = oCompany_TB_CD
                                    'Class_Report.Report_Name = "AE_RP005_PagingSign.rpt"
                                    'Class_Report.Report_Parameter = "@OrderChitNo"
                                    Class_Report.sGuestName = Guestname
                                    If DA1_Thread.IsAlive Then
                                        Oapplication_CD.MessageBox("Report is already open....")
                                    Else
                                        DA1_Thread.TrySetApartmentState(Threading.ApartmentState.STA)
                                        Oapplication_CD.StatusBar.SetText("Paging Sign Report Opening in process ......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        DA1_Thread.Start()
                                    End If

                                Catch ex As Exception
                                    'DA1_Thread.Abort()
                                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                End Try
                                Exit Sub
                            End If

                            If pVal.ItemUID = "45" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                Dim RMG_Thread As System.Threading.Thread
                                Try
                                    Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                    Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific
                                    Docnum = 0
                                    For mjs As Integer = 1 To oMatrix.RowCount
                                        If oMatrix.IsRowSelected(mjs) = True Then
                                            If oMatrix.Columns.Item("Col_10").Cells.Item(mjs).Specific.String = "" Then
                                                Oapplication_CD.StatusBar.SetText("Order Chit Number should not be empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                                BubbleEvent = False
                                                Exit Sub
                                            Else
                                                Docnum = CInt(oMatrix.Columns.Item("Col_10").Cells.Item(mjs).Specific.String)
                                            End If

                                        End If
                                    Next mjs

                                    If Docnum = 0 Then
                                        Oapplication_CD.StatusBar.SetText("Order Chit Number should not be empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If

                                    RMG_Thread = New System.Threading.Thread(AddressOf Class_Report.Report_ExportToPDF)
                                    Class_Report.oApplication = Oapplication_CD
                                    Class_Report.oCompany = oCompany_TB_CD
                                    Class_Report.Report_Name = "AE_RP001_OrderChit.rpt"
                                    Class_Report.Report_Parameter = "@OrderChitNo"
                                    Class_Report.Docnum = Docnum
                                    Class_Report.FileName = "OCNO"
                                    Class_Report.Report_Title = "RMG Report"
                                    If RMG_Thread.IsAlive Then
                                        Oapplication_CD.MessageBox("Report is already open....")
                                    Else
                                        RMG_Thread.TrySetApartmentState(Threading.ApartmentState.STA)
                                        Oapplication_CD.StatusBar.SetText("RMG Report Opening in process ......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        RMG_Thread.Start()
                                    End If



                                Catch ex As Exception
                                    RMG_Thread.Abort()
                                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                End Try

                            End If

                            ''If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            ''    Try
                            ''        Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                            ''        Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific

                            ''        If Trim(oform.Items.Item("Item_16").Specific.value) = "Dropped" Then
                            ''            If oMAtrix.Columns.Item("Col_9").Cells.Item(1).Specific.String = "" Then
                            ''                Oapplication_CD.StatusBar.SetText("Drop Time should not be Empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            ''                BubbleEvent = False
                            ''                Exit Try
                            ''            End If

                            ''            If oMAtrix.Columns.Item("Col_13").Cells.Item(1).Specific.String = "" Then
                            ''                Oapplication_CD.StatusBar.SetText("Vehicle No should not be empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            ''                BubbleEvent = False
                            ''                Exit Try
                            ''            End If

                            ''            If oMAtrix.Columns.Item("Col_14").Cells.Item(1).Specific.String = "" Then
                            ''                Oapplication_CD.StatusBar.SetText("Driver Name should not be Empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            ''                BubbleEvent = False
                            ''                Exit Try
                            ''            End If

                            ''            oform.Items.Item("CT").Enabled = True
                            ''            oform.Items.Item("Item_23").Enabled = True
                            ''            oform.Items.Item("33").Enabled = True

                            ''        End If
                            ''    Catch ex As Exception
                            ''    End Try
                            ''End If
                        Catch ex As Exception
                            Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            BubbleEvent = False
                            Exit Try
                        End Try
                        Exit Sub
                    End If
                End If
            End If

            '----------------------- New --------------------------------------
            If pVal.FormUID = "CDBR" Then
                If pVal.Before_Action = False Then

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pVal
                        Dim sCFL_ID As String
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        Dim oForm As SAPbouiCOM.Form = Oapplication_CD.Forms.Item(FormUID)
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)

                        If oCFLEvento.BeforeAction = False Then
                            Dim oDataTable As SAPbouiCOM.DataTable
                            Dim oedit1 As SAPbouiCOM.EditText
                            oDataTable = oCFLEvento.SelectedObjects

                            If pVal.ItemUID = "11" Then
                                oForm.Items.Item("3").Specific.string = oDataTable.GetValue("CardName", 0)
                                oForm.Items.Item("11").Specific.string = oDataTable.GetValue("CardCode", 0)
                            End If
                        End If
                    End If

                End If
                If pVal.Before_Action = True Then

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED Then
                        If pVal.ItemUID = "8" And pVal.ColUID = "Booking No" Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                Dim ogrid As SAPbouiCOM.Grid = oform.Items.Item("8").Specific
                                Dim str As String = ogrid.DataTable.GetValue("Booking No", pVal.Row)
                                LoadFromXML("ChaufferDriver_Booking.srf", Oapplication_CD)
                                oform = Oapplication_CD.Forms.Item("CDB")
                                Dim ocombobutton As SAPbouiCOM.ButtonCombo = oform.Items.Item("CT").Specific
                                ocombobutton.ValidValues.Add("Copy To", "Copy To A/R Invoice")
                                oform.Visible = True
                                oform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                oform.Items.Item("Item_13").Enabled = True
                                oform.Items.Item("Item_13").Specific.String = str
                                oform.Items.Item("Item_4").Specific.active = True
                                oform.Items.Item("Item_13").Enabled = False

                                oform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                            Catch ex As Exception
                                Oapplication_CD.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                                BubbleEvent = False
                                Exit Sub
                            End Try
                        End If
                    End If


                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        If pVal.ItemUID = "9" Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                Dim customer, event_, status As String

                                If oform.Items.Item("11").Specific.String <> "" Then
                                    customer = oform.Items.Item("11").Specific.String
                                Else
                                    customer = "%"
                                    oform.Items.Item("3").Specific.String = ""
                                End If

                                If oform.Items.Item("5").Specific.String <> "" Then
                                    event_ = "%" & oform.Items.Item("5").Specific.String & "%"
                                Else
                                    event_ = "%"
                                End If

                                If Trim(oform.Items.Item("7").Specific.value) <> "All" Then
                                    status = Trim(oform.Items.Item("7").Specific.value)
                                Else
                                    status = "%"
                                End If


                                Try
                                    oform.DataSources.DataTables.Add("CDBR")
                                Catch ex As Exception

                                End Try
                                Dim ogrid As SAPbouiCOM.Grid = oform.Items.Item("8").Specific

                                ogrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto
                                Dim ss

                                If status = "Closed" Then
                                    ss = "SELECT T0.[DocNum] as 'Booking No', T1.[U_AE_Invno] as 'Invoice No', T0.[U_AE_Bname] as 'Customer', T0.[U_AE_Event] as 'Event', " & _
"T0.[U_AE_CDate] as 'Document Date', T0.[U_AE_Order] as 'Order By', T0.[U_AE_Statu] as 'Status' FROM [dbo].[@AE_CDRIVER]  T0 inner join [@AE_CDRIVER_R] T1 on T0.Docentry = T1.Docentry " & _
"where T0.[U_AE_Bcode] like '" & customer & "' and  isnull(T0.[U_AE_Event], '') like '" & event_ & "' and  isnull(T0.[U_AE_Statu], '') like 'Closed' " & _
"group by T0.[DocNum] , T1.[U_AE_Invno] , T0.[U_AE_Bname], T0.[U_AE_Event], " & _
"T0.[U_AE_CDate] , T0.[U_AE_Order] , T0.[U_AE_Statu] "
                                   
                                Else
                                    ss = "SELECT T0.[DocNum] as 'Booking No', T0.[U_AE_Bname] as 'Customer', T0.[U_AE_Event] as 'Event', T0.[U_AE_CDate] as 'Document Date', T0.[U_AE_Order] as 'Order By', T0.[U_AE_Statu] as 'Status' FROM [dbo].[@AE_CDRIVER]  T0 " & _
                                                                                  "where T0.[U_AE_Bcode] like '" & customer & "' and  isnull(T0.[U_AE_Event], '') like '" & event_ & "' and  isnull(T0.[U_AE_Statu], '') like '" & status & "'"
                                End If

                                oform.DataSources.DataTables.Item(0).ExecuteQuery(ss)
                                ogrid.DataTable = oform.DataSources.DataTables.Item("CDBR")
                                ogrid.AutoResizeColumns()

                                Dim ocol As SAPbouiCOM.EditTextColumn = ogrid.Columns.Item("Booking No")
                                ocol.LinkedObjectType = "AE_Cdriver"

                                ocol = ogrid.Columns.Item("Invoice No")
                                ocol.LinkedObjectType = 13

                                ogrid.Columns.Item("Booking No").ForeColor = RGB(20, 20, 200)
                                ogrid.Columns.Item("Status").ForeColor = RGB(20, 20, 200)
                                ogrid.Columns.Item("Booking No").TextStyle = 1
                                ogrid.Columns.Item("Customer").TextStyle = 1
                                ogrid.Columns.Item("Event").TextStyle = 1
                                ogrid.Columns.Item("Document Date").TextStyle = 1
                                ogrid.Columns.Item("Order By").TextStyle = 1

                                ogrid.Columns.Item("Status").TextStyle = 1
                                oform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
                                oform.Visible = True

                            Catch ex As Exception

                            End Try

                        End If
                    End If
                End If
            End If

            '-------------------------------------------------------------------------------------------------------------

            If pVal.FormUID = "DM" Then
                If pVal.Before_Action = True Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm

                                If DriverMAster(oform, Oapplication_CD) = False Then
                                    BubbleEvent = False
                                    Exit Try
                                End If

                            Catch ex As Exception

                            End Try
                            Exit Sub
                        End If
                    End If
                ElseIf pVal.Before_Action = False Then

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then

                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pVal
                        Dim sCFL_ID As String
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        Dim oForm As SAPbouiCOM.Form = Oapplication_CD.Forms.Item(FormUID)
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                        Try

                            If oCFLEvento.BeforeAction = False Then
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim oedit1 As SAPbouiCOM.EditText
                                oDataTable = oCFLEvento.SelectedObjects
                                If pVal.ItemUID = "44" Then
                                    oForm.Items.Item("44").Specific.string = oDataTable.GetValue("ItemCode", 0)
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then
                        Try
                            Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                            oform.Items.Item("49").Specific.String = oform.Items.Item("51").Specific.selected.description
                        Catch ex As Exception

                        End Try
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            If pVal.Action_Success = True Then
                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                orset.DoQuery("SELECT max(cast(isnull(T0.[U_AE_Dcode],0)as integer))  as 'M' FROM [dbo].[@AE_DRIVERM]  T0")
                                ' oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                                Dim max As String = orset.Fields.Item("M").Value + 1
                                oform.Items.Item("8").Specific.String = max.ToString.PadLeft(7, "0"c)
                                oform.Items.Item("40").Specific.String = Format(Now.Date, "yyyyMMdd") 'Now.Date
                                oform.DataBrowser.BrowseBy = "8"
                            End If

                        End If

                        If pVal.ItemUID = "42" Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                Dim ocheck As SAPbouiCOM.CheckBox = oform.Items.Item("42").Specific

                                If ocheck.Checked = True Then
                                    oform.Items.Item("44").Enabled = True
                                Else
                                    oform.Items.Item("44").Specific.String = ""
                                    oform.Items.Item("40").Specific.active = True
                                    oform.Items.Item("44").Enabled = False
                                End If
                            Catch ex As Exception

                            End Try
                            Exit Sub
                        End If
                    End If
                End If

            End If


            If pVal.FormUID = "Extended" Then
                If pVal.Before_Action = True Then

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        If pVal.ItemUID = "11" Then
                            Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                            Dim oform_CD As SAPbouiCOM.Form = Oapplication_CD.Forms.Item("CDB")
                            Dim oMAtrix As SAPbouiCOM.Matrix = oform_CD.Items.Item("Item_17").Specific

                            Try
                                Dim obutton As SAPbouiCOM.Button = oform.Items.Item("11").Specific
                                Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                If obutton.Caption = "Update" Then
                                    orset.DoQuery("update [@AE_EXTENDED] set  [U_AE_Rem] = '" & oform.Items.Item("10").Specific.String & "' WHERE [Code] = '" & oform.Items.Item("6").Specific.String & "'")

                                ElseIf obutton.Caption = "Add" Then
                                    orset.DoQuery("select max(cast(isnull(T0.code,0) as integer)) as 'm' from [@AE_EXTENDED] T0")
                                    Dim code As String = orset.Fields.Item("m").Value + 1
                                    orset.DoQuery("insert into [@AE_EXTENDED] ([Code],  [Name], [U_AE_Rem] , [U_AE_Dno], [U_AE_Lno], [U_AE_Object], [U_AE_ColID])  values ('" & code & "', '" & code & "', '" & oform.Items.Item("10").Specific.String & "' , '" & oform.Items.Item("2").Specific.String & "' , '" & oform.Items.Item("3").Specific.String & "' , '" & oform.Items.Item("5").Specific.String & "', '" & oform.Items.Item("7").Specific.String & "')")

                                End If

                                oMAtrix.Columns.Item(oform.Items.Item("7").Specific.String).Cells.Item(CInt(oform.Items.Item("3").Specific.String)).Specific.String = oform.Items.Item("10").Specific.String

                                oform.Close()
                            Catch ex As Exception
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                        End If

                        If pVal.ItemUID = "8" Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                oform.Close()
                            Catch ex As Exception
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                        End If
                    End If

                ElseIf pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE Then
                    Try
                        Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.Item("Extended")
                        oform.Visible = True

                    Catch ex As Exception
                        Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        BubbleEvent = False
                        Exit Try
                    End Try

                End If

            End If

            If pVal.FormUID = "CDSA" Then

                If pVal.Before_Action = False Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then

                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pVal
                        Dim sCFL_ID As String
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        Dim oForm As SAPbouiCOM.Form = Oapplication_CD.Forms.Item(FormUID)
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                        Try

                            If oCFLEvento.BeforeAction = False Then
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim oedit1 As SAPbouiCOM.EditText
                                oDataTable = oCFLEvento.SelectedObjects
                                If pVal.ItemUID = "Item_12" And pVal.ColUID = "V_0mjs" Then 'Driver Code
                                    Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("Item_12").Specific
                                    Try
                                        oMatrix.Columns.Item("V_1sjr").Cells.Item(pVal.Row).Specific.String = oDataTable.GetValue("U_AE_Vno", 0)
                                    Catch ex As Exception
                                    End Try
                                    oMatrix.Columns.Item("V_0sjr").Cells.Item(pVal.Row).Specific.String = oDataTable.GetValue("U_AE_Aliase", 0)
                                    oMatrix.Columns.Item("V_1mjs").Cells.Item(pVal.Row).Specific.String = oDataTable.GetValue("U_AE_Hphone", 0)
                                    oMatrix.Columns.Item("Team").Cells.Item(pVal.Row).Specific.String = oDataTable.GetValue("U_AE_TName", 0)

                                    oMatrix.Columns.Item("V_0mjs").Cells.Item(pVal.Row).Specific.String = oDataTable.GetValue("U_AE_Dcode", 0)

                                End If

                                If pVal.ItemUID = "Item_12" And pVal.ColUID = "V_1sjr" Then
                                    Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("Item_12").Specific
                                    Try
                                        oMatrix.Columns.Item("V_1sjr").Cells.Item(pVal.Row).Specific.String = oDataTable.GetValue("ItemCode", 0)
                                    Catch ex As Exception
                                    End Try
                                End If


                                If pVal.ItemUID = "19" Then

                                    Try
                                        oForm.Items.Item("20").Specific.String = oDataTable.GetValue("CardName", 0)
                                        oForm.Items.Item("19").Specific.String = oDataTable.GetValue("CardCode", 0)

                                    Catch ex As Exception
                                    End Try
                                End If

                            End If
                        Catch ex As Exception
                        End Try
                    End If

                    If pVal.ItemUID = "Item_7" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then
                        Try

                            Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                            oform.Items.Item("17").Specific.String = oform.Items.Item("Item_7").Specific.selected.description


                        Catch ex As Exception

                        End Try

                    End If
                End If


                If pVal.Before_Action = True Then

                    ''If pVal.EventType = SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED And pVal.ItemUID = "Item_12" And pVal.ColUID = "V_1sjr" Then
                    ''    Try
                    ''        Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                    ''        Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_12").Specific
                    ''        Dim oCFLs As SAPbouiCOM.ChooseFromList
                    ''        Dim oCons As SAPbouiCOM.Conditions
                    ''        Dim oCon As SAPbouiCOM.Condition
                    ''        Dim empty As New SAPbouiCOM.Conditions
                    ''        Dim vtype As String = ""

                    ''        For mjs As Integer = 1 To oMAtrix.RowCount
                    ''            If oMAtrix.IsRowSelected(mjs) = True Then
                    ''                vtype = oMAtrix.Columns.Item("V_0vtype").Cells.Item(mjs).Specific.value
                    ''                Exit For
                    ''            End If

                    ''        Next mjs

                    ''        oCFLs = oform.ChooseFromLists.Item("CFL_3")
                    ''        oCFLs.SetConditions(empty)
                    ''        oCons = oCFLs.GetConditions()
                    ''        oCon = oCons.Add()
                    ''        oCon.Alias = "QryGroup1"
                    ''        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    ''        oCon.CondVal = "Y"
                    ''        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    ''        oCon = oCons.Add()
                    ''        oCon.Alias = "QryGroup2"
                    ''        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    ''        oCon.CondVal = "Y"
                    ''        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    ''        oCon = oCons.Add()
                    ''        oCon.Alias = "U_AE_MODEL"
                    ''        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    ''        oCon.CondVal = vtype
                    ''        oCFLs.SetConditions(oCons)



                    ''    Catch ex As Exception

                    ''    End Try
                    ''End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvento = pVal
                        Dim sCFL_ID As String
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        Dim oForm As SAPbouiCOM.Form = Oapplication_CD.Forms.Item(FormUID)
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                        Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        If oCFLEvento.BeforeAction = True And pVal.ItemUID = "Item_12" And pVal.ColUID = "V_1sjr" Then
                            Dim oCFLs As SAPbouiCOM.ChooseFromList
                            Dim oCons As SAPbouiCOM.Conditions
                            Dim oCon As SAPbouiCOM.Condition
                            Dim empty As New SAPbouiCOM.Conditions
                            Dim vtype As String = ""

                            Dim oMAtrix As SAPbouiCOM.Matrix = oForm.Items.Item("Item_12").Specific
                            For mjs As Integer = 1 To oMAtrix.RowCount
                                If oMAtrix.IsRowSelected(mjs) = True Then
                                    vtype = oMAtrix.Columns.Item("V_0vtype").Cells.Item(mjs).Specific.value
                                    Exit For
                                End If

                            Next mjs


                            oCFLs = oForm.ChooseFromLists.Item("CFL_3")
                            oCFLs.SetConditions(empty)
                            oCons = oCFLs.GetConditions()
                            oCon = oCons.Add()
                            oCon.Alias = "QryGroup1"
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            oCon.CondVal = "Y"
                            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                            oCon = oCons.Add()
                            oCon.Alias = "QryGroup2"
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            oCon.CondVal = "Y"
                            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                            oCon = oCons.Add()
                            oCon.Alias = "U_AE_MODEL"
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            oCon.CondVal = vtype
                            oCFLs.SetConditions(oCons)

                        End If
                    End If



                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        If pVal.ItemUID = "Item_9" Then
                            Try

                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                ''If oform.Items.Item("Item_1").Specific.String = "" Then
                                ''    oform.Items.Item("Item_1").Specific.active = True
                                ''    Oapplication_CD.StatusBar.SetText("From Date should not be Empty ........... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                ''    BubbleEvent = False
                                ''    Exit Try
                                ''End If

                                ''If oform.Items.Item("Item_3").Specific.String = "" Then
                                ''    oform.Items.Item("Item_3").Specific.active = True
                                ''    Oapplication_CD.StatusBar.SetText("To Date should not be Empty ............. !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                ''    BubbleEvent = False
                                ''    Exit Try
                                ''End If

                                If Driver_assign_Load(oform) = False Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If


                            Catch ex As Exception
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub
                        End If

                        If pVal.ItemUID = "Item_14" Then

                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                Dim oMatrix1 As SAPbouiCOM.Matrix = oform.Items.Item("Item_12").Specific
                                Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim Docentry, LineID, Vno, Dname, Dcode, orno, Dhp, Team As String

                                For mjs As Integer = 1 To oMatrix1.RowCount
                                    If String.IsNullOrEmpty(oMatrix1.Columns.Item("Col_3").Cells.Item(mjs).Specific.String) And Not String.IsNullOrEmpty(oMatrix1.Columns.Item("V_0mjs").Cells.Item(mjs).Specific.String) Then

                                        If String.IsNullOrEmpty(oMatrix1.Columns.Item("V_1sjr").Cells.Item(mjs).Specific.String) Then
                                            Oapplication_CD.StatusBar.SetText("Vehicle no should not be empty .......... ! Line no " & mjs, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If

                                    End If

                                    Oapplication_CD.StatusBar.SetText("Validation Process Started ....... " & mjs, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                Next mjs

                                For mjs As Integer = 1 To oMatrix1.RowCount
                                    If String.IsNullOrEmpty(oMatrix1.Columns.Item("Col_3").Cells.Item(mjs).Specific.String) And Not String.IsNullOrEmpty(oMatrix1.Columns.Item("V_0mjs").Cells.Item(mjs).Specific.String) Then

                                        Docentry = oMatrix1.Columns.Item("Col_0").Cells.Item(mjs).Specific.String
                                        LineID = oMatrix1.Columns.Item("V_02SJR").Cells.Item(mjs).Specific.String
                                        ' orno = oMatrix1.Columns.Item("Col_3").Cells.Item(mjs).Specific.String
                                        Vno = oMatrix1.Columns.Item("V_1sjr").Cells.Item(mjs).Specific.String
                                        Dname = oMatrix1.Columns.Item("V_0sjr").Cells.Item(mjs).Specific.String
                                        Dcode = oMatrix1.Columns.Item("V_0mjs").Cells.Item(mjs).Specific.String
                                        Dhp = oMatrix1.Columns.Item("V_1mjs").Cells.Item(mjs).Specific.String
                                        Team = oMatrix1.Columns.Item("Team").Cells.Item(mjs).Specific.String

                                        orset.DoQuery("SELECT max(cast(isnull(T0.[U_AE_Ono],50000) as integer)) as 'no' FROM [dbo].[@AE_CDRIVER_R]  T0")
                                        'Dim orno1 As Integer
                                        orno = orset.Fields.Item("no").Value + 1
                                        ' MsgBox(orno.ToString.PadLeft(10, "0"c))
                                        orset.DoQuery("update [@AE_CDRIVER_R] set U_AE_Vno = '" & Vno & "', U_AE_Dname = '" & Dname & "', U_AE_Dcode = '" & Dcode & "', U_AE_Ono = '" & orno.ToString.PadLeft(6, "0"c) & "', U_AE_DHP = '" & Dhp & "', U_AE_Tref = '" & Team & "' where DocEntry = '" & Docentry & "' and  LineId = '" & LineID & "'")

                                        Oapplication_CD.StatusBar.SetText("Assigning Driver for CD Booking  ....... " & Docentry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    End If
                                Next mjs

                                Oapplication_CD.StatusBar.SetText("Assigning Driver Has Been Completed Successfully  ....... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                If Driver_assign_Load(oform) = False Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub
                        End If

                        If pVal.ItemUID = "Item_16" Then

                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                Dim oMatrix1 As SAPbouiCOM.Matrix = oform.Items.Item("Item_12").Specific
                                Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim Docentry, LineID, Vno, Dname, Dcode, orno As String

                                For mjs As Integer = 1 To oMatrix1.RowCount
                                    If oMatrix1.IsRowSelected(mjs) = True Then
                                        Docentry = oMatrix1.Columns.Item("Col_0").Cells.Item(mjs).Specific.String
                                        LineID = oMatrix1.Columns.Item("V_02SJR").Cells.Item(mjs).Specific.String
                                        orno = oMatrix1.Columns.Item("Col_3").Cells.Item(mjs).Specific.String
                                        Exit For
                                    End If
                                Next mjs

                                If orno = "" Then
                                    Oapplication_CD.MessageBox("No vehicle has been assigned for this booking ........ ! ", 1, "Ok")
                                    BubbleEvent = False
                                    Exit Try
                                End If


                                ' MsgBox(orno.ToString.PadLeft(10, "0"c))
                                orset.DoQuery("update [@AE_CDRIVER_R] set U_AE_Vno = '', U_AE_Dname = '', U_AE_Dcode = '', U_AE_Ono = '', U_AE_DHP = '', U_AE_Tref = '' where DocEntry = '" & Docentry & "' and  LineId = '" & LineID & "'")


                                If Driver_assign_Load(oform) = False Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub
                        End If



                        If pVal.ItemUID = "1000001" And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            Dim DA_Thread As System.Threading.Thread
                            Try

                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_12").Specific
                                Dim OCNO As String
                                Docnum = 0
                                For mjs As Integer = 1 To oMatrix.RowCount
                                    If oMatrix.IsRowSelected(mjs) = True Then
                                        Docnum = CInt(oMatrix.Columns.Item("Col_0").Cells.Item(mjs).Specific.String)
                                        OCNO = oMatrix.Columns.Item("Col_3").Cells.Item(mjs).Specific.String
                                        Exit For
                                    End If
                                Next mjs

                                If Docnum = 0 Then
                                    Oapplication_CD.StatusBar.SetText("Kindly select the to view .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Try
                                Else
                                    If OCNO = "" Then
                                        Oapplication_CD.StatusBar.SetText("Order Chit no should not be empty  .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Try
                                    End If
                                End If

                                DA_Thread = New System.Threading.Thread(AddressOf Class_Report.Report_CallingFunction)
                                Class_Report.oApplication = Oapplication_CD
                                Class_Report.oCompany = oCompany_TB_CD
                                Class_Report.Report_Name = "AE_RP001_OrderChit.rpt"
                                Class_Report.Report_Parameter = "@OrderChitNo"
                                Class_Report.Docnum = OCNO
                                If DA_Thread.IsAlive Then
                                    Oapplication_CD.MessageBox("Report is already open....")
                                Else
                                    DA_Thread.TrySetApartmentState(Threading.ApartmentState.STA)
                                    Oapplication_CD.StatusBar.SetText("RMG Report Opening in process ......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    DA_Thread.Start()
                                End If

                            Catch ex As Exception
                                DA_Thread.Abort()
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End Try
                            Exit Sub
                        End If


                    End If


                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED Then
                        If pVal.ItemUID = "Item_12" And pVal.ColUID = "Col_0" Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_12").Specific
                                Dim docentry As String = omatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.String

                                LoadFromXML("ChaufferDriver_Booking.srf", Oapplication_CD)
                                oform = Oapplication_CD.Forms.Item("CDB")
                                oform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                oform.Items.Item("Item_13").Enabled = True
                                oform.Items.Item("Item_13").Specific.String = docentry
                                oform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                oform.Items.Item("Item_13").Enabled = False
                                Dim ocombobutton As SAPbouiCOM.ButtonCombo = oform.Items.Item("CT").Specific
                                ocombobutton.ValidValues.Add("Copy To", "Copy To A/R Invoice")

                                ' oform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                oform.Visible = True
                            Catch ex As Exception
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub
                        End If

                        If pVal.ItemUID = "Item_12" And (pVal.ColUID = "Col_5" Or pVal.ColUID = "Col_7") Then

                            Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                            Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_12").Specific
                            Dim ColID, Title As String
                            Dim obutton As SAPbouiCOM.Button
                            Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            If pVal.ColUID = "Col_5" Then
                                ColID = "Col_3"
                                Title = "Guest Name"
                            Else
                                ColID = "Col_7"
                                Title = "Pickup Location"
                            End If

                            orset.DoQuery("SELECT T0.[code], T0.[U_AE_Rem], T0.[U_AE_Dno], T0.[U_AE_Lno], T0.[U_AE_Object] FROM [dbo].[@AE_EXTENDED]  T0 WHERE T0.[U_AE_Dno] = '" & omatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.String & "' and  T0.[U_AE_Lno] = '" & omatrix.Columns.Item("V_02SJR").Cells.Item(pVal.Row).Specific.String & "' and  T0.[U_AE_Object] = '" & "CD" & "' and T0.[U_AE_ColID] = '" & ColID & "'")

                            LoadFromXML("ExtendedBox.srf", Oapplication_CD)
                            oform = Oapplication_CD.Forms.Item("Extended")
                            oform.Freeze(True)
                            oform.Title = Title
                            obutton = oform.Items.Item("1").Specific
                            oform.Items.Item("10").Specific.string = orset.Fields.Item("U_AE_Rem").Value
                            obutton.Caption = "Ok"
                            oform.Freeze(False)
                            oform.Visible = True

                        End If

                    End If
                End If
            End If

            If pVal.FormUID = "CDPL" Then

                If pVal.Before_Action = False Then

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                        If pVal.ItemUID = "1" And pVal.Action_Success = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                oform.Close()
                                LoadFromXML("Pricelist.srf", Oapplication_CD)
                                oform = Oapplication_CD.Forms.Item("CDPL")
                                oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                                If Price_List_Binding(oform) = False Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                oform.Visible = True

                            Catch ex As Exception
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub
                        End If
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.InnerEvent = False Then
                        If pVal.ItemUID = "Item_6" And pVal.ColUID = "Col_0" Then
                            Dim oform As SAPbouiCOM.Form
                            Try
                                oform = Oapplication_CD.Forms.Item("CDPL")
                                oform.Freeze(True)
                                Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_6").Specific
                                Dim ocombo As SAPbouiCOM.ComboBox
                                If oMAtrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.value <> "-" Then
                                    oMAtrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.String = oMAtrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific.selected.description
                                    If pVal.Row = oMAtrix.RowCount Then
                                        oMAtrix.AddRow()
                                        oMAtrix.Columns.Item("#").Cells.Item(oMAtrix.RowCount).Specific.String = oMAtrix.RowCount
                                        ocombo = oMAtrix.Columns.Item("Col_0").Cells.Item(oMAtrix.RowCount).Specific
                                        ocombo.Select("-")
                                    End If
                                End If
                                oform.Freeze(False)
                            Catch ex As Exception
                                oform.Freeze(False)
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                            Exit Sub
                        End If


                    End If

                End If

                If pVal.Before_Action = True Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then


                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            Try

                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_6").Specific

                                If oform.Items.Item("13").Specific.value = "" Then
                                    oform.Items.Item("13").Specific.active = True
                                    Oapplication_CD.StatusBar.SetText("Default option should not be empty ........ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Try
                                End If
                                For mjs As Integer = 1 To oMAtrix.RowCount
                                    oMAtrix.Columns.Item("#").Cells.Item(mjs).Specific.string = mjs
                                Next mjs
                            Catch ex As Exception

                            End Try
                            Exit Sub
                        End If
                        If pVal.ItemUID = "1" And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Try
                                Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                                If oform.Items.Item("Item_1").Specific.String = "" Then
                                    oform.Items.Item("Item_1").Specific.active = True
                                    Oapplication_CD.StatusBar.SetText("Price List Code should not be Empty ........ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Try
                                End If

                                If oform.Items.Item("Item_3").Specific.String = "" Then
                                    oform.Items.Item("Item_3").Specific.active = True
                                    Oapplication_CD.StatusBar.SetText("Price List Name should not be Empty ........ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Try
                                End If

                                If oform.Items.Item("13").Specific.value = "" Then
                                    oform.Items.Item("13").Specific.active = True
                                    Oapplication_CD.StatusBar.SetText("Default option should not be empty ........ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Exit Try
                                End If

                                Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_6").Specific

                                If oMAtrix.RowCount = 1 Then

                                    oMAtrix.Columns.Item("#").Cells.Item(1).Specific.string = "1"

                                    If oMAtrix.Columns.Item("Col_0").Cells.Item(1).Specific.value = "-" Then
                                        oMAtrix.Columns.Item("Col_0").Cells.Item(1).Specific.active = True
                                        Oapplication_CD.StatusBar.SetText("Vehicle Type Should not be Empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Try
                                    End If

                                    If oMAtrix.Columns.Item("Col_1").Cells.Item(1).Specific.String = "0.00" Then
                                        oMAtrix.Columns.Item("Col_1").Cells.Item(1).Specific.active = True
                                        Oapplication_CD.StatusBar.SetText("Hourly Rate should not be Empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Try
                                    End If

                                    If oMAtrix.Columns.Item("Col_2").Cells.Item(1).Specific.String = "0.00" Then
                                        oMAtrix.Columns.Item("Col_2").Cells.Item(1).Specific.active = True
                                        Oapplication_CD.StatusBar.SetText("Surcharge One way should not be Empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Try
                                    End If

                                    If oMAtrix.Columns.Item("Col_3").Cells.Item(1).Specific.String = "0.00" Then
                                        oMAtrix.Columns.Item("Col_3").Cells.Item(1).Specific.active = True
                                        Oapplication_CD.StatusBar.SetText("Surcharge One Disposal should not be Empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Try

                                    End If
                                End If

                                For mjs As Integer = 1 To oMAtrix.RowCount - 1
                                    oMAtrix.Columns.Item("#").Cells.Item(mjs).Specific.string = mjs

                                    If oMAtrix.Columns.Item("Col_0").Cells.Item(mjs).Specific.value <> "-" Then
                                        If oMAtrix.Columns.Item("Col_0").Cells.Item(mjs).Specific.value = "-" Then
                                            oMAtrix.Columns.Item("Col_0").Cells.Item(mjs).Specific.active = True
                                            Oapplication_CD.StatusBar.SetText("Vehicle Type Should not be Empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Try
                                        End If

                                        If oMAtrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.String = "0.00" Then
                                            oMAtrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.active = True
                                            Oapplication_CD.StatusBar.SetText("Hourly Rate should not be Empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Try
                                        End If

                                        If oMAtrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.String = "0.00" Then
                                            oMAtrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.active = True
                                            Oapplication_CD.StatusBar.SetText("Surcharge One way should not be Empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Try
                                        End If

                                        If oMAtrix.Columns.Item("Col_3").Cells.Item(mjs).Specific.String = "0.00" Then
                                            oMAtrix.Columns.Item("Col_3").Cells.Item(mjs).Specific.active = True
                                            Oapplication_CD.StatusBar.SetText("Surcharge One Disposal should not be Empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Try

                                        End If
                                    End If
                                Next
                            Catch ex As Exception
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                BubbleEvent = False
                                Exit Try
                            End Try
                        End If
                    End If
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Oapplication_CD_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles Oapplication_CD.MenuEvent

        Try
            If pVal.MenuUID = "CDVB" And pVal.BeforeAction = True Then
                Dim oform As SAPbouiCOM.Form
                Try
                    LoadFromXML("ChaufferDriver_Booking.srf", Oapplication_CD)
                    oform = Oapplication_CD.Forms.Item("CDB")
                    oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                    oform.Freeze(True)
                    If Chauffer_Driver_binding(oform) = False Then
                        oform.Freeze(False)
                        BubbleEvent = False
                        Exit Try
                    End If
                    oform.Freeze(False)
                    oform.Visible = True

                Catch ex As Exception
                    oform.Freeze(False)
                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try
                Exit Sub
            End If


            '------------------------------ New -------------------------------------------------------------
            If pVal.MenuUID = "CDBR" And pVal.BeforeAction = True Then

                Try
                    LoadFromXML("ChaufferDriver_BookingReport.srf", Oapplication_CD)
                    Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.Item("CDBR")
                    oform.Visible = True
                    Try
                        oform.DataSources.DataTables.Add("CDBR")
                    Catch ex As Exception

                    End Try
                    Dim ocombo As SAPbouiCOM.ComboBox = oform.Items.Item("7").Specific
                    ocombo.ValidValues.Add("All", "")
                    ocombo.Select("All")
                    Dim ogrid As SAPbouiCOM.Grid = oform.Items.Item("8").Specific

                    ogrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto
                    oform.DataSources.DataTables.Item(0).ExecuteQuery("SELECT T0.[DocNum] as 'Booking No', T0.[U_AE_Bname] as 'Customer', T0.[U_AE_Event] as 'Event', T0.[U_AE_CDate] as 'Document Date', T0.[U_AE_Order] as 'Order By', T0.[U_AE_Statu] as 'Status' FROM [dbo].[@AE_CDRIVER]  T0 where T0.[DocNum] = ''")
                    ogrid.DataTable = oform.DataSources.DataTables.Item("CDBR")
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
                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try

            End If

            ' ----------------------------------------------------------------------------------------------------


            If pVal.MenuUID = "CDSP" And pVal.BeforeAction = True Then

                Try
                    LoadFromXML("Scheduling&Assigning.srf", Oapplication_CD)
                    Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.Item("CDSA")
                    oform.Visible = True
                    If Driver_assign_Binding(oform) = False Then
                        BubbleEvent = False
                        Exit Try
                    End If

                Catch ex As Exception
                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try

            End If

            If pVal.MenuUID = "CDPS" And pVal.BeforeAction = True Then

                Try
                    LoadFromXML("Pricelist.srf", Oapplication_CD)
                    Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.Item("CDPL")
                    oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                    oform.Visible = True
                    If Price_List_Binding(oform) = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If


                Catch ex As Exception
                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try
            End If

            If pVal.MenuUID = "VTYPE" And pVal.BeforeAction = True Then
                Try
                    Oapplication_CD.ActivateMenuItem(UDTform("AE_VTYPE - AE_Vehicle Type", Oapplication_CD))
                Catch ex As Exception
                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try
            End If

            If pVal.MenuUID = "VMAS" And pVal.BeforeAction = True Then
                Try
                    Oapplication_CD.ActivateMenuItem("3073")
                    Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                    oform.Title = "Vehicle Master Data"

                Catch ex As Exception
                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try
            End If

            If pVal.MenuUID = "VTM" And pVal.BeforeAction = True Then
                Try
                    Oapplication_CD.ActivateMenuItem(UDTform("AE_TEAMM - AE_Team Master", Oapplication_CD))

                Catch ex As Exception
                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try
            End If


            If pVal.MenuUID = "COMM" And pVal.BeforeAction = True Then
                Try
                    Oapplication_CD.ActivateMenuItem(UDTform("AE_COMMISION - RMG Sales Commision", Oapplication_CD))
                Catch ex As Exception
                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try
            End If

            If pVal.MenuUID = "STYP" And pVal.BeforeAction = True Then
                Try
                    Oapplication_CD.ActivateMenuItem(UDTform("AE_SERVICETYPE - RMG Service Type", Oapplication_CD))
                Catch ex As Exception
                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try
            End If

            If pVal.MenuUID = "AGYDTL" And pVal.BeforeAction = True Then
                Try
                    Oapplication_CD.ActivateMenuItem(UDTform("AE_AGENCY - AE_Agency Details", Oapplication_CD))
                Catch ex As Exception
                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try
            End If

            If pVal.MenuUID = "GLD" And pVal.BeforeAction = True Then
                Try
                    Oapplication_CD.ActivateMenuItem(UDTform("AE_GLACC - AE_GL Determination", Oapplication_CD))
                Catch ex As Exception
                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try
            End If

            If pVal.MenuUID = "SCT" And pVal.BeforeAction = True Then
                Try
                    Oapplication_CD.ActivateMenuItem(UDTform("AE_NTIME - AE_Surcharge Time", Oapplication_CD))
                Catch ex As Exception
                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try
            End If

            If pVal.MenuUID = "VDMAS" And pVal.BeforeAction = True Then
                Try
                    LoadFromXML("DriverMaster.srf", Oapplication_CD)
                    Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.Item("DM")
                    Dim ocombo As SAPbouiCOM.ComboBox = oform.Items.Item("51").Specific

                    Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    orset.DoQuery("SELECT max(cast(isnull(T0.[U_AE_Dcode],0)as integer))  as 'M' FROM [dbo].[@AE_DRIVERM]  T0")
                    oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                    Dim max As String = orset.Fields.Item("M").Value + 1
                    oform.Visible = True
                    orset.DoQuery("SELECT T0.[Code], T0.[Name] FROM [dbo].[@AE_TEAMM]  T0")

                    For mjs As Integer = 1 To orset.RecordCount
                        ocombo.ValidValues.Add(orset.Fields.Item("Code").Value, orset.Fields.Item("Name").Value)
                        orset.MoveNext()
                    Next mjs

                    oform.Items.Item("8").Specific.String = max.ToString.PadLeft(7, "0"c)
                    oform.Items.Item("40").Specific.String = Format(Now.Date, "yyyyMMdd") 'Now.Date
                    oform.DataBrowser.BrowseBy = "8"

                Catch ex As Exception
                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try
            End If

            If pVal.MenuUID = "1281" And pVal.BeforeAction = False Then
                Try
                    Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                    If oform.UniqueID = "CDB" Then
                        oform.Items.Item("Item_13").Enabled = True
                        oform.Items.Item("Item_16").Enabled = True
                    End If
                Catch ex As Exception
                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try
            End If

            If (pVal.MenuUID = "1288" Or pVal.MenuUID = "1289" Or pVal.MenuUID = "1290" Or pVal.MenuUID = "1291") And pVal.BeforeAction = False Then
                Try
                    Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                    If oform.UniqueID = "CDB" Then
                        oform.Freeze(True)
                        CB_Navigation(oform)
                        oform.Freeze(False)
                        ' oform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                        ' oform.Items.Item("Item_16").Enabled = True
                    End If
                Catch ex As Exception
                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try
            End If

            If pVal.MenuUID = "1282" And pVal.BeforeAction = False Then

                Try
                    Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm

                    Select Case oform.UniqueID

                        Case "CDB"
                            Try
                                oform.Freeze(True)
                                If Chauffer_Driver_binding_AddMode(oform) = False Then
                                    BubbleEvent = False
                                    Exit Try
                                End If
                                oform.Items.Item("1000003").Enabled = True
                                oform.Items.Item("Item_4").Enabled = True
                                oform.Items.Item("42").Enabled = True
                                oform.Items.Item("Item_11").Enabled = True
                                oform.Items.Item("Item_16").Enabled = True
                                oform.Items.Item("Item_19").Enabled = True
                                oform.Freeze(False)
                                oform.Visible = True
                            Catch ex As Exception
                                oform.Freeze(False)
                            End Try


                        Case "SDB"
                            Try
                                oform.Freeze(True)
                                Dim Tmp_val As String = oform.BusinessObject.GetNextSerialNumber(NextSerialNo(oCompany_TB_CD, Oapplication_CD, "AE_Sbooking"))
                                oform.Items.Item("Item_14").Specific.String = Tmp_val

                                'oform.Items.Item("Item_20").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                oform.Items.Item("Item_16").Specific.String = Format(Now.Date, "yyyyMMdd") 'Now.Date
                                oform.Items.Item("Item_14").Enabled = False
                                oform.Visible = True
                                Dim ocombo As SAPbouiCOM.ComboBox = oform.Items.Item("Item_18").Specific
                                ocombo.Select("Open")
                                oform.PaneLevel = 4
                                oform.Items.Item("Item_154").Specific.String = "9999"
                                oform.PaneLevel = 3

                                ocombo = oform.Items.Item("226").Specific
                                ocombo.Select("No")
                                ocombo = oform.Items.Item("228").Specific
                                ocombo.Select("-")


                                Dim opt As SAPbouiCOM.OptionBtn = oform.Items.Item("195").Specific
                                Dim opt1 As SAPbouiCOM.OptionBtn = oform.Items.Item("196").Specific
                                Dim opt2 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_104").Specific
                                Dim opt3 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_105").Specific
                                Dim opt4 As SAPbouiCOM.OptionBtn = oform.Items.Item("Item_106").Specific

                                opt1.GroupWith("195")
                                opt1.Selected = True

                                opt3.GroupWith("Item_104")
                                opt4.GroupWith("Item_104")
                                opt3.Selected = True
                                opt2.Selected = True

                                oform.DataBrowser.BrowseBy = "Item_14"
                                oform.Items.Item("233").Enabled = False
                               
                                oform.Items.Item("217").Specific.String = oCompany_TB_CD.UserName
                                oform.Items.Item("218").Specific.String = Company_Name
                               

                                oform.Items.Item("Item_22").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                oform.Freeze(False)
                                oform.Items.Item("188").Specific.String = "SO"

                            Catch ex As Exception
                                oform.Freeze(False)
                                Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Exit Sub
                            End Try
                           
                        Case "CDPL"


                            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_6").Specific
                            Dim ocombo As SAPbouiCOM.ComboBox
                            oform.Items.Item("Item_5").Specific.String = Format(Now.Date, "yyyyMMdd") 'Now.Date
                            Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            orset.DoQuery("SELECT *  FROM [dbo].[@AE_VTYPE]  T0")
                            oMatrix.AddRow()
                            oMatrix.Columns.Item("#").Cells.Item(oMatrix.RowCount).Specific.String = oMatrix.RowCount
                            ocombo = oMatrix.Columns.Item("Col_0").Cells.Item(1).Specific
                            ocombo.Select("-")
                            oMatrix.Columns.Item("Col_0").Editable = True
                            oMatrix.Columns.Item("Col_1").Editable = True
                            oMatrix.Columns.Item("Col_2").Editable = True
                            oMatrix.Columns.Item("Col_3").Editable = True
                            oform.DataBrowser.BrowseBy = "Item_1"

                        Case "TPOD"

                            Dim Tmp_val As String = oform.BusinessObject.GetNextSerialNumber(NextSerialNo(oCompany_TB_CD, Oapplication_CD, "AE_TrafficO"))
                            oform.Items.Item("Item_14").Specific.String = Tmp_val

                            oform.Items.Item("105").Specific.select("Open")
                            oform.Items.Item("Item_20").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oform.Visible = True
                            oform.Items.Item("Item_16").Specific.String = Format(Now.Date, "yyyyMMdd") ' Now.Date
                            oform.Items.Item("1000001").Enabled = False
                            oform.Items.Item("Item_14").Enabled = False

                            oform.Items.Item("108").Specific.String = oCompany_TB_CD.UserName
                            oform.Items.Item("109").Specific.String = Company_Name


                        Case "ACD"

                            Dim Tmp_val As String = oform.BusinessObject.GetNextSerialNumber(NextSerialNo(oCompany_TB_CD, Oapplication_CD, "AE_Accident"))
                            oform.Items.Item("Item_14").Specific.String = Tmp_val

                            oform.Items.Item("Item_20").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            oform.Visible = True
                            oform.Items.Item("Item_16").Specific.String = Format(Now.Date, "yyyyMMdd") 'Now.Date
                            oform.Items.Item("134").Specific.select("Open")
                            oform.Items.Item("Item_14").Enabled = False
                            oform.Items.Item("1000001").Enabled = False
                            oform.Items.Item("138").Specific.String = oCompany_TB_CD.UserName
                            oform.Items.Item("139").Specific.String = Company_Name


                        Case "SM"

                            Dim Tmp_val As String = oform.BusinessObject.GetNextSerialNumber(NextSerialNo(oCompany_TB_CD, Oapplication_CD, "AE_SM"))
                            oform.Items.Item("Item_7").Specific.String = Tmp_val
                            oform.Items.Item("Item_9").Specific.String = Format(Now.Date, "yyyyMMdd") 'Now.Date
                            oform.Items.Item("Item_7").Enabled = False
                            oform.PaneLevel = 1
                            oform.Items.Item("21").Enabled = False
                            Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_13").Specific
                            omatrix.Columns.Item("Col_0").Editable = True
                            Dim oCombo As SAPbouiCOM.ComboBox = oform.Items.Item("23").Specific
                            oCombo.Select("Open")
                            omatrix.AddRow()
                            omatrix.Columns.Item("#").Cells.Item(omatrix.RowCount).Specific.String = omatrix.RowCount
                            oform.Items.Item("1000002").Specific.String = oCompany_TB_CD.UserName
                            oform.Items.Item("24").Specific.String = Company_Name

                        Case "DM"

                            ' Dim ocombo As SAPbouiCOM.ComboBox = oform.Items.Item("51").Specific

                            Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            orset.DoQuery("SELECT max(cast(isnull(T0.[U_AE_Dcode],0)as integer))  as 'M' FROM [dbo].[@AE_DRIVERM]  T0")
                            oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                            Dim max As String = orset.Fields.Item("M").Value + 1
                            oform.Visible = True
                            orset.DoQuery("SELECT T0.[Code], T0.[Name] FROM [dbo].[@AE_TEAMM]  T0")

                            'For mjs As Integer = 1 To orset.RecordCount
                            '    ocombo.ValidValues.Add(orset.Fields.Item("Code").Value, orset.Fields.Item("Name").Value)
                            '    orset.MoveNext()
                            'Next mjs

                            oform.Items.Item("8").Specific.String = max.ToString.PadLeft(7, "0"c)
                            oform.Items.Item("40").Specific.String = Format(Now.Date, "yyyyMMdd") 'Now.Date
                            oform.DataBrowser.BrowseBy = "8"

                    End Select



                Catch ex As Exception
                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try

            End If

            If pVal.MenuUID = "3045645" And pVal.BeforeAction = True Then
                Try
                    Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm
                    If oform.UniqueID = "CDB" Then
                        Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific
                        For mjs As Integer = 1 To omatrix.RowCount
                            If omatrix.IsRowSelected(mjs) = True Then
                                If omatrix.Columns.Item("Col_10").Cells.Item(mjs).Specific.String = "" Then
                                    omatrix.DeleteRow(mjs)
                                    If oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    Exit For
                                Else
                                    Oapplication_CD.StatusBar.SetText("Can`t delete the row, the Order chit no. is generated ........ !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                End If
                            End If
                        Next mjs
                    End If

                    If oform.UniqueID = "CDPL" Then
                        Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_6").Specific
                        For mjs As Integer = 1 To omatrix.RowCount
                            If omatrix.IsRowSelected(mjs) = True Then
                                If omatrix.Columns.Item("Col_0").Cells.Item(omatrix.RowCount).Specific.value <> "" Then
                                    omatrix.DeleteRow(mjs)
                                    If oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    Exit For
                                End If
                            End If
                        Next mjs
                    End If


                Catch ex As Exception
                    Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    BubbleEvent = False
                    Exit Try
                End Try
            End If


        Catch ex As Exception
            Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            BubbleEvent = False
            Exit Try
        End Try
    End Sub

    Private Function Chauffer_Driver_Validation(ByRef oform As SAPbouiCOM.Form) As Boolean
        Try

            If oform.Items.Item("Item_4").Specific.String = "" Then
                oform.Items.Item("Item_4").Specific.active = True
                Oapplication_CD.StatusBar.SetText("Billing To should not be Empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False
            End If

            If oform.Items.Item("Item_16").Specific.value = "" Then
                oform.Items.Item("Item_16").Specific.active = True
                Oapplication_CD.StatusBar.SetText("Document Status should be Open .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False
            End If

            If oform.Items.Item("Item_1").Specific.String = "" Then
                oform.Items.Item("Item_1").Specific.active = True
                Oapplication_CD.StatusBar.SetText("Issued by should not be Empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False
            End If


            If oform.Items.Item("Item_19").Specific.String = "" Then
                oform.Items.Item("Item_19").Specific.active = True
                Oapplication_CD.StatusBar.SetText("Document Date Should not be Empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False
            End If

            Dim oMAtrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific

            For mjs As Integer = 1 To oMAtrix.RowCount
                'oMAtrix.Columns.Item("Date").Cells.Item(mjs).Specific.String = mjs
                ' MsgBox(oMAtrix.Columns.Item("Date").Cells.Item(mjs).Specific.String)
                If oMAtrix.Columns.Item("Date").Cells.Item(mjs).Specific.String <> "" Then

                    If Trim(oform.Items.Item("Item_16").Specific.value) = "Open" Then
                        If oMAtrix.Columns.Item("Col_0").Cells.Item(mjs).Specific.value = "-" Then
                            oMAtrix.Columns.Item("Col_0").Cells.Item(mjs).Specific.active = True
                            Oapplication_CD.StatusBar.SetText("Vehicle Type Should not be Empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            Return False
                        End If

                        If oMAtrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.value = "-" Then
                            oMAtrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.active = True
                            Oapplication_CD.StatusBar.SetText("Service Type Should not be Empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            Return False

                        ElseIf Trim(oMAtrix.Columns.Item("Col_1").Cells.Item(mjs).Specific.value) = "Disposal" Then

                            If oMAtrix.Columns.Item("Col_9").Cells.Item(mjs).Specific.value = "" Then
                                oMAtrix.Columns.Item("Col_9").Cells.Item(mjs).Specific.active = True
                                Oapplication_CD.StatusBar.SetText("Est. End time Should not be Empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Return False
                            End If


                        End If
                        If oMAtrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.value = "" Then
                            oMAtrix.Columns.Item("Col_2").Cells.Item(mjs).Specific.active = True
                            Oapplication_CD.StatusBar.SetText("Pickup Time Should not be Empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            Return False
                        End If

                        If oMAtrix.Columns.Item("Col_3").Cells.Item(mjs).Specific.value = "." Then
                            ' oMAtrix.Columns.Item("Col_3").Cells.Item(mjs).Specific.active = True
                            Oapplication_CD.StatusBar.SetText("Guest Name Should not be Empty .......... ! Line " & mjs, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            Return False
                        End If

                        ''If oMAtrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.value = "" Then
                        ''    oMAtrix.Columns.Item("Col_4").Cells.Item(mjs).Specific.active = True
                        ''    Oapplication_CD.StatusBar.SetText("Guest HP Should not be Empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        ''    Return False
                        ''End If

                        If oMAtrix.Columns.Item("Col_7").Cells.Item(mjs).Specific.value = "." Then
                            'oMAtrix.Columns.Item("Col_7").Cells.Item(mjs).Specific.active = True
                            Oapplication_CD.StatusBar.SetText("Pickup Location Should not be Empty .......... ! Line " & mjs, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            Return False
                        End If

                        If oMAtrix.Columns.Item("Col_8").Cells.Item(mjs).Specific.value = "." Then
                            'oMAtrix.Columns.Item("Col_8").Cells.Item(mjs).Specific.active = True
                            Oapplication_CD.StatusBar.SetText("Drop Location Should not be Empty .......... ! Line " & mjs, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            Return False
                        End If

                        If oMAtrix.Columns.Item("Col_9").Cells.Item(mjs).Specific.value = "" Then
                            ' oMAtrix.Columns.Item("Col_9").Cells.Item(mjs).Specific.active = True
                            Oapplication_CD.StatusBar.SetText("Drop Time Should be Empty .......... !Line  " & mjs, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            Return False
                        End If
                    End If

                    If Trim(oform.Items.Item("Item_16").Specific.value) = "Billing" Then
                        If oMAtrix.Columns.Item("Col_9").Cells.Item(mjs).Specific.String = "" Then
                            Oapplication_CD.StatusBar.SetText("Drop Time should not be Empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            Return False
                        End If

                        If oMAtrix.Columns.Item("Col_13").Cells.Item(mjs).Specific.String = "" Then
                            Oapplication_CD.StatusBar.SetText("Vehicle No should not be empty ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            Return False
                        End If

                        If oMAtrix.Columns.Item("Col_14").Cells.Item(mjs).Specific.String = "" Then
                            Oapplication_CD.StatusBar.SetText("Driver Name should not be Empty .......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            Return False
                        End If
                    End If

                End If

            Next mjs


            If oform.Items.Item("Item_16").Specific.selected.value = "Closed" And oform.Items.Item("Item_16").Enabled = True Then
                oform.Items.Item("Item_16").Specific.active = True
                Oapplication_CD.StatusBar.SetText("Can`t set status manually to Closed ......... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False
            End If

            If Trim(oform.Items.Item("Item_16").Specific.value) = "Billing" Then
                oform.Items.Item("Item_4").Specific.active = True
                oform.Items.Item("CT").Enabled = True
                oform.Items.Item("Item_16").Enabled = False
            Else
                oform.Items.Item("Item_4").Specific.active = True
                oform.Items.Item("CT").Enabled = False
                oform.Items.Item("Item_16").Enabled = True
            End If


            If oMAtrix.RowCount > 1 And oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                If oMAtrix.Columns.Item("Date").Cells.Item(oMAtrix.RowCount).Specific.String = "" Then
                    oMAtrix.DeleteRow(oMAtrix.RowCount)
                End If
            End If

            Return True

        Catch ex As Exception
            Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function

    Private Function Chauffer_Driver_binding(ByRef oform As SAPbouiCOM.Form) As Boolean

        Try
            oform.Freeze(True)
            Dim Tmp_val As String = oform.BusinessObject.GetNextSerialNumber(NextSerialNo(oCompany_TB_CD, Oapplication_CD, "AE_Cdriver"))
            oform.Items.Item("Item_13").Specific.String = Tmp_val
            Dim ocombo As SAPbouiCOM.ComboBox
            ocombo = oform.Items.Item("Item_16").Specific
            ocombo.Select("Open")
            oform.Items.Item("Item_19").Specific.String = Format(Now.Date, "yyyyMMdd") 'Format(Now.Date, "dd MMM yyyy,ddd")
            ' oform.Items.Item("Item_19").Specific.String = Now.Date

            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific
            Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset.DoQuery("SELECT *  FROM [dbo].[@AE_VTYPE]  T0")
            oMatrix.AddRow()
            oMatrix.Columns.Item("#").Cells.Item(1).Specific.string = "1"
            ocombo = oMatrix.Columns.Item("Col_0").Cells.Item(1).Specific
            ocombo.ValidValues.Add("-", "-")
            For mjs As Integer = 1 To orset.RecordCount
                ocombo.ValidValues.Add(orset.Fields.Item("Name").Value, orset.Fields.Item("Code").Value)
                orset.MoveNext()
            Next mjs
            ocombo.Select("-")
            ocombo = oMatrix.Columns.Item("Col_1").Cells.Item(1).Specific
            ocombo.Select("-")
            oMatrix.Columns.Item("Col_9").Editable = False
            Dim ocombobutton As SAPbouiCOM.ButtonCombo = oform.Items.Item("CT").Specific
            ocombobutton.ValidValues.Add("Copy To", "Copy To A/R Invoice")
            oform.Items.Item("CT").Enabled = False
            oMatrix.Columns.Item("Date").Editable = True
            oMatrix.Columns.Item("Col_0").Editable = True
            oMatrix.Columns.Item("Col_1").Editable = True
            oMatrix.Columns.Item("Col_2").Editable = True
            oMatrix.Columns.Item("Col_4").Editable = True
            oMatrix.Columns.Item("Col_5").Editable = True
            oMatrix.Columns.Item("Col_6").Editable = True
            oMatrix.Columns.Item("V_7").Editable = False
            oMatrix.Columns.Item("Col_3").Cells.Item(1).Specific.String = "."
            oMatrix.Columns.Item("Col_7").Cells.Item(1).Specific.String = "."
            oMatrix.Columns.Item("Col_11").Cells.Item(1).Specific.String = "."
            oMatrix.Columns.Item("Col_12").Cells.Item(1).Specific.String = "."
            oMatrix.Columns.Item("Col_8").Cells.Item(1).Specific.String = "."
            oform.Items.Item("Item_1").Specific.String = oCompany_TB_CD.UserName
            oform.Items.Item("31").Specific.String = Company_Name
            oform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            oMatrix.Columns.Item("Date").Width = 90
            oMatrix.Columns.Item("Col_3").Width = 150
            oMatrix.Columns.Item("Col_11").Width = 150
            oMatrix.Columns.Item("Col_12").Width = 150
            oform.DataBrowser.BrowseBy = "Item_13"

            'oMatrix.Columns.Item("V_2").Visible = False
            'oMatrix.Columns.Item("V_1").Visible = False
            'oMatrix.Columns.Item("V_6").Visible = False
            'oMatrix.Columns.Item("V_5").Visible = False
            'oMatrix.Columns.Item("V_4").Visible = False
            'oMatrix.Columns.Item("V_3").Visible = False

            Dim oCFLs As SAPbouiCOM.ChooseFromList
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim empty As New SAPbouiCOM.Conditions
            oCFLs = oform.ChooseFromLists.Item("Bill")
            oCFLs.SetConditions(empty)
            oCons = oCFLs.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFLs.SetConditions(oCons)

            oform.Freeze(False)
            Return True
        Catch ex As Exception
            oform.Freeze(False)
            Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function

    Private Function Chauffer_Driver_binding_AddMode(ByRef oform As SAPbouiCOM.Form) As Boolean

        Try
            Dim Tmp_val As String = oform.BusinessObject.GetNextSerialNumber(NextSerialNo(oCompany_TB_CD, Oapplication_CD, "AE_Cdriver"))
            oform.Items.Item("Item_13").Specific.String = Tmp_val
            Dim ocombo As SAPbouiCOM.ComboBox
            ocombo = oform.Items.Item("Item_16").Specific
            ocombo.Select("Open")
            oform.Items.Item("Item_13").Enabled = False

            oform.Items.Item("Item_19").Specific.String = Format(Now.Date, "yyyyMMdd") 'Now.Date 'Format(Now.Date, "dd MMM yyyy, ddd")
            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_17").Specific
            oMatrix.AddRow()
            oMatrix.Columns.Item("#").Cells.Item(1).Specific.string = "1"
            ocombo = oMatrix.Columns.Item("Col_0").Cells.Item(1).Specific
            ocombo.Select("-")
            ocombo = oMatrix.Columns.Item("Col_1").Cells.Item(1).Specific
            ocombo.Select("-")
            oMatrix.Columns.Item("Col_9").Editable = False
            If oMatrix.RowCount > 1 Then
                If oMatrix.Columns.Item("Date").Cells.Item(oMatrix.RowCount).Specific.String = "" Then
                    oMatrix.DeleteRow(oMatrix.RowCount)
                End If
            End If

            ''Dim ocombobutton As SAPbouiCOM.ButtonCombo = oform.Items.Item("CT").Specific
            ''ocombobutton.ValidValues.Add("Copy To", "Copy To A/R Invoice")
            oform.Items.Item("CT").Enabled = False

            oMatrix.Columns.Item("Date").Editable = True
            oMatrix.Columns.Item("Col_0").Editable = True
            oMatrix.Columns.Item("Col_1").Editable = True
            oMatrix.Columns.Item("Col_2").Editable = True
            oMatrix.Columns.Item("Col_4").Editable = True
            oMatrix.Columns.Item("Col_5").Editable = True
            oMatrix.Columns.Item("Col_6").Editable = True

            oMatrix.CommonSetting.SetCellEditable(1, 1, True)
            oMatrix.CommonSetting.SetCellEditable(1, 2, True)
            oMatrix.CommonSetting.SetCellEditable(1, 3, True)
            oMatrix.CommonSetting.SetCellEditable(1, 4, True)
            oMatrix.CommonSetting.SetCellEditable(1, 6, True)
            oMatrix.CommonSetting.SetCellEditable(1, 7, True)
            oMatrix.CommonSetting.SetCellEditable(1, 8, True)

            oMatrix.Columns.Item("Col_3").Cells.Item(1).Specific.String = "."
            oMatrix.Columns.Item("Col_7").Cells.Item(1).Specific.String = "."
            oMatrix.Columns.Item("Col_11").Cells.Item(1).Specific.String = "."
            oMatrix.Columns.Item("Col_12").Cells.Item(1).Specific.String = "."
            oMatrix.Columns.Item("Col_8").Cells.Item(1).Specific.String = "."
            '' oMatrix.AutoResizeColumns()
            oform.Items.Item("Item_1").Specific.String = oCompany_TB_CD.UserName
            oform.Items.Item("31").Specific.String = Company_Name
            oform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
            oMatrix.Columns.Item("Date").Width = 90
            oform.DataBrowser.BrowseBy = "Item_13"
            Return True
        Catch ex As Exception
            Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function

    Private Function Driver_assign_Binding(ByRef oform As SAPbouiCOM.Form) As Boolean

        Try

            oform.DataSources.UserDataSources.Add("Date1", SAPbouiCOM.BoDataType.dt_DATE)
            oform.DataSources.UserDataSources.Add("Date2", SAPbouiCOM.BoDataType.dt_DATE)
            oform.Items.Item("Item_1").Specific.databind.setbound(True, "", "Date1")
            oform.Items.Item("Item_3").Specific.databind.setbound(True, "", "Date2")
            oform.Items.Item("Item_11").Specific.String = Format(Now.Date, "yyyyMMdd") 'Now.Date
            Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset.DoQuery("SELECT *  FROM [dbo].[@AE_VTYPE]  T0")
            Dim ocombo As SAPbouiCOM.ComboBox = oform.Items.Item("Item_7").Specific
            ocombo.ValidValues.Add("-", "-")
            For mjs As Integer = 1 To orset.RecordCount
                ocombo.ValidValues.Add(orset.Fields.Item("Name").Value, orset.Fields.Item("Code").Value)
                orset.MoveNext()
            Next mjs
            ocombo.Select("-")

            ocombo = oform.Items.Item("Item_8").Specific
            ocombo.Select("-")

            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_12").Specific
            oMatrix.Columns.Item("V_1sjr").Editable = True
            oMatrix.Columns.Item("V_0mjs").Editable = True
            oform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized

            Return True
        Catch ex As Exception
            Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function

    Private Function Driver_assign_Load(ByRef oform As SAPbouiCOM.Form) As Boolean
        Try
            Dim Vehicle_Type, Service_Type, Company, Event_ As String

            If Trim(oform.Items.Item("17").Specific.value) = "-" Then
                Vehicle_Type = "%"
            Else
                Vehicle_Type = Trim(oform.Items.Item("17").Specific.String)
            End If

            If Trim(oform.Items.Item("Item_8").Specific.value) = "-" Then
                Service_Type = "%"
            Else
                Service_Type = Trim(oform.Items.Item("Item_8").Specific.value)
            End If

            If oform.Items.Item("19").Specific.String = "" Then
                Company = "%"
                oform.Items.Item("20").Specific.String = ""
            Else
                Company = oform.Items.Item("19").Specific.value
            End If

            If oform.Items.Item("22").Specific.String = "" Then
                Event_ = "%"
            Else
                Event_ = "%" & oform.Items.Item("22").Specific.value & "%"
            End If

            oform.Freeze(True)


            Dim SQL_String As String
            Dim Sql_String1 As String

            If oform.Items.Item("Item_1").Specific.String = "" And oform.Items.Item("Item_3").Specific.String = "" Then
                SQL_String = "SELECT T0.[DocNum] as 'Object', T1.[DocEntry], T1.[U_AE_Date], " & _
               "T0.[U_AE_Bname] as 'U_AE_Dloc', T0.[U_AE_Order] as 'U_AE_Rem2', T1.[U_AE_Vtype], T1.[U_AE_Gname], T1.[U_AE_Ptime], T1.[U_AE_Ploc],  " & _
               "T1.[U_AE_Stype], T1.[U_AE_Ono], T1.[U_AE_Rem1], T1.[LineId], T1.[U_AE_Vno] , T1.[U_AE_Dname], T1.[U_AE_Vtcod], T1.[U_AE_Dcode], T1.[U_AE_DHP], T1.[U_AE_Tref] FROM [dbo].[@AE_CDRIVER]  T0 inner join  [dbo].[@AE_CDRIVER_R]  T1 " & _
               "on T0.docentry = T1.docentry WHERE T0.[U_AE_Statu] = 'Open' and (T1.U_AE_Date <> '' or T1.U_AE_Date is not null) and  isnull(T1.[U_AE_Vtcod],'') like '" & Vehicle_Type & "' " & _
                                   " and isnull(T1.[U_AE_Stype],'') like '" & Service_Type & "' and isnull(T0.[U_AE_Bcode],'') like '" & Company & "' and isnull(T0.[U_AE_Event],'') like '" & Event_ & "'"

            Else
                ''SQL_String = "SELECT T0.[DocNum] as 'Object', T1.[DocEntry], T1.[U_AE_Date], " & _
                ''   "T0.[U_AE_Bname] as 'U_AE_Dloc', T0.[U_AE_Order] as 'U_AE_Rem2', T1.[U_AE_Vtype], T1.[U_AE_Gname], T1.[U_AE_Ptime], T1.[U_AE_Ploc],  " & _
                ''   "T1.[U_AE_Stype], T1.[U_AE_Ono], T1.[U_AE_Rem1], T1.[LineId], T1.[U_AE_Vno] , T1.[U_AE_Dname], T1.[U_AE_Orno], T1.[U_AE_Vtcod], T1.[U_AE_Dcode] FROM [dbo].[@AE_CDRIVER]  T0 inner join  [dbo].[@AE_CDRIVER_R]  T1 " & _
                ''   "on T0.docentry = T1.docentry WHERE T1.[U_AE_Date] >= '" & System.DateTime.Parse(oform.Items.Item("Item_1").Specific.String, format1, Globalization.DateTimeStyles.None) & "' " & _
                ''   " and T1.[U_AE_Date] <= '" & System.DateTime.Parse(oform.Items.Item("Item_3").Specific.String, format1, Globalization.DateTimeStyles.None) & "' and  T1.[U_AE_Vtcod] like '" & Vehicle_Type & "' " & _
                ''   " and T1.[U_AE_Stype] like '" & Service_Type & "'" ' and (T1.[U_AE_Vno] = '' or T1.[U_AE_Vno] is null) "

                SQL_String = "SELECT T0.[DocNum] as 'Object', T1.[DocEntry], T1.[U_AE_Date], " & _
                                   "T0.[U_AE_Bname] as 'U_AE_Dloc', T0.[U_AE_Order] as 'U_AE_Rem2', T1.[U_AE_Vtype], T1.[U_AE_Gname], T1.[U_AE_Ptime], T1.[U_AE_Ploc],  " & _
                                   "T1.[U_AE_Stype], T1.[U_AE_Ono], T1.[U_AE_Rem1], T1.[LineId], T1.[U_AE_Vno] , T1.[U_AE_Dname], T1.[U_AE_Vtcod], T1.[U_AE_Dcode], T1.[U_AE_DHP], T1.[U_AE_Tref] FROM [dbo].[@AE_CDRIVER]  T0 inner join  [dbo].[@AE_CDRIVER_R]  T1 " & _
                                   "on T0.docentry = T1.docentry WHERE T0.[U_AE_Statu] = 'Open' and T1.[U_AE_Date] >= '" & GateDate(oform.Items.Item("Item_1").Specific.String, oCompany_TB_CD) & "' " & _
                                   " and T1.[U_AE_Date] <= '" & GateDate(oform.Items.Item("Item_3").Specific.String, oCompany_TB_CD) & "' and  isnull(T1.[U_AE_Vtcod],'') like '" & Vehicle_Type & "' " & _
                                   " and isnull(T1.[U_AE_Stype],'') like '" & Service_Type & "' and isnull(T0.[U_AE_Bcode],'') like '" & Company & "' and isnull(T0.[U_AE_Event],'') like '" & Event_ & "'" ' and (T1.[U_AE_Vno] = '' or T1.[U_AE_Vno] is null) "
            End If



            Try
                oform.DataSources.DataTables.Add("@AE_CDRIVER_R")
            Catch ex As Exception

            End Try

            oform.DataSources.DataTables.Item("@AE_CDRIVER_R").ExecuteQuery(SQL_String)
            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_12").Specific
            oMatrix.Clear()
            oform.Items.Item("Item_12").Specific.columns.item("Col_0").databind.bind("@AE_CDRIVER_R", "DocEntry")
            oform.Items.Item("Item_12").Specific.columns.item("Col_1").databind.bind("@AE_CDRIVER_R", "U_AE_Date")
            oform.Items.Item("Item_12").Specific.columns.item("Col_2").databind.bind("@AE_CDRIVER_R", "U_AE_Dloc")
            oform.Items.Item("Item_12").Specific.columns.item("V_0SJR").databind.bind("@AE_CDRIVER_R", "U_AE_Rem2")
            oform.Items.Item("Item_12").Specific.columns.item("Col_3").databind.bind("@AE_CDRIVER_R", "U_AE_Ono")
            oform.Items.Item("Item_12").Specific.columns.item("Col_4").databind.bind("@AE_CDRIVER_R", "U_AE_Vtype")
            oform.Items.Item("Item_12").Specific.columns.item("Col_5").databind.bind("@AE_CDRIVER_R", "U_AE_Gname")
            oform.Items.Item("Item_12").Specific.columns.item("Col_6").databind.bind("@AE_CDRIVER_R", "U_AE_Ptime")
            oform.Items.Item("Item_12").Specific.columns.item("Col_7").databind.bind("@AE_CDRIVER_R", "U_AE_Ploc")
            oform.Items.Item("Item_12").Specific.columns.item("Col_8").databind.bind("@AE_CDRIVER_R", "U_AE_Stype")
            'oform.Items.Item("Item_12").Specific.columns.item("Col_9").databind.bind("@AE_CDRIVER_R", "U_AE_Orno")
            oform.Items.Item("Item_12").Specific.columns.item("Col_10").databind.bind("@AE_CDRIVER_R", "U_AE_Rem1")
            oform.Items.Item("Item_12").Specific.columns.item("V_02SJR").databind.bind("@AE_CDRIVER_R", "LineId")
            oform.Items.Item("Item_12").Specific.columns.item("V_1sjr").databind.bind("@AE_CDRIVER_R", "U_AE_Vno")
            oform.Items.Item("Item_12").Specific.columns.item("V_0mjs").databind.bind("@AE_CDRIVER_R", "U_AE_Dcode")
            oform.Items.Item("Item_12").Specific.columns.item("V_0sjr").databind.bind("@AE_CDRIVER_R", "U_AE_Dname")
            oform.Items.Item("Item_12").Specific.columns.item("V_0vtype").databind.bind("@AE_CDRIVER_R", "U_AE_Vtcod")

            oform.Items.Item("Item_12").Specific.columns.item("V_1mjs").databind.bind("@AE_CDRIVER_R", "U_AE_DHP")
            oform.Items.Item("Item_12").Specific.columns.item("Team").databind.bind("@AE_CDRIVER_R", "U_AE_Tref")


            oform.Items.Item("Item_12").Specific.LoadFromDataSource()
            'oform.Items.Item("Item_12").Specific.AutoResizeColumns()

            oMatrix.Columns.Item("Col_2").Width = 150
            oMatrix.Columns.Item("V_0SJR").Width = 150
            oform.Freeze(False)

            Return True

        Catch ex As Exception
            oform.Freeze(False)
            Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try

    End Function

    Private Function Price_List_Binding(ByRef oform As SAPbouiCOM.Form) As Boolean
        Try
            Dim oMatrix As SAPbouiCOM.Matrix = oform.Items.Item("Item_6").Specific
            Dim ocombo As SAPbouiCOM.ComboBox
            oform.Items.Item("Item_5").Specific.String = Format(Now.Date, "yyyyMMdd") 'Now.Date
            Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset.DoQuery("SELECT *  FROM [dbo].[@AE_VTYPE]  T0")
            oMatrix.AddRow()
            oMatrix.Columns.Item("#").Cells.Item(oMatrix.RowCount).Specific.String = oMatrix.RowCount
            ocombo = oMatrix.Columns.Item("Col_0").Cells.Item(1).Specific
            ocombo.ValidValues.Add("-", "-")

            For mjs As Integer = 1 To orset.RecordCount
                ocombo.ValidValues.Add(orset.Fields.Item("Code").Value, orset.Fields.Item("Name").Value)
                orset.MoveNext()
            Next mjs
            ocombo.Select("-")
            oMatrix.Columns.Item("Col_0").Editable = True
            oMatrix.Columns.Item("Col_1").Editable = True
            oMatrix.Columns.Item("Col_2").Editable = True
            oMatrix.Columns.Item("Col_3").Editable = True
            oform.DataBrowser.BrowseBy = "Item_1"
            Return True
        Catch ex As Exception
            Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function


    ''Private Sub RMG_Report()
    ''    Try
    ''        Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    ''        orset.DoQuery("SELECT T0.[Name] FROM [dbo].[@AE_CRYSTAL]  T0")
    ''        Dim cryRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
    ''        Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
    ''        Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
    ''        Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
    ''        Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
    ''        Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table
    ''        Dim sPath As String
    ''        sPath = IO.Directory.GetParent(Application.StartupPath).ToString

    ''        'MsgBox(System.Windows.Forms.Application.StartupPath.ToString)
    ''        cryRpt.Load(sPath & "\AE_FleetMangement\AE_RP001_OrderChit.rpt")
    ''        'cryRpt.Load("PUT CRYSTAL REPORT PATH HERE\CrystalReport1.rpt")

    ''        Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
    ''        Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
    ''        Dim crParameterValues As New CrystalDecisions.Shared.ParameterValues
    ''        Dim crParameterDiscreteValue As New CrystalDecisions.Shared.ParameterDiscreteValue

    ''        crParameterDiscreteValue.Value = Convert.ToInt32(Docnum)
    ''        crParameterFieldDefinitions = _
    ''    cryRpt.DataDefinition.ParameterFields
    ''        crParameterFieldDefinition = _
    ''    crParameterFieldDefinitions.Item("@OrderChitNo")
    ''        crParameterValues = crParameterFieldDefinition.CurrentValues

    ''        crParameterValues.Clear()
    ''        crParameterValues.Add(crParameterDiscreteValue)
    ''        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
    ''        Dim Server As String = oCompany_TB_CD.Server
    ''        Dim DB As String = oCompany_TB_CD.CompanyDB
    ''        Dim pwd As String = orset.Fields.Item("Name").Value

    ''        With crConnectionInfo
    ''            .ServerName = Server
    ''            .DatabaseName = DB
    ''            .UserID = "sa"
    ''            .Password = pwd
    ''        End With

    ''        CrTables = cryRpt.Database.Tables
    ''        For Each CrTable In CrTables
    ''            crtableLogoninfo = CrTable.LogOnInfo
    ''            crtableLogoninfo.ConnectionInfo = crConnectionInfo
    ''            CrTable.ApplyLogOnInfo(crtableLogoninfo)
    ''        Next


    ''        Dim RptFrm As Viewer
    ''        RptFrm = New Viewer
    ''        RptFrm.CrystalReportViewer1.ReportSource = cryRpt
    ''        RptFrm.CrystalReportViewer1.Refresh()
    ''        RptFrm.Text = "RMG Report for Chaffuer Driver"
    ''        RptFrm.TopMost = True

    ''        RptFrm.Activate()
    ''        RptFrm.ShowDialog()
    ''        System.Threading.Thread.Sleep(100)

    ''    Catch ex As Exception
    ''        Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    ''    End Try





    ''    ''Dim frm1 As Viewer
    ''    ''frm1 = New Viewer(Docnum, adoOleDbConnection)
    ''    ''frm1.TopMost = True
    ''    ''frm1.Activate()
    ''    ''frm1.ShowDialog()

    ''End Sub

   

   

    Public Sub Oapplication_CD_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles Oapplication_CD.RightClickEvent
        Try
            Dim oform As SAPbouiCOM.Form = Oapplication_CD.Forms.ActiveForm

            If oform.UniqueID = "CDB" Or oform.UniqueID = "CDPL" Then


                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                ' Create menu popup MyUserMenu01 and add it to Tools menu
                Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                oCreationPackage = Oapplication_CD.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oMenuItem = Oapplication_CD.Menus.Item("1280") 'Data'
                oMenus = oMenuItem.SubMenus
                'Create sub menu MySubMenu1 and add it to popup MyUserMenu01
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.UniqueID = "3045645"
                oCreationPackage.String = "Delete Row"
                oCreationPackage.Enabled = True
                oMenus.AddEx(oCreationPackage)





            End If
        Catch ex As Exception
            '   MessageBox.Show(ex.Message)
        End Try
    End Sub

    ' ''Private Function CDBooking_ExcelUpload(ByVal sFileName As String, ByRef oForm As SAPbouiCOM.Form) As Boolean

    ' ''    Try
    ' ''        Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("Item_17").Specific
    ' ''        Dim dDate As Date
    ' ''        Dim ocombo As SAPbouiCOM.ComboBox
    ' ''        Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

    ' ''        Dim boolFileExists As Boolean = False
    ' ''        Dim MyConnection As System.Data.OleDb.OleDbConnection = Nothing
    ' ''        boolFileExists = My.Computer.FileSystem.FileExists(sFileName)
    ' ''        Dim workbook As String = "Chauffer Driver Booking"

    ' ''        If boolFileExists = True Then

    ' ''            Dim DtSet As System.Data.DataSet
    ' ''            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
    ' ''            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; " & _
    ' ''            "data source='" & sFileName & " '; " & "Extended Properties=Excel 8.0;")
    ' ''            ' Select the data from Sheet1 of the workbook.
    ' ''            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [" & workbook & "$]", MyConnection)
    ' ''            MyCommand.TableMappings.Add("Table", "CDBooking")
    ' ''            DtSet = New System.Data.DataSet
    ' ''            MyCommand.Fill(DtSet)
    ' ''            oMatrix.Clear()

    ' ''            oForm.Freeze(True)
    ' ''            oMatrix.AddRow(DtSet.Tables.Item(0).Rows.Count - 2)

    ' ''            For i As Integer = 1 To DtSet.Tables.Item(0).Rows.Count - 1
    ' ''                ' Sno
    ' ''                oMatrix.Columns.Item("#").Cells.Item(i).Specific.String = i
    ' ''                ' Date
    ' ''                If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(i - 1).Item(0).ToString) Then
    ' ''                    dDate = DtSet.Tables.Item(0).Rows(i - 1).Item(0).ToString
    ' ''                    oMatrix.Columns.Item("Date").Cells.Item(i).Specific.String = dDate
    ' ''                End If
    ' ''                ' Vehicle Type
    ' ''                If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(i - 1).Item(1).ToString) Then
    ' ''                    ocombo = oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific
    ' ''                    ocombo.Select(Trim(DtSet.Tables.Item(0).Rows(i - 1).Item(1).ToString))
    ' ''                End If
    ' ''                ' Service Type
    ' ''                If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(i - 1).Item(2).ToString) Then
    ' ''                    ocombo = oMatrix.Columns.Item("Col_1").Cells.Item(i).Specific
    ' ''                    ocombo.Select(Trim(DtSet.Tables.Item(0).Rows(i - 1).Item(2).ToString))
    ' ''                End If
    ' ''                ' PickUp Time
    ' ''                If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(i - 1).Item(3).ToString) Then
    ' ''                    dDate = DtSet.Tables.Item(0).Rows(i - 1).Item(3).ToString
    ' ''                    oMatrix.Columns.Item("Col_2").Cells.Item(i).Specific.String = Format(dDate, "HH:mm")
    ' ''                End If
    ' ''                ' Gest Name
    ' ''                If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(i - 1).Item(4).ToString) Then
    ' ''                    orset.DoQuery("select max(cast(isnull(T0.code,0) as integer)) as 'm' from [@AE_EXTENDED] T0")
    ' ''                    Dim code As String = orset.Fields.Item("m").Value + 1
    ' ''                    orset.DoQuery("insert into [@AE_EXTENDED] ([Code],  [Name], [U_AE_Rem] , [U_AE_Dno], [U_AE_Lno], [U_AE_Object], [U_AE_ColID])  values ('" & code & "', '" & code & "', '" & Trim(DtSet.Tables.Item(0).Rows(i - 1).Item(4).ToString) & "' , '" & oForm.Items.Item("Item_13").Specific.String & "' , '" & i & "' , 'CD', 'Col_3')")
    ' ''                    oMatrix.Columns.Item("Col_3").Cells.Item(i).Specific.String = DtSet.Tables.Item(0).Rows(i - 1).Item(4).ToString
    ' ''                Else
    ' ''                    oMatrix.Columns.Item("Col_3").Cells.Item(i).Specific.String = "."
    ' ''                End If
    ' ''                ' Guest HP
    ' ''                oMatrix.Columns.Item("Col_4").Cells.Item(i).Specific.String = DtSet.Tables.Item(0).Rows(i - 1).Item(5).ToString
    ' ''                ' Flight No
    ' ''                oMatrix.Columns.Item("Col_5").Cells.Item(i).Specific.String = DtSet.Tables.Item(0).Rows(i - 1).Item(6).ToString
    ' ''                ' Flight Time
    ' ''                If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(i - 1).Item(7).ToString) Then
    ' ''                    dDate = DtSet.Tables.Item(0).Rows(i - 1).Item(7).ToString
    ' ''                    oMatrix.Columns.Item("Col_6").Cells.Item(i).Specific.String = Format(dDate, "HH:mm")
    ' ''                End If
    ' ''                ' Pickup Location
    ' ''                If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(i - 1).Item(8).ToString) Then
    ' ''                    orset.DoQuery("select max(cast(isnull(T0.code,0) as integer)) as 'm' from [@AE_EXTENDED] T0")
    ' ''                    Dim code As String = orset.Fields.Item("m").Value + 1
    ' ''                    orset.DoQuery("insert into [@AE_EXTENDED] ([Code],  [Name], [U_AE_Rem] , [U_AE_Dno], [U_AE_Lno], [U_AE_Object], [U_AE_ColID])  values ('" & code & "', '" & code & "', '" & Trim(DtSet.Tables.Item(0).Rows(i - 1).Item(8).ToString) & "' , '" & oForm.Items.Item("Item_13").Specific.String & "' , '" & i & "' , 'CD', 'Col_7')")

    ' ''                    oMatrix.Columns.Item("Col_7").Cells.Item(i).Specific.String = DtSet.Tables.Item(0).Rows(i - 1).Item(8).ToString ' Pickup Location
    ' ''                Else
    ' ''                    oMatrix.Columns.Item("Col_7").Cells.Item(i).Specific.String = "."
    ' ''                End If
    ' ''                ' Drop Location
    ' ''                If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(i - 1).Item(9).ToString) Then
    ' ''                    orset.DoQuery("select max(cast(isnull(T0.code,0) as integer)) as 'm' from [@AE_EXTENDED] T0")
    ' ''                    Dim code As String = orset.Fields.Item("m").Value + 1
    ' ''                    orset.DoQuery("insert into [@AE_EXTENDED] ([Code],  [Name], [U_AE_Rem] , [U_AE_Dno], [U_AE_Lno], [U_AE_Object], [U_AE_ColID])  values ('" & code & "', '" & code & "', '" & Trim(DtSet.Tables.Item(0).Rows(i - 1).Item(9).ToString) & "' , '" & oForm.Items.Item("Item_13").Specific.String & "' , '" & i & "' , 'CD', 'Col_8')")
    ' ''                    oMatrix.Columns.Item("Col_8").Cells.Item(i).Specific.String = DtSet.Tables.Item(0).Rows(i - 1).Item(9).ToString ' Drop Location
    ' ''                Else
    ' ''                    oMatrix.Columns.Item("Col_8").Cells.Item(i).Specific.String = "."
    ' ''                End If
    ' ''                'Est End Time
    ' ''                If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(i - 1).Item(10).ToString) Then
    ' ''                    dDate = DtSet.Tables.Item(0).Rows(i - 1).Item(10).ToString
    ' ''                    oMatrix.Columns.Item("Col_9").Cells.Item(i).Specific.String = Format(dDate, "HH:mm")
    ' ''                End If
    ' ''                ' Remark to Driver
    ' ''                If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(i - 1).Item(11).ToString) Then
    ' ''                    orset.DoQuery("select max(cast(isnull(T0.code,0) as integer)) as 'm' from [@AE_EXTENDED] T0")
    ' ''                    Dim code As String = orset.Fields.Item("m").Value + 1
    ' ''                    orset.DoQuery("insert into [@AE_EXTENDED] ([Code],  [Name], [U_AE_Rem] , [U_AE_Dno], [U_AE_Lno], [U_AE_Object], [U_AE_ColID])  values ('" & code & "', '" & code & "', '" & Trim(DtSet.Tables.Item(0).Rows(i - 1).Item(11).ToString) & "' , '" & oForm.Items.Item("Item_13").Specific.String & "' , '" & i & "' , 'CD', 'Col_11')")
    ' ''                    oMatrix.Columns.Item("Col_11").Cells.Item(i).Specific.String = DtSet.Tables.Item(0).Rows(i - 1).Item(11).ToString
    ' ''                Else
    ' ''                    oMatrix.Columns.Item("Col_11").Cells.Item(i).Specific.String = "."
    ' ''                End If
    ' ''                ' Remark to Billing
    ' ''                If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(i - 1).Item(12).ToString) Then
    ' ''                    orset.DoQuery("select max(cast(isnull(T0.code,0) as integer)) as 'm' from [@AE_EXTENDED] T0")
    ' ''                    Dim code As String = orset.Fields.Item("m").Value + 1
    ' ''                    orset.DoQuery("insert into [@AE_EXTENDED] ([Code],  [Name], [U_AE_Rem] , [U_AE_Dno], [U_AE_Lno], [U_AE_Object], [U_AE_ColID])  values ('" & code & "', '" & code & "', '" & Trim(DtSet.Tables.Item(0).Rows(i - 1).Item(12).ToString) & "' , '" & oForm.Items.Item("Item_13").Specific.String & "' , '" & i & "' , 'CD', 'Col_12')")
    ' ''                    oMatrix.Columns.Item("Col_12").Cells.Item(i).Specific.String = DtSet.Tables.Item(0).Rows(i - 1).Item(12).ToString
    ' ''                Else
    ' ''                    oMatrix.Columns.Item("Col_12").Cells.Item(i).Specific.String = "."
    ' ''                End If

    ' ''                Oapplication_CD.StatusBar.SetText("Uploading data from the excel file is in process .....  " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    ' ''            Next i

    ' ''            oForm.Freeze(False)
    ' ''            Oapplication_CD.StatusBar.SetText("Uploading data from the excel file is Completed ..... ! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

    ' ''        End If
    ' ''        Return True
    ' ''    Catch ex As Exception
    ' ''        oForm.Freeze(False)
    ' ''        Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    ' ''        Return False
    ' ''    End Try
    ' ''    Exit Function
    ' ''End Function


    Private Function CDBooking_ExcelUpload(ByVal sFileName As String, ByRef oForm As SAPbouiCOM.Form) As Boolean

        Try
            Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("Item_17").Specific
            Dim dDate As Date
            Dim ocombo As SAPbouiCOM.ComboBox
            Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim boolFileExists As Boolean = False
            Dim MyConnection As System.Data.OleDb.OleDbConnection = Nothing
            boolFileExists = My.Computer.FileSystem.FileExists(sFileName)
            Dim workbook As String = "Chauffer Driver Booking"

            If boolFileExists = True Then

                Dim DtSet As System.Data.DataSet
                Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
                MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; " & _
                "data source='" & sFileName & " '; " & "Extended Properties=Excel 8.0;")
                ' Select the data from Sheet1 of the workbook.
                MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [" & workbook & "$]", MyConnection)
                MyCommand.TableMappings.Add("Table", "CDBooking")
                DtSet = New System.Data.DataSet
                MyCommand.Fill(DtSet)
                oMatrix.Clear()
                oForm.Freeze(True)
                Try
                    oForm.DataSources.DataTables.Add("@AE_CDRIVER_R")
                Catch ex As Exception
                End Try

                Dim oDBDatasource As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item(1)
                Dim oDataTable As SAPbouiCOM.DataTable = oForm.DataSources.DataTables.Item("@AE_CDRIVER_R")

                oDBDatasource.Clear()

                For row As Integer = 0 To DtSet.Tables.Item(0).Rows.Count - 1

                    Dim offset As Integer = oDBDatasource.Size
                    oDBDatasource.InsertRecord(row)

                    If String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(row).Item(0).ToString) Then
                        Exit For
                    End If

                    oDBDatasource.SetValue("LineID", offset, row + 1)
                    ' Date
                    If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(row).Item(0).ToString) Then
                        dDate = DtSet.Tables.Item(0).Rows(row).Item(0).ToString
                        oDBDatasource.SetValue("U_AE_Date", offset, dDate.ToString("yyyyMMdd"))
                    End If

                    ' Vehicle Type
                    oDBDatasource.SetValue("U_AE_Vtype", offset, Trim(DtSet.Tables.Item(0).Rows(row).Item(1).ToString))
                    orset.DoQuery("SELECT T0.[Code] FROM [dbo].[@AE_VTYPE]  T0 WHERE T0.[Name]  = '" & Replace(Trim(DtSet.Tables.Item(0).Rows(row).Item(1).ToString), "'", "''") & "'")
                    oDBDatasource.SetValue("U_AE_Vtcod", offset, orset.Fields.Item("Code").Value)
                    ' Service Type
                    If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(row).Item(2).ToString) Then
                        If DtSet.Tables.Item(0).Rows(row).Item(2).ToString.Trim = "1 way" Then
                            oDBDatasource.SetValue("U_AE_Stype", offset, "One Way")
                        Else
                            oDBDatasource.SetValue("U_AE_Stype", offset, Trim(DtSet.Tables.Item(0).Rows(row).Item(2).ToString))
                        End If
                    End If

                    ' PickUp Time
                    If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(row).Item(3).ToString) Then
                        dDate = DtSet.Tables.Item(0).Rows(row).Item(3).ToString
                        oDBDatasource.SetValue("U_AE_Ptime", offset, Format(dDate, "HH:mm"))
                    End If
                    ' Gest Name
                    If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(row).Item(4).ToString) Then
                        orset.DoQuery("select max(cast(isnull(T0.code,0) as integer)) as 'm' from [@AE_EXTENDED] T0")
                        Dim code As String = orset.Fields.Item("m").Value + 1
                        orset.DoQuery("insert into [@AE_EXTENDED] ([Code],  [Name], [U_AE_Rem] , [U_AE_Dno], [U_AE_Lno], [U_AE_Object], [U_AE_ColID])  values ('" & code & "', '" & code & "', '" & Trim(DtSet.Tables.Item(0).Rows(row).Item(4).ToString) & "' , '" & oForm.Items.Item("Item_13").Specific.String & "' , '" & row + 1 & "' , 'CD', 'Col_3')")
                        oDBDatasource.SetValue("U_AE_Gname", offset, DtSet.Tables.Item(0).Rows(row).Item(4).ToString)
                    Else
                        oDBDatasource.SetValue("U_AE_Gname", offset, ".")
                    End If
                    ' Guest HP
                    oDBDatasource.SetValue("U_AE_GHP", offset, DtSet.Tables.Item(0).Rows(row).Item(6).ToString)
                    ' Flight No
                    oDBDatasource.SetValue("U_AE_Fno", offset, DtSet.Tables.Item(0).Rows(row).Item(6).ToString)
                    ' Flight Time
                    If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(row).Item(7).ToString) Then
                        dDate = DtSet.Tables.Item(0).Rows(row).Item(7).ToString
                        oDBDatasource.SetValue("U_AE_Ftime", offset, Format(dDate, "HH:mm"))
                    End If
                    ' Pickup Location
                    If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(row).Item(8).ToString) Then
                        orset.DoQuery("select max(cast(isnull(T0.code,0) as integer)) as 'm' from [@AE_EXTENDED] T0")
                        Dim code As String = orset.Fields.Item("m").Value + 1
                        orset.DoQuery("insert into [@AE_EXTENDED] ([Code],  [Name], [U_AE_Rem] , [U_AE_Dno], [U_AE_Lno], [U_AE_Object], [U_AE_ColID])  values ('" & code & "', '" & code & "', '" & Trim(DtSet.Tables.Item(0).Rows(row).Item(8).ToString) & "' , '" & oForm.Items.Item("Item_13").Specific.String & "' , '" & row + 1 & "' , 'CD', 'Col_7')")
                        oDBDatasource.SetValue("U_AE_Ploc", offset, DtSet.Tables.Item(0).Rows(row).Item(8).ToString) ' Pickup Location
                    Else
                        oDBDatasource.SetValue("U_AE_Ploc", offset, ".")
                    End If
                    ' Drop Location
                    If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(row).Item(9).ToString) Then
                        orset.DoQuery("select max(cast(isnull(T0.code,0) as integer)) as 'm' from [@AE_EXTENDED] T0")
                        Dim code As String = orset.Fields.Item("m").Value + 1
                        orset.DoQuery("insert into [@AE_EXTENDED] ([Code],  [Name], [U_AE_Rem] , [U_AE_Dno], [U_AE_Lno], [U_AE_Object], [U_AE_ColID])  values ('" & code & "', '" & code & "', '" & Trim(DtSet.Tables.Item(0).Rows(row).Item(9).ToString) & "' , '" & oForm.Items.Item("Item_13").Specific.String & "' , '" & row + 1 & "' , 'CD', 'Col_8')")
                        oDBDatasource.SetValue("U_AE_Dloc", offset, DtSet.Tables.Item(0).Rows(row).Item(9).ToString) ' Drop Location
                    Else
                        oDBDatasource.SetValue("U_AE_Dloc", offset, ".")
                    End If
                    ''Est End Time
                    If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(row).Item(10).ToString) Then
                        dDate = DtSet.Tables.Item(0).Rows(row).Item(10).ToString
                        oDBDatasource.SetValue("U_AE_Dtime", offset, Format(dDate, "HH:mm"))
                    End If
                    ' Remark to Driver
                    If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(row).Item(11).ToString) And DtSet.Tables.Item(0).Rows(row).Item(11).ToString <> "." Then
                        orset.DoQuery("select max(cast(isnull(T0.code,0) as integer)) as 'm' from [@AE_EXTENDED] T0")
                        Dim code As String = orset.Fields.Item("m").Value + 1
                        orset.DoQuery("insert into [@AE_EXTENDED] ([Code],  [Name], [U_AE_Rem] , [U_AE_Dno], [U_AE_Lno], [U_AE_Object], [U_AE_ColID])  values ('" & code & "', '" & code & "', '" & Trim(DtSet.Tables.Item(0).Rows(row).Item(11).ToString) & "' , '" & oForm.Items.Item("Item_13").Specific.String & "' , '" & row + 1 & "' , 'CD', 'Col_11')")
                        oDBDatasource.SetValue("U_AE_Rem1", offset, DtSet.Tables.Item(0).Rows(row).Item(11).ToString)
                    Else
                        oDBDatasource.SetValue("U_AE_Rem1", offset, ".")
                    End If
                    ' Remark to Billing
                    If Not String.IsNullOrEmpty(DtSet.Tables.Item(0).Rows(row).Item(12).ToString) And DtSet.Tables.Item(0).Rows(row).Item(12).ToString <> "." Then
                        orset.DoQuery("select max(cast(isnull(T0.code,0) as integer)) as 'm' from [@AE_EXTENDED] T0")
                        Dim code As String = orset.Fields.Item("m").Value + 1
                        orset.DoQuery("insert into [@AE_EXTENDED] ([Code],  [Name], [U_AE_Rem] , [U_AE_Dno], [U_AE_Lno], [U_AE_Object], [U_AE_ColID])  values ('" & code & "', '" & code & "', '" & Trim(DtSet.Tables.Item(0).Rows(row).Item(12).ToString) & "' , '" & oForm.Items.Item("Item_13").Specific.String & "' , '" & row + 1 & "' , 'CD', 'Col_12')")
                        oDBDatasource.SetValue("U_AE_Rem2", offset, DtSet.Tables.Item(0).Rows(row).Item(12).ToString)
                    Else
                        oDBDatasource.SetValue("U_AE_Rem2", offset, ".")
                    End If

                    Oapplication_CD.StatusBar.SetText("Uploading data from the excel file is in process .....  " & row + 1, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Next row
                oMatrix.LoadFromDataSource()
                oMatrix.AutoResizeColumns()
                oForm.DataSources.DBDataSources.Item(1).Clear()
                oMatrix.AddRow()
                oMatrix.Columns.Item("#").Cells.Item(oMatrix.RowCount).Specific.String = oMatrix.RowCount

                oMatrix.Columns.Item("Col_3").Cells.Item(oMatrix.RowCount).Specific.String = "."
                oMatrix.Columns.Item("Col_7").Cells.Item(oMatrix.RowCount).Specific.String = "."
                oMatrix.Columns.Item("Col_8").Cells.Item(oMatrix.RowCount).Specific.String = "."
                oMatrix.Columns.Item("Col_11").Cells.Item(oMatrix.RowCount).Specific.String = "."
                oMatrix.Columns.Item("Col_12").Cells.Item(oMatrix.RowCount).Specific.String = "."

                oForm.Freeze(False)
                Oapplication_CD.StatusBar.SetText("Uploading data from the excel file is Completed ..... ! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            End If
            Return True
        Catch ex As Exception
            oForm.Freeze(False)
            Oapplication_CD.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
        Exit Function
    End Function

    Private Function showOpenFileDialog() As String

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


    Private Sub ShowFolderBrowser()

        Dim MyProcs() As System.Diagnostics.Process
        FileName = ""
        Dim OpenFile As New OpenFileDialog
        ' Dim stFilePathAndName As String
        Dim orset As SAPbobsCOM.Recordset = oCompany_TB_CD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
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
