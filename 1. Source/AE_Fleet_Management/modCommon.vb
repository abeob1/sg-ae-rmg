Imports System.Globalization

Module modCommon

    Public Const RTN_SUCCESS As Int16 = 1
    Public Const RTN_ERROR As Int16 = 0
    ' Debug Value Variable Control
    Public Const DEBUG_ON As Int16 = 1
    Public Const DEBUG_OFF As Int16 = 0

    ' Global variables group
    Public p_iDebugMode As Int16 = DEBUG_ON
    Public p_iErrDispMethod As Int16
    Public p_iDeleteDebugLog As Int16
    Public p_oCompDef As CompanyDefault



    Public Structure CompanyDefault

        Public sPath As String
    End Structure

    Public Function GateDate(ByVal sDate As String, ByRef oCompany As SAPbobsCOM.Company) As String

        Dim dateValue As DateTime
        Dim DateString As String = String.Empty
        Dim sSQL As String = String.Empty
        Dim oRs As SAPbobsCOM.Recordset
        Dim sDatesep As String

        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        sSQL = "SELECT DateFormat,DateSep FROM OADM"

        oRs.DoQuery(sSQL)

        If Not oRs.EoF Then
            sDatesep = oRs.Fields.Item("DateSep").Value

            Select Case oRs.Fields.Item("DateFormat").Value
                Case 0
                    If Date.TryParseExact(sDate, "dd" & sDatesep & "MM" & sDatesep & "yy", _
                       New CultureInfo("en-US"), _
                       DateTimeStyles.None, _
                       dateValue) Then
                        DateString = dateValue.ToString("yyyyMMdd")
                    End If
                Case 1
                    If Date.TryParseExact(sDate, "dd" & sDatesep & "MM" & sDatesep & "yyyy", _
                       New CultureInfo("en-US"), _
                       DateTimeStyles.None, _
                       dateValue) Then
                        DateString = dateValue.ToString("yyyyMMdd")
                    End If
                Case 2
                    If Date.TryParseExact(sDate, "MM" & sDatesep & "dd" & sDatesep & "yy", _
                        New CultureInfo("en-US"), _
                        DateTimeStyles.None, _
                        dateValue) Then
                        DateString = dateValue.ToString("yyyyMMdd")
                    End If
                Case 3
                    If Date.TryParseExact(sDate, "MM" & sDatesep & "dd" & sDatesep & "yyyy", _
                        New CultureInfo("en-US"), _
                        DateTimeStyles.None, _
                        dateValue) Then
                        DateString = dateValue.ToString("yyyyMMdd")
                    End If
                Case 4
                    If Date.TryParseExact(sDate, "yyyy" & sDatesep & "MM" & sDatesep & "dd", _
                        New CultureInfo("en-US"), _
                        DateTimeStyles.None, _
                        dateValue) Then
                        DateString = dateValue.ToString("yyyyMMdd")
                    End If
                Case 5
                    If Date.TryParseExact(sDate, "dd" & sDatesep & "MMMM" & sDatesep & "yyyy", _
                        New CultureInfo("en-US"), _
                        DateTimeStyles.None, _
                        dateValue) Then
                        DateString = dateValue.ToString("yyyyMMdd")
                    End If
                Case 6
                    If Date.TryParseExact(sDate, "yy" & sDatesep & "MM" & sDatesep & "dd", _
                        New CultureInfo("en-US"), _
                        DateTimeStyles.None, _
                        dateValue) Then
                        DateString = dateValue.ToString("yyyyMMdd")
                    End If
                Case Else
                    DateString = dateValue.ToString("yyyyMMdd")
            End Select

        End If

        Return DateString

    End Function

    Public Function PostDate(ByVal iday As Integer, ByVal imonth As Integer, ByVal iyear As Integer, ByRef oApplication As SAPbouiCOM.Application) As Date
        Try
            Dim dtmpdate As String = iyear & "/" & Format(imonth, "00") & "/" & Format(iday, "00")
            Dim dreturndate As DateTime = Convert.ToDateTime(dtmpdate)
            PostDate = dreturndate

        Catch ex As Exception
            oApplication.MessageBox(ex.Message)
        End Try

    End Function

    Public Function AppConfigInfo(ByRef oCompDef As CompanyDefault, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Try
            oCompDef.sPath = IO.Directory.GetCurrentDirectory
            AppConfigInfo = RTN_SUCCESS
        Catch ex As Exception
            AppConfigInfo = RTN_ERROR
        End Try
    End Function

    Public Function NavigationValidation_SelfDriver(ByVal oForm As SAPbouiCOM.Form, ByVal oCompany As SAPbobsCOM.Company, ByVal oApplication As SAPbouiCOM.Application) As Boolean

        Try

            Dim opt As SAPbouiCOM.OptionBtn = oForm.Items.Item("195").Specific
            Dim opt1 As SAPbouiCOM.OptionBtn = oForm.Items.Item("196").Specific
            Dim oopt As SAPbouiCOM.OptionBtn = oForm.Items.Item("Item_104").Specific
            Dim oopt1 As SAPbouiCOM.OptionBtn = oForm.Items.Item("Item_105").Specific
            Dim oopt3 As SAPbouiCOM.OptionBtn = oForm.Items.Item("Item_106").Specific
            Dim orset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim ocombo As SAPbouiCOM.ComboBox
            Dim sAttention As String
            Dim bfalg As Boolean = False
            Dim oMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("1000001").Specific
            oForm.Freeze(True)

         

            Dim ss = "SELECT T0.[FirstName] + ' ' + T0.[LastName] as 'Name' FROM OCPR T0 WHERE T0.[CardCode]  = '" & oForm.Items.Item("Item_2").Specific.string & "'"
            orset.DoQuery("SELECT isnull(T0.[FirstName],'') + ' ' + isnull(T0.[LastName],'') as 'Name' FROM OCPR T0 WHERE T0.[CardCode]  = '" & oForm.Items.Item("Item_2").Specific.string & "'")
            ' " & _                                            "and T0.[FirstName] + ' ' + T0.[LastName] <> '" & oform.Items.Item("235").Specific.value.ToString.Trim & "'")
            sAttention = oForm.Items.Item("235").Specific.value.ToString.Trim

            ocombo = oForm.Items.Item("235").Specific

            oMatrix.Columns.Item("V_2").Visible = False

            Try
                For mjs As Integer = ocombo.ValidValues.Count To 1 Step -1
                    ocombo.ValidValues.Remove(mjs - 1, SAPbouiCOM.BoSearchKey.psk_Index)
                Next mjs
            Catch ex As Exception

            End Try

            Try
                For mjs As Integer = 1 To orset.RecordCount
                    If Not String.IsNullOrEmpty(orset.Fields.Item("Name").Value) Then
                        ocombo.ValidValues.Add(orset.Fields.Item("Name").Value, "")
                        If sAttention = orset.Fields.Item("Name").Value Then
                            bfalg = True
                        End If
                        orset.MoveNext()
                    End If

                Next mjs

            Catch ex As Exception
            End Try

            Try
                If bfalg = True Then
                    ocombo.Select(sAttention, SAPbouiCOM.BoSearchKey.psk_ByValue)
                Else
                    ocombo.Select(0)
                End If

            Catch ex As Exception
            End Try

            oForm.Items.Item("201").Enabled = True
            '------------------- PAI 
            If oForm.Items.Item("226").Specific.value.ToString.Trim = "Yes" Then
                oForm.Items.Item("Item_115").Enabled = True
            Else
                oForm.Items.Item("Item_115").Enabled = False
            End If

            '-------------------- CDW
            If oForm.Items.Item("228").Specific.selected.value.ToString.Trim = "Yes" Then
                oForm.Items.Item("Item_123").Enabled = True
            Else
                oForm.Items.Item("Item_123").Enabled = False
            End If

            If opt.Selected = True Then
                oForm.Items.Item("210").Enabled = False
                'oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                If Trim(oForm.Items.Item("Item_18").Specific.value) = "Open" Then
                    oForm.Items.Item("Item_18").Enabled = True
                Else
                    oForm.Items.Item("Item_18").Enabled = False
                End If
            Else
                If Trim(oForm.Items.Item("Item_18").Specific.value) = "Billing" Then
                    ''   oForm.Items.Item("Item_2").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    oForm.Items.Item("Item_18").Enabled = False
                    oForm.Items.Item("210").Enabled = True
                    ''oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                ElseIf Trim(oForm.Items.Item("Item_18").Specific.value) <> "Open" Then
                    oForm.Items.Item("Item_18").Enabled = False
                    oForm.Items.Item("210").Enabled = False
                ElseIf Trim(oForm.Items.Item("Item_18").Specific.value) = "Open" Then

                    '' oForm.Items.Item("Item_2").Specific.active = True
                    ''  oForm.Items.Item("Item_2").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    oForm.Items.Item("210").Enabled = False
                    oForm.Items.Item("Item_18").Enabled = True
                    'oform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                End If
            End If

            If oopt.Selected = True Then
                oForm.Items.Item("Item_108").Specific.caption = "Daily Rates"
                oForm.Items.Item("Item_110").Specific.caption = "Number of Days"
                oForm.Items.Item("Item_114").Specific.caption = "PAI Per Day"
                oForm.Items.Item("Item_122").Specific.caption = "CDW Per Day"
            ElseIf oopt1.Selected = True Then
                oForm.Items.Item("Item_108").Specific.caption = "Weekly Rates"
                oForm.Items.Item("Item_110").Specific.caption = "Number of Days"
                oForm.Items.Item("Item_114").Specific.caption = "PAI Per Day"
                oForm.Items.Item("Item_122").Specific.caption = "CDW Per Day"
            ElseIf oopt3.Selected = True Then
                oForm.Items.Item("Item_108").Specific.caption = "Monthly Rates"
                oForm.Items.Item("Item_110").Specific.caption = "Number of Months"
                oForm.Items.Item("Item_114").Specific.caption = "PAI Per Month"
                oForm.Items.Item("Item_122").Specific.caption = "CDW Per Month"
            End If

            ocombo = oForm.Items.Item("237").Specific
            If ocombo.ValidValues.Count = 0 Then
                For imjs As Integer = 1 To 31
                    ocombo.ValidValues.Add(imjs, imjs)
                Next
                ocombo.ValidValues.Add("", "")
            End If
            ' oform.PaneLevel = 3
            If p_bSDBooking = False Then
                oForm.Items.Item("Item_22").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            End If

            oForm.Items.Item("Item_14").Enabled = False
            oForm.Items.Item("233").Enabled = True
            oForm.Items.Item("237").Enabled = False
            oForm.Freeze(False)
            Return True
        Catch ex As Exception
            oApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, False)
            oForm.Freeze(False)
            Return False
        End Try

    End Function

    Public Function CB_Navigation(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim oMAtrix As SAPbouiCOM.Matrix = oForm.Items.Item("Item_17").Specific
            oForm.Freeze(True)
            oForm.Items.Item("1000003").Enabled = False
            If Trim(oForm.Items.Item("Item_16").Specific.value) = "Open" Then
                oForm.Items.Item("Item_16").Enabled = True
                oForm.Items.Item("Item_15").Enabled = True

                oForm.Items.Item("Item_4").Enabled = True
                oForm.Items.Item("42").Enabled = True
                'oForm.Items.Item("30").Enabled = True
                oForm.Items.Item("Item_11").Enabled = True
                oForm.DataSources.DBDataSources.Item(1).Clear()
                If oMAtrix.RowCount > 0 Then
                    If Not String.IsNullOrEmpty(oMAtrix.Columns.Item("Date").Cells.Item(oMAtrix.RowCount).Specific.String) Then
                        oMAtrix.AddRow()
                        oMAtrix.CommonSetting.SetRowEditable(oMAtrix.RowCount, True)
                        oMAtrix.Columns.Item("#").Cells.Item(oMAtrix.RowCount).Specific.String = oMAtrix.RowCount
                        oMAtrix.Columns.Item("Col_3").Cells.Item(oMAtrix.RowCount).Specific.String = "."
                        oMAtrix.Columns.Item("Col_7").Cells.Item(oMAtrix.RowCount).Specific.String = "."
                        oMAtrix.Columns.Item("Col_8").Cells.Item(oMAtrix.RowCount).Specific.String = "."
                        oMAtrix.Columns.Item("Col_11").Cells.Item(oMAtrix.RowCount).Specific.String = "."
                        oMAtrix.Columns.Item("Col_12").Cells.Item(oMAtrix.RowCount).Specific.String = "."
                    End If
                End If

                oForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                For mjs As Integer = 1 To oMAtrix.RowCount
                    If oMAtrix.Columns.Item("Col_13").Cells.Item(mjs).Specific.value <> "" Then
                        oMAtrix.CommonSetting.SetRowEditable(mjs, False)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 13, False)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 14, False)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 16, False)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 20, False)
                        oMAtrix.Columns.Item("Col_11").Editable = False

                    Else
                        oMAtrix.CommonSetting.SetRowEditable(mjs, True)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 5, False)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 9, False)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 10, False)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 13, False)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 14, False)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 11, False)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 12, False)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 15, False)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 16, False)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 17, False)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 18, False)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 19, False)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 20, False)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 21, False)
                        oMAtrix.CommonSetting.SetCellEditable(mjs, 22, False)
                    End If
                    oForm.Items.Item("CT").Enabled = False

                Next

                oMAtrix.Columns.Item("V_7").Editable = False

            ElseIf Trim(oForm.Items.Item("Item_16").Specific.value) = "Billing" Then
                oForm.Items.Item("48").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                oForm.Items.Item("Item_4").Enabled = True
                oForm.Items.Item("42").Enabled = True
                'oForm.Items.Item("30").Enabled = True
                oForm.Items.Item("Item_11").Enabled = True
                oForm.Items.Item("Item_15").Enabled = False
                oForm.Items.Item("Item_16").Enabled = True
                oMAtrix.Columns.Item("V_7").Editable = False
                For mjs As Integer = 1 To oMAtrix.RowCount
                    oMAtrix.CommonSetting.SetRowEditable(mjs, False)
                Next
                oForm.Items.Item("CT").Enabled = True

            ElseIf Trim(oForm.Items.Item("Item_16").Specific.value) = "Closed" Then

                '' oForm.Items.Item("Item_13").Specific.active = True
                oForm.Items.Item("Item_4").Enabled = False
                oForm.Items.Item("Item_5").Enabled = False
                oForm.Items.Item("42").Enabled = False
                'oForm.Items.Item("30").Enabled = False
                oForm.Items.Item("Item_19").Enabled = False
                oForm.Items.Item("Item_11").Enabled = False
                oForm.Items.Item("Item_16").Enabled = False
                oForm.Items.Item("Item_1").Enabled = False
                oForm.Items.Item("Item_15").Enabled = False

                'oForm.Items.Item("Item_17").Enabled = False
                oForm.Items.Item("CT").Enabled = False
                oMAtrix.Columns.Item("V_7").Editable = True
                oMAtrix.CommonSetting.SetCellEditable(1, 1, True)
                oMAtrix.Columns.Item("Date").Editable = False
                oMAtrix.Columns.Item("Col_0").Editable = False
                oMAtrix.Columns.Item("Col_1").Editable = False
                oMAtrix.Columns.Item("Col_2").Editable = False
                oMAtrix.Columns.Item("Col_3").Editable = False
                oMAtrix.Columns.Item("Col_4").Editable = False
                oMAtrix.Columns.Item("Col_5").Editable = False
                oMAtrix.Columns.Item("Col_6").Editable = False
                oMAtrix.Columns.Item("Col_7").Editable = False
                oMAtrix.Columns.Item("Col_8").Editable = False
                oMAtrix.Columns.Item("Col_9").Editable = False
                oMAtrix.Columns.Item("Col_10").Editable = False
                oMAtrix.Columns.Item("Col_11").Editable = False
                oMAtrix.Columns.Item("Col_12").Editable = False
                oMAtrix.Columns.Item("Col_13").Editable = False
                oMAtrix.Columns.Item("Col_14").Editable = False
                oMAtrix.Columns.Item("Col_15").Editable = False
                oMAtrix.Columns.Item("Col_16").Editable = False
                oMAtrix.Columns.Item("Col_17").Editable = False
                oMAtrix.Columns.Item("V_0").Editable = False
                oForm.Items.Item("CT").Enabled = False
            End If
            oMAtrix.Columns.Item("#").Editable = False
            oForm.Freeze(False)
            Return True
        Catch ex As Exception
            oForm.Freeze(False)
            Return False
        End Try
    End Function


End Module
