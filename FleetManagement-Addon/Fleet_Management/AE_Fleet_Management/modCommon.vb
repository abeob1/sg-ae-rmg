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
                Case 2
                    If Date.TryParseExact(sDate, "MM" & sDatesep & "dd" & sDatesep & "yy", _
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



End Module
