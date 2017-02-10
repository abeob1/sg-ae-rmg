Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Data.Common




Module modDBHelper

   

    Public Function AppConfigInfo(ByRef oCompDef As CompanyDefault, ByRef sErrDesc As String) As Long

        ' ***********************************************************************************
        '   Function   :    AppConfigInfo()
        '   Purpose    :    Get the Key values in the App.config file and stored in the Structure Named "CompanyDefault"
        '
        '   Parameters :    ByRef oCompDef As CompanyDefault
        '                       oCompDef = set the Structure Object
        '                       ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    JOHN
        '   Date       :    07/05/2014 
        '   Change     :   
        '                   
        ' ***********************************************************************************

        Dim sFuncName As String = String.Empty
        Try

            sFuncName = "AppConfigInfo()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Console.WriteLine("Starting Function " & sFuncName)

            '------------------ Set structure Fields to empty 
            oCompDef.sDBName = String.Empty
            oCompDef.sServer = String.Empty
            oCompDef.sDBUser = String.Empty
            oCompDef.sDBPwd = String.Empty
            oCompDef.sDebug = String.Empty

            oCompDef.sEmailFrom = String.Empty
            oCompDef.sEmailTo_CD = String.Empty
            oCompDef.sEmailTo_SD = String.Empty
            oCompDef.sEmailT0_SDContractor = String.Empty
            oCompDef.sEmailT0_SM = String.Empty

            oCompDef.sEmailSubject = String.Empty
            oCompDef.sEmailBody = String.Empty

            oCompDef.sSMTPServer = String.Empty
            oCompDef.sSMTPPort = String.Empty

            oCompDef.sCDTime01 = String.Empty
            oCompDef.sCDTime02 = String.Empty
            oCompDef.sCDTime03 = String.Empty
            oCompDef.sSDTime01 = String.Empty
            oCompDef.sSMTime01 = String.Empty

            oCompDef.sPath = String.Empty

            '------------------ Passing values from App.config file to structure fields 
            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Server")) Then
                oCompDef.sServer = ConfigurationManager.AppSettings("Server")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBName")) Then
                oCompDef.sDBName = ConfigurationManager.AppSettings("DBName")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBUser")) Then
                oCompDef.sDBUser = ConfigurationManager.AppSettings("DBUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("DBPwd")) Then
                oCompDef.sDBPwd = ConfigurationManager.AppSettings("DBPwd")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Debug")) Then
                oCompDef.sDebug = ConfigurationManager.AppSettings("Debug")
            End If


            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailFrom")) Then
                oCompDef.sEmailFrom = ConfigurationManager.AppSettings("EmailFrom")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPUser")) Then
                oCompDef.sSMTPUser = ConfigurationManager.AppSettings("SMTPUser")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPPassword")) Then
                oCompDef.sSMTPPwd = ConfigurationManager.AppSettings("SMTPPassword")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailTo_CD")) Then
                oCompDef.sEmailTo_CD = ConfigurationManager.AppSettings("EmailTo_CD")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailTo_SD")) Then
                oCompDef.sEmailTo_SD = ConfigurationManager.AppSettings("EmailTo_SD")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailTo_SDContract")) Then
                oCompDef.sEmailT0_SDContractor = ConfigurationManager.AppSettings("EmailTo_SDContract")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailTo_SM")) Then
                oCompDef.sEmailT0_SM = ConfigurationManager.AppSettings("EmailTo_SM")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailSubject")) Then
                oCompDef.sEmailSubject = ConfigurationManager.AppSettings("EmailSubject")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("EmailBody")) Then
                oCompDef.sEmailBody = ConfigurationManager.AppSettings("EmailBody")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPServer")) Then
                oCompDef.sSMTPServer = ConfigurationManager.AppSettings("SMTPServer")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTPPort")) Then
                oCompDef.sSMTPPort = ConfigurationManager.AppSettings("SMTPPort")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CDTime01")) Then
                oCompDef.sCDTime01 = ConfigurationManager.AppSettings("CDTime01")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CDTime02")) Then
                oCompDef.sCDTime02 = ConfigurationManager.AppSettings("CDTime02")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("CDTime03")) Then
                oCompDef.sCDTime03 = ConfigurationManager.AppSettings("CDTime03")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SDTime01")) Then
                oCompDef.sSDTime01 = ConfigurationManager.AppSettings("SDTime01")
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("Path")) Then
                oCompDef.sPath = IO.Directory.GetCurrentDirectory
            End If

            If Not String.IsNullOrEmpty(ConfigurationManager.AppSettings("SMTime01")) Then
                oCompDef.sSMTime01 = ConfigurationManager.AppSettings("SMTime01")
            End If


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Console.WriteLine("Completed with SUCCESS " & sFuncName)
            AppConfigInfo = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with Error " & sFuncName)
            AppConfigInfo = RTN_ERROR
        End Try
    End Function

    Public Function ExecuteSQLQuery(ByRef oreturnp As ReturnParameters, ByVal sQuery As String) As ReturnParameters

        ' ***********************************************************************************
        '   Function   :    ExecuteSQLQuery()
        '   Purpose    :    This function is handles the query/store procedure execution and return the datas in the dataset, bflag, recordcount
        '
        '   Parameters :    ByRef oreturnp As ReturnParameters
        '                       oreturnp = set the Structure Object to get the bflag, oDateset, iRecordcount
        '                   ByVal sQuery As String
        '                       sQuery   = Passing the store procedure for execution
        '   Return     :    oreturnp.bflag        - True / False (if it returns true the query/store procedure executed without issues)
        '                   oreturnp.oDateset     - Dataset with executed data
        '                   oreturnp.iRecordcount - Recounts of executed query / store procedure
        '                   oreturnp.sFpath       - Empty

        '   Author     :    JOHN
        '   Date       :    07/05/2014 
        '   Change     :   
        '                   
        ' ***********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim sConstr As String = "Data Source=" & p_oCompDef.sServer & ";Initial Catalog=" & p_oCompDef.sDBName & ";User ID=" & p_oCompDef.sDBUser & "; Password=" & p_oCompDef.sDBPwd
        Dim oCon As New SqlConnection(sConstr)
        Dim oCmd As New SqlCommand
        Dim oDs As New DataSet

        Try
            sFuncName = "ExecuteQuery() " & sQuery
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            Console.WriteLine("Starting Function... " & sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(" SQL " & sQuery, sFuncName)
            oCon.ConnectionString = sConstr

            oCon.Open()
            oCmd.CommandType = CommandType.Text
            oCmd.CommandText = sQuery
            oCmd.Connection = oCon
            oCmd.CommandTimeout = 0
            Dim da As New SqlDataAdapter(oCmd)
            da.Fill(oDs)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)
            Console.WriteLine("Function Completed Successfully. " & sFuncName)

            oreturnp.oDateset = oDs
            oreturnp.bflag = True
            oreturnp.iRecordcount = oDs.Tables(0).Rows.Count
            oreturnp.sFpath = ""

        Catch ex As Exception
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Starting Function... " & sFuncName)
            oreturnp.bflag = False
        Finally

            oCon.Dispose()
            oCmd.Dispose()
        End Try

    End Function

End Module
