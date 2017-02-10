Module modMain
    Public Structure CompanyDefault
        Public sDBName As String
        Public sServer As String
        Public sDBUser As String
        Public sDBPwd As String
        Public sDebug As String

        Public sEmailFrom As String
        Public sSMTPUser As String
        Public sSMTPPwd As String
        Public sEmailTo_CD As String
        Public sEmailTo_SD As String
        Public sEmailT0_SDContractor As String
        Public sEmailT0_SM As String

        Public sEmailSubject As String
        Public sEmailBody As String

        Public sSMTPServer As String
        Public sSMTPPort As String

        Public sCDTime01 As String
        Public sCDTime02 As String
        Public sCDTime03 As String
        Public sSDTime01 As String
        Public sSMTime01 As String

        Public sPath As String
    End Structure

    Public Structure ReturnParameters

        Public oDateset As DataSet
        Public iRecordcount As Integer
        Public bflag As Boolean
        Public sFpath As String
    End Structure


    ' Return Value Variable Control
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
    Public p_oReturnPArameters As ReturnParameters

    Public p_oDtSuccess As DataTable
    Public p_oDtError As DataTable
    Public p_SyncDateTime As String




    Sub Main()

        Dim sFuncName As String = String.Empty
        Dim sErrDesc As String = String.Empty

        Try
            p_oCompDef.sPath = IO.Directory.GetCurrentDirectory

            sFuncName = "Main"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            Console.Title = "RMG Notification Module"
            Console.WriteLine("Starting Function " & sFuncName)

            If AppConfigInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If CDNotification(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If SDNotification(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If SMNotification(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Console.WriteLine("Completed with SUCCESS " & sFuncName)

        Catch ex As Exception
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with Error " & sFuncName)
            End
        End Try

    End Sub

End Module
