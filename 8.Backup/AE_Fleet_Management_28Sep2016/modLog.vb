Option Explicit On
Imports System.IO

Module modLog

    Private Const MAXFILESIZE_IN_MB As Int16 = 5
    Private Const LOG_FILE_ERROR As String = "ErrorLog"
    Private Const LOG_FILE_ERROR_ARCH As String = "ErrorLog_"
    Private Const LOG_FILE_DEBUG As String = "DebugLog"
    Private Const LOG_FILE_DEBUG_ARCH As String = "DebugLog_"
    Private Const FILE_SIZE_CHECK_ENABLE As Int16 = 1
    Private Const FILE_SIZE_CHECK_DISABLE As Int16 = 0




    Public Function WriteToLogFile(ByVal strErrText As String, ByVal strSourceName As String, Optional ByVal intCheckFileForDelete As Int16 = 1) As Long

        Dim oStreamWriter As StreamWriter = Nothing
        Dim strFileName As String = String.Empty
        Dim strArchFileName As String = String.Empty
        Dim strTempString As String = String.Empty
        Dim lngFileSizeInMB As Double

        Try
            strTempString = Space(IIf(Len(strSourceName) > 30, 0, 30 - Len(strSourceName)))
            strSourceName = strTempString & strSourceName
            strErrText = "[" & Format(Now, "MM/dd/yyyy HH:mm:ss") & "]" & "[" & strSourceName & "] " & strErrText

            strFileName = p_oCompDef.sPath & "\" & LOG_FILE_ERROR & ".log"
            strArchFileName = p_oCompDef.sPath & "\" & LOG_FILE_ERROR_ARCH & Format(Now(), "yyMMddHHMMss") & ".log"

            If intCheckFileForDelete = FILE_SIZE_CHECK_ENABLE Then
                If File.Exists(strFileName) Then
                    lngFileSizeInMB = (FileLen(strFileName) / 1024) / 1024
                    If lngFileSizeInMB >= MAXFILESIZE_IN_MB Then
                        If p_iDeleteDebugLog = 1 Then
                            For Each sFileName As String In Directory.GetFiles(System.IO.Directory.GetCurrentDirectory(), LOG_FILE_DEBUG_ARCH & "*")
                                File.Delete(sFileName)
                            Next
                        End If
                        File.Move(strFileName, strArchFileName)
                    End If
                End If
            End If
            oStreamWriter = File.AppendText(strFileName)
            oStreamWriter.WriteLine(strErrText)
            WriteToLogFile = RTN_SUCCESS
        Catch exc As Exception
            WriteToLogFile = RTN_ERROR
        Finally
            If Not IsNothing(oStreamWriter) Then
                oStreamWriter.Flush()
                oStreamWriter.Close()
                oStreamWriter = Nothing
            End If
        End Try

    End Function

    Public Function WriteToLogFile_Debug(ByVal strErrText As String, ByVal strSourceName As String, Optional ByVal intCheckFileForDelete As Int16 = 1) As Long
        Dim oStreamWriter As StreamWriter = Nothing
        Dim strFileName As String = String.Empty
        Dim strArchFileName As String = String.Empty
        Dim strTempString As String = String.Empty
        Dim lngFileSizeInMB As Double
        Dim iFileCount As Integer = 0

        Try
            strTempString = Space(IIf(Len(strSourceName) > 30, 0, 30 - Len(strSourceName)))
            strSourceName = strTempString & strSourceName
            strErrText = "[" & Format(Now, "MM/dd/yyyy HH:mm:ss") & "]" & "[" & strSourceName & "] " & strErrText

            strFileName = p_oCompDef.sPath & "\" & LOG_FILE_DEBUG & ".log"
            strArchFileName = p_oCompDef.sPath & "\" & LOG_FILE_DEBUG_ARCH & Format(Now(), "yyMMddHHMMss") & ".log"

            If intCheckFileForDelete = FILE_SIZE_CHECK_ENABLE Then
                If File.Exists(strFileName) Then
                    lngFileSizeInMB = (FileLen(strFileName) / 1024) / 1024
                    If lngFileSizeInMB >= MAXFILESIZE_IN_MB Then
                        If p_iDeleteDebugLog = 1 Then
                            For Each sFileName As String In Directory.GetFiles(System.IO.Directory.GetCurrentDirectory(), LOG_FILE_DEBUG_ARCH & "*")
                                File.Delete(sFileName)
                            Next
                        End If
                        File.Move(strFileName, strArchFileName)
                    End If
                End If
            End If
            oStreamWriter = File.AppendText(strFileName)
            oStreamWriter.WriteLine(strErrText)
            WriteToLogFile_Debug = RTN_SUCCESS
        Catch exc As Exception
            WriteToLogFile_Debug = RTN_ERROR
        Finally
            If Not IsNothing(oStreamWriter) Then
                oStreamWriter.Flush()
                oStreamWriter.Close()
                oStreamWriter = Nothing
            End If
        End Try

    End Function


End Module

