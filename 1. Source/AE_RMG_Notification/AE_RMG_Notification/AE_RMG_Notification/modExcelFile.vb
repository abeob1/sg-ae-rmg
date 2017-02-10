Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports excel = Microsoft.Office.Interop.Excel
Imports System.Net.Mail
Imports System.Net.Mime




Module modExcelFile


    Public Function UploadToExcel(ByVal sFileName As String, ByRef oDataset As DataSet, ByVal HeaderName As String) As ReturnParameters

        ' ***********************************************************************************
        '   Function   :    UploadToExcel()
        '   Purpose    :    This function is handles the data upload from the dataset to excel file
        '   Parameters :    ByVal sFileName As String
        '                       sFileName = Passing file name
        '                   ByRef oDataset As DataSet
        '                       oDataset   = Passing dataset
        '                   ByVal HeaderName As String
        '                       HeaderName = Passing header name
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
        Dim oAppXL As excel.Application = Nothing
        Dim oWbXl As excel.Workbook = Nothing
        Dim oShXL As excel.Worksheet = Nothing
        Dim oRaXL As excel.Range = Nothing
        Dim oDataTable As System.Data.DataTable = oDataset.Tables(0)
        Dim oDTColumn As System.Data.DataColumn
        Dim oDTRow As System.Data.DataRow
        Dim iColIndex As Integer = 0
        Dim iRowIndex As Integer = 3

        Try
            sFuncName = "UploadToExcel()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            Console.WriteLine("Starting Function... " & sFuncName)

            ' Start Excel and get Application object.
            oAppXL = CreateObject("Excel.Application")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Excel Object Created .", sFuncName)
            ' Add a new workbook.
            oWbXl = oAppXL.Workbooks.Add
            oShXL = oWbXl.ActiveSheet
            ' Add table headers going cell by cell.
            oShXL.Range(oShXL.Cells(1, 1), oShXL.Cells(1, 10)).Merge()
            oShXL.Cells(1, 1) = HeaderName
            oShXL.Cells(1, 1).Font.Bold = True

            ' Adding column names in excel 
            For Each oDTColumn In oDataTable.Columns
                iColIndex += 1
                oShXL.Cells(3, iColIndex) = oDTColumn.ColumnName
                oShXL.Cells(3, iColIndex).Font.Bold = True
                ' oShXL.Range(oShXL.Cells(3, iColIndex)).AutoFit()
                oShXL.Cells.VerticalAlignment = excel.XlVAlign.xlVAlignCenter
                oShXL.Cells.HorizontalAlignment = excel.XlVAlign.xlVAlignCenter
            Next
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Columns Created .", sFuncName)
            ' Uploading datas in the excel rows cell by cell
            For Each oDTRow In oDataTable.Rows
                iRowIndex += 1
                iColIndex = 0

                For Each oDTColumn In oDataTable.Columns
                    iColIndex += 1
                    oShXL.Cells(iRowIndex + 1, iColIndex) = oDTRow(oDTColumn.ColumnName)
                Next
            Next
            oShXL.Columns.AutoFit()
            oShXL.Name = sFileName
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rows Created .", sFuncName)
            oShXL.SaveAs(p_oCompDef.sPath & sFileName & " " & Format(Now.Date, "dd,MM,yyyy,ddd") & ".xlsx") '(p_oCompDef.sPath & sFileName & ".xlsx")

            p_oReturnPArameters.bflag = True
            p_oReturnPArameters.sFpath = p_oCompDef.sPath & sFileName & " " & Format(Now.Date, "dd,MM,yyyy,ddd") & ".xlsx"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Console.WriteLine("Completed with SUCCESS " & sFuncName)

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with Error " & sFuncName)
            p_oReturnPArameters.bflag = False
        Finally
            oWbXl.Close()
            oWbXl = Nothing
            oShXL = Nothing
            oRaXL = Nothing
            oAppXL.Quit()
            oAppXL = Nothing
            oDataTable = Nothing
            oDTColumn = Nothing
            oDTRow = Nothing
            oDataset = Nothing

        End Try

    End Function

    Public Function SendEmailNotification(ByVal sfileName As String, ByVal sSenderEmail As String, ByVal sErrDesc As String) As Long

        ' ***********************************************************************************
        '   Function   :    SendEmailNotification()
        '   Purpose    :    This function is handles - Sending notification mails
        '   Parameters :    ByVal sFileName As String
        '                       sFileName = Passing file name
        '                   ByRef sErrDesc As String
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    JOHN
        '   Date       :    07/05/2014 
        '   Change     :   
        '                   
        ' ***********************************************************************************

        Dim sFuncName As String = String.Empty
        Dim oSmtpServer As New SmtpClient()
        Dim oMail As New MailMessage
        Dim sBody As String = String.Empty

        Try
            sFuncName = "SendEmailNotification()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function...", sFuncName)
            Console.WriteLine("Starting Function... " & sFuncName)
            '------------  Date format
            p_SyncDateTime = Format(Now, "dddd") & ", " & Format(Now, "MMM") & " " & Format(Now, "dd") & ", " & Format(Now, "yyyy") & " " & Format(Now, "HH:mm:ss")
            '--------- Message Content in HTML tags
            sBody = sBody & "<div align=left style='font-size:10.0pt;font-family:Arial'>"
            sBody = sBody & " Dear Valued Customer,<br /><br />"
            sBody = sBody & p_SyncDateTime & " <br /><br />"
            sBody = sBody & " Please find the attached alert notification document.<br /><br />"
            sBody = sBody & "<br/> Note: This email message is computer generated and it will be used internal purpose usage only.<div/>"

            oSmtpServer.Credentials = New Net.NetworkCredential(p_oCompDef.sSMTPUser, p_oCompDef.sSMTPPwd)
            oSmtpServer.Port = p_oCompDef.sSMTPPort
            oSmtpServer.Host = p_oCompDef.sSMTPServer
            oSmtpServer.EnableSsl = True
            oMail.From = New MailAddress(p_oCompDef.sEmailFrom)
            oMail.To.Add(sSenderEmail)
            oMail.Attachments.Add(New Attachment(sfileName))
            oMail.Subject = p_oCompDef.sEmailSubject
            'oMail.Body = "Greetings ....... " & vbNewLine & p_oCompDef.sEmailBody
            oMail.Body = sBody
            oMail.IsBodyHtml = True
            oSmtpServer.Send(oMail)
            oMail.Dispose()
            '    My.Computer.FileSystem.DeleteFile(AttPAth,
            'Microsoft.VisualBasic.FileIO.UIOption.AllDialogs,
            'Microsoft.VisualBasic.FileIO.RecycleOption.DeletePermanently,
            'Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Console.WriteLine("Completed with SUCCESS " & sFuncName)
            SendEmailNotification = RTN_SUCCESS
        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with Error " & sFuncName)
            SendEmailNotification = RTN_ERROR
        Finally
            File.Delete(sfileName)
        End Try

    End Function

End Module
