Module modSDNotification

    Public Function SDNotification(ByVal sErrDesc As String) As Long

        ' ***********************************************************************************
        '   Function   :    SDNotification()
        '   Purpose    :    This function is handles the Email notification for Self Driver Booking
        '                   This notification triggres in four conditions
        '                   Condition 1 : 
        '                     vehicle assignments Notification - Booking without vehicle assignments ( suppose to deliver in next 2 days )
        '                   Condition 2 :
        '                     Contract Expiry Notification - 3months before Contract Expiry - Sale Person get email notification
        '                   Condition 3 :
        '                     Invoice Generation Notification - Invoice for the next rental period (Start of the rental date + 3 days) is not generated then send an email
        '                   Condition 4 :
        '                     Driver Licence Expiry Notification - Driving License Expiry within 2months
        '   Parameters :    ByRef sErrDesc As String
        '                     sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    JOHN
        '   Date       :    07/05/2014 
        '   Change     :   
        '                   
        ' ***********************************************************************************



        Dim sFuncName As String = String.Empty
        Dim sSqlQuery(4) As String
        Dim oCDDateset As DataSet = Nothing
        Dim sHeading(4) As String

        Try
            sFuncName = "Attempting SD Notification"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Console.WriteLine("Starting Function " & sFuncName)

            '-----------------------------------------------------------------------------------------------------
            '------- SD Email Notification Time 1- Triggers when the clock and the notification times are same
            '-----------------------------------------------------------------------------------------------------

            If Format(DateTime.Now, "HH:mm") = p_oCompDef.sSDTime01 Then

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting the Timer 1", sFuncName)

                '------- SD Email Notification Time 1 - Booking without vehicle assignments ( suppose to deliver in next 2 days )

                sSqlQuery(0) = "[AE_SP011_SDNotification_WithoutVehicle]'" & Format(Now.Date.AddDays(2), "MM-dd-yyyy") & "'"
                sHeading(0) = "Self Driver - Booking without vehicle assignments " & p_oCompDef.sSDTime01

                '------- SD Email Notification Time 1 -  3months before Contract Expiry

                sSqlQuery(1) = "[AE_SP013_SDNotification_ContractExpiry]'" & Format(Now.Date.AddMonths(3), "MM-dd-yyyy") & "'"
                sHeading(1) = "Self Driver - Contract Expiry After 3 Months " & p_oCompDef.sSDTime01

                '------- SD Email Notification Time 1 -  Invoice Alert

                sSqlQuery(2) = "[AE_SP012_SDNotification_InvoiceAlert]'" & Format(Now.Date.AddDays(-3), "MM-dd-yyyy") & "'"
                sHeading(2) = "Self Driver - Invoice Alert " & p_oCompDef.sSDTime01

                For ivar As Integer = 0 To 2

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery() " & sSqlQuery(ivar), sFuncName)
                    Console.WriteLine("Calling ExecuteSQLQuery() " & sSqlQuery(ivar) & " " & sFuncName)
                    ' This function will execute the SQL query and return bflag = true/false, irecordcount, oDataset
                    ExecuteSQLQuery(p_oReturnPArameters, sSqlQuery(ivar))

                    ' This "if" condition will check whether the bflag value true/false (If the ExecuteSQLQuery function completed without error the flag "bflag" 
                    'switch to true other wise false ) and also returns the recordcount of SQL query
                    If p_oReturnPArameters.bflag = True And p_oReturnPArameters.iRecordcount > 0 Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling UploadToExcel() " & sHeading(ivar), sFuncName)
                        Console.WriteLine("Calling UploadToExcel() " & sHeading(ivar) & " " & sFuncName)
                        ' This function will upload the datas for the dataset in to excel file and return bflag = true/false
                        UploadToExcel("SD Notification " & p_oCompDef.sCDTime01.ToString.Replace(":", "."), p_oReturnPArameters.oDateset, sHeading(ivar))

                        ' This "if" condition will check whether the bflag value true/false (If the UploadToExcel function completed without error the flag "bflag" 
                        'switch to true other wise false ) 
                        If p_oReturnPArameters.bflag = True Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SendEmailNotification()", sFuncName)
                            Console.WriteLine("Calling SendEmailNotification()" & sFuncName)
                            ' This function will send the notification mail to the concern persons
                            If SendEmailNotification(p_oReturnPArameters.sFpath, p_oCompDef.sEmailTo_SD, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If

                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Values Found ....... !", sFuncName)
                    End If
                Next ivar

                '------- SD Email Notification Time 1 -  Driver Licence Expiry


                If Day(Now.Date) = 1 Then
                    ReDim sSqlQuery(0)
                    ReDim sHeading(0)

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting the Driver Licence Expiry", sFuncName)

                    sSqlQuery(0) = "[AE_SP014_SDNotification_LicenseExpiry]'" & Format(Now.Date.AddMonths(2), "MM-dd-yyyy") & "'"
                    sHeading(0) = "Self Driver - Driver Licence Expiry " & p_oCompDef.sSDTime01

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery() " & sSqlQuery(0), sFuncName)
                    Console.WriteLine("Calling ExecuteSQLQuery() " & sSqlQuery(0) & " " & sFuncName)
                    ' This function will execute the SQL query and return bflag = true/false, irecordcount, oDataset
                    ExecuteSQLQuery(p_oReturnPArameters, sSqlQuery(0))

                    ' This "if" condition will check whether the bflag value true/false (If the ExecuteSQLQuery function completed without error the flag "bflag" 
                    'switch to true other wise false ) and also returns the recordcount of SQL query
                    If p_oReturnPArameters.bflag = True And p_oReturnPArameters.iRecordcount > 0 Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling UploadToExcel() " & sHeading(0), sFuncName)
                        Console.WriteLine("Calling UploadToExcel() " & sHeading(0) & " " & sFuncName)
                        ' This function will upload the datas for the dataset in to excel file and return bflag = true/false
                        UploadToExcel("SD Notification " & p_oCompDef.sCDTime01.ToString.Replace(":", "."), p_oReturnPArameters.oDateset, sHeading(0))

                        ' This "if" condition will check whether the bflag value true/false (If the UploadToExcel function completed without error the flag "bflag" 
                        'switch to true other wise false ) 
                        If p_oReturnPArameters.bflag = True Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SendEmailNotification()", sFuncName)
                            Console.WriteLine("Calling SendEmailNotification()" & sFuncName)
                            ' This function will send the notification mail to the concern persons
                            If SendEmailNotification(p_oReturnPArameters.sFpath, p_oCompDef.sEmailT0_SDContractor, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If

                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Values Found ........ !", sFuncName)
                    End If
                End If

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Console.WriteLine("Completed with SUCCESS" & sFuncName)
            SDNotification = RTN_SUCCESS

        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with Error " & sFuncName)
            SDNotification = RTN_ERROR
        End Try
    End Function

End Module
