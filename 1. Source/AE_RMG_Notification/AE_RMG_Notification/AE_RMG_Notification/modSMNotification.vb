Module modSMNotification


    Public Function SMNotification(ByVal sErrDesc As String) As Long


        ' ***********************************************************************************
        '   Function   :    SMNotification()
        '   Purpose    :    This function is handles the Email notification for Self Driver Booking
        '                   This notification triggres in three conditions
        '                   Condition 1 : 
        '                     General Service Notification - General Service running KM greater then 10,000 or service date exceeds 1 yr.
        '                   Condition 2 :
        '                     Battery Notification - Battery serviced date exceeds 2yrs.
        '                   Condition 3 :
        '                     Tire Notification - every 40,000 KMs
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
            sFuncName = "Attempting SM Notification"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Console.WriteLine("Starting Function " & sFuncName)

            '-----------------------------------------------------------------------------------------------------
            '------- SM Email Notification Time 1 - Triggers when the clock and the notification times are same
            '-----------------------------------------------------------------------------------------------------

            If Format(DateTime.Now, "HH:mm") = p_oCompDef.sSMTime01 Then

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting the Timer 1", sFuncName)

                '------- SM Email Notification Time 1 - General Service ( change oil , filters and engine maintenances) - >=10000 (Mileage) or 1 Yr

                sSqlQuery(0) = "[AE_SP015_SMNotification_GeneralService]'" & Format(Now.Date.AddYears(-1), "MM-dd-yyyy") & "'"
                sHeading(0) = "Service & Maintenance - Mileage >=10000  or more then 1 Yr  " & p_oCompDef.sSMTime01

                '------- SM Email Notification Time 1 - Battery -  Based on time (2yrs)

                sSqlQuery(1) = "[AE_SP016_SMNotification_BattertyAlert]'" & Format(Now.Date.AddYears(-2), "MM-dd-yyyy") & "'"
                sHeading(1) = "Service & Maintenance - Battery -  Based on time (2yrs)  " & p_oCompDef.sSMTime01

                '------- SM Email Notification Time 1 - Tires – Based on mileage 40,000

                sSqlQuery(2) = "[AE_SP017_SMNotification_TireAlert]"
                sHeading(2) = "Service & Maintenance - Tires – Based on mileage 40,000  " & p_oCompDef.sSMTime01

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
                        UploadToExcel("SM Notification " & p_oCompDef.sCDTime01.ToString.Replace(":", "."), p_oReturnPArameters.oDateset, sHeading(ivar))

                        ' This "if" condition will check whether the bflag value true/false (If the UploadToExcel function completed without error the flag "bflag" 
                        'switch to true other wise false )  
                        If p_oReturnPArameters.bflag = True Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SendEmailNotification()", sFuncName)
                            Console.WriteLine("Calling SendEmailNotification()" & sFuncName)
                            ' This function will send the notification mail to the concern persons
                            If SendEmailNotification(p_oReturnPArameters.sFpath, p_oCompDef.sEmailT0_SM, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If
                    Else
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Values Found ........ !", sFuncName)
                    End If
                Next ivar
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Console.WriteLine("Completed with SUCCESS" & sFuncName)
            SMNotification = RTN_SUCCESS


        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with Error " & sFuncName)
            SMNotification = RTN_ERROR
        End Try

    End Function


End Module
