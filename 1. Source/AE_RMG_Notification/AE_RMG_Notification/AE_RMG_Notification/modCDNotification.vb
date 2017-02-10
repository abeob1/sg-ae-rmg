Module modCDNotification



    Public Function CDNotification(ByVal sErrDesc As String) As Long

        ' ***********************************************************************************
        '   Function   :    CDNotification()
        '   Purpose    :    This function is handles the Email notification for Chauffer Driver Booking
        '                   This notification triggres Three times when the system clock and time in App.config file are same
        '                   Time1 Functionality : 
        '                     Contidition 1 Whether Monday to Friday
        '                         Send email notiication - [Complete today’s] list of un-assigned driver CD booking
        '                     Contidition 2 Whether Saturday
        '                         Send email notiication - [9am Saturday to Monday 12pm] list of un-assigned driver CD booking

        '                   Time2 Functionality : 
        '                     Contidition 1 Whether Monday to Thursday
        '                         Send email notiication - [From 12pm today to next day 6pm] list of un-assigned driver CD booking
        '                     Contidition 2 Whether Friday
        '                         Send email notiication - [From 12pm today to Monday 12pm] list of un-assigned driver CD booking

        '                   Time3 Functionality:
        '                     Contidition 1 Whether Monday to Thursday
        '                         Send email notiication - [From 4pm today to next day 11:59pm] list of un-assigned driver CD booking
        '                     Contidition 3 Whether Friday
        '                         Send email notiication - [4pm Friday to Monday 12pm] list of un-assigned driver CD booking
        '
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
        Dim sSqlQuery As String = String.Empty
        Dim oCDDateset As DataSet = Nothing
        Dim sHeading As String = String.Empty

        Try
            sFuncName = "Attempting CD Notification"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Console.WriteLine("Starting Function " & sFuncName)
            '-----------------------------------------------------------------------------------------------------
            '------- CD Email Notification Time 1- Triggers when the clock and the notification times are same
            '-----------------------------------------------------------------------------------------------------

            If Format(DateTime.Now, "HH:mm") = p_oCompDef.sCDTime01 Then

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting the Timer 1", sFuncName)

                '------- Nofitication for Monday to Friday
                If Now.DayOfWeek >= DayOfWeek.Monday And Now.DayOfWeek <= DayOfWeek.Friday Then
                    sSqlQuery = "[AE_SP009_CDNotification_1]'" & Format(Now.Date, "MM-dd-yyyy") & "'"

                    '------- Nofitication for Saturday
                ElseIf Now.DayOfWeek = DayOfWeek.Saturday Then
                    sSqlQuery = "[AE_SP010_CDNotification_2]'" & Format(Now.Date, "M-dd-yyyy") & " 09:00" & "', '" & Format(Now.Date.AddDays(2), "M-d-yyyy") & " 12:00" & "'"
                End If

                sHeading = "Chaffure Driver - Unassigned Driver for Vehicles " & p_oCompDef.sCDTime01

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery() " & sSqlQuery, sFuncName)
                Console.WriteLine("Calling ExecuteSQLQuery() " & sSqlQuery & " " & sFuncName)
                ' This function will execute the SQL query and return bflag = true/false, irecordcount, oDataset
                ExecuteSQLQuery(p_oReturnPArameters, sSqlQuery)

                ' This "if" condition will check whether the bflag value true/false (If the ExecuteSQLQuery function completed without error the flag "bflag" 
                'switch to true other wise false ) and also returns the recordcount of SQL query
                If p_oReturnPArameters.bflag = True And p_oReturnPArameters.iRecordcount > 0 Then

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling UploadToExcel() " & sHeading, sFuncName)
                    Console.WriteLine("Calling UploadToExcel() " & sHeading & " " & sFuncName)
                    ' This function will upload the datas for the dataset in to excel file and return bflag = true/false
                    UploadToExcel("CD Notification " & p_oCompDef.sCDTime01.ToString.Replace(":", "."), p_oReturnPArameters.oDateset, sHeading)

                    ' This "if" condition will check whether the bflag value true/false (If the UploadToExcel function completed without error the flag "bflag" 
                    'switch to true other wise false ) 
                    If p_oReturnPArameters.bflag = True Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SendEmailNotification()", sFuncName)
                        Console.WriteLine("Calling SendEmailNotification()" & sFuncName)
                        ' This function will send the notification mail to the concern persons
                        If SendEmailNotification(p_oReturnPArameters.sFpath, p_oCompDef.sEmailTo_CD, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Values Found", sFuncName)
                End If

                '-----------------------------------------------------------------------------------------------------
                '------- CD Email Notification Time 2 - Triggers when the clock and the notification times are same
                '-----------------------------------------------------------------------------------------------------

            ElseIf Format(DateTime.Now, "HH:mm") = p_oCompDef.sCDTime02 Then

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting the Timer 2", sFuncName)

                '------- Nofitication for Monday to Thursday
                If Now.DayOfWeek >= DayOfWeek.Monday And Now.DayOfWeek <= DayOfWeek.Thursday Then

                    sSqlQuery = "[AE_SP010_CDNotification_2]'" & Format(Now.Date, "M-dd-yyyy") & " 12:00" & "', '" & Format(Now.Date.AddDays(1), "M-d-yyyy") & " 18:00" & "'"

                    '------- Nofitication for Friday
                ElseIf Now.DayOfWeek = DayOfWeek.Friday Then
                    sSqlQuery = "[AE_SP010_CDNotification_2]'" & Format(Now.Date, "M-dd-yyyy") & " 12:00" & "', '" & Format(Now.Date.AddDays(3), "M-d-yyyy") & " 12:00" & "'"

                End If
                sHeading = "Chaffure Driver - Unassigned Driver for Vehicles " & p_oCompDef.sCDTime02

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery() " & sSqlQuery, sFuncName)
                Console.WriteLine("Calling ExecuteSQLQuery() " & sSqlQuery & " " & sFuncName)
                ' This function will execute the SQL query and return bflag = true/false, irecordcount, oDataset
                ExecuteSQLQuery(p_oReturnPArameters, sSqlQuery)

                ' This "if" condition will check whether the bflag value true/false (If the ExecuteSQLQuery function completed without error the flag "bflag" 
                'switch to true other wise false ) and also returns the recordcount of SQL query
                If p_oReturnPArameters.bflag = True And p_oReturnPArameters.iRecordcount > 0 Then

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling UploadToExcel() " & sHeading, sFuncName)
                    Console.WriteLine("Calling UploadToExcel() " & sHeading & " " & sFuncName)
                    ' This function will upload the datas for the dataset in to excel file and return bflag = true/false
                    UploadToExcel("CD Notification " & p_oCompDef.sCDTime01.ToString.Replace(":", "."), p_oReturnPArameters.oDateset, sHeading)

                    ' This "if" condition will check whether the bflag value true/false (If the UploadToExcel function completed without error the flag "bflag" 
                    'switch to true other wise false ) 
                    If p_oReturnPArameters.bflag = True Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SendEmailNotification()", sFuncName)
                        Console.WriteLine("Calling SendEmailNotification()" & sFuncName)
                        ' This function will send the notification mail to the concern persons
                        If SendEmailNotification(p_oReturnPArameters.sFpath, p_oCompDef.sEmailTo_CD, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Values Found .......... !", sFuncName)
                End If

                '-----------------------------------------------------------------------------------------------------
                '------- CD Email Notification Time 3 - Triggers when the clock and the notification times are same
                '-----------------------------------------------------------------------------------------------------

            ElseIf Format(DateTime.Now, "HH:mm") = p_oCompDef.sCDTime03 Then

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting the Timer 3", sFuncName)

                '------- Nofitication for Monday to Thursday
                If Now.DayOfWeek >= DayOfWeek.Monday And Now.DayOfWeek <= DayOfWeek.Thursday Then
                    sSqlQuery = "[AE_SP010_CDNotification_2]'" & Format(Now.Date, "M-dd-yyyy") & " 16:00" & "', '" & Format(Now.Date.AddDays(1), "M-d-yyyy") & " 23:59" & "'"

                    '------- Nofitication for Friday
                ElseIf Now.DayOfWeek = DayOfWeek.Friday Then
                    sSqlQuery = "[AE_SP010_CDNotification_2]'" & Format(Now.Date, "M-dd-yyyy") & " 16:00" & "', '" & Format(Now.Date.AddDays(3), "M-d-yyyy") & " 12:00" & "'"

                End If
                sHeading = "Chaffure Driver - Unassigned Driver for Vehicles " & p_oCompDef.sCDTime03

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery() " & sSqlQuery, sFuncName)
                Console.WriteLine("Calling ExecuteSQLQuery() " & sSqlQuery & " " & sFuncName)
                ' This function will execute the SQL query and return bflag = true/false, irecordcount, oDataset
                ExecuteSQLQuery(p_oReturnPArameters, sSqlQuery)

                ' This "if" condition will check whether the bflag value true/false (If the ExecuteSQLQuery function completed without error the flag "bflag" 
                'switch to true other wise false ) and also returns the recordcount of SQL query
                If p_oReturnPArameters.bflag = True And p_oReturnPArameters.iRecordcount > 0 Then

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling UploadToExcel() " & sHeading, sFuncName)
                    Console.WriteLine("Calling UploadToExcel() " & sHeading & " " & sFuncName)
                    ' This function will upload the datas for the dataset in to excel file and return bflag = true/false
                    UploadToExcel("CD Notification " & p_oCompDef.sCDTime01.ToString.Replace(":", "."), p_oReturnPArameters.oDateset, sHeading)

                    ' This "if" condition will check whether the bflag value true/false (If the UploadToExcel function completed without error the flag "bflag" 
                    'switch to true other wise false ) 
                    If p_oReturnPArameters.bflag = True Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SendEmailNotification()", sFuncName)
                        Console.WriteLine("Calling SendEmailNotification()" & sFuncName)
                        ' This function will send the notification mail to the concern persons
                        If SendEmailNotification(p_oReturnPArameters.sFpath, p_oCompDef.sEmailTo_CD, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Values Found ........... !", sFuncName)
                End If

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            Console.WriteLine("Completed with SUCCESS" & sFuncName)
            CDNotification = RTN_SUCCESS
        Catch ex As Exception
            WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Completed with Error " & sFuncName)
            CDNotification = RTN_ERROR
        End Try
        Exit Function 
    End Function

End Module
