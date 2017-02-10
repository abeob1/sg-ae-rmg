Public Class TableCreation

    Dim WithEvents Oapplication_TB As SAPbouiCOM.Application
    Dim oCompany_TB As New SAPbobsCOM.Company

    ' Error handling variables
    Public sErrMsg As String
    Public lErrCode As Integer
    Public lRetCode As Integer

    Public Sub New(ByRef oApplication As SAPbouiCOM.Application, ByRef oCompany As SAPbobsCOM.Company)

        Oapplication_TB = oApplication
        oCompany_TB = oCompany

        Create_Table()

    End Sub



    Public Sub Create_Table()

        AddUserTable("AE_Cdriver", "AE_Chaffuer Booking", SAPbobsCOM.BoUTBTableType.bott_Document)
        AddUserTable("AE_Cdriver_R", "AE_Chaffuer Row", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        AddUserTable("AE_Plist", "AE_Price List", SAPbobsCOM.BoUTBTableType.bott_Document)
        AddUserTable("AE_Plist_R", "AE_Price Row", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)


        AddUserTable("AE_Sbooking", "AE_Self Booking", SAPbobsCOM.BoUTBTableType.bott_Document)
        AddUserTable("AE_Sbooking_R", "AE_Self Booking Row", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)


        AddUserTable("AE_SM", "AE_Service Maint", SAPbobsCOM.BoUTBTableType.bott_Document)
        AddUserTable("AE_SM_R", "AE_Service MaintRow", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        AddUserTable("AE_SM_R1", "AE_Service MaintRow1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)


        AddUserTable("AE_SMMS", "AE_SM Master", SAPbobsCOM.BoUTBTableType.bott_Document)
        AddUserTable("AE_SMMS_R1", "AE_SM Master Row1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        AddUserTable("AE_SMMS_R2", "AE_SM Master Row2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        AddUserTable("AE_SMMS_R3", "AE_SM Master Row3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

        AddUserTable("AE_TrafficO", "AE_Traffic", SAPbobsCOM.BoUTBTableType.bott_Document)
        AddUserTable("AE_TrafficO_R1", "AE_Traffic Row", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)


        AddUserTable("AE_Accident", "AE_Accident", SAPbobsCOM.BoUTBTableType.bott_Document)
        AddUserTable("AE_Accident_R1", "AE_Accident Row1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)


        AddUserTable("AE_Vtrack", "AE_Vehicle Tracking", SAPbobsCOM.BoUTBTableType.bott_Document)

        Add_AE_Cdriver_Fields()
        Add_AE_Cdriver_Rows_Fields()
        Add_AE_Sdriver_Fields()
        Add_AE_Sdriver_Row_Fields()
        Add_AE_PriceList_Fields()
        Add_AE_PriceList_Row_Fields()
        Add_AE_SM_Fields()
        Add_AE_SM_R_Fields()
        Add_AE_SM_R1_Fields()
        Add_AE_SMMS_Fields()
        Add_AE_SMMS_R1_Fields()
        Add_AE_SMMS_R2_Fields()
        Add_AE_SMMS_R3_Fields()
        Add_AE_TO_Fields()
        Add_AE_TO_R_Fields()
        Add_AE_AC_Fields()
        Add_AE_AC_R2_Fields()
        Add_AE_AC_R1_Fields()
        Add_AE_VehicleTrack_Fields()
        Add_AE_OHEM_Fields()
        Add_AE_OITM_Fields()
        AddUDO_Cdriver()
        AddUDO_Sdriver()
        AddUDO_PriceList()
        AddUDO_ServiceMAintenance()
        AddUDO_ServiceMAintenanceSetup()
        AddUDO_TrafficeOffense()
        AddUDO_VehicleTracking()
        AddUDO_AccidentClaim()


        Oapplication_TB.StatusBar.SetText("Table Creation has been Finished ....... !", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)


    End Sub




     Private Sub AddUserTable(ByVal Name As String, ByVal Description As String, _
       ByVal Type As SAPbobsCOM.BoUTBTableType)

        Try

            '//****************************************************************************
            '// The UserTablesMD represents a meta-data object which allows us
            '// to add\remove tables, change a table name etc.
            '//****************************************************************************

            Dim oUserTablesMD As SAPbobsCOM.UserTablesMD

            '//****************************************************************************
            '// In any meta-data operation there should be no other object "alive"
            '// but the meta-data object, otherwise the operation will fail.
            '// This restriction is intended to prevent a collisions
            '//****************************************************************************

            '// the meta-data object needs to be initialized with a
            '// regular UserTables object
            oUserTablesMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)

            '//**************************************************
            '// when adding user tables or fields to the SBO DB
            '// use a prefix identifying your partner name space
            '// this will prevent collisions between different
            '// partners add-ons
            '//
            '// SAP's name space prefix is "BE_"
            '//**************************************************		

            '// set the table parameters
            oUserTablesMD.TableName = Name
            oUserTablesMD.TableDescription = Description
            oUserTablesMD.TableType = Type

            '// Add the table
            '// This action add an empty table with 2 default fields
            '// 'Code' and 'Name' which serve as the key
            '// in order to add your own User Fields
            '// see the AddUserFields.frm in this project
            '// a privat, user defined, key may be added
            '// see AddPrivateKey.frm in this project

            lRetCode = oUserTablesMD.Add
            '// check for errors in the process
            If lRetCode <> 0 Then
                If lRetCode = -1 Then
                Else
                    oCompany_TB.GetLastError(lRetCode, sErrMsg)
                    MsgBox(sErrMsg)
                End If
            Else
                Oapplication_TB.StatusBar.SetText("Table: " & oUserTablesMD.TableName & " was added successfully")
            End If

            oUserTablesMD = Nothing

            GC.Collect() 'Release the handle to the table


        Catch ex As Exception

        End Try

    End Sub

    Private Sub Add_AE_Cdriver_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD.TableName = "@AE_Cdriver"
        oUserFieldsMD.Name = "AE_Bcode"
        oUserFieldsMD.Description = "Billing Code"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        End If
        '************************************
        ' Adding "Room" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver"
        oUserFieldsMD.Name = "AE_Bname"
        oUserFieldsMD.Description = "Billing Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver"
        oUserFieldsMD.Name = "AE_Orderby"
        oUserFieldsMD.Description = "Order by"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver"
        oUserFieldsMD.Name = "AE_Cno"
        oUserFieldsMD.Description = "Contact No."
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver"
        oUserFieldsMD.Name = "AE_Event"
        oUserFieldsMD.Description = "Event"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 200

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver"
        oUserFieldsMD.Name = "AE_Status"
        oUserFieldsMD.Description = "Status"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 30

        oUserFieldsMD.ValidValues.Value = "Open"
        oUserFieldsMD.ValidValues.Description = "Open"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "Cancel C"
        oUserFieldsMD.ValidValues.Description = "With Charges"
        oUserFieldsMD.ValidValues.Add()


        oUserFieldsMD.ValidValues.Value = "Cancel NC"
        oUserFieldsMD.ValidValues.Description = "Without Charges"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "Close"
        oUserFieldsMD.ValidValues.Description = "Close"
        oUserFieldsMD.ValidValues.Add()



        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver"
        oUserFieldsMD.Name = "AE_Issued"
        oUserFieldsMD.Description = "Issued By"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver"
        oUserFieldsMD.Name = "AE_Semployee"
        oUserFieldsMD.Description = "Sales Employees"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 50

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText("Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver"
        oUserFieldsMD.Name = "AE_Amount"
        oUserFieldsMD.Description = "Amount"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price

        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver"
        oUserFieldsMD.Name = "AE_Acharges"
        oUserFieldsMD.Description = "Additional Charges"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver"
        oUserFieldsMD.Name = "AE_Tamount"
        oUserFieldsMD.Description = "Total Amount"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_Cdriver_Rows_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_Stype"
        oUserFieldsMD.Description = "Service Type"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 25

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Room" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_Ptime"
        oUserFieldsMD.Description = "Pickup Time"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_Gname"
        oUserFieldsMD.Description = "Guest Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_GHP"
        oUserFieldsMD.Description = "Guest HP"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_Fno"
        oUserFieldsMD.Description = "Flight No."
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_Ftime"
        oUserFieldsMD.Description = "Flight Time"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_Ploc"
        oUserFieldsMD.Description = "Pickup Location"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_Dloc"
        oUserFieldsMD.Description = "Drop Location"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_Dtime"
        oUserFieldsMD.Description = "Drop Time"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None

        oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_Date"
        oUserFieldsMD.Description = "Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_Vtype"
        oUserFieldsMD.Description = "Vehicle Type"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 25

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_Ono"
        oUserFieldsMD.Description = "Order Chit No"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 25

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_Remarks2"
        oUserFieldsMD.Description = "Remarks2"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 200

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_Remarks1"
        oUserFieldsMD.Description = "Remarks1"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 200

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_Vno"
        oUserFieldsMD.Description = "Vehicle No"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_Dname"
        oUserFieldsMD.Description = "Driver"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 40

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_Dcode"
        oUserFieldsMD.Description = "Driver Code"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_DHP"
        oUserFieldsMD.Description = "Driver HP"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_Tref"
        oUserFieldsMD.Description = "Team Reference"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Cdriver_R"
        oUserFieldsMD.Name = "AE_Tdriver"
        oUserFieldsMD.Description = "Team Driver"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 35

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_Sdriver_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Bcode"
        oUserFieldsMD.Description = "Billing To"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Room" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Bname"
        oUserFieldsMD.Description = "Billing Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Address"
        oUserFieldsMD.Description = "Address"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Cno"
        oUserFieldsMD.Description = "Contact No"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Contract"
        oUserFieldsMD.Description = "Contract No"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 25

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Rno"
        oUserFieldsMD.Description = "Rental Agreement No"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Status"
        oUserFieldsMD.Description = "Status"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Dcode"
        oUserFieldsMD.Description = "Driver Code"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_DName"
        oUserFieldsMD.Description = "Driver Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None

        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Dadd"
        oUserFieldsMD.Description = "Address"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Dcno"
        oUserFieldsMD.Description = "Contact No"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Occuption"
        oUserFieldsMD.Description = "Occupation"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 50

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Nation"
        oUserFieldsMD.Description = "Nationality"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 35

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_DOB"
        oUserFieldsMD.Description = "Date Of Birth"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 200

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_License"
        oUserFieldsMD.Description = "License Number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 35

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Pissue"
        oUserFieldsMD.Description = "Place Of Issue"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 40

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Exdate"
        oUserFieldsMD.Description = "Expiry Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Passno"
        oUserFieldsMD.Description = "Passport Number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 40

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Pissuepno"
        oUserFieldsMD.Description = "Place Of Issue"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 35

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Pexdate"
        oUserFieldsMD.Description = "Expiry Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        'oUserFieldsMD.EditSize = 35

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Dcode1"
        oUserFieldsMD.Description = "Driver Code"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_DName1"
        oUserFieldsMD.Description = "Driver Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Dadd1"
        oUserFieldsMD.Description = "Address"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 200

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Dcno1"
        oUserFieldsMD.Description = "Contact No"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Occuption1"
        oUserFieldsMD.Description = "Occupation"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 50

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Nation1"
        oUserFieldsMD.Description = "Nationality"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 35

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_DOB1"
        oUserFieldsMD.Description = "Date Of Birth"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 200

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_License1"
        oUserFieldsMD.Description = "License Number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 35

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Pissue1"
        oUserFieldsMD.Description = "Place Of Issue"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 40

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Exdate1"
        oUserFieldsMD.Description = "Expiry Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Passno1"
        oUserFieldsMD.Description = "Passport Number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 40

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Pissuepno1"
        oUserFieldsMD.Description = "Place Of Issue"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 35

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Pexdate1"
        oUserFieldsMD.Description = "Expiry Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        'oUserFieldsMD.EditSize = 35

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Vregno"
        oUserFieldsMD.Description = "Registration Number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Vdes"
        oUserFieldsMD.Description = "Description"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Vmodel"
        oUserFieldsMD.Description = "Vehicle Model"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 25

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_expecD"
        oUserFieldsMD.Description = "Date Expected To Return"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        'oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_expecT"
        oUserFieldsMD.Description = "Time Expected To Return"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Vexten"
        oUserFieldsMD.Description = "Extension"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 45

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Vout"
        oUserFieldsMD.Description = "Vehicle Checked Out at"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Vin"
        oUserFieldsMD.Description = "Vehicle Checked In at"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Vkmin"
        oUserFieldsMD.Description = "KM in"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Vkmout"
        oUserFieldsMD.Description = "KM Out"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Vdatein"
        oUserFieldsMD.Description = "Date  In"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Vdatetout"
        oUserFieldsMD.Description = "Date Out"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Vtimein"
        oUserFieldsMD.Description = "Time In"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Vtimeout"
        oUserFieldsMD.Description = "Time Out"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_charges"
        oUserFieldsMD.Description = "Charges For"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 2

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_rate"
        oUserFieldsMD.Description = "Daily Rates"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_dwm"
        oUserFieldsMD.Description = "Number of Days"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 4

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_stot"
        oUserFieldsMD.Description = "Sub Total"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_PAI"
        oUserFieldsMD.Description = "PAI Per Day"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_CDW"
        oUserFieldsMD.Description = "CDW Per Day"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Dcfees"
        oUserFieldsMD.Description = "Delivery / Collection Fees"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Ocharges"
        oUserFieldsMD.Description = "Other Charges"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Des"
        oUserFieldsMD.Description = "Other Charges Description"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Vkmout"
        oUserFieldsMD.Description = "KM Out"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Rcharg"
        oUserFieldsMD.Description = "Monthly Recurring Other Charges"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Rdesc"
        oUserFieldsMD.Description = "Recurring Other Charges Descrip"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_petrol"
        oUserFieldsMD.Description = "Petrol"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_BGST"
        oUserFieldsMD.Description = "Total Before GST"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_GST"
        oUserFieldsMD.Description = "GST"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_GSTP"
        oUserFieldsMD.Description = "GST Per"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Netc"
        oUserFieldsMD.Description = "Net Charge"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Pay"
        oUserFieldsMD.Description = "Form of Payment"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Tamount"
        oUserFieldsMD.Description = "Total Amount"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Exliability"
        oUserFieldsMD.Description = "Excess Liability (SG)"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_SPD"
        oUserFieldsMD.Description = "Surcharge Per Day"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_SPT"
        oUserFieldsMD.Description = "Surchage Total"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_SPGST"
        oUserFieldsMD.Description = "Surchage GST"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_SPNET"
        oUserFieldsMD.Description = "Surchage Net Charge"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_SPLIB"
        oUserFieldsMD.Description = "Excess Liability (MY)"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Perpared"
        oUserFieldsMD.Description = "SA Perpared By"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 60

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Percode"
        oUserFieldsMD.Description = "SA Perpared code"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Invoice"
        oUserFieldsMD.Description = "SA Invoice By"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 60

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Invcode"
        oUserFieldsMD.Description = "SA Invoice Code"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_SPRemarks"
        oUserFieldsMD.Description = "Remarks"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_VMOD"
        oUserFieldsMD.Description = "Out Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        'oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_VMOT"
        oUserFieldsMD.Description = "Out Time"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_VMOdo"
        oUserFieldsMD.Description = "Out Odometer"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_VMFL"
        oUserFieldsMD.Description = "Fuel Level Out"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        oUserFieldsMD.ValidValues.Value = "E"
        oUserFieldsMD.ValidValues.Description = "Empty"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "1/8"
        oUserFieldsMD.ValidValues.Description = "1/8"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "1/4"
        oUserFieldsMD.ValidValues.Description = "1/4"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "3/8"
        oUserFieldsMD.ValidValues.Description = "3/8"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "1/2"
        oUserFieldsMD.ValidValues.Description = "1/2"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "5/8"
        oUserFieldsMD.ValidValues.Description = "5/8"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "3/4"
        oUserFieldsMD.ValidValues.Description = "3/4"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "7/8"
        oUserFieldsMD.ValidValues.Description = "7/8"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "F"
        oUserFieldsMD.ValidValues.Description = "Full"
        oUserFieldsMD.ValidValues.Add()

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_VMSCO"
        oUserFieldsMD.Description = "Staff Check Out"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 60

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_VMSCOC"
        oUserFieldsMD.Description = "Staff Check Out C"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_VMID"
        oUserFieldsMD.Description = "In Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_VMIT"
        oUserFieldsMD.Description = "In Time"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_VMIODO"
        oUserFieldsMD.Description = "In Odometer"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_VMIFL"
        oUserFieldsMD.Description = "Fuel Level In"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        oUserFieldsMD.ValidValues.Value = "E"
        oUserFieldsMD.ValidValues.Description = "Empty"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "1/8"
        oUserFieldsMD.ValidValues.Description = "1/8"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "1/4"
        oUserFieldsMD.ValidValues.Description = "1/4"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "3/8"
        oUserFieldsMD.ValidValues.Description = "3/8"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "1/2"
        oUserFieldsMD.ValidValues.Description = "1/2"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "5/8"
        oUserFieldsMD.ValidValues.Description = "5/8"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "3/4"
        oUserFieldsMD.ValidValues.Description = "3/4"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "7/8"
        oUserFieldsMD.ValidValues.Description = "7/8"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "F"
        oUserFieldsMD.ValidValues.Description = "Full"
        oUserFieldsMD.ValidValues.Add()

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_VMISC"
        oUserFieldsMD.Description = "Staff Check In"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 60

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_VMISCC"
        oUserFieldsMD.Description = "Staff Check In C"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_VMRem"
        oUserFieldsMD.Description = "Remarks"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        oUserFieldsMD.TableName = "@AE_Sbooking"
        oUserFieldsMD.Name = "AE_Gremark"
        oUserFieldsMD.Description = "Remarks"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 200

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_Sdriver_Row_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD.TableName = "@AE_Sbooking_R"
        oUserFieldsMD.Name = "AE_Adate"
        oUserFieldsMD.Description = "Attachment Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        'oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Room" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking_R"
        oUserFieldsMD.Name = "AE_Apath"
        oUserFieldsMD.Description = "Attachment Path"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Memo
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Link
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Sbooking_R"
        oUserFieldsMD.Name = "AE_Afile"
        oUserFieldsMD.Description = "File Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 50

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        GC.Collect() 'Release the handle to the User Fields
    End Sub


    Private Sub Add_AE_PriceList_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD.TableName = "@AE_Plist"
        oUserFieldsMD.Name = "AE_Pcode"
        oUserFieldsMD.Description = "Price List Code"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Room" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Plist"
        oUserFieldsMD.Name = "AE_Pname"
        oUserFieldsMD.Description = "Price List Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_PriceList_Row_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD.TableName = "@AE_Plist_R"
        oUserFieldsMD.Name = "AE_Vtype"
        oUserFieldsMD.Description = "Vehicle Type"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 35

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Room" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Plist_R"
        oUserFieldsMD.Name = "AE_Hrate"
        oUserFieldsMD.Description = "Hourly Rate"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Plist_R"
        oUserFieldsMD.Name = "AE_Surcharge"
        oUserFieldsMD.Description = "Surcharge One Way"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Plist_R"
        oUserFieldsMD.Name = "AE_Surdisposal"
        oUserFieldsMD.Description = "Surcharge Disposal"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_SM_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD.TableName = "@AE_SM"
        oUserFieldsMD.Name = "AE_Vno"
        oUserFieldsMD.Description = "Vehicle Number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Room" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SM"
        oUserFieldsMD.Name = "AE_Vdesc"
        oUserFieldsMD.Description = "Vehicle Description"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SM"
        oUserFieldsMD.Name = "AE_VID"
        oUserFieldsMD.Description = "Vehicle ID"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 25

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SM"
        oUserFieldsMD.Name = "AE_Vmileage"
        oUserFieldsMD.Description = "Vehicle Mileage"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_SM_R_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

      

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SM_R"
        oUserFieldsMD.Name = "AE_Idate"
        oUserFieldsMD.Description = "IN Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 25

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SM_R"
        oUserFieldsMD.Name = "AE_Odate"
        oUserFieldsMD.Description = "Out Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SM_R"
        oUserFieldsMD.Name = "AE_Stype"
        oUserFieldsMD.Description = "Service Type"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 35

        oUserFieldsMD.ValidValues.Value = "Accessories"
        oUserFieldsMD.ValidValues.Description = "Accessories"
        oUserFieldsMD.Add()

        oUserFieldsMD.ValidValues.Value = "Aircon"
        oUserFieldsMD.ValidValues.Description = "Aircon Related Repair"
        oUserFieldsMD.Add()

        oUserFieldsMD.ValidValues.Value = "Battery"
        oUserFieldsMD.ValidValues.Description = "Battery Change"
        oUserFieldsMD.Add()

        oUserFieldsMD.ValidValues.Value = "Body Work"
        oUserFieldsMD.ValidValues.Description = "Body Work"
        oUserFieldsMD.Add()

        oUserFieldsMD.ValidValues.Value = "Brake"
        oUserFieldsMD.ValidValues.Description = "Brake Related"
        oUserFieldsMD.Add()

        oUserFieldsMD.ValidValues.Value = "General Service"
        oUserFieldsMD.ValidValues.Description = "General Service"
        oUserFieldsMD.Add()

        oUserFieldsMD.ValidValues.Value = "Tyre"
        oUserFieldsMD.ValidValues.Description = "Tyre Change"
        oUserFieldsMD.Add()




        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SM_R"
        oUserFieldsMD.Name = "AE_Desc"
        oUserFieldsMD.Description = "Description"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SM_R"
        oUserFieldsMD.Name = "AE_Scode"
        oUserFieldsMD.Description = "Supplier Code"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SM_R"
        oUserFieldsMD.Name = "AE_Sname"
        oUserFieldsMD.Description = "Supplier Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 200

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)


            oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

            oUserFieldsMD.TableName = "@AE_SM_R"
            oUserFieldsMD.Name = "AE_Remark"
            oUserFieldsMD.Description = "Remarks"
            oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
            oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
            oUserFieldsMD.EditSize = 150

            '// Adding the Field to the Table
            lRetCode = oUserFieldsMD.Add

            '// Check for errors
            If lRetCode <> 0 Then
                If lRetCode = -1 Then
                Else
                    oCompany_TB.GetLastError(lRetCode, sErrMsg)
                    MsgBox(sErrMsg)
                End If
            Else
                Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If
        End If
        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_SM_R1_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD.TableName = "@AE_SM_R1"
        oUserFieldsMD.Name = "AE_Adate"
        oUserFieldsMD.Description = "Attachment Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        'oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Room" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SM_R1"
        oUserFieldsMD.Name = "AE_Apath"
        oUserFieldsMD.Description = "Attachment Path"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Memo
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Link
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SM_R1"
        oUserFieldsMD.Name = "AE_Afile"
        oUserFieldsMD.Description = "File Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 50

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_SMMS_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD.TableName = "@AE_SMMS"
        oUserFieldsMD.Name = "AE_Vno"
        oUserFieldsMD.Description = "Vehicle Number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Room" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS"
        oUserFieldsMD.Name = "AE_Vdesc"
        oUserFieldsMD.Description = "Vehicle Description"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS"
        oUserFieldsMD.Name = "AE_GSKM"
        oUserFieldsMD.Description = "Starting KM"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS"
        oUserFieldsMD.Name = "AE_Gfre"
        oUserFieldsMD.Description = "Frequency"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS"
        oUserFieldsMD.Name = "AE_GSerKM"
        oUserFieldsMD.Description = "Service KM"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS"
        oUserFieldsMD.Name = "AE_GSdate"
        oUserFieldsMD.Description = "Start Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        'oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS"
        oUserFieldsMD.Name = "AE_Gdays"
        oUserFieldsMD.Description = "Days"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 5

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS"
        oUserFieldsMD.Name = "AE_GSdate"
        oUserFieldsMD.Description = "Service Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS"
        oUserFieldsMD.Name = "AE_Bsdate"
        oUserFieldsMD.Description = "Start Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS"
        oUserFieldsMD.Name = "AE_Bdays"
        oUserFieldsMD.Description = "Days"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 5

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS"
        oUserFieldsMD.Name = "AE_Bserdate"
        oUserFieldsMD.Description = "Service Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None

        ' oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS"
        oUserFieldsMD.Name = "AE_TSKM"
        oUserFieldsMD.Description = "Starting KM"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS"
        oUserFieldsMD.Name = "AE_Gfre"
        oUserFieldsMD.Description = "Frequency"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS"
        oUserFieldsMD.Name = "AE_TSerKM"
        oUserFieldsMD.Description = "Service KM"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_SMMS_R1_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS_R1"
        oUserFieldsMD.Name = "AE_SKM"
        oUserFieldsMD.Description = "Service KM"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS_R1"
        oUserFieldsMD.Name = "AE_ASKM"
        oUserFieldsMD.Description = "Actual Service KM"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS_R1"
        oUserFieldsMD.Name = "AE_Date"
        oUserFieldsMD.Description = "Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        'oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS_R1"
        oUserFieldsMD.Name = "AE_Adate"
        oUserFieldsMD.Description = "Actual Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 5

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS_R1"
        oUserFieldsMD.Name = "AE_Sdesc"
        oUserFieldsMD.Description = "Service Description"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS_R1"
        oUserFieldsMD.Name = "AE_Docnum"
        oUserFieldsMD.Description = "DocNum"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_SMMS_R2_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS_R1"
        oUserFieldsMD.Name = "AE_Sdate"
        oUserFieldsMD.Description = "Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS_R1"
        oUserFieldsMD.Name = "AE_Adate"
        oUserFieldsMD.Description = "Actual Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        'oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS_R1"
        oUserFieldsMD.Name = "AE_Sdesc"
        oUserFieldsMD.Description = "Service Description"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS_R1"
        oUserFieldsMD.Name = "AE_Docnum"
        oUserFieldsMD.Description = "DocNum"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        End If



        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_SMMS_R3_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS_R3"
        oUserFieldsMD.Name = "AE_SKM"
        oUserFieldsMD.Description = "Service KM"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS_R3"
        oUserFieldsMD.Name = "AE_ASKM"
        oUserFieldsMD.Description = "Actual Service KM"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS_R3"
        oUserFieldsMD.Name = "AE_Date"
        oUserFieldsMD.Description = "Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        'oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS_R3"
        oUserFieldsMD.Name = "AE_Adate"
        oUserFieldsMD.Description = "Actual Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 5

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS_R3"
        oUserFieldsMD.Name = "AE_Sdesc"
        oUserFieldsMD.Description = "Service Description"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_SMMS_R3"
        oUserFieldsMD.Name = "AE_Docnum"
        oUserFieldsMD.Description = "DocNum"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_TO_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Cate"
        oUserFieldsMD.Description = "Category"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 40

        oUserFieldsMD.ValidValues.Value = "Staff (CD or Errand)"
        oUserFieldsMD.ValidValues.Description = "Chauffer Driver"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "RA (Self Drive)"
        oUserFieldsMD.ValidValues.Description = "Rental Agreement"
        oUserFieldsMD.ValidValues.Add()



        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Room" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_RA"
        oUserFieldsMD.Description = "Enter RA / CD"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 25

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_DriverC1"
        oUserFieldsMD.Description = "DriverC"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Driver1"
        oUserFieldsMD.Description = "Driver"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_DriverC2"
        oUserFieldsMD.Description = "DriverC2"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Driver2"
        oUserFieldsMD.Description = "Driver2"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_DriverC3"
        oUserFieldsMD.Description = "DriverC3"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Driver3"
        oUserFieldsMD.Description = "Driver3"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Dcode"
        oUserFieldsMD.Description = "Driver Code"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_DName"
        oUserFieldsMD.Description = "Driver Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Dadd"
        oUserFieldsMD.Description = "Address"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 200

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Dcno"
        oUserFieldsMD.Description = "Contact No"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Occuption"
        oUserFieldsMD.Description = "Occupation"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Nation"
        oUserFieldsMD.Description = "Nationality"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_DOB"
        oUserFieldsMD.Description = "Date Of Birth"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 5

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_License"
        oUserFieldsMD.Description = "License Number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None

        oUserFieldsMD.EditSize = 30

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Pissue"
        oUserFieldsMD.Description = "Lic Place Of Issue"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Exdate"
        oUserFieldsMD.Description = "Expiry Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Passno"
        oUserFieldsMD.Description = "Passport Number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 35

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Pissuepno"
        oUserFieldsMD.Description = "Pass Place Of Issue"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Pexdate"
        oUserFieldsMD.Description = "Expiry Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Dcode1"
        oUserFieldsMD.Description = "Driver Code1"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_DName1"
        oUserFieldsMD.Description = "Driver Name1"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Dadd1"
        oUserFieldsMD.Description = "Address"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Dcno1"
        oUserFieldsMD.Description = "Contact Person"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Dcp1"
        oUserFieldsMD.Description = "Contact No"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Vno"
        oUserFieldsMD.Description = "Vehicle Number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Vdesc"
        oUserFieldsMD.Description = "Vehicle Description"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Vbrand"
        oUserFieldsMD.Description = "Brand"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 40

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Vmodel"
        oUserFieldsMD.Description = "Vehicle Model"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 40

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_make"
        oUserFieldsMD.Description = "Year of Make"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 7

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_chassis"
        oUserFieldsMD.Description = "Chasis Number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 30

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Agency"
        oUserFieldsMD.Description = "Agency"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)


        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Offense"
        oUserFieldsMD.Description = "Type of Offence"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 200

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Odate"
        oUserFieldsMD.Description = "Out Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        'oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Otime"
        oUserFieldsMD.Description = "Out Time"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Oloc"
        oUserFieldsMD.Description = "Location"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Dmerit"
        oUserFieldsMD.Description = "Demerit Point"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_nno"
        oUserFieldsMD.Description = "Notice Number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 40

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_fine"
        oUserFieldsMD.Description = "Fine Amount"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Edate"
        oUserFieldsMD.Description = "Expiry Date of Notice"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Submit"
        oUserFieldsMD.Description = "Submitted By"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_SubmitC"
        oUserFieldsMD.Description = "Submitted By C"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Sdate"
        oUserFieldsMD.Description = "Submission Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Smode"
        oUserFieldsMD.Description = "Submission Mode"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 40

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO"
        oUserFieldsMD.Name = "AE_Oremark"
        oUserFieldsMD.Description = "Remark"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 200

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If





        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_TO_R_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD.TableName = "@AE_TrafficO_R1"
        oUserFieldsMD.Name = "AE_Adate"
        oUserFieldsMD.Description = "Attachment Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        'oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Room" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO_R1"
        oUserFieldsMD.Name = "AE_Apath"
        oUserFieldsMD.Description = "Attachment Path"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Memo
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Link
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_TrafficO_R1"
        oUserFieldsMD.Name = "AE_Afile"
        oUserFieldsMD.Description = "File Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 50

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_AC_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Cate"
        oUserFieldsMD.Description = "Category"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 40

        oUserFieldsMD.ValidValues.Value = "Staff (CD or Errand)"
        oUserFieldsMD.ValidValues.Description = "Chauffer Driver"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "RA (Self Drive)"
        oUserFieldsMD.ValidValues.Description = "Rental Agreement"
        oUserFieldsMD.ValidValues.Add()


        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Room" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_RA"
        oUserFieldsMD.Description = "Enter RA / CD"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 25

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_DriverC1"
        oUserFieldsMD.Description = "DriverC"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Driver1"
        oUserFieldsMD.Description = "Driver"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_DriverC2"
        oUserFieldsMD.Description = "DriverC2"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Driver2"
        oUserFieldsMD.Description = "Driver2"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_DriverC3"
        oUserFieldsMD.Description = "DriverC3"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Driver3"
        oUserFieldsMD.Description = "Driver3"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Dcode"
        oUserFieldsMD.Description = "Driver Code"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_DName"
        oUserFieldsMD.Description = "Driver Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Dadd"
        oUserFieldsMD.Description = "Address"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 200

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Dcno"
        oUserFieldsMD.Description = "Contact No"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Occuption"
        oUserFieldsMD.Description = "Occupation"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Nation"
        oUserFieldsMD.Description = "Nationality"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_DOB"
        oUserFieldsMD.Description = "Date Of Birth"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 5

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_License"
        oUserFieldsMD.Description = "License Number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None

        oUserFieldsMD.EditSize = 30

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Pissue"
        oUserFieldsMD.Description = "Lic Place Of Issue"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Exdate"
        oUserFieldsMD.Description = "Expiry Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Passno"
        oUserFieldsMD.Description = "Passport Number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 35

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Pissuepno"
        oUserFieldsMD.Description = "Pass Place Of Issue"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Pexdate"
        oUserFieldsMD.Description = "Expiry Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 10

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Dcode1"
        oUserFieldsMD.Description = "Driver Code1"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_DName1"
        oUserFieldsMD.Description = "Driver Name1"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Dadd1"
        oUserFieldsMD.Description = "Address"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Dcno1"
        oUserFieldsMD.Description = "Contact Person"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Dcp1"
        oUserFieldsMD.Description = "Contact No"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Vno"
        oUserFieldsMD.Description = "Vehicle Number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Vdesc"
        oUserFieldsMD.Description = "Vehicle Description"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Vbrand"
        oUserFieldsMD.Description = "Brand"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 40

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Vmodel"
        oUserFieldsMD.Description = "Vehicle Model"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 40

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_make"
        oUserFieldsMD.Description = "Year of Make"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 7

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_chassis"
        oUserFieldsMD.Description = "Chasis Number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 30

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Adate"
        oUserFieldsMD.Description = "Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        ' oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)


        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Atime"
        oUserFieldsMD.Description = "Time"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Aloc"
        oUserFieldsMD.Description = "Location"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Ddesc"
        oUserFieldsMD.Description = "Description"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 200

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_payable"
        oUserFieldsMD.Description = "Payable"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 9

        oUserFieldsMD.ValidValues.Value = "Y"
        oUserFieldsMD.ValidValues.Description = "Y"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "N"
        oUserFieldsMD.ValidValues.Description = "N"
        oUserFieldsMD.ValidValues.Add()


        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Eamount"
        oUserFieldsMD.Description = "Excess Amount"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Invoice"
        oUserFieldsMD.Description = "Notice Number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 30

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_OPVN"
        oUserFieldsMD.Description = "Other Party Vehicle No."
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_name"
        oUserFieldsMD.Description = "Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Cno"
        oUserFieldsMD.Description = "Contact No."
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_DrivLicno"
        oUserFieldsMD.Description = "Driving License No."
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 25

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_OPassno"
        oUserFieldsMD.Description = "IC / Passport No."
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 40

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Damage"
        oUserFieldsMD.Description = "Damage Description"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 200

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_accident"
        oUserFieldsMD.Description = "Accident and Scene Description"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 200

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_apayable"
        oUserFieldsMD.Description = "Amount Payable"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Ppayable"
        oUserFieldsMD.Description = "Payable Party"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Rmode"
        oUserFieldsMD.Description = "Report Mode (Y/N)"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 25

        oUserFieldsMD.ValidValues.Value = "Y"
        oUserFieldsMD.ValidValues.Description = "Y"
        oUserFieldsMD.Add()

        oUserFieldsMD.ValidValues.Value = "Y"
        oUserFieldsMD.ValidValues.Description = "Y"
        oUserFieldsMD.Add()


        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_recno"
        oUserFieldsMD.Description = "PV or Receipt No."
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 25

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Iuser"
        oUserFieldsMD.Description = "Insurance User"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 45

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Status"
        oUserFieldsMD.Description = "Insurance Status"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 25

        oUserFieldsMD.ValidValues.Value = "Open"
        oUserFieldsMD.ValidValues.Description = "Open"
        oUserFieldsMD.Add()

        oUserFieldsMD.ValidValues.Value = "Pending Doc"
        oUserFieldsMD.ValidValues.Description = "Pending Document"
        oUserFieldsMD.Add()

        oUserFieldsMD.ValidValues.Value = "Pending Inves"
        oUserFieldsMD.ValidValues.Description = "Pending Investigation"
        oUserFieldsMD.Add()

        oUserFieldsMD.ValidValues.Value = "Pending Outcome"
        oUserFieldsMD.ValidValues.Description = "Pending Outcome"
        oUserFieldsMD.Add()

        oUserFieldsMD.ValidValues.Value = "Closed"
        oUserFieldsMD.ValidValues.Description = "Closed"
        oUserFieldsMD.Add()

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Rdate"
        oUserFieldsMD.Description = "Report Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        'oUserFieldsMD.EditSize = 25

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_IRdate"
        oUserFieldsMD.Description = "Report Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 25

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Idate"
        oUserFieldsMD.Description = "Insurance Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 25

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Rtime"
        oUserFieldsMD.Description = "Report Time"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident"
        oUserFieldsMD.Name = "AE_Irtime"
        oUserFieldsMD.Description = "Report Time"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If





        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_AC_R2_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD.TableName = "@AE_Accident_R1"
        oUserFieldsMD.Name = "AE_Adate"
        oUserFieldsMD.Description = "Attachment Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        'oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Room" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident_R1"
        oUserFieldsMD.Name = "AE_Apath"
        oUserFieldsMD.Description = "Attachment Path"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Memo
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Link
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident_R1"
        oUserFieldsMD.Name = "AE_Afile"
        oUserFieldsMD.Description = "File Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 50

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_AC_R1_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD.TableName = "@AE_Accident_R1"
        oUserFieldsMD.Name = "AE_Adate"
        oUserFieldsMD.Description = "Attachment Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        'oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Room" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident_R1"
        oUserFieldsMD.Name = "AE_Apath"
        oUserFieldsMD.Description = "Attachment Path"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Memo
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Link
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Accident_R1"
        oUserFieldsMD.Name = "AE_Afile"
        oUserFieldsMD.Description = "File Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 50

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If




        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_VehicleTrack_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD.TableName = "@AE_Vtrack"
        oUserFieldsMD.Name = "AE_Loc"
        oUserFieldsMD.Description = "Local Address"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 200

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Room" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Vtrack"
        oUserFieldsMD.Name = "AE_Vno"
        oUserFieldsMD.Description = "Vehicle Number"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Vtrack"
        oUserFieldsMD.Name = "AE_Vdesc"
        oUserFieldsMD.Description = "Vehicle Description"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 50

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Vtrack"
        oUserFieldsMD.Name = "AE_Mileage"
        oUserFieldsMD.Description = "Mileage"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Vtrack"
        oUserFieldsMD.Name = "AE_Petrol"
        oUserFieldsMD.Description = "Petrol"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 5

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Vtrack"
        oUserFieldsMD.Name = "AE_Date"
        oUserFieldsMD.Description = "Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        'oUserFieldsMD.EditSize = 50

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Vtrack"
        oUserFieldsMD.Name = "AE_Time"
        oUserFieldsMD.Description = "Time"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Vtrack"
        oUserFieldsMD.Name = "AE_NRIC"
        oUserFieldsMD.Description = "NRIC"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 25

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Vtrack"
        oUserFieldsMD.Name = "AE_Name"
        oUserFieldsMD.Description = "Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 100

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If





        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "@AE_Vtrack"
        oUserFieldsMD.Name = "AE_Remark"
        oUserFieldsMD.Description = "Remark"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 240

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_OHEM_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD.TableName = "OHEM"
        oUserFieldsMD.Name = "AE_NRICno"
        oUserFieldsMD.Description = "NRIC / FIN No."
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 35

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Room" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OHEM"
        oUserFieldsMD.Name = "AE_NRICname"
        oUserFieldsMD.Description = "NRIC Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OHEM"
        oUserFieldsMD.Name = "AE_Iuno"
        oUserFieldsMD.Description = "In Vehicle Unit No."
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 35

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OHEM"
        oUserFieldsMD.Name = "AE_Aliase"
        oUserFieldsMD.Description = "Aliase Name"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OHEM"
        oUserFieldsMD.Name = "AE_Vno"
        oUserFieldsMD.Description = "Vehicle No."
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 50

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OHEM"
        oUserFieldsMD.Name = "AE_DOB"
        oUserFieldsMD.Description = "Date of Birth"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        'oUserFieldsMD.EditSize = 50

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OHEM"
        oUserFieldsMD.Name = "AE_Team"
        oUserFieldsMD.Description = "Team"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 150

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OHEM"
        oUserFieldsMD.Name = "AE_hp"
        oUserFieldsMD.Description = "Handphone No."
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OHEM"
        oUserFieldsMD.Name = "AE_Address"
        oUserFieldsMD.Description = "Local Address"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 230

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        GC.Collect() 'Release the handle to the User Fields
    End Sub

    Private Sub Add_AE_OITM_Fields()

        Dim oUserFieldsMD As SAPbobsCOM.UserFieldsMD
        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        '************************************
        ' Adding "Name" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_MODEL"
        oUserFieldsMD.Description = "Vehicle Model"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 35

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Room" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_TRANS"
        oUserFieldsMD.Description = "Transmission Type"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 30

        oUserFieldsMD.ValidValues.Value = "Auto"
        oUserFieldsMD.ValidValues.Description = "Auto"
        oUserFieldsMD.ValidValues.Add()

        oUserFieldsMD.ValidValues.Value = "Manual"
        oUserFieldsMD.ValidValues.Description = "Manual"
        oUserFieldsMD.ValidValues.Add()


        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If
        '************************************
        ' Adding "Price" field
        '************************************
        '// Setting the Field's properties

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_YEAR_Make"
        oUserFieldsMD.Description = "Year of Make"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 5

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_COLOR"
        oUserFieldsMD.Description = "Vehicle Color"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 30

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_ENG_CAP"
        oUserFieldsMD.Description = "Engine Capacity"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 35

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_CHASSIS_NO"
        oUserFieldsMD.Description = "Chassis No"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 30

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_ENGINE_NO"
        oUserFieldsMD.Description = "Engine No"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.EditSize = 35

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_REG_DATE"
        oUserFieldsMD.Description = "Registration Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        'oUserFieldsMD.EditSize = 20

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If



        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_TRSF_DATE"
        oUserFieldsMD.Description = "Ownership Transfer Date"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Date
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        'oUserFieldsMD.EditSize = 230

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If

        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_COST"
        oUserFieldsMD.Description = "Net Cost of Purchase"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_DISCOUNT"
        oUserFieldsMD.Description = "Discount"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_COE_QP"
        oUserFieldsMD.Description = "Quotation Premium Paid"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_OMV"
        oUserFieldsMD.Description = "OMV Amount"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_PARF"
        oUserFieldsMD.Description = "Parf Value"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_ANNL_RD_TAX"
        oUserFieldsMD.Description = "Road Tax (12 Months)"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Price
        oUserFieldsMD.EditSize = 9

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_COST"
        oUserFieldsMD.Description = "Net Cost of Purchase"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 30

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_IU NO"
        oUserFieldsMD.Description = "In Vehicle Unit No."
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 30

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_BATTERY"
        oUserFieldsMD.Description = "Battery Capacity"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 30

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_TYRE"
        oUserFieldsMD.Description = "Tyre Model"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 30

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_WARRANTY"
        oUserFieldsMD.Description = "Warranty Period"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 15

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_RHS_WIPER"
        oUserFieldsMD.Description = "Driver Side Wiper Size"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 30

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If


        oUserFieldsMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

        oUserFieldsMD.TableName = "OITM"
        oUserFieldsMD.Name = "AE_LHS_WIPER"
        oUserFieldsMD.Description = "Passenger Side Wiper Size"
        oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha
        oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_None
        oUserFieldsMD.EditSize = 30

        '// Adding the Field to the Table
        lRetCode = oUserFieldsMD.Add

        '// Check for errors
        If lRetCode <> 0 Then
            If lRetCode = -1 Then
            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText( "Field: '" & oUserFieldsMD.Name & "' was added successfuly to " & oUserFieldsMD.TableName & " Table", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End If






        GC.Collect() 'Release the handle to the User Fields
    End Sub



    Private Sub AddUDO_Cdriver()


        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD

        oUserObjectMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

        oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.ChildTables.TableName = "AE_Cdriver_R"
        oUserObjectMD.Code = "AE_Cdriver"
        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.Name = "AE_Chaffuer Booking"
        oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document
        oUserObjectMD.TableName = "AE_Cdriver"

        ' Handle UDO Form

        oUserObjectMD.FormColumns.FormColumnAlias = "DocEntry"
        oUserObjectMD.FormColumns.FormColumnDescription = "DocEntry"
        oUserObjectMD.FormColumns.Add()

        oUserObjectMD.FormColumns.FormColumnAlias = "AE_Bcode"
        oUserObjectMD.FormColumns.FormColumnDescription = "Billing Code"
        oUserObjectMD.FormColumns.Add()

        lRetCode = oUserObjectMD.Add()

        '// check for errors in the process
        If lRetCode <> 0 Then
            If lRetCode = -1 Then

            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText("UDO: " & oUserObjectMD.Name & " was added successfully")
        End If

        oUserObjectMD = Nothing

        GC.Collect() 'Release the handle to the table
    End Sub

    Private Sub AddUDO_Sdriver()


        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD

        oUserObjectMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

        oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.ChildTables.TableName = "AE_Sbooking_R"
        oUserObjectMD.Code = "AE_Sbooking"
        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.Name = "AE_Self Booking"
        oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document
        oUserObjectMD.TableName = "AE_Sbooking"

        ' Handle UDO Form

        oUserObjectMD.FormColumns.FormColumnAlias = "DocEntry"
        oUserObjectMD.FormColumns.FormColumnDescription = "DocEntry"
        oUserObjectMD.FormColumns.Add()

        oUserObjectMD.FormColumns.FormColumnAlias = "AE_Bcode"
        oUserObjectMD.FormColumns.FormColumnDescription = "Billing Code"
        oUserObjectMD.FormColumns.Add()

        lRetCode = oUserObjectMD.Add()

        '// check for errors in the process
        If lRetCode <> 0 Then
            If lRetCode = -1 Then

            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText("UDO: " & oUserObjectMD.Name & " was added successfully")
        End If

        oUserObjectMD = Nothing

        GC.Collect() 'Release the handle to the table
    End Sub

    Private Sub AddUDO_PriceList()


        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD

        oUserObjectMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

        oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.ChildTables.TableName = "AE_Plist_R"
        oUserObjectMD.Code = "AE_Plist"
        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.Name = "AE_Price List"
        oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document
        oUserObjectMD.TableName = "AE_Plist"

        ' Handle UDO Form

        oUserObjectMD.FormColumns.FormColumnAlias = "DocEntry"
        oUserObjectMD.FormColumns.FormColumnDescription = "DocEntry"
        oUserObjectMD.FormColumns.Add()

        oUserObjectMD.FormColumns.FormColumnAlias = "AE_Pcode"
        oUserObjectMD.FormColumns.FormColumnDescription = "Price List Code"
        oUserObjectMD.FormColumns.Add()

        lRetCode = oUserObjectMD.Add()

        '// check for errors in the process
        If lRetCode <> 0 Then
            If lRetCode = -1 Then

            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText("UDO: " & oUserObjectMD.Name & " was added successfully")
        End If

        oUserObjectMD = Nothing

        GC.Collect() 'Release the handle to the table
    End Sub

    Private Sub AddUDO_ServiceMAintenance()


        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD

        oUserObjectMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

        oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.ChildTables.TableName = "AE_SM_R"
        oUserObjectMD.ChildTables.Add()
        oUserObjectMD.ChildTables.TableName = "AE_SM_R1"
        oUserObjectMD.ChildTables.Add()

        oUserObjectMD.Code = "AE_SM"
        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.Name = "AE_Service Maint"
        oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document
        oUserObjectMD.TableName = "AE_SM"

        ' Handle UDO Form

        oUserObjectMD.FormColumns.FormColumnAlias = "DocEntry"
        oUserObjectMD.FormColumns.FormColumnDescription = "DocEntry"
        oUserObjectMD.FormColumns.Add()

        oUserObjectMD.FormColumns.FormColumnAlias = "AE_Vno"
        oUserObjectMD.FormColumns.FormColumnDescription = "Vehicle Number"
        oUserObjectMD.FormColumns.Add()

        lRetCode = oUserObjectMD.Add()

        '// check for errors in the process
        If lRetCode <> 0 Then
            If lRetCode = -1 Then

            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText("UDO: " & oUserObjectMD.Name & " was added successfully")
        End If

        oUserObjectMD = Nothing

        GC.Collect() 'Release the handle to the table
    End Sub

    Private Sub AddUDO_ServiceMAintenanceSetup()


        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD

        oUserObjectMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

        oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.ChildTables.TableName = "AE_SMMS_R1"
        oUserObjectMD.ChildTables.Add()
        oUserObjectMD.ChildTables.TableName = "AE_SMMS_R2"
        oUserObjectMD.ChildTables.Add()
        oUserObjectMD.ChildTables.TableName = "AE_SMMS_R3"
        oUserObjectMD.ChildTables.Add()

        oUserObjectMD.Code = "AE_SMMS"
        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.Name = "AE_SM Master"
        oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document
        oUserObjectMD.TableName = "AE_SMMS"

        ' Handle UDO Form

        oUserObjectMD.FormColumns.FormColumnAlias = "DocEntry"
        oUserObjectMD.FormColumns.FormColumnDescription = "DocEntry"
        oUserObjectMD.FormColumns.Add()

        oUserObjectMD.FormColumns.FormColumnAlias = "AE_Vno"
        oUserObjectMD.FormColumns.FormColumnDescription = "Vehicle Number"
        oUserObjectMD.FormColumns.Add()

        lRetCode = oUserObjectMD.Add()

        '// check for errors in the process
        If lRetCode <> 0 Then
            If lRetCode = -1 Then

            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText("UDO: " & oUserObjectMD.Name & " was added successfully")
        End If

        oUserObjectMD = Nothing

        GC.Collect() 'Release the handle to the table
    End Sub

    Private Sub AddUDO_TrafficeOffense()


        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD

        oUserObjectMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

        oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.ChildTables.TableName = "AE_TrafficO_R1"

        oUserObjectMD.Code = "AE_TrafficO"
        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.Name = "AE_Traffic"
        oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document
        oUserObjectMD.TableName = "AE_TrafficO"

        ' Handle UDO Form

        oUserObjectMD.FormColumns.FormColumnAlias = "DocEntry"
        oUserObjectMD.FormColumns.FormColumnDescription = "DocEntry"
        oUserObjectMD.FormColumns.Add()

        oUserObjectMD.FormColumns.FormColumnAlias = "AE_Cate"
        oUserObjectMD.FormColumns.FormColumnDescription = "Category"
        oUserObjectMD.FormColumns.Add()

        lRetCode = oUserObjectMD.Add()

        '// check for errors in the process
        If lRetCode <> 0 Then
            If lRetCode = -1 Then

            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText("UDO: " & oUserObjectMD.Name & " was added successfully")
        End If

        oUserObjectMD = Nothing

        GC.Collect() 'Release the handle to the table
    End Sub

    Private Sub AddUDO_AccidentClaim()


        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD

        oUserObjectMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

        oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.ChildTables.TableName = "AE_Accident_R1"

        oUserObjectMD.Code = "AE_Accident"
        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.Name = "AE_Accident"
        oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document
        oUserObjectMD.TableName = "AE_Accident"

        ' Handle UDO Form

        oUserObjectMD.FormColumns.FormColumnAlias = "DocEntry"
        oUserObjectMD.FormColumns.FormColumnDescription = "DocEntry"
        oUserObjectMD.FormColumns.Add()

        oUserObjectMD.FormColumns.FormColumnAlias = "AE_Cate"
        oUserObjectMD.FormColumns.FormColumnDescription = "Category"
        oUserObjectMD.FormColumns.Add()

        lRetCode = oUserObjectMD.Add()

        '// check for errors in the process
        If lRetCode <> 0 Then
            If lRetCode = -1 Then

            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText("UDO: " & oUserObjectMD.Name & " was added successfully")
        End If

        oUserObjectMD = Nothing

        GC.Collect() 'Release the handle to the table
    End Sub

    Private Sub AddUDO_VehicleTracking()


        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD

        oUserObjectMD = oCompany_TB.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)

        oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tNO
        oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO

        oUserObjectMD.Code = "AE_Vtrack"
        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
        oUserObjectMD.Name = "AE_Vehicle Tracking"
        oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document
        oUserObjectMD.TableName = "AE_Vtrack"

        ' Handle UDO Form

        lRetCode = oUserObjectMD.Add()

        '// check for errors in the process
        If lRetCode <> 0 Then
            If lRetCode = -1 Then

            Else
                oCompany_TB.GetLastError(lRetCode, sErrMsg)
                MsgBox(sErrMsg)
            End If
        Else
            Oapplication_TB.StatusBar.SetText("UDO: " & oUserObjectMD.Name & " was added successfully")
        End If

        oUserObjectMD = Nothing

        GC.Collect() 'Release the handle to the table
    End Sub


End Class
