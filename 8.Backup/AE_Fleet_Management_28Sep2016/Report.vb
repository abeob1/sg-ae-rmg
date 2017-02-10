Imports System.Data
Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class Report

    Public WithEvents oApplication As SAPbouiCOM.Application
    Public oCompany As New SAPbobsCOM.Company
    Public Docnum As Integer
    Public DocKey As String
    Public Report_Name As String
    Public Report_Parameter As String
    Public Report_Title As String
    Public sGuestName As String
    Public FileName As String
    Public PrinterName As String



    Public Sub Report_CallingFunction()
        Try

            Dim orset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset.DoQuery("SELECT T0.[Name], T0.[U_path] FROM [dbo].[@AE_CRYSTAL]  T0")
            Dim cryRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
            Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
            Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
            Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
            Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table
            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.StartupPath).ToString

            'MsgBox(System.Windows.Forms.Application.StartupPath.ToString)
            cryRpt.Load(orset.Fields.Item("U_path").Value & "\" & Report_Name)
            'cryRpt.Load("PUT CRYSTAL REPORT PATH HERE\CrystalReport1.rpt")

            Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
            Dim crParameterValues As New CrystalDecisions.Shared.ParameterValues
            Dim crParameterDiscreteValue As New CrystalDecisions.Shared.ParameterDiscreteValue
          

            crParameterDiscreteValue.Value = Convert.ToInt32(Docnum)
            crParameterFieldDefinitions = _
        cryRpt.DataDefinition.ParameterFields
            crParameterFieldDefinition = _
        crParameterFieldDefinitions.Item(Report_Parameter)
            crParameterValues = crParameterFieldDefinition.CurrentValues

            crParameterValues.Clear()
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
            Dim Server As String = oCompany.Server
            Dim DB As String = oCompany.CompanyDB
            Dim pwd As String = orset.Fields.Item("Name").Value

            With crConnectionInfo
                .ServerName = Server
                .DatabaseName = DB
                .UserID = "sa"
                .Password = pwd
            End With

            CrTables = cryRpt.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
            Next

          

            Dim RptFrm As Viewer
            RptFrm = New Viewer
            RptFrm.CrystalReportViewer1.ReportSource = cryRpt
            RptFrm.CrystalReportViewer1.Refresh()
            RptFrm.Text = Report_Title
            RptFrm.TopMost = True

            RptFrm.Activate()
            RptFrm.ShowDialog()
            System.Threading.Thread.Sleep(100)

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Sub


    Public Sub GroupTaxInvoice_CallingFunction()
        Try

            Dim orset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset.DoQuery("SELECT T0.[Name], T0.[U_path] FROM [dbo].[@AE_CRYSTAL]  T0")
            Dim cryRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
            Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
            Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
            Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
            Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table
            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.StartupPath).ToString

            'MsgBox(System.Windows.Forms.Application.StartupPath.ToString)
            cryRpt.Load(orset.Fields.Item("U_path").Value & "\" & Report_Name)
            'cryRpt.Load("PUT CRYSTAL REPORT PATH HERE\CrystalReport1.rpt")

            Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
            Dim crParameterValues As New CrystalDecisions.Shared.ParameterValues
            Dim crParameterDiscreteValue As New CrystalDecisions.Shared.ParameterDiscreteValue

            crParameterDiscreteValue.Value = DocKey
            crParameterFieldDefinitions = _
        cryRpt.DataDefinition.ParameterFields
            crParameterFieldDefinition = _
        crParameterFieldDefinitions.Item(Report_Parameter)
            crParameterValues = crParameterFieldDefinition.CurrentValues

            crParameterValues.Clear()
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
            Dim Server As String = oCompany.Server
            Dim DB As String = oCompany.CompanyDB
            Dim pwd As String = orset.Fields.Item("Name").Value

            With crConnectionInfo
                .ServerName = Server
                .DatabaseName = DB
                .UserID = "sa"
                .Password = pwd
            End With

            CrTables = cryRpt.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
            Next


            Dim RptFrm As Viewer
            RptFrm = New Viewer
            RptFrm.CrystalReportViewer1.ReportSource = cryRpt
            RptFrm.CrystalReportViewer1.Refresh()
            RptFrm.Text = Report_Title
            RptFrm.TopMost = True

            RptFrm.Activate()
            RptFrm.ShowDialog()
            System.Threading.Thread.Sleep(100)

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Sub

    Public Sub TaxInvoice_CallingFunction()
        Try

            Dim orset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset.DoQuery("SELECT T0.[Name], T0.[U_path] FROM [dbo].[@AE_CRYSTAL]  T0")
            Dim cryRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
            Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
            Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
            Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
            Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table
            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.StartupPath).ToString

            'MsgBox(System.Windows.Forms.Application.StartupPath.ToString)
            cryRpt.Load(orset.Fields.Item("U_path").Value & "\" & Report_Name)
            'cryRpt.Load("PUT CRYSTAL REPORT PATH HERE\CrystalReport1.rpt")

            Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
            Dim crParameterValues As New CrystalDecisions.Shared.ParameterValues
            Dim crParameterDiscreteValue As New CrystalDecisions.Shared.ParameterDiscreteValue

            crParameterDiscreteValue.Value = DocKey
            crParameterFieldDefinitions = _
        cryRpt.DataDefinition.ParameterFields
            crParameterFieldDefinition = _
        crParameterFieldDefinitions.Item(Report_Parameter)
            crParameterValues = crParameterFieldDefinition.CurrentValues

            crParameterValues.Clear()
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
            Dim Server As String = oCompany.Server
            Dim DB As String = oCompany.CompanyDB
            Dim pwd As String = orset.Fields.Item("Name").Value

            With crConnectionInfo
                .ServerName = Server
                .DatabaseName = DB
                .UserID = "sa"
                .Password = pwd
            End With

            CrTables = cryRpt.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
            Next

            Dim RptFrm As Viewer
            RptFrm = New Viewer
            RptFrm.CrystalReportViewer1.ReportSource = cryRpt
            RptFrm.CrystalReportViewer1.Refresh()
            RptFrm.Text = Report_Title
            RptFrm.TopMost = True

            RptFrm.Activate()
            RptFrm.ShowDialog()
            System.Threading.Thread.Sleep(100)


            ''    Dim RptFrm As Viewer
            ''    RptFrm = New Viewer
            ''    RptFrm.CrystalReportViewer1.ReportSource = cryRpt
            ''    RptFrm.CrystalReportViewer1.Refresh()
            ''    RptFrm.Text = Report_Title
            ''    RptFrm.TopMost = True

            ''    RptFrm.Activate()
            ''    RptFrm.ShowDialog()
            ''    'System.Threading.Thread.Sleep(100)

            ''    cryRpt.Load(orset.Fields.Item("U_path").Value & "\" & "AE_FRM01_PrePrint_TaxInvoice_SD_Copy.rpt")
            ''    crParameterDiscreteValue.Value = DocKey
            ''    crParameterFieldDefinitions = _
            ''cryRpt.DataDefinition.ParameterFields
            ''    crParameterFieldDefinition = _
            ''crParameterFieldDefinitions.Item(Report_Parameter)
            ''    crParameterValues = crParameterFieldDefinition.CurrentValues

            ''    crParameterValues.Clear()
            ''    crParameterValues.Add(crParameterDiscreteValue)
            ''    crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
            ''    With crConnectionInfo
            ''        .ServerName = Server
            ''        .DatabaseName = DB
            ''        .UserID = "sa"
            ''        .Password = pwd
            ''    End With

            ''    CrTables = cryRpt.Database.Tables
            ''    For Each CrTable In CrTables
            ''        crtableLogoninfo = CrTable.LogOnInfo
            ''        crtableLogoninfo.ConnectionInfo = crConnectionInfo
            ''        CrTable.ApplyLogOnInfo(crtableLogoninfo)
            ''    Next
            ''    RptFrm.CrystalReportViewer1.ReportSource = cryRpt
            ''    RptFrm.CrystalReportViewer1.Refresh()
            ''    RptFrm.Text = Report_Title
            ''    RptFrm.TopMost = True

            ''    RptFrm.Activate()
            ''    RptFrm.ShowDialog()

            '' System.Threading.Thread.Sleep(100)

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Sub

    Public Sub TaxInvoice_CallingFunctionOLD()
        Try

            Dim orset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset.DoQuery("SELECT T0.[Name], T0.[U_path] FROM [dbo].[@AE_CRYSTAL]  T0")
            Dim cryRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
            Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
            Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
            Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
            Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table
            Dim sPath As String
            sPath = IO.Directory.GetParent(Application.StartupPath).ToString

            'MsgBox(System.Windows.Forms.Application.StartupPath.ToString)
            cryRpt.Load(orset.Fields.Item("U_path").Value & "\" & Report_Name)
            'cryRpt.Load("PUT CRYSTAL REPORT PATH HERE\CrystalReport1.rpt")

            Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
            Dim crParameterValues As New CrystalDecisions.Shared.ParameterValues
            Dim crParameterDiscreteValue As New CrystalDecisions.Shared.ParameterDiscreteValue

            crParameterDiscreteValue.Value = DocKey
            crParameterFieldDefinitions = _
        cryRpt.DataDefinition.ParameterFields
            crParameterFieldDefinition = _
        crParameterFieldDefinitions.Item(Report_Parameter)
            crParameterValues = crParameterFieldDefinition.CurrentValues

            crParameterValues.Clear()
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
            Dim Server As String = oCompany.Server
            Dim DB As String = oCompany.CompanyDB
            Dim pwd As String = orset.Fields.Item("Name").Value

            With crConnectionInfo
                .ServerName = Server
                .DatabaseName = DB
                .UserID = "sa"
                .Password = pwd
            End With

            CrTables = cryRpt.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
            Next

            Dim RptFrm As Viewer
            RptFrm = New Viewer
            RptFrm.CrystalReportViewer1.ReportSource = cryRpt
            RptFrm.CrystalReportViewer1.Refresh()
            If Not String.IsNullOrEmpty(PrinterName) Then
                cryRpt.PrintOptions.PrinterName = PrinterName
                cryRpt.PrintToPrinter(1, False, 1, 100)
            Else
                cryRpt.PrintToPrinter(1, True, 1, 100)
            End If
            System.Threading.Thread.Sleep(100)

            cryRpt.Load(orset.Fields.Item("U_path").Value & "\" & "AE_FRM01_PrePrint_TaxInvoice_SD_Copy.rpt")
            crParameterDiscreteValue.Value = DocKey
            crParameterFieldDefinitions = _
        cryRpt.DataDefinition.ParameterFields
            crParameterFieldDefinition = _
        crParameterFieldDefinitions.Item(Report_Parameter)
            crParameterValues = crParameterFieldDefinition.CurrentValues

            crParameterValues.Clear()
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
            With crConnectionInfo
                .ServerName = Server
                .DatabaseName = DB
                .UserID = "sa"
                .Password = pwd
            End With

            CrTables = cryRpt.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
            Next
            RptFrm.CrystalReportViewer1.ReportSource = cryRpt
            RptFrm.CrystalReportViewer1.Refresh()
            If Not String.IsNullOrEmpty(PrinterName) Then
                cryRpt.PrintOptions.PrinterName = PrinterName
                cryRpt.PrintToPrinter(1, False, 1, 100)
            Else
                cryRpt.PrintToPrinter(1, True, 1, 100)
            End If
            System.Threading.Thread.Sleep(100)

            ''    Dim RptFrm As Viewer
            ''    RptFrm = New Viewer
            ''    RptFrm.CrystalReportViewer1.ReportSource = cryRpt
            ''    RptFrm.CrystalReportViewer1.Refresh()
            ''    RptFrm.Text = Report_Title
            ''    RptFrm.TopMost = True

            ''    RptFrm.Activate()
            ''    RptFrm.ShowDialog()
            ''    'System.Threading.Thread.Sleep(100)

            ''    cryRpt.Load(orset.Fields.Item("U_path").Value & "\" & "AE_FRM01_PrePrint_TaxInvoice_SD_Copy.rpt")
            ''    crParameterDiscreteValue.Value = DocKey
            ''    crParameterFieldDefinitions = _
            ''cryRpt.DataDefinition.ParameterFields
            ''    crParameterFieldDefinition = _
            ''crParameterFieldDefinitions.Item(Report_Parameter)
            ''    crParameterValues = crParameterFieldDefinition.CurrentValues

            ''    crParameterValues.Clear()
            ''    crParameterValues.Add(crParameterDiscreteValue)
            ''    crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
            ''    With crConnectionInfo
            ''        .ServerName = Server
            ''        .DatabaseName = DB
            ''        .UserID = "sa"
            ''        .Password = pwd
            ''    End With

            ''    CrTables = cryRpt.Database.Tables
            ''    For Each CrTable In CrTables
            ''        crtableLogoninfo = CrTable.LogOnInfo
            ''        crtableLogoninfo.ConnectionInfo = crConnectionInfo
            ''        CrTable.ApplyLogOnInfo(crtableLogoninfo)
            ''    Next
            ''    RptFrm.CrystalReportViewer1.ReportSource = cryRpt
            ''    RptFrm.CrystalReportViewer1.Refresh()
            ''    RptFrm.Text = Report_Title
            ''    RptFrm.TopMost = True

            ''    RptFrm.Activate()
            ''    RptFrm.ShowDialog()

            '' System.Threading.Thread.Sleep(100)

        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try

    End Sub

    Public Sub Report_ExportToPDF()
        Dim sTargetFile As String

        Try

            Dim orset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            orset.DoQuery("SELECT T0.[Name], T0.[U_path] FROM [dbo].[@AE_CRYSTAL]  T0")
            Dim cryRpt As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
            Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
            Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
            Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
            Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table
            Dim sPath As String


            sPath = IO.Directory.GetParent(Application.StartupPath).ToString

            'MsgBox(System.Windows.Forms.Application.StartupPath.ToString)
            cryRpt.Load(orset.Fields.Item("U_path").Value & "\" & Report_Name)
            'cryRpt.Load("PUT CRYSTAL REPORT PATH HERE\CrystalReport1.rpt")

            Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
            Dim crParameterValues As New CrystalDecisions.Shared.ParameterValues
            Dim crParameterDiscreteValue As New CrystalDecisions.Shared.ParameterDiscreteValue
            Dim CrExportOptions As CrystalDecisions.Shared.ExportOptions
            Dim CrDiskFileDestinationOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions()
            Dim CrFormatTypeOptions As New CrystalDecisions.Shared.PdfRtfWordFormatOptions()


            crParameterDiscreteValue.Value = Convert.ToInt32(Docnum)
            crParameterFieldDefinitions = _
        cryRpt.DataDefinition.ParameterFields
            crParameterFieldDefinition = _
        crParameterFieldDefinitions.Item(Report_Parameter)
            crParameterValues = crParameterFieldDefinition.CurrentValues

            crParameterValues.Clear()
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
            Dim Server As String = oCompany.Server
            Dim DB As String = oCompany.CompanyDB
            Dim pwd As String = orset.Fields.Item("Name").Value

            With crConnectionInfo
                .ServerName = Server
                .DatabaseName = DB
                .UserID = "sa"
                .Password = pwd
            End With

            CrTables = cryRpt.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
            Next

            sTargetFile = FileName & Docnum & ".pdf"
            sTargetFile = (orset.Fields.Item("U_path").Value) & "\" & sTargetFile

            Dim DirInfo As New System.IO.DirectoryInfo(orset.Fields.Item("U_path").Value)
            Dim files() As System.IO.FileInfo

            files = DirInfo.GetFiles("*.pdf")

            For Each File As System.IO.FileInfo In files
                Try
                    If File.IsReadOnly = False Then File.Delete()
                Catch ex As Exception
                End Try
            Next

            CrDiskFileDestinationOptions.DiskFileName = sTargetFile
            CrExportOptions = cryRpt.ExportOptions
            With CrExportOptions
                'Set the destination to a disk file 
                .ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                'Set the format to PDF 
                .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                'Set the destination options to DiskFileDestinationOptions object 
                .DestinationOptions = CrDiskFileDestinationOptions
                .FormatOptions = CrFormatTypeOptions
            End With
            'Export the report 
            cryRpt.Export()

            System.Diagnostics.Process.Start(sTargetFile)


        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)


        Finally
           
        End Try


    End Sub


    Public Sub VT_Live_TimerScript()
        Try

            Dim rorset_VT As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            Dim oform As SAPbouiCOM.Form = oApplication.Forms.Item("VTLR")
            Dim omatrix As SAPbouiCOM.Matrix = oform.Items.Item("7").Specific
            omatrix.Clear()

            Dim str_query = "VehicleTrackingLiveReport"


            Try
                oform.DataSources.DataTables.Add("@AE_VTRACK")
            Catch ex As Exception

            End Try

            oform.DataSources.DataTables.Item("@AE_VTRACK").ExecuteQuery(Str_query)

            omatrix.Clear()
            oform.Items.Item("7").Specific.columns.item("V_0mjs").databind.bind("@AE_VTRACK", "U_AE_Loc")
            oform.Items.Item("7").Specific.columns.item("V_3mjs").databind.bind("@AE_VTRACK", "U_AE_Vno")
            oform.Items.Item("7").Specific.columns.item("V_2mjs").databind.bind("@AE_VTRACK", "U_AE_Vdesc")
            oform.Items.Item("7").Specific.columns.item("V_1mjs").databind.bind("@AE_VTRACK", "U_AE_Mileage")
            oform.Items.Item("7").Specific.columns.item("V_7mjs").databind.bind("@AE_VTRACK", "U_AE_NRIC")
            oform.Items.Item("7").Specific.columns.item("V_6mjs").databind.bind("@AE_VTRACK", "U_AE_Name")
            oform.Items.Item("7").Specific.columns.item("V_5mjs").databind.bind("@AE_VTRACK", "U_AE_Remark")
            oform.Items.Item("7").Specific.columns.item("V_4mjs").databind.bind("@AE_VTRACK", "U_AE_Loc1")

            oform.Items.Item("7").Specific.LoadFromDataSource()
            oform.Items.Item("7").Specific.AutoResizeColumns()
            oform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized



        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message & oCompany.CompanyName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub

    Public Sub OpenPagingSign()

        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim sPath As String

        'Dim orset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'orset.DoQuery("SELECT T0.[Name], T0.[U_path] FROM [dbo].[@AE_CRYSTAL]  T0")

        'Start Word and open the document template.
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        ' oDoc = oWord.Documents.Add(orset.Fields.Item("U_path").Value & "\Paging Template.docx")

        sPath = IO.Directory.GetParent(Application.StartupPath).ToString
        oDoc = oWord.Documents.Add(sPath & "\AE_FleetMangement\Paging Template.docx")

        oDoc.Bookmarks("Name_1").Range.Text = sGuestName

    End Sub








End Class
