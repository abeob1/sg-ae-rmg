Imports System.Data
Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine

Public Class Viewer

    Inherits System.Windows.Forms.Form
    ''Dim crReportDocument1
    '' ''ADO.NET Variables
    ''Dim RMG_DS As RMGDS

    ''Public Sub New(ByVal DocNum As String, ByVal adoOleDbConnection As OleDbConnection)

    ''    MyBase.New()
    ''    Try

    ''        InitializeComponent()
    ''        Dim sqlString1 As String = "AE_SP001_OrderChit'" & DocNum & "'"
    ''        adoOleDbDataAdapter = New OleDbDataAdapter(sqlString1, adoOleDbConnection)
    ''        RMG_DS = New RMGDS
    ''        adoOleDbDataAdapter.Fill(RMG_DS, "RMG")
    ''        crReportDocument1 = New AE_RP001_OrderChit
    ''        crReportDocument1.SetDataSource(RMG_DS)
    ''        CrystalReportViewer1.ReportSource = crReportDocument1


    ''    Catch ex As Exception
    ''        MsgBox(ex.Message)

    ''    End Try

    ''End Sub

    Public Sub New()

        MyBase.New()
        Try

            InitializeComponent()


        Catch ex As Exception
            MsgBox(ex.Message)

        End Try

    End Sub


    Private Sub Viewer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class