
Imports System.Configuration
Imports System.Data.SqlClient



Public Class WebForm1

   
        Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If LoadGrid() = False Then
            Exit Sub
        End If
        Timer1.Enabled = True
    End Sub

    Private Function LoadGrid() As Boolean

        'Dim sConstr As String = "Data Source=john-pc;Initial Catalog=AddonTesting;User ID=sa; Password=mjs20082010"
        Dim sConstr As String = ConfigurationManager.ConnectionStrings("conn").ConnectionString
        Dim oCon As New SqlConnection(sConstr)
        Dim oCmd As New SqlCommand
        Dim oCmd1 As New SqlCommand
        Dim oReader As SqlDataReader = Nothing
        Dim oDs As New DataSet
        Dim sLocation As String = String.Empty

        Try
            DateL.Text = Now.ToLongDateString
            TimeL.Text = Now.ToLongTimeString

            oCmd1.CommandType = CommandType.Text
            oCmd1.CommandText = "Select T0.name from [@AE_LOCATION] T0 order by T0.Code  "
            oCmd1.Connection = oCon
            oCon.Open()
            oCmd1.ExecuteNonQuery()
            oReader = oCmd1.ExecuteReader
            If oReader.HasRows Then
                Do While oReader.Read
                    sLocation = sLocation & "'" & oReader(0) & "', "
                Loop
            End If
            oCon.Close()

            oCmd.CommandType = CommandType.Text
            Dim ssqltext As String = "[VehicleTrackingLiveReport_Web]" & sLocation
            oCmd.CommandText = ssqltext.Substring(0, ssqltext.Length - 2)
            Dim oDT As New DataTable("@AE_Location")
            oCmd.Connection = oCon
            oCon.Open()
            oCmd.CommandTimeout = 0
            Dim da As New SqlDataAdapter(oCmd)
            da.Fill(oDs)
            VehicleTrackingGRID.DataSource = oDs.Tables(0)
            VehicleTrackingGRID.DataBind()

            ErrorL.Text = "Connected Successfully ...... !"
            oCon.Close()
            Return True
        Catch ex As Exception
            ErrorL.Text = ex.Message
            Return False
        Finally

            oReader.Close()
            oCmd1.Dispose()
            oCon.Dispose()
            oCmd.Dispose()
        End Try

    End Function

        Protected Sub Timer1_Tick(ByVal sender As Object, ByVal e As EventArgs) Handles Timer1.Tick
            LoadGrid()
        End Sub

        Private Sub WebForm1_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload
        Timer1.Enabled = False
        'test
        End Sub

    Private Sub VehicleTrackingGRID_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles VehicleTrackingGRID.RowDataBound
        Dim cell As TableCell = e.Row.Cells(0)
        cell.Width = New Unit("50px")
        cell.HorizontalAlign = HorizontalAlign.Center
    End Sub

End Class