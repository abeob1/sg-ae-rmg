<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="VehicleTrackingReport.aspx.vb" Inherits="VehicleTrackingReport.WebForm1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title> RMG Vehicle Tracking Live Report </title>
    <style type="text/css">
        .style1
        {
            width: 100%;
            height: 1046px;
        }
        .style2
        {
            height: 52px;
            background-color:#6699CC;   
            width: 100%;  
                  
        }
        .style3
        {
            height: 933px;
            width: 100%;
        }
        .style4
        {
            width: 100%;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
     <table class="style1" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td class="style2">
                <asp:Panel ID="Panel1" runat="server" Height="55px" 
                    style="margin-top: 0px" Width="1112px">
                    &nbsp;
                    <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Size="XX-Large"  
                        ForeColor="White" Height="51px" Text="Vehicle Tracking Live Report" 
                        Width="433px"></asp:Label>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;
                </asp:Panel>
            </td>
           
        </tr>
        <tr>
            <td class="style3" style="vertical-align: top">
                <asp:ScriptManager ID="ScriptManager1" runat="server">
                </asp:ScriptManager>
                <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                    <ContentTemplate>
                        <asp:Timer ID="Timer1" runat="server" Interval="7000">
                        </asp:Timer>
                        <asp:Panel ID="Panel2" runat="server" BackColor="#666666" Height="35px" 
                            Width="100%">
                           <asp:Label ID="TimeL" runat="server" Font-Bold="True" Font-Size="Larger"
                                ForeColor="White" Height="41px" Width="450px" style="margin-top: 7px"> </asp:Label>
                           <asp:Label ID="DateL" runat="server" Font-Bold="True" Font-Size="Larger" 
                                ForeColor="White" Height="41px" style="margin-top: 7px" Width="450px"></asp:Label>
                           
                                  &nbsp;&nbsp;&nbsp;&nbsp;                    
                                                                              
                        </asp:Panel>
                        <br />
                        <asp:GridView ID="VehicleTrackingGRID" runat="server" CellPadding="4" ForeColor="#333333" AutoGenerateColumns="true" 
                            style="table-layout:fixed;" Height="168px" Width="100%" Font-Size="20px" 
                              >
                            <AlternatingRowStyle BackColor="White" />
                            <EditRowStyle BackColor="#2461BF" />
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                            <RowStyle BackColor="#EFF3FB" />
                            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                            <SortedAscendingCellStyle BackColor="#F5F7FB" />
                            <SortedAscendingHeaderStyle BackColor="#6D95E1" />
                            <SortedDescendingCellStyle BackColor="#E9EBEF" />
                            <SortedDescendingHeaderStyle BackColor="#4870BE" />
                         </asp:GridView>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </td>
            <td class="style3" style="vertical-align: top">
                &nbsp;</td>
            <td class="style3" style="vertical-align: top">
                &nbsp;</td>
            <td class="style3" style="vertical-align: top">
                &nbsp;</td>
        </tr>
        <tr>
            <td>
                <asp:Label ID="ErrorL" runat="server" Font-Bold="True" Font-Size="Small"></asp:Label>
            </td>
            <td>
                &nbsp;</td>
            <td class="style4">
                &nbsp;</td>
            <td>
                &nbsp;</td>
        </tr>
    </table>
    </form>
</body>
</html>


