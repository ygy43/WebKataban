<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUC_Stopper.ascx.vb"
    Inherits="WebKataban.WebUC_Stopper" %>
<div class="stopper">
    <asp:Panel ID="pnlMain" runat="server" Height="610px" CssClass="mainContainer">
        <center>
            <asp:Panel ID="Panel1" runat="server" Height="5px">
            </asp:Panel>
            <div class="title">
                <asp:Label ID="lblSeriesNm" runat="server" ViewStateMode="Enabled"></asp:Label>
            </div>
            <asp:Panel ID="Panel8" runat="server" Height="10px">
            </asp:Panel>
            <asp:Image ID="img1" runat="server" ImageUrl="~/KHImage/LCGStop.gif" Visible="False" />&nbsp;
            <br />
            <asp:Image ID="img2" runat="server" ImageUrl="~/KHImage/LCRStop.gif" BorderStyle="Outset"
                Visible="False" TabIndex="1" />
        </center>
        <asp:Panel ID="Panel5" runat="server" Height="10px">
        </asp:Panel>
    </asp:Panel>
    <asp:Panel ID="Panel4" runat="server" Height="40px" Width="100%" BackColor="#C7EDCC"
        HorizontalAlign="Center" BorderStyle="None">
        <center>
            <asp:Panel ID="Panel2" runat="server" Height="40px" Width="98%" HorizontalAlign="Left">
                <table style="width: 100%">
                    <tr>
                        <td style="width: 20%">
                        </td>
                        <td style="width: 20%">
                            &nbsp;
                        </td>
                        <td style="width: 20%">
                            &nbsp;
                        </td>
                        <td style="width: 20%">
                            <asp:Button ID="btnOK" runat="server" Font-Size="Medium" Height="28px" Text="OK"
                                Width="100px" onfocus="this.style.color = 'white';" onblur="this.style.color = 'black';"
                                CssClass="button" Font-Bold="True" />
                        </td>
                        <td style="width: 20%">
                        </td>
                    </tr>
                </table>
            </asp:Panel>
        </center>
    </asp:Panel>
</div>
