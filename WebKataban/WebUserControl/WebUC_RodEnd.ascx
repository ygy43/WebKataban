<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUC_RodEnd.ascx.vb"
    Inherits="WebKataban.WebUC_RodEnd" %>
<div class="rodEnd">
    <asp:Panel ID="pnlMain" runat="server" Height="550px" CssClass="mainContainer">
        <center>
            <div class="title" style="width: 98%;">
                <asp:Label ID="lblSeriesNm" runat="server" ViewStateMode="Enabled"></asp:Label>
            </div>
            <div style="height: 480px; overflow-y: scroll;">
                <table>
                    <tr>
                        <td>
                            <asp:Table ID="TblRodLst" runat="server">
                            </asp:Table>
                        </td>
                    </tr>
                </table>
            </div>
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
                        </td>
                        <td style="width: 20%">
                        </td>
                        <td style="width: 20%">
                            <asp:Button ID="btnOK" runat="server" Text="OK" Width="100px" OnClientClick="f_RodEnd_OK('ctl00_ContentDetail_WebUC_RodEnd','ctl00_ContentDetail_WebUC_RodEnd_wucRodEndOrder')"
                                onfocus="this.style.color = 'white';" onblur="this.style.color = 'black';" CssClass="button" />
                        </td>
                        <td style="width: 20%">
                            <asp:Button ID="btnCancel" runat="server" Width="100px" Text="Cancel" onfocus="this.style.color = 'white';"
                                onblur="this.style.color = 'black';" CssClass="button" />
                        </td>
                    </tr>
                </table>
            </asp:Panel>
        </center>
    </asp:Panel>
    <asp:HiddenField ID="HidMessage" runat="server" />
    <asp:HiddenField ID="HdnPtnCnt" runat="server" />
    <asp:HiddenField ID="HdnSelProdSize" runat="server" />
</div>
