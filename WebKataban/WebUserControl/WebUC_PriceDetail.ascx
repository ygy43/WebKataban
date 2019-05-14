<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUC_PriceDetail.ascx.vb"
    Inherits="WebKataban.WebUC_PriceDetail" %>
<div class="priceDetail">
    <asp:Panel ID="pnlMain" runat="server" CssClass="mainContainer">
        <div style="width: 98%;">
            <div class="title">
                <div>
                    <asp:Label ID="lblSeriesNm" runat="server" ViewStateMode="Enabled"></asp:Label>
                </div>
                <div>
                    <asp:Label runat="server" ID="lblSeriesKat" ViewStateMode="Enabled"></asp:Label>
                </div>
                <div style="margin-top: 10px;">
                    <asp:Label ID="Label1" runat="server" ViewStateMode="Enabled">価格詳細</asp:Label>
                </div>
            </div>
            <asp:Panel ID="Panel2" runat="server">
                <table style="width: 100%">
                    <tr>
                        <td style="width: 60%">
                            <asp:Label ID="Label2" runat="server" Text="クリップボードにコピーしますか？" Font-Bold="True" Font-Size="11pt"></asp:Label>
                        </td>
                        <td style="width: 20%">
                            <asp:Button ID="Button1" runat="server" Text="OK" Width="100px" onfocus="this.style.color = 'white';"
                                onblur="this.style.color = 'black';" CssClass="button" />
                        </td>
                        <td style="width: 20%">
                            <asp:Button ID="Button2" runat="server" Width="100px" Text="Cancel" onfocus="this.style.color = 'white';"
                                onblur="this.style.color = 'black';" CssClass="button" />
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:GridView ID="grdPriceDetail" runat="server" CellPadding="2" Width="100%" >
                <HeaderStyle CssClass="gridHeader" />
                <AlternatingRowStyle CssClass="gridAltRow" />
                <RowStyle CssClass="gridRow left" Height="22px" />
            </asp:GridView>
        </div>
    </asp:Panel>
    <asp:HiddenField ID="HidMode" runat="server" />
</div>
