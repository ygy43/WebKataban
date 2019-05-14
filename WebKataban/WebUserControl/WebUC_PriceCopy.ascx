<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUC_PriceCopy.ascx.vb"
    Inherits="WebKataban.WebUC_PriceCopy" %>
<div class="priceCopy">
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
                    <asp:Label ID="Label10" runat="server" ViewStateMode="Enabled">単価積上げ明細</asp:Label>
                </div>
            </div>
            <asp:Panel ID="Panel2" runat="server">
                <table style="width: 100%">
                    <tr>
                        <td style="width: 60%">
                            <asp:Label ID="Label9" runat="server" Text="クリップボードにコピーしますか？" Font-Bold="True" Font-Size="11pt"></asp:Label>
                        </td>
                        <td style="width: 20%">
                            <asp:Button ID="btnOK" runat="server" Text="OK" Width="100px" onfocus="this.style.color = 'white';"
                                onblur="this.style.color = 'black';" CssClass="button" />
                        </td>
                        <td style="width: 20%">
                            <asp:Button ID="btnCancel" runat="server" Width="100px" Text="Cancel" onfocus="this.style.color = 'white';"
                                onblur="this.style.color = 'black';" CssClass="button" />
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Table ID="tblPriceList" runat="server" CellSpacing="0" CellPadding="0" border="1"
                BackColor="White" Width="100%">
            </asp:Table>
        </div>
    </asp:Panel>
    <asp:HiddenField ID="HidMode" runat="server" />
</div>
