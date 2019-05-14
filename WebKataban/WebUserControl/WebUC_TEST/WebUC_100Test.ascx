<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUC_100Test.ascx.vb"
    Inherits="WebKataban.WebUC_100Test" %>
<asp:Panel ID="pnlMain" runat="server" CssClass="mainContainer">
    <div class="100Test" style="min-height: 500px;">
        <div>
            <asp:Label ID="lblTitle" runat="server" CssClass="title">テスト機能</asp:Label>
            <asp:Button ID="btnBack" runat="server" TabIndex="-1" Text="戻る" CssClass="button" />
            <asp:LinkButton ID="btnMFTest" runat="server"></asp:LinkButton>
        </div>
        <div style="text-align: left; width: 98%; margin-top: 20px;">
            <table>
                <tr>
                    <td>
                        <asp:Label ID="lblPriceTest" runat="server" CssClass="label">価格テスト</asp:Label>
                    </td>
                    <td>
                        <asp:Button ID="btnPriceTest" runat="server"  Text="価格テスト" CssClass="button" />
                    </td>
                    <td>
                        <asp:Label ID="lblMessage1" runat="server" CssClass="label">「kh_price_test」=>「D:\Log_Net\PriceTest.txt」</asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblShiyouTest" runat="server" CssClass="label" Visible="False">仕様テスト</asp:Label>
                    </td>
                    <td>
                        <asp:Button ID="btnShiyouTest" runat="server"  Text="仕様テスト" CssClass="button" 
                            Enabled="False" Visible="False" />
                    </td>
                    <td>
                        <asp:Label ID="lblMessage2" runat="server" CssClass="label" Visible="False">「kh_shiyou_test」=>「D:\Log_Net\ShiyouTest.txt」</asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lbl100Test" runat="server" CssClass="label" Visible="False">100万件テスト</asp:Label>
                    </td>
                    <td>
                        <asp:Button ID="btn100Test" runat="server"  Text="100万件テスト" CssClass="button" 
                            Enabled="False" Visible="False" />
                    </td>
                    <td>
                        <asp:Label ID="lblMessage3" runat="server" CssClass="label" Visible="False">「kh_TEST_NEW」=>「D:\Log_Net\100Test_グループ番号.txt」</asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblHistory" runat="server" CssClass="label" Visible="False">仕様取込(TXT)テスト</asp:Label>
                    </td>
                    <td>
                        <asp:Button ID="btnMFSiyou" runat="server" Text="仕様取込(TXT)" CssClass="button" 
                            Enabled="False" Visible="False" />
                    </td>
                    <td>
                        <asp:Label ID="lblMessage4" runat="server" CssClass="label" Visible="False">「MF_Siyou」=>「D:\Log_Net\MFShiyouTest.txt」</asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblISOHistory" runat="server" CssClass="label" Visible="False">ISO仕様取込(TXT)テスト</asp:Label>
                    </td>
                    <td>
                        <asp:Button ID="btnMFISO" runat="server"  Text="仕様(ISO)取込(TXT)" 
                            CssClass="button" Enabled="False" Visible="False" />
                    </td>
                    <td>
                        <asp:Label ID="lblMessage5" runat="server" CssClass="label" Visible="False">「MF_Siyou_ISO」=>「D:\Log_Net\MFShiyouTestISO.txt」</asp:Label>
                    </td>
                </tr>
            </table>
        </div>
        <div>
            <asp:Panel ID="pnlName" runat="server" HorizontalAlign="Left" Font-Bold="True" ScrollBars="Auto">
                <asp:GridView ID="GVYouso" runat="server" AutoGenerateColumns="False" CellPadding="2"
                    Width="100%">
                    <Columns>
                        <asp:BoundField DataField="UpdateDate" HeaderText="登録時間" />
                        <asp:BoundField DataField="UpdateUser" HeaderText="ユーザー" />
                        <asp:BoundField DataField="Kataban" HeaderText="形番" />
                        <asp:BoundField DataField="GSPrice" HeaderText="GS価格" />
                        <asp:BoundField DataField="ErrorMsgCd" HeaderText="エラーコード" />
                    </Columns>
                    <HeaderStyle CssClass="gridHeader" />
                    <AlternatingRowStyle CssClass="gridAltRow" />
                    <RowStyle CssClass="gridRow" />
                </asp:GridView>
            </asp:Panel>
        </div>
        <asp:UpdateProgress ID="updateProgress" runat="server" DisplayAfter="500">
            <ProgressTemplate>
                <span class="loadMessage">Loading ...</span>
            </ProgressTemplate>
        </asp:UpdateProgress>
    </div>
</asp:Panel>
<asp:HiddenField ID="HidSelKey" runat="server" EnableViewState="False" />
<asp:HiddenField ID="HidSelRowID" runat="server" EnableViewState="False" />
