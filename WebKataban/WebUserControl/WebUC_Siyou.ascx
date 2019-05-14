<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUC_Siyou.ascx.vb"
    Inherits="WebKataban.WebUC_Siyou" %>
<div class="siyou">
    <asp:Panel ID="pnlMain" runat="server" CssClass="mainContainer">
        <center>
            <div class="title" style="width: 98%;">
                <asp:Label ID="lblSeriesNm" runat="server" ViewStateMode="Enabled"></asp:Label>
                <table style="width: 100%">
                    <tr>
                        <td style="width: 80%">
                            <asp:Label ID="lblSeriesKat" runat="server" ViewStateMode="Enabled"></asp:Label>
                        </td>
                        <td style="width: 20%">
                            <asp:Button ID="btnOK" runat="server" Font-Size="Medium" Text="OK" UseSubmitBehavior="false"
                                OnClientClick="Siyou_OK('ctl00_ContentDetail_WebUC_Siyou')" Width="100px" Height="28px"
                                onfocus="this.style.color = 'white';" onblur="this.style.color = 'black';" CssClass="button"
                                Font-Bold="True" />
                        </td>
                    </tr>
                </table>
            </div>
            <asp:Panel ID="Panel8" runat="server" Height="5px">
            </asp:Panel>
            <asp:Panel ID="PnlDetail" runat="server" Width="98%" HorizontalAlign="Left">
                <div>
                    <table border="0" cellpadding="0" cellspacing="0" style="width: 100%; border-style: none;">
                        <tr>
                            <td>
                                <asp:GridView ID="GridViewTitle" runat="server" AutoGenerateColumns="False" Font-Bold="True"
                                    Font-Size="12pt" HorizontalAlign="Left" PageSize="50" CellPadding="0" CellSpacing="0">
                                    <HeaderStyle BackColor="Green" ForeColor="White" Height="28px" HorizontalAlign="Center"
                                        VerticalAlign="Middle" />
                                </asp:GridView>
                                <asp:Panel ID="PnlDetail2" runat="server" Width="680px" HorizontalAlign="Left" 
                                    Wrap="False" CssClass="Detail">
                                    <asp:GridView ID="GridViewDetail" runat="server" AutoGenerateColumns="False" Font-Bold="True"
                                        Font-Size="12pt" HorizontalAlign="Left" PageSize="50" CellPadding="0" CellSpacing="0">
                                        <HeaderStyle BackColor="Green" ForeColor="White" Height="28px" HorizontalAlign="Center"
                                            VerticalAlign="Middle" />
                                    </asp:GridView>
                                </asp:Panel>

                            </td>
                        </tr>
                    </table>
                </div>
                <asp:UpdateProgress ID="updateProgress" runat="server" DisplayAfter="500">
                    <ProgressTemplate>
                        <span style="border-width: 0px; position: fixed; padding: 30px; background-color: #FFFFFF;
                            font-size: 36px; left: 40%; top: 30%;">Loading ...</span>
                    </ProgressTemplate>
                </asp:UpdateProgress>
            </asp:Panel>
        </center>
    </asp:Panel>
    <asp:HiddenField ID="HidColMerge" runat="server" />
    <asp:HiddenField ID="HidManifoldMode" runat="server" />
    <asp:HiddenField ID="HidColCount" runat="server" />
    <asp:HiddenField ID="HidClick" runat="server" />
    <asp:HiddenField ID="HidStdNum" runat="server" />
    <asp:HiddenField ID="HidOther" runat="server" />
    <asp:HiddenField ID="HidSelect" runat="server" />
    <asp:HiddenField ID="HidUse" runat="server" />
    <!-- CX操作区分 -->
    <asp:HiddenField ID="HidSetCX" runat="server" />
    <asp:HiddenField ID="HidStartID" runat="server" />
    <asp:HiddenField ID="HidSimpleOther" runat="server" />
    <asp:HiddenField ID="HidCXA" runat="server" />
    <asp:HiddenField ID="HidCXB" runat="server" />
    <asp:HiddenField ID="HidTube" runat="server" />
    <asp:HiddenField ID="HidPostBack" runat="server" Value="0" />
    <div style="height: 0px;">
        <asp:HiddenField ID="HidRailChangeFlg" runat="server" Value="0" />
        <asp:LinkButton ID="btnClick" runat="server" Height="0px" Width="0px"></asp:LinkButton>
    </div>
</div>
