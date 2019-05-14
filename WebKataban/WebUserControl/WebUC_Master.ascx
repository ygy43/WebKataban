<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUC_Master.ascx.vb"
    Inherits="WebKataban.WebUC_Master" %>
<%@ Register Assembly="AspNetPager" Namespace="Wuqi.Webdiyer" TagPrefix="webdiyer" %>
<div class="master">
    <asp:Panel ID="pnlMain" runat="server" Width="1000px" BackColor="#C7EDCC" HorizontalAlign="Center"
        BorderStyle="None">
        <center>
            <div class="title" style="width: 98%;">
                <asp:Label ID="Title1" runat="server" ViewStateMode="Enabled"></asp:Label>
            </div>
            <asp:Panel ID="pnlMainTitle" runat="server" Width="98%" HorizontalAlign="Left">
                <table style="width: 100%">
                    <tr>
                        <td style="width: 40%">
                            <asp:Label ID="Title2" runat="server" CssClass="title2"></asp:Label>
                            <asp:Label ID="lblUserID" runat="server" CssClass="title2"></asp:Label>
                        </td>
                        <td style="width: 60%" align="left">
                            <asp:RadioButton ID="RadioButton1" runat="server" AutoPostBack="True" />
                            <asp:RadioButton ID="RadioButton2" runat="server" AutoPostBack="True" />
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Panel ID="pnlSelect" runat="server" ViewStateMode="Disabled" CssClass="selectPanel">
            </asp:Panel>
            <asp:Panel ID="pnlEdit" runat="server" Width="98%" HorizontalAlign="Left">
                <table style="width: 100%" cellpadding="0" cellspacing="0">
                    <tr>
                        <td>
                            <asp:Panel ID="pnlEditTitle" runat="server" CssClass="leftNoBorder">
                            </asp:Panel>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:Panel ID="pnlEditInput" runat="server" ViewStateMode="Disabled" CssClass="leftNoBorder">
                            </asp:Panel>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:Panel ID="pnlWebLog" runat="server" Visible="true" CssClass="leftNoBorder">
                                <table id="tblWebLog" runat="server" width="100%" cellpadding="0" cellspacing="0"
                                    rules="all" border="1">
                                    <tr>
                                        <td class="webLogTitle">
                                            <asp:Label ID="lblPassword" runat="server" Text="パスワード"></asp:Label>
                                        </td>
                                        <td class="webLogTitle">
                                            <asp:Label ID="lblMacAddress" runat="server" Text="マックアドレス"></asp:Label>
                                        </td>
                                        <td class="webLogTitle">
                                            <asp:Label ID="lblSerial" runat="server" Text="シリアルNo"></asp:Label>
                                        </td>
                                        <td class="webLogTitle">
                                            <asp:Label ID="lblLastUsedTime" runat="server" Text="最終利用時間"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="width: 25%">
                                            <asp:TextBox ID="txtPassword" runat="server" CssClass="webLogInput"></asp:TextBox>
                                        </td>
                                        <td style="width: 25%">
                                            <asp:TextBox ID="txtMacAddress" runat="server" CssClass="webLogInput"></asp:TextBox>
                                        </td>
                                        <td style="width: 25%">
                                            <asp:TextBox ID="txtSerial" runat="server" CssClass="webLogInput"></asp:TextBox>
                                        </td>
                                        <td style="width: 25%">
                                            <asp:TextBox ID="txtLastUsedTime" runat="server" CssClass="webLogInput" Enabled="False"></asp:TextBox>
                                        </td>
                                    </tr>
                                </table>
                            </asp:Panel>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:Panel ID="pnlEditInput1" runat="server" CssClass="leftNoBorder">
                            </asp:Panel>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:Panel ID="pnlEditInput2" runat="server" CssClass="leftNoBorder">
                            </asp:Panel>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:Panel ID="pnlEditInput3" runat="server" CssClass="leftNoBorder">
                            </asp:Panel>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:Panel ID="Panel6" runat="server" Height="5px" Width="100%">
                            </asp:Panel>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:Panel ID="pnlEditButton" runat="server" CssClass="centerNoBorder">
                            </asp:Panel>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Panel ID="Panel5" runat="server" Height="5px" Width="98%">
            </asp:Panel>
            <asp:Panel ID="pnlList" runat="server" CssClass="leftNoBorder">
                <table cellpadding="0" cellspacing="0" width="100%">
                    <tr>
                        <td>
                            <asp:Table ID="tblTitle" runat="server" CellPadding="0" CellSpacing="0" Width="98%"
                                CssClass="margin1percent">
                            </asp:Table>
                            <div class="gridViewMain">
                                <asp:GridView ID="GridViewMain" runat="server" AutoGenerateColumns="False" Font-Bold="False"
                                    Font-Size="11pt" HorizontalAlign="center" PageSize="20" CellPadding="0" CellSpacing="0"
                                    GridLines="None" ShowHeader="False">
                                </asp:GridView>
                            </div>
                            <div>
                                <webdiyer:AspNetPager runat="server" ID="AspNetPager1" PageSize="5" Visible="false">
                                </webdiyer:AspNetPager>
                            </div>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
        </center>
    </asp:Panel>
    <asp:HiddenField ID="HidMode" runat="server" />
    <asp:HiddenField ID="HidSelID" runat="server" />
    <asp:HiddenField ID="HidTableKey" runat="server" />
    <asp:HiddenField ID="HidRateDiv" runat="server" />
</div>
