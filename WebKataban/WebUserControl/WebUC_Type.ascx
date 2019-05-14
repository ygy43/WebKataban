<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUC_Type.ascx.vb"
    Inherits="WebKataban.WebUC_Type" %>
<div class="type">
    <asp:Panel ID="PnlSelect" runat="server" CssClass="mainContainer">
        <center>
            <div class="searchPanel">
                <table style="width: 100%; margin-top: 5px;">
                    <tr>
                        <td align="center" style="width: 7%">
                            <asp:Label ID="Label1" runat="server" CssClass="label"></asp:Label>
                        </td>
                        <td align="left" style="width: 93%">
                            <asp:TextBox ID="txtKataban" runat="server" MaxLength="60" ViewStateMode="Enabled"
                                TabIndex="1" CssClass="text"></asp:TextBox>
                        </td>
                        <td align="left" style="width: 10%">
                        </td>
                    </tr>
                </table>
                <table width="100%" cellspacing="0">
                    <tr>
                        <td align="left" style="width: 40%">
                            <table>
                                <tr>
                                    <td>
                                        <asp:Panel ID="PnlSearch" runat="server" Width="100%" HorizontalAlign="Center">
                                            <fieldset class="fieldInset">
                                                <legend align="left">
                                                    <asp:Label ID="Label4" runat="server" Text="Label"></asp:Label>
                                                </legend>
                                                <asp:RadioButtonList ID="RadioButtonList1" runat="server" RepeatDirection="Horizontal"
                                                    RepeatColumns="4" Height="25px" TabIndex="2" Width="400px" Enabled="True">
                                                    <asp:ListItem>0</asp:ListItem>
                                                    <asp:ListItem>1</asp:ListItem>
                                                    <asp:ListItem>2</asp:ListItem>
                                                    <asp:ListItem>3</asp:ListItem>
                                                </asp:RadioButtonList>
                                            </fieldset>
                                        </asp:Panel>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Button ID="Button4" runat="server" TabIndex="3" CssClass="button" />
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td align="right" style="width: 60%">
                            <asp:Image ID="ImgFixedMessage1" runat="server" ImageUrl="~/KHImage/FixedMessage1.gif"
                                Visible="False" />
                        </td>
                    </tr>
                </table>
            </div>
            <div class="grid">
                <asp:GridView ID="GVDetail" runat="server" AutoGenerateColumns="False" CellPadding="2"
                    Width="100%">
                    <HeaderStyle CssClass="gridHeader" />
                    <AlternatingRowStyle CssClass="gridAltRow" />
                    <RowStyle CssClass="gridRow" />
                </asp:GridView>
            </div>
        </center>
        <asp:Panel ID="Panel7" runat="server" Height="40px" Width="100%" BackColor="#C7EDCC"
            HorizontalAlign="Center" BorderStyle="None" Visible="False">
            <center>
                <asp:Panel ID="Panel6" runat="server" HorizontalAlign="Left" Width="90%">
                    <table width="100%" style="padding-top: 5px;">
                        <tr>
                            <td style="width: 80%">
                                <asp:Button ID="Button2" runat="server" TabIndex="19" CssClass="button" />
                                <asp:Button ID="Button3" runat="server" TabIndex="20" CssClass="button" />
                                <asp:Button ID="btnKatOut" runat="server" Text="組合せ出力" Visible="False" CssClass="button" />
                                <asp:Button ID="btnKatsepchk" runat="server" Text="形番分解" Visible="False" CssClass="button" />
                                <asp:Button ID="btn100Test" runat="server" Text="逆展開" Visible="False" CssClass="button" />
                            </td>
                            <td style="width: 20%" align="right">
                                <asp:Button ID="btnOK" runat="server" TabIndex="21" Text="OK" onfocus="this.style.color = 'white';"
                                    onblur="this.style.color = 'black';" CssClass="button" />
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
            </center>
        </asp:Panel>
    </asp:Panel>
    <asp:HiddenField ID="HidSelPage" runat="server" EnableViewState="False" />
    <asp:HiddenField ID="HidSelRowID" runat="server" EnableViewState="False" />
    <asp:HiddenField ID="HidRowCount" runat="server" EnableViewState="False" />
    <asp:HiddenField ID="HidKeyKatabans" runat="server" EnableViewState="False" />
    <asp:HiddenField ID="HidCurrency" runat="server" EnableViewState="false" />
</div>
