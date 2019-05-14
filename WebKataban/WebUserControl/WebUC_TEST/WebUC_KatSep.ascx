<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUC_KatSep.ascx.vb"
    Inherits="WebKataban.WebUC_KatSep" %>
<asp:Panel ID="pnlMain" runat="server" CssClass="mainContainer">
    <div class="katSep">
        <table width="100%">
            <tr>
                <td>
                    <asp:Label ID="lblKataban" runat="server" CssClass="title">形番</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtKata" runat="server" CssClass="textBox" TabIndex="1"></asp:TextBox>
                    <asp:Button ID="btnKatSep" runat="server" CssClass="button" Text="形番分解" TabIndex="3" />
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lblKatabanFilePath" runat="server" CssClass="title">ファイルパス</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtKatabanFilePath" runat="server" CssClass="textBox" TabIndex="2"></asp:TextBox>
                    <%--<asp:FileUpload ID="FileUpload1" runat="server" CssClass="textBox" TabIndex="2" />--%>
                    <asp:Button ID="btnKatSepAll" runat="server" CssClass="button" Text="形番一括分解" TabIndex="4" />
                    <asp:Button ID="btnKatSepAllWithNetPrice" runat="server" CssClass="button" Text="形番一括分解(購入価格込み)" TabIndex="4" />
                </td>
                
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lbl" runat="server" CssClass="title">名前</asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblKataName" runat="server" CssClass="title"></asp:Label>
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td style="width: 15%">
                    <asp:Button ID="btnClear" runat="server" CssClass="button" Text="クリア" TabIndex="-1" />
                </td>
                <td style="width: 15%">
                    <asp:Button ID="btnBack" runat="server" CssClass="button" Text="戻る" TabIndex="5" />
                </td>
                <td style="width: 70%">
                </td>
            </tr>
        </table>
        <div>
            <asp:Label ID="lblSeparator" runat="server" CssClass="title"></asp:Label>
            <asp:Label ID="lblPrice" runat="server" CssClass="title"></asp:Label>
        </div>
        <table style="width: 100%;">
            <tr style="vertical-align: top;">
                <td style="width: 25%" align="right">
                    <asp:Panel ID="Panel3" runat="server">
                        <asp:GridView ID="GVTitle" runat="server" AutoGenerateColumns="False" CellPadding="2"
                            Width="100%">
                            <Columns>
                                <asp:BoundField DataField="title_nm" HeaderText="項目名" ReadOnly="True">
                                    <ItemStyle CssClass="gridTitle" Width="50%" HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="colValue" HeaderText="内容" ReadOnly="True">
                                    <ItemStyle Font-Bold="True" Font-Size="Medium" HorizontalAlign="Center" Width="50%" />
                                </asp:BoundField>
                            </Columns>
                            <HeaderStyle BackColor="Green" Font-Bold="True" Font-Size="Large" ForeColor="White" />
                            <RowStyle Font-Bold="True" Font-Size="Medium" />
                        </asp:GridView>
                    </asp:Panel>
                </td>
                <td style="width: 20%;">
                    <asp:Panel ID="Panel8" runat="server" HorizontalAlign="Center">
                        <asp:GridView ID="GVDetail" runat="server" AutoGenerateColumns="False" CellPadding="2"
                            Visible="False" Width="100%">
                            <Columns>
                                <asp:BoundField DataField="title_nm" HeaderText="区分" ReadOnly="True">
                                    <ItemStyle CssClass="gridTitle" Width="50%" HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="colValue" HeaderText="価格" ReadOnly="True">
                                    <ItemStyle Font-Bold="True" Font-Size="Medium" HorizontalAlign="Right" Width="50%" />
                                </asp:BoundField>
                            </Columns>
                            <HeaderStyle BackColor="Green" Font-Bold="True" Font-Size="Large" ForeColor="White" />
                            <RowStyle Font-Bold="True" Font-Size="Medium" />
                        </asp:GridView>
                    </asp:Panel>
                </td>
                <td style="width: 5%">
                </td>
                <td style="width: 50%">
                    <asp:Panel ID="pnlName" runat="server" HorizontalAlign="Center" Font-Bold="True">
                        <asp:GridView ID="GVYouso" runat="server" AutoGenerateColumns="False" CellPadding="2"
                            Width="100%">
                            <Columns>
                                <asp:BoundField DataField="title_nm" HeaderText="要素名" ReadOnly="True">
                                    <ItemStyle CssClass="gridTitle" Width="70%" HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="colValue" HeaderText="要素内容" ReadOnly="True">
                                    <ItemStyle Font-Bold="True" Font-Size="Medium" HorizontalAlign="Center" Width="20%" />
                                </asp:BoundField>
                                <asp:BoundField DataField="colHyphen" HeaderText="ﾊｲﾌﾝ" ReadOnly="True">
                                    <ItemStyle Font-Bold="True" Font-Size="Medium" HorizontalAlign="Center" Width="10%" />
                                </asp:BoundField>
                            </Columns>
                            <HeaderStyle CssClass="gridHeader" />
                        </asp:GridView>
                    </asp:Panel>
                </td>
            </tr>
        </table>
        <asp:UpdateProgress ID="updateProgress" runat="server" DisplayAfter="500">
            <ProgressTemplate>
                <span class="loadMessage">Loading ...</span></ProgressTemplate>
        </asp:UpdateProgress>
    </div>
</asp:Panel>
