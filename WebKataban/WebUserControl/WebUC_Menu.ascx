<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUC_Menu.ascx.vb"
    Inherits="WebKataban.WebUC_Menu" %>
<div class="menu">
    <asp:Panel ID="pnlMain" runat="server" Height="580px" CssClass="mainContainer">
        <div class="divTitle">
            <asp:Label ID="Label1" runat="server" CssClass="title"></asp:Label>
        </div>
        <center>
            <asp:Panel ID="pnlGrid" runat="server" ViewStateMode="Disabled" Wrap="true" CssClass="grid">
                <div class="divSubTitle">
                    <asp:Label ID="Label2" runat="server"></asp:Label>
                </div>
                <%--<asp:ListBox ID="ListMsg" runat="server" Rows="20" CssClass="list"></asp:ListBox>--%>
                <asp:TextBox ID="ListMsg" runat="server" TextMode="MultiLine" Rows="20" CssClass="list"></asp:TextBox>
            </asp:Panel>
        </center>
    </asp:Panel>
</div>
