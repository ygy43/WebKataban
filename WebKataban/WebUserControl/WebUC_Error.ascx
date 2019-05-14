<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUC_Error.ascx.vb"
    Inherits="WebKataban.WebUC_Error" %>
<asp:Panel ID="pnlMain" runat="server" Height="650px" CssClass="mainContainer">
    <asp:Panel ID="Panel2" runat="server" Height="40px">
    </asp:Panel>
    <center>
        <asp:Panel ID="PnlHead" runat="server" Width="80%" HorizontalAlign="Center">
            <asp:Panel ID="Panel3" runat="server" HorizontalAlign="Center">
                <asp:Label ID="lblTitle" runat="server" Font-Bold="False" Font-Size="Larger" ForeColor="Black"
                    Width="98%" ViewStateMode="Enabled"></asp:Label>
            </asp:Panel>
            <asp:Panel ID="Panel1" runat="server" Height="30px">
            </asp:Panel>
            <asp:Panel ID="Panel7" runat="server" HorizontalAlign="Center">
                <asp:Label runat="server" Font-Bold="True" Font-Size="Large" ForeColor="Red" Width="98%"
                    ID="lblMessage" Style="margin-top: 0px" ViewStateMode="Enabled"></asp:Label>
            </asp:Panel>
            <asp:Panel ID="Panel8" runat="server" Height="30px">
            </asp:Panel>
        </asp:Panel>
        <asp:Panel ID="Panel6" runat="server" Width="80%" Height="40px" HorizontalAlign="Center">
            <asp:Button ID="btnLogin" runat="server" Text="Login" Width="100px" onfocus="this.style.color = 'white';"
                onblur="this.style.color = 'black';" CssClass="button" Visible="False" />
            <asp:Button ID="btnClose" runat="server" Text="Close" Width="100px" onfocus="this.style.color = 'white';"
                onblur="this.style.color = 'black';" OnClientClick="window.open('about:blank','_self').close();return false;"
                CssClass="button" />
            <asp:HiddenField ID="HidErrMsg" runat="server" />
        </asp:Panel>
    </center>
</asp:Panel>
