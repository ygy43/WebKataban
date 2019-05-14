<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUC_Login.ascx.vb"
    Inherits="WebKataban.WebUC_Login" %>
<div class="login">
    <asp:Panel ID="pnlMain" runat="server" CssClass="mainContainer">
        <div class="title">
            <asp:Label ID="Label1" runat="server" Font-Size="XX-Large"></asp:Label>
        </div>
        <center>
            <asp:Panel ID="pnlTitle" runat="server" ViewStateMode="Disabled" CssClass="authentic">
                <div class="panel">
                    <div>
                        <asp:Label ID="Label2" runat="server"></asp:Label>
                    </div>
                    <div>
                        <asp:TextBox ID="txtUserID" runat="server" class="textbox"></asp:TextBox>
                    </div>
                    <div>
                        <asp:Label ID="Label3" runat="server"></asp:Label>
                    </div>
                    <div>
                        <asp:TextBox ID="txtPasswd" runat="server" class="textbox" TextMode="Password"></asp:TextBox>
                    </div>
                    <div class="divButton">
                        <asp:Button ID="Button1" runat="server" CssClass="button" Width="140px" />
                    </div>
                </div>
            </asp:Panel>
            <asp:Panel ID="pnlGrid" runat="server" ViewStateMode="Disabled" CssClass="authentic">
                <div class="panel">
                    <div class="title">
                        <asp:Label ID="Label4" runat="server"></asp:Label>
                    </div>
                    <div>
                        <asp:Label ID="Label5" runat="server"></asp:Label>
                    </div>
                    <div>
                        <asp:TextBox ID="txtNewPasswd" runat="server" class="textbox" TextMode="Password"></asp:TextBox>
                    </div>
                    <div>
                        <asp:Label ID="Label6" runat="server"></asp:Label>
                    </div>
                    <div>
                        <asp:TextBox ID="txtNewPasswdRe" runat="server" class="textbox" TextMode="Password"></asp:TextBox>
                    </div>
                    <div class="divButton">
                        <asp:Button ID="Button3" runat="server" CssClass="button" Width="140px" />
                    </div>
                </div>
            </asp:Panel>
            <div class="bottom">
                <div>
                    <asp:Label ID="Label7" runat="server" CssClass="label"></asp:Label>
                </div>
                <div>
                    <asp:Label ID="Label8" runat="server" CssClass="label"></asp:Label>
                </div>
            </div>
        </center>
    </asp:Panel>
    <asp:HiddenField ID="hiddenCurrentDatetime" runat="server" />
    <asp:HiddenField ID="hiddenOver" runat="server" />
    <asp:HiddenField ID="hiddenPasswd" runat="server" />
    <asp:HiddenField ID="hiddenNewPasswd" runat="server" />
</div>
