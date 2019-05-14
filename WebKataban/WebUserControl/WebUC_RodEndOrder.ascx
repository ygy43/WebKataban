<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUC_RodEndOrder.ascx.vb"
    Inherits="WebKataban.WebUC_RodEndOrder" %>
<div class="rodEndOrder">
    <asp:Panel ID="pnlMain" runat="server" Height="610px" CssClass="mainContainer">
        <asp:Panel ID="PnlSelect" runat="server" Height="80px" Width="100%" HorizontalAlign="Center">
            <asp:Panel ID="Panel1" runat="server" Height="5px">
            </asp:Panel>
            <center>
                <div class="title">
                    <asp:Label runat="server" ID="lblSeriesNm" ViewStateMode="Enabled"></asp:Label>
                </div>
            </center>
            <asp:Panel ID="Panel7" runat="server" Height="30px" HorizontalAlign="Left">
            </asp:Panel>
            <asp:Panel ID="Panel8" runat="server" Height="10px">
            </asp:Panel>
        </asp:Panel>
        <center>
            <asp:Panel ID="Panel2" runat="server" Height="80px" HorizontalAlign="Left" Width="98%">
                <asp:Label ID="Label1" runat="server" Font-Bold="False" Font-Size="X-Large" ForeColor="Black"
                    Height="28px" ViewStateMode="Enabled"></asp:Label>
                <asp:TextBox ID="txtRodEndSize" runat="server" BackColor="#FFFF99" Font-Size="X-Large"
                    Height="28px" Width="70%"></asp:TextBox>
            </asp:Panel>
            <asp:Panel ID="pnlGrid" runat="server" Height="440px" Width="98%" ViewStateMode="Disabled"
                Wrap="true">
                <asp:Panel ID="Panel4" runat="server" Height="30px" HorizontalAlign="Left">
                    <asp:Label ID="Label2" runat="server" Font-Bold="False" Font-Size="Small" Height="20px"
                        ViewStateMode="Enabled"></asp:Label>
                </asp:Panel>
                <asp:Panel ID="Panel9" runat="server" Height="30px" HorizontalAlign="Left">
                    <asp:Label ID="Label3" runat="server" Font-Bold="False" Font-Size="Small" Height="20px"
                        ViewStateMode="Enabled"></asp:Label>
                    <asp:Label ID="Label4" runat="server" Font-Bold="False" Font-Size="Small" ForeColor="#0099FF"
                        Height="20px" ViewStateMode="Enabled"></asp:Label>
                </asp:Panel>
                <asp:Panel ID="Panel10" runat="server" Height="20px" HorizontalAlign="Left">
                </asp:Panel>
                <asp:Panel ID="Panel11" runat="server" Height="30px" HorizontalAlign="Left">
                    <asp:Label ID="Label5" runat="server" Font-Bold="False" Font-Size="Small" Height="20px"
                        ViewStateMode="Enabled"></asp:Label>
                </asp:Panel>
                <asp:Panel ID="Panel12" runat="server" Height="30px" HorizontalAlign="Left">
                    <asp:Label ID="Label6" runat="server" Font-Bold="False" Font-Size="Small" Height="20px"
                        ViewStateMode="Enabled"></asp:Label>
                    <asp:Label ID="Label7" runat="server" Font-Bold="False" Font-Size="Small" ForeColor="#0099FF"
                        Height="20px" ViewStateMode="Enabled"></asp:Label>
                </asp:Panel>
                <asp:Panel ID="Panel13" runat="server" Height="20px" HorizontalAlign="Left">
                </asp:Panel>
                <asp:Panel ID="Panel14" runat="server" Height="30px" HorizontalAlign="Left">
                    <asp:Label ID="Label8" runat="server" Font-Bold="False" Font-Size="Small" Height="20px"
                        ViewStateMode="Enabled"></asp:Label>
                </asp:Panel>
            </asp:Panel>
        </center>
        <asp:Panel ID="Panel5" runat="server" Height="10px">
        </asp:Panel>
    </asp:Panel>
    <asp:Panel ID="Panel15" runat="server" Height="40px" Width="100%" BackColor="#C7EDCC"
        HorizontalAlign="Center" BorderStyle="None">
        <center>
            <asp:Panel ID="Panel6" runat="server" Width="98%" Height="40px" HorizontalAlign="Left">
                <asp:Button ID="btnOK" runat="server" Text="OK" Width="100px" onfocus="this.style.color = 'white';"
                    onblur="this.style.color = 'black';" CssClass="button" />
                <asp:Button ID="Button1" runat="server" Width="120px" Visible="False" CssClass="button" />
                <asp:Button ID="Button2" runat="server" Width="120px" Visible="False" CssClass="button" />
                <asp:Button ID="Button3" runat="server" Width="120px" Visible="False" CssClass="button" />
                <asp:Button ID="Button5" runat="server" Width="120px" Visible="False" CssClass="button" />
                <asp:Button ID="Button6" runat="server" Width="120px" Visible="False" CssClass="button" />
            </asp:Panel>
        </center>
    </asp:Panel>
</div>
