<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUC_Motor.ascx.vb"
    Inherits="WebKataban.WebUC_Motor" %>
<div class="motor">
    <div style="background-color: #C7EDCC; padding-top: 5px;">
        <div class="title" style="width: 98%">
            <asp:Label ID="lblSeriesName" runat="server" ViewStateMode="Enabled"></asp:Label>
        </div>
        <center>
            <asp:Image ID="ImageJA" runat="server" ImageUrl="../KHImage/ETSMotorJapan.gif" Visible="False" />
            <asp:Image ID="ImageEN" runat="server" ImageUrl="../KHImage/ETSMotorEnglish.gif"
                Visible="False" />
            <asp:Image ID="ImageKO" runat="server" ImageUrl="../KHImage/ETSMotorKorea.gif" Visible="False" />
            <asp:Image ID="ImageTW" runat="server" ImageUrl="../KHImage/ETSMotorTaiwan.gif" Visible="False" />
            <asp:Image ID="ImageZH" runat="server" ImageUrl="../KHImage/ETSMotorChina.gif" Visible="False" />
            <asp:Image ID="ImageIAVB" runat="server" ImageUrl="../KHImage/IAVBPortPosition.gif" Visible="False" />
        </center>
    </div>
    <div style="background-color: #C7EDCC; padding-top: 5px;">
        <table style="width: 100%">
            <tr>
                <td style="width: 80%">
                </td>
                <td style="width: 20%">
                    <asp:Button ID="btnOK" runat="server" Font-Size="Medium" Height="28px" Text="OK"
                        Width="100px" onfocus="this.style.color = 'white';" onblur="this.style.color = 'black';"
                        CssClass="button" Font-Bold="True" />
                </td>
            </tr>
        </table>
    </div>
</div>
