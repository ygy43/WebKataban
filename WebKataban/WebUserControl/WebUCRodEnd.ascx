<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUCRodEnd.ascx.vb"
    Inherits="WebKataban.WebUCRodEnd" %>
<%@ Register Assembly="WebKataban" Namespace="WebKataban" TagPrefix="cc1" %>
<div class="rodEnd">
    <table style="border-style: double; border-width: 1px; border-color: black; width: 420px;
        height: 180px; vertical-align: bottom;">
        <tr>
            <td colspan="2">
                <asp:Label ID="Label5" runat="server" CssClass="RodListLabel"></asp:Label>
                <asp:Label ID="Label6" runat="server" CssClass="RodListLabel"></asp:Label>
            </td>
        </tr>
        <tr>
            <td>
                <asp:Table ID="tblImg" runat="server" CellSpacing="0" CellPadding="0">
                </asp:Table>
            </td>
            <td>
                <asp:Table ID="TblLst" runat="server" CellSpacing="0" CellPadding="0">
                </asp:Table>
            </td>
        </tr>
        <tr>
            <td colspan="2">
                <table id="TblOther" cellspacing="0" cellpadding="0">
                    <tr>
                        <td>
                            <asp:TextBox ID="Label4" runat="server" Width="380px" ReadOnly="true" CssClass="labelOther"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <cc1:CtlCharText ID="CtlCharText1" runat="server" CharCasing="Upper" ToAlp="toHankaku"
                                ToKana="toHankaku" ToKatakana="True" ToKigou="toHankaku" Width="380px"></cc1:CtlCharText>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <asp:HiddenField ID="HdnStdA" runat="server" />
                <asp:HiddenField ID="HdnStdKK" runat="server" />
                <asp:HiddenField ID="HdnStdC" runat="server" />
                <asp:HiddenField ID="HdnSltKK" runat="server" />
                <asp:HiddenField ID="HdnActSltKK" runat="server" />
                <asp:HiddenField ID="HdnRowA" runat="server" />
                <asp:HiddenField ID="HdnRowKK" runat="server" />
                <asp:HiddenField ID="HdnRowC" runat="server" />
            </td>
        </tr>
    </table>
</div>
