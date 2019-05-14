<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="UC_ISOTanka.ascx.vb"
    Inherits="WebKataban.UC_ISOTanka" %>
<%@ Register Assembly="WebKataban" Namespace="WebKataban" TagPrefix="cc1" %>
<div class="ucIsoTanka" id ="tbl">
    <table style="border-style: ridge; border-width: 1px;" width="770px">
        <tr>
            <td align="center">
                <asp:Label ID="lblNo" runat="server" CssClass="title" Width="45px"></asp:Label>
            </td>
            <td align="left">
                <table style="width: 245px; background-color: Green;" cellpadding="0" cellspacing="0">
                    <tr>
                        <td>
                            <asp:Label ID="lblName" runat="server" CssClass="title" Width="100%"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:Label ID="lblKataName" runat="server" CssClass="title" Width="100%"></asp:Label>
                        </td>
                    </tr>
                </table>
            </td>
            <td align="left">
                <table cellpadding="1" cellspacing="0" class="priceListTable">
                    <tr>
                        <td align="left">
                            <asp:Label ID="Label29" runat="server" CssClass="label"></asp:Label>
                            <asp:Label ID="Label30" runat="server" CssClass="label"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <div class="verticalMiddle">
                                <asp:CheckBox ID="ChkUnitList" runat="server" CssClass="label" />
                                <asp:Label ID="Label31" runat="server" CssClass="label"></asp:Label>
                            </div>
                        </td>
                    </tr>
                </table>
            </td>
            <td align="right" style="width: 160px">
                <table cellspacing="1px" cellpadding="0px">
                    <tr>
                        <td align="center">
                            <asp:Label ID="Label21" runat="server" CssClass="columnTitle" Width="100%"></asp:Label>
                        </td>
                        <td align="center">
                            <asp:Label ID="Label22" runat="server" CssClass="columnTitle" Width="100%"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td align="center" style="border: 1px solid #C0C0C0; white-space: nowrap;" bgcolor="White" class="ClsChk">
                            <cc1:CtlNumText ID="txt_ChkZ" runat="server" ReadOnly="True" ViewStateMode="Enabled" CssClass="textReadOnly" BorderStyle="None"></cc1:CtlNumText><cc1:CtlNumText ID="txt_KtbnChk" runat="server" ReadOnly="True" 
                                ViewStateMode="Enabled" CssClass="textReadOnly" BorderStyle="None" 
                                style="margin-left: 0px"></cc1:CtlNumText>
                        </td>
                        <td>
                            <cc1:CtlNumText ID="txt_Place" runat="server" ReadOnly="True" ViewStateMode="Enabled"
                                CssClass="textReadOnly">
                            </cc1:CtlNumText>
                        </td>
                        <td>
                          <asp:Label ID="lblStrageLocation" runat="server" CssClass="label" Width="100%"></asp:Label>
                        </td>
                        <td>
                            &nbsp;</td>
                        <%--                        <td align="left" style="font-weight: 700;">
                            <asp:Label ID="Label15" runat="server" BackColor="Green" Font-Bold="True" ForeColor="White"
                                Height="15px" Width="100%" Visible="False"></asp:Label>
                        </td>--%>
                    </tr>
                </table>
            </td>
        </tr>
        <tr style="height: 86px">
            <td>
            </td>
            <td>
                <table style="width: 100%" cellpadding="0" cellspacing="0">
                    <tr>
                        <td align="left">
                            <asp:Label ID="Label23" runat="server" CssClass="columnTitle" Width="100%"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td align="left">
                            <asp:Panel ID="pnlPrice" runat="server" Height="70px" ScrollBars="Vertical">
                                <asp:GridView ID="GVPrice" runat="server" AutoGenerateColumns="False" Font-Bold="True"
                                    Font-Size="12pt" HorizontalAlign="Left" CellPadding="0" ShowHeader="False" Width="100%"
                                    BackColor="White">
                                    <Columns>
                                        <asp:BoundField DataField="strText">
                                            <ItemStyle Height="20px" HorizontalAlign="Left" Width="50%" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="strPrice">
                                            <ItemStyle Height="20px" HorizontalAlign="Right" Width="50%" />
                                        </asp:BoundField>
                                        <asp:TemplateField ItemStyle-CssClass="displaynone">
                                            <ItemTemplate>
                                                <asp:HiddenField runat="server" ID="ColumnKbn" Value='<%#Eval("ColumnKBN") %>'></asp:HiddenField>
                                            </ItemTemplate>
                                            <HeaderStyle CssClass="displaynone" />
                                            <ItemStyle CssClass="displaynone" />
                                        </asp:TemplateField>
                                    </Columns>
                                </asp:GridView>
                            </asp:Panel>
                        </td>
                    </tr>
                </table>
            </td>
            <td colspan="2" align="left">
                <table cellspacing="1px" cellpadding="0px" style="width: 470px;">
                    <tr style="height: 16px">
                        <td>
                            <asp:Label ID="Label24" runat="server" CssClass="columnTitle" Width="100%"></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="Label25" runat="server" CssClass="columnTitle" Width="100%"></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="Label26" runat="server" CssClass="columnTitle" Width="100%"></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="Label27" runat="server" CssClass="columnTitle" Width="100%"></asp:Label>
                        </td>
                        <td>
                            <asp:Label ID="Label28" runat="server" CssClass="columnTitle" Width="100%"></asp:Label>
                        </td>
                        <td>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <cc1:CtlNumText ID="txt_Rate" runat="server" ViewStateMode="Enabled" CssClass="textNum"
                                DecLen="3" DispComma="True"></cc1:CtlNumText>
                        </td>
                        <td>
                            <cc1:CtlNumText ID="txt_UnitPrc" runat="server" ViewStateMode="Enabled" CssClass="textNum"
                                DecLen="2" DispComma="True"></cc1:CtlNumText>
                        </td>
                        <td>
                            <cc1:CtlNumText ID="txt_Amount" runat="server" ViewStateMode="Enabled" CssClass="textNum"></cc1:CtlNumText>
                        </td>
                        <td>
                            <cc1:CtlNumText ID="txt_Price" runat="server" ViewStateMode="Enabled" CssClass="textNum"
                                ReadOnly="True"></cc1:CtlNumText>
                        </td>
                        <td>
                            <cc1:CtlNumText ID="txt_Total" runat="server" ViewStateMode="Enabled" CssClass="textNum"
                                ReadOnly="True"></cc1:CtlNumText>
                        </td>
                        <td>
                        </td>
                    </tr>
                    <tr style="height: 27px">
                        <td>
                        </td>
                        <td>
                            <cc1:CtlNumText ID="txt_DtlPrc" runat="server" Height="15px" Font-Bold="True" Width="94px"
                                ViewStateMode="Enabled" ReadOnly="True" DecLen="3" DispComma="True"></cc1:CtlNumText>
                        </td>
                        <td>
                        </td>
                        <td>
                            <cc1:CtlNumText ID="txt_Tax" runat="server" Height="15px" Font-Bold="True" Width="94px"
                                ViewStateMode="Enabled" ReadOnly="True"></cc1:CtlNumText>
                        </td>
                        <td colspan="2">
                        </td>
                    </tr>
                    <tr>
                        <td colspan="6">
                            <asp:TextBox ID="SelUnitValue" runat="server" Height="0px" Width="0px" Style="border: 0px;
                                padding: 0px;" Wrap="False"></asp:TextBox>
                            <asp:TextBox ID="SelCurrValue" runat="server" Height="0px" Style="border-right: 0px;
                                padding-right: 0px; border-top: 0px; padding-left: 0px; padding-bottom: 0px;
                                border-left: 0px; padding-top: 0px; border-bottom: 0px" Width="3px" Wrap="False"></asp:TextBox>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</div>
<asp:HiddenField ID="HidSelRowID" runat="server" />
