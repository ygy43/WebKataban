<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUC_Tanka.ascx.vb"
    Inherits="WebKataban.WebUC_Tanka" %>
<%@ Register Assembly="WebKataban" Namespace="WebKataban" TagPrefix="cc1" %>
<div class="tanka">
    <asp:Panel ID="pnlMain" runat="server" Height="580px" CssClass="mainContainer">
        <center>
            <div style="width: 98%;">
                <asp:Panel ID="PnlHead" runat="server">
                    <div class="title">
                        <div>
                            <asp:Label ID="lblSeriesNm" runat="server" ViewStateMode="Enabled"></asp:Label>
                        </div>
                        <div>
                            <asp:Label ID="lblSeriesKat" runat="server" ViewStateMode="Enabled"></asp:Label>
                        </div>
                    </div>
                </asp:Panel>
            </div>
        </center>
        <asp:Panel ID="Panel2" runat="server">
            <center>
                <div class="additionInfo">
                    <table style="width: 100%;">
                        <tr>
                            <td class="additionCol10 center">
                            </td>
                            <td class="additionCol20 center">
                                <asp:Label ID="Label1" runat="server" ViewStateMode="Enabled" CssClass="additionLabel"></asp:Label>
                            </td>
                            <td class="additionCol20 center">
                                <asp:Label ID="Label2" runat="server" ViewStateMode="Enabled" CssClass="additionLabel"></asp:Label>
                            </td>
                            <td class="additionCol20 center">
                                <asp:Label ID="Label3" runat="server" ViewStateMode="Enabled" CssClass="additionLabel"></asp:Label>
                            </td>
                            <td class="additionCol15 center">
                                <asp:Label ID="Label4" runat="server" ViewStateMode="Enabled" CssClass="additionLabel"></asp:Label>
                            </td>
                            <td class="additionCol15 center">
                                <asp:Label ID="Label5" runat="server" ViewStateMode="Enabled" CssClass="additionLabel"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="center">
                            </td>
                            <td class="ClsChk center" style="border: 1px solid #C0C0C0; " bgcolor="White" width="14%">
                            <asp:TextBox ID="lblCheckZ" runat="server" ReadOnly="True" ViewStateMode="Enabled" 
                                    TabIndex="-1" CssClass="additionText" BorderStyle="None" Text="Z" Width="35%"></asp:TextBox><asp:TextBox ID="lblCheck" runat="server" ReadOnly="True" ViewStateMode="Enabled" TabIndex="-1" CssClass="additionText" 
                                    BorderStyle="None" Width="35%"></asp:TextBox>
                            </td>
                            <td class="center">
                                <asp:DropDownList ID="cmbPlace" runat="server" Width="80%" AutoPostBack="True" DataTextField="PlaceName"
                                    DataValueField="PlaceID" CssClass="additionText">
                                </asp:DropDownList>
                            </td>
                            <td class="center">
                                <asp:TextBox ID="lblNouki" runat="server" ReadOnly="True" ViewStateMode="Enabled"
                                    Width="90%" TabIndex="-1" CssClass="additionText">納期お問い合わせ下さい</asp:TextBox>
                            </td>
                            <td class="center">
                                <asp:TextBox ID="lblKosuu" runat="server" ReadOnly="True" ViewStateMode="Enabled"
                                    Width="80%" TabIndex="-1" CssClass="additionText"></asp:TextBox>
                            </td>
                            <td class="center">
                                <asp:TextBox ID="lblEL" runat="server" ReadOnly="True" ViewStateMode="Enabled" Width="80%"
                                    TabIndex="-1" CssClass="additionText"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td style="text-align: center;" colspan="2">
                                <asp:Label ID="Label7" runat="server" ViewStateMode="Enabled" Visible="False" CssClass="placeText"></asp:Label>
                                <asp:Label ID="Label10" runat="server" ViewStateMode="Enabled" Visible="False" CssClass="placeText"></asp:Label>
                                <asp:Label ID="Label11" runat="server" ViewStateMode="Enabled" Visible="False" CssClass="placeText"></asp:Label>
                            </td>
                            <td class="center">
                             <asp:DropDownList ID="cmbStrageEvaluation" runat="server" Width="80%" AutoPostBack="True" DataTextField="StrageEvaluationName"
                                    DataValueField="StrageEvaluationID" CssClass="additionText">
                                </asp:DropDownList>
                                <br/>
                                <asp:Label ID="Label8" runat="server" ViewStateMode="Enabled" CssClass="placeMessage"></asp:Label>
                            </td>
                            <td style="text-align: center;" colspan="2">
                                <asp:Button ID="Button10" runat="server" CssClass="button" Visible="false" />
                            </td>
                            <td style="text-align: center;">
                                <asp:Button ID="Button11" runat="server" CssClass="button" Visible="false" />
                            </td>
                        </tr>
                        <tr>
                            <td style="text-align: left;" colspan="2">
                                <asp:Button ID="Button13" runat="server" CssClass="button" Visible="false" />
                            </td>
                        </tr>
                    </table>
                </div>
                <asp:Panel ID="Panel12" runat="server">
                    <table width="98%">
                        <tr>
                            <td align="left">
                                <asp:Label ID="lblAction" runat="server" Visible="False" ViewStateMode="Enabled"
                                    CssClass="message1">
                                    ＜注意事項＞標準納期は生産状況により実際の納期と異なる場合がございます。適用個数を超える場合は別途納期をご相談ください。
                                </asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="left">
                                <asp:Label ID="Label9" runat="server" Visible="False" ViewStateMode="Enabled" CssClass="message2"></asp:Label>
                                <asp:Label ID="Label27" runat="server" Visible="False" ViewStateMode="Enabled" CssClass="message2"></asp:Label>
                                <asp:Label ID="Label6" runat="server" Visible="False" ViewStateMode="Enabled" CssClass="message2"></asp:Label>
                                <asp:Label ID="Label36" runat="server" CssClass="message2" ViewStateMode="Enabled"
                                    Visible="False"></asp:Label>
                                <asp:Label ID="Label37" runat="server" CssClass="message2" ViewStateMode="Enabled"
                                    Visible="False"></asp:Label>
                                <asp:Label ID="Label40" runat="server" CssClass="message2" ViewStateMode="Enabled"
                                    Visible="False"></asp:Label>
                                <asp:Label ID="Label42" runat="server" Visible="False" ViewStateMode="Enabled" CssClass="message2"></asp:Label>
                                <asp:Label ID="Label43" runat="server" Visible="False" ViewStateMode="Enabled" CssClass="message2" Font-Size="10pt"></asp:Label>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
            </center>
        </asp:Panel>
        <div style="width: 100%">
            <center>
                <table width="98%">
                    <tr>
                        <td class="tableColPrice">
                            <asp:Panel ID="pnlPriceList" runat="server">
                                <asp:GridView ID="GVPrice" runat="server" AutoGenerateColumns="False" CellPadding="2"
                                    Width="100%" TabIndex="1">
                                    <Columns>
                                        <asp:BoundField DataField="Kubun">
                                            <ItemStyle Height="22px" HorizontalAlign="Left" Width="50%" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="ViewPrice">
                                            <ItemStyle Height="22px" HorizontalAlign="Right" Width="50%" />
                                        </asp:BoundField>
                                        <asp:TemplateField ItemStyle-CssClass="displaynone">
                                            <ItemTemplate>
                                                <asp:HiddenField runat="server" ID="ColumnKbn" Value='<%#Eval("ColumnKBN") %>'></asp:HiddenField>
                                            </ItemTemplate>
                                            <HeaderStyle CssClass="displaynone" />
                                            <ItemStyle CssClass="displaynone" />
                                        </asp:TemplateField>
                                    </Columns>
                                    <HeaderStyle CssClass="gridHeader" />
                                    <AlternatingRowStyle CssClass="gridAltRow" />
                                    <RowStyle CssClass="gridRow" />
                                </asp:GridView>
                                <asp:Label ID="lblQtyUnit1" runat="server" CssClass="qtyUnit"></asp:Label>
                            </asp:Panel>
                            <div class="buttonDiv">
                                <asp:Button ID="Button2" runat="server" CssClass="button" TabIndex="4" />
                                <asp:Button ID="Button3" runat="server" CssClass="button" TabIndex="5" Style="margin-right: 0px" />
                                <asp:Button ID="Button6" runat="server" OnClientClick="f_Tanka_File('ctl00_ContentDetail_WebUC_Tanka_')"
                                    CssClass="button" TabIndex="6" />
                                <asp:Button ID="Button9" runat="server" CssClass="button" TabIndex="7" Visible="false" />
                                <asp:Button ID="Button12" runat="server" CssClass="button" TabIndex="8" />
                            </div>
                        </td>
                        <td class="tableColInfo">
                            <asp:Panel ID="PnlInput" runat="server">
                                <div class="calculate">
                                    <table width="100%" cellpadding="5px">
                                        <tr>
                                            <td>
                                                <asp:Label ID="Label14" runat="server" ViewStateMode="Enabled" CssClass="calculateLabel"></asp:Label>
                                            </td>
                                            <td>
                                                <cc1:CtlNumText ID="txt_Rate" runat="server" IntLen="1" ViewStateMode="Enabled" TabIndex="1"
                                                    CssClass="calculateText"></cc1:CtlNumText>
                                            </td>
                                            <td>
                                                <asp:Label ID="Label15" runat="server" ViewStateMode="Enabled" CssClass="calculateLabel"></asp:Label>
                                            </td>
                                            <td>
                                                <cc1:CtlNumText ID="TextMoney" runat="server" AllowZero="False" DispComma="False"
                                                    TabIndex="-1" ViewStateMode="Enabled" CssClass="calculateTextReadOnly"></cc1:CtlNumText>
                                            </td>
                                        </tr>
                                        <tr style="width: 100%;">
                                            <td>
                                                <asp:Label ID="Label16" runat="server" ViewStateMode="Enabled" CssClass="calculateLabel"></asp:Label>
                                            </td>
                                            <td>
                                                <cc1:CtlNumText ID="TextUnitPrice" runat="server" AllowZero="True" DispComma="True"
                                                    TabIndex="2" ViewStateMode="Enabled" CssClass="calculateText"></cc1:CtlNumText>
                                            </td>
                                            <td>
                                                <asp:Label ID="Label17" runat="server" ViewStateMode="Enabled" CssClass="calculateLabel"></asp:Label>
                                            </td>
                                            <td>
                                                <cc1:CtlNumText ID="TextTax" runat="server" AllowZero="False" DispComma="False" TabIndex="-1"
                                                    ViewStateMode="Enabled" CssClass="calculateTextReadOnly"></cc1:CtlNumText>
                                            </td>
                                        </tr>
                                        <tr style="width: 100%; height: 30px">
                                            <td>
                                            </td>
                                            <td>
                                                <cc1:CtlNumText ID="TextRateUnitPrice" runat="server" AllowZero="False" DispComma="False"
                                                    TabIndex="-1" ViewStateMode="Enabled" CssClass="calculateTextReadOnly"></cc1:CtlNumText>
                                            </td>
                                            <td colspan="2">
                                            </td>
                                        </tr>
                                        <tr style="width: 100%;">
                                            <td>
                                                <asp:Label ID="Label18" runat="server" ViewStateMode="Enabled" CssClass="calculateLabel"></asp:Label>
                                            </td>
                                            <td>
                                                <cc1:CtlNumText ID="TextCnt" runat="server" DispComma="True" ViewStateMode="Enabled"
                                                    TabIndex="3" CssClass="calculateText"></cc1:CtlNumText>
                                            </td>
                                            <td>
                                                <asp:Label ID="Label19" runat="server" ViewStateMode="Enabled" CssClass="calculateLabel"></asp:Label>
                                            </td>
                                            <td>
                                                <cc1:CtlNumText ID="TextAmount" runat="server" AllowZero="False" DispComma="False"
                                                    TabIndex="-1" ViewStateMode="Enabled" CssClass="calculateTextReadOnly"></cc1:CtlNumText>
                                            </td>
                                        </tr>
                                    </table>
                                    <div>
                                        <asp:Label ID="lblQtyUnit" runat="server" CssClass="qtyUnit"></asp:Label>
                                    </div>
                                </div>
                                <asp:Panel ID="PnlSelect" runat="server">
                                    <div class="select">
                                        <table width="100%">
                                            <tr>
                                                <td style="text-align: left;">
                                                    <asp:Panel ID="Pnl10" runat="server" Visible="False">
                                                        <asp:Label ID="Label21" runat="server" ViewStateMode="Enabled" CssClass="message1"></asp:Label>
                                                    </asp:Panel>
                                                    <asp:Panel ID="Pnl14" runat="server" Visible="False">
                                                        <asp:Label ID="Label25" runat="server" ViewStateMode="Enabled" CssClass="message1"></asp:Label>
                                                    </asp:Panel>
                                                    <asp:Panel ID="Pnl11" runat="server" Visible="False">
                                                        <asp:Label ID="Label22" runat="server" ViewStateMode="Enabled" CssClass="message1"></asp:Label>
                                                    </asp:Panel>
                                                    <asp:Panel ID="Pnl15" runat="server" Visible="False">
                                                        <asp:Label ID="Label26" runat="server" ViewStateMode="Enabled" CssClass="message1"></asp:Label>
                                                    </asp:Panel>
                                                    <asp:Panel ID="Pnl17" runat="server" Visible="False">
                                                        <asp:Label ID="Label28" runat="server" ViewStateMode="Enabled" CssClass="message1"></asp:Label>
                                                    </asp:Panel>
                                                    <asp:Panel ID="Pnl18" runat="server" Visible="False">
                                                        <asp:Label ID="Label29" runat="server" ViewStateMode="Enabled" CssClass="message1"></asp:Label>
                                                    </asp:Panel>
                                                    <asp:Panel ID="Pnl19" runat="server" Visible="False">
                                                        <asp:Label ID="Label30" runat="server" ViewStateMode="Enabled" CssClass="message1"></asp:Label>
                                                    </asp:Panel>
                                                    <asp:Panel ID="Pnl20" runat="server" Visible="False">
                                                        <asp:Label ID="Label31" runat="server" ViewStateMode="Enabled" CssClass="message1"></asp:Label>
                                                    </asp:Panel>
                                                    <asp:Panel ID="Pnl21" runat="server" Visible="False">
                                                        <asp:Label ID="Label32" runat="server" ViewStateMode="Enabled" CssClass="message1"></asp:Label>
                                                    </asp:Panel>
                                                    <asp:Panel ID="Pnl22" runat="server" Visible="False">
                                                        <asp:Label ID="Label33" runat="server" ViewStateMode="Enabled" CssClass="message1"></asp:Label>
                                                    </asp:Panel>
                                                    <asp:Panel ID="Pnl23" runat="server" Visible="False">
                                                        <asp:Label ID="Label34" runat="server" ViewStateMode="Enabled" CssClass="message1"></asp:Label>
                                                    </asp:Panel>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="width: 100%;">
                                                    <asp:Panel ID="Panel10" runat="server" Width="100%">
                                                        <table width="100%">
                                                            <tr>
                                                                <td style="width: 25%;">
                                                                    <asp:Label ID="Label23" runat="server" ViewStateMode="Enabled" CssClass="calculateLabel"></asp:Label>
                                                                </td>
                                                                <td style="width: 25%;">
                                                                    <asp:TextBox ID="txtSelKosu" runat="server" ReadOnly="True" TabIndex="-1" CssClass="calculateText"></asp:TextBox>
                                                                </td>
                                                                <td style="width: 20%;">
                                                                    <asp:Label ID="Label24" runat="server" ViewStateMode="Enabled" CssClass="calculateLabel"></asp:Label>
                                                                </td>
                                                                <td style="width: 25%;">
                                                                    <asp:TextBox ID="txtSelNoki" runat="server" ReadOnly="True" TabIndex="-1" CssClass="calculateText"></asp:TextBox>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </asp:Panel>
                                                </td>
                                            </tr>
                                        </table>
                                    </div>
                                </asp:Panel>
                                 
                                    <div style="text-align: left;">
                                        <asp:Label ID="Label41" runat="server" ViewStateMode="Enabled" CssClass="message1"></asp:Label>
                                    </div>
                             
                                <asp:Panel ID="Pnl24" runat="server" Visible="False">
                                    <div style="text-align: left;">
                                        <asp:Label ID="Label35" runat="server" ViewStateMode="Enabled" CssClass="message1"></asp:Label>
                                    </div>
                                </asp:Panel>
                                <asp:Panel ID="Pnl25" runat="server" Visible="False">
                                    <div style="text-align: left;">
                                        <asp:Label ID="Label38" runat="server" ViewStateMode="Enabled" CssClass="message1"></asp:Label>
                                    </div>
                                </asp:Panel>
                                <asp:Panel ID="Pnl26" runat="server" Visible="False">
                                    <div style="text-align: left;">
                                        <asp:Label ID="Label39" runat="server" ViewStateMode="Enabled" CssClass="message1"></asp:Label>
                                    </div>
                                </asp:Panel>
                            </asp:Panel>
                        </td>
                    </tr>
                    <tr>
                        <td>
                        </td>
                        <td align="right">
                            <asp:Button ID="Button5" runat="server" CssClass="button" TabIndex="7" />
                        </td>
                    </tr>
                </table>
            </center>
        </div>
    </asp:Panel>
    <asp:HiddenField ID="HdnSetYFlg" runat="server" />
    <asp:HiddenField ID="HidShiftD" runat="server" />
    <asp:HiddenField ID="SelUnitValue" runat="server" Value="0" />
    <asp:HiddenField ID="SelCurrValue" runat="server" />
    <asp:HiddenField ID="HidSelRowID" runat="server" />
    <asp:HiddenField ID="HidPriceForFile" runat="server" />
    <asp:HiddenField ID="HidPriceList" runat="server" />
    <asp:HiddenField ID="HidNewPlace" runat="server" />
    <asp:HiddenField ID="HidPriceDetail" runat="server" />
    <asp:HiddenField ID="HidJsonData" runat="server" />
    <div style="height: 0px;">
        <asp:TextBox ID="txt_EditNormal" runat="server" Height="0px" Style="border: 0px; padding: 0px;"
            Width="0px" Wrap="False"></asp:TextBox>
        <asp:LinkButton ID="btnCopy" runat="server" Height="0px" Width="0px" />
        <asp:LinkButton ID="btnClear" runat="server" Height="0px" Width="0px" />
        <asp:LinkButton ID="btnClick" runat="server" Height="0px" Width="0px" />
    </div>
</div>
