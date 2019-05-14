<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUC_ISOTanka.ascx.vb"
    Inherits="WebKataban.WebUC_ISOTanka" %>
<%@ Register Assembly="WebKataban" Namespace="WebKataban" TagPrefix="cc1" %>
<div class="isoTanka">
    <asp:Panel ID="pnlMain" runat="server" Height="700px" CssClass="mainContainer">
        <center>
            <div class="title" style="width: 98%;">
                <div>
                    <asp:Label ID="lblSeriesNm" runat="server" ViewStateMode="Enabled"></asp:Label>
                </div>
                <div>
                    <asp:Label ID="lblSeriesKat" runat="server" ViewStateMode="Enabled"></asp:Label>
                </div>
            </div>
            <asp:Panel ID="Panel2" runat="server" CssClass="fixedMessage">
                <asp:Image ID="imgFixedMessage1" runat="server" ImageUrl="~/KHImage/FixedMessage2.gif" />
            </asp:Panel>
            <asp:Panel ID="PnlProductionPlace" runat="server">
                <table style="width: 100%; text-align: center;">
                    <tr>
                        <td width="30%" rowspan="2" style="vertical-align: bottom;">
                            <asp:Label ID="Label15" runat="server" CssClass="placeText"></asp:Label>
                            <asp:Label ID="Label19" runat="server" CssClass="placeText"></asp:Label>
                        </td>
                        <td width="20%">
                            <asp:Label ID="Label22" runat="server" CssClass="additionLabel"></asp:Label>
                        </td>
                        <td width="50%">
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:DropDownList ID="cmbPlace" runat="server" AutoPostBack="True" DataTextField="PlaceName"
                                DataValueField="PlaceID" CssClass="additionText">
                            </asp:DropDownList>
                        </td>
                        <td style="text-align: left;">
                            <asp:Label ID="Label16" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3" style="text-align: left;">
                            <asp:Label ID="Label17" runat="server"></asp:Label>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Panel ID="PnlDetail" runat="server" CssClass="contents">
                <table cellpadding="1" cellspacing="0" rules="all" border="0" style="text-align: center;
                    margin-bottom: 5px; width: 100%;">                    
                
                        <tr style="background-color: Green; height: 20px;">
                        <td style="background-color: #C7EDCC;">
                        </td>
                        <td style="width: 14%;">
                            <asp:Label ID="Label9" runat="server" CssClass="columnTitle labelTitle"></asp:Label>
                        </td>
                        <td style="width: 14%;">
                            <asp:Label ID="Label14" runat="server" CssClass="columnTitle labelTitle"></asp:Label>
                        </td>
                        <td style="width: 14%;">
                            <asp:Label ID="Label10" runat="server" CssClass="columnTitle labelTitle"></asp:Label>
                        </td>
                        <td style="width: 14%;">
                            <asp:Label ID="Label11" runat="server" CssClass="columnTitle labelTitle" Visible="False"></asp:Label>
                        </td>
                        <td style="width: 14%;">
                            <asp:Label ID="Label12" runat="server" CssClass="columnTitle labelTitle" Visible="False"></asp:Label>
                        </td>
                        <td style="width: 14%;">
                            <asp:Label ID="Label13" runat="server" CssClass="columnTitle labelTitle" Visible="False"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td style="background-color: Green;" align="center">
                            <div style="line-height: 22px;">
                                <asp:Label ID="Label8" runat="server" CssClass="columnTitle labelTitle"></asp:Label>
                            </div>
                        </td>
                        <td>
                            <cc1:CtlNumText ID="txt_AmtPrice" runat="server" ViewStateMode="Enabled" CssClass="numText right"
                                DecLen="2" DispComma="True" ReadOnly="True"></cc1:CtlNumText>
                        </td>
                        <td>
                            <cc1:CtlNumText ID="txt_AmtTax" runat="server" ViewStateMode="Enabled" CssClass="numText right"
                                DecLen="2" DispComma="True" ReadOnly="True"></cc1:CtlNumText>
                        </td>
                        <td>
                            <cc1:CtlNumText ID="txt_SumTotal" runat="server" ViewStateMode="Enabled" CssClass="numText right"
                                DecLen="2" DispComma="True" ReadOnly="True"></cc1:CtlNumText>
                        </td>
                        <td>
                            <cc1:CtlNumText ID="txt_DelDate" runat="server" ViewStateMode="Enabled" CssClass="numText"
                                Visible="False" ReadOnly="True"></cc1:CtlNumText>
                        </td>
                        <td>
                            <cc1:CtlNumText ID="txt_PrpAmt" runat="server" ViewStateMode="Enabled" CssClass="numText center"
                                Visible="False" ReadOnly="True"></cc1:CtlNumText>
                        </td>
                        <td>
                            <cc1:CtlNumText ID="txt_ELPrd" runat="server" ViewStateMode="Enabled" CssClass="numText center"
                                Visible="False" ReadOnly="True"></cc1:CtlNumText>
                        </td>
                    </tr>
                </table>
                <div class="priceList" id="priceList">
                    <asp:Panel ID="PnlTankaList" runat="server">
                    </asp:Panel>
                </div>
                <asp:Label ID="Label41" runat="server" ViewStateMode="Enabled" CssClass="message1"></asp:Label>
                <asp:Panel ID="pnlIndMessage" runat="server" Visible="false">
                    <asp:Label ID="Label32" runat="server" ViewStateMode="Enabled" CssClass="message1"></asp:Label>
                </asp:Panel>
                <asp:Panel ID="pnlEurMessage" runat="server" Visible="false">
                    <asp:Label ID="Label33" runat="server" ViewStateMode="Enabled" CssClass="message1"></asp:Label>
                </asp:Panel>
            </asp:Panel>
        </center>
    </asp:Panel>
    <asp:Panel ID="Panel16" runat="server" Height="30px" Width="100%" BackColor="#C7EDCC"
        HorizontalAlign="Center" BorderStyle="None">
        <center>
            <asp:Panel ID="Panel6" runat="server" Width="98%" Height="40px" HorizontalAlign="Left">
                <asp:Button ID="Button2" runat="server" CssClass="button" OnClientClick="f_ISOTanka_File('ctl00_ContentDetail_WebUC_ISOTanka_')" />
                <asp:Button ID="Button3" runat="server" CssClass="button" OnClientClick="f_ISOTanka_File('ctl00_ContentDetail_WebUC_ISOTanka_')" />
                <asp:Button ID="Button5" runat="server" CssClass="button" OnClientClick="f_ISOTanka_File('ctl00_ContentDetail_WebUC_ISOTanka_')" />
                <asp:Button ID="Button6" runat="server" CssClass="button" OnClientClick="f_ISOTanka_File('ctl00_ContentDetail_WebUC_ISOTanka_')" />
                <asp:Button ID="Button9" runat="server" CssClass="button" Visible="false" />
            </asp:Panel>
        </center>
    </asp:Panel>
    <asp:HiddenField ID="strHiddenKbn" runat="server" Value="" />
    <asp:HiddenField ID="intItemRow" runat="server" Value="-1" />
    <asp:HiddenField ID="intItemCnt" runat="server" Value="" />
    <asp:HiddenField ID="intListRowCnt" runat="server" Value="" />
    <asp:HiddenField ID="HidShiftD" runat="server" />
    <asp:HiddenField ID="HidPriceForFile" runat="server" />
    <asp:HiddenField ID="HidPriceList" runat="server" />
    <asp:HiddenField ID="HidPriceDetail" runat="server" />
    <div style="height: 0px;">
        <asp:TextBox ID="txt_EditNormal" runat="server" Height="0px" Style="border: 0px;
            padding: 0px;" Width="0px" Wrap="False" />
        <asp:LinkButton ID="btnCopy" runat="server" Height="0px" Width="0px" />
    </div>
</div>
