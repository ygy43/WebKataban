<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUC_OutOfOption.ascx.vb"
    Inherits="WebKataban.WebUC_OutOfOption" %>
<div class="outOfOption">
    <asp:Panel ID="pnlMain" runat="server" Height="610px" CssClass="mainContainer">
        <div class="title" style="width: 98%">
            <asp:Label ID="lblSeriesNm" runat="server" ViewStateMode="Enabled"></asp:Label>
        </div>
        <div style="width: 98%; height: 550px; overflow-y: scroll;">
            <asp:Panel ID="pnlPortCushon" runat="server" Visible="true" Width="98%" HorizontalAlign="Left">
                <table style="width: 100%;">
                    <tr>
                        <td style="width: 40%;">
                            <asp:Label ID="Label1" runat="server" CssClass="titleLabel" BorderStyle="None" Width="90%"></asp:Label>
                        </td>
                        <td style="width: 30%;">
                            <asp:DropDownList ID="cmbPortCushon" runat="server" CssClass="dropDown">
                            </asp:DropDownList>
                        </td>
                        <td>
                            <asp:TextBox ID="txtPortCuchon" runat="server" ReadOnly="true" CssClass="textReadOnly"></asp:TextBox>
                        </td>
                    </tr>
                </table>
                <table style="width: 100%;">
                    <tr>
                        <td rowspan="2">
                            <asp:Image ID="img1" runat="server" ImageUrl="~/KHImage/outOp1.gif" />
                        </td>
                        <td style="width: 20%">
                        </td>
                        <td style="width: 50%">
                            <asp:Label ID="Label2" runat="server" CssClass="defaultLabel"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 20%; vertical-align: bottom;">
                            <table>
                                <tr>
                                    <td>
                                        <asp:Label ID="Label15" runat="server" CssClass="defaultLabel"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Label ID="Label16" runat="server" CssClass="defaultLabel"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Label ID="Label17" runat="server" CssClass="defaultLabel"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Label ID="Label18" runat="server" CssClass="defaultLabel"></asp:Label>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td style="width: 50%">
                            <table class="dispTable">
                                <tr>
                                    <td colspan="2" style="text-align: center">
                                        <asp:Label ID="Label3" runat="server" CssClass="defaultLabel"></asp:Label>
                                    </td>
                                    <td style="width: 7%">
                                        <asp:Label ID="Label4" Text="１" runat="server" CssClass="defaultLabel"></asp:Label>
                                    </td>
                                    <td style="width: 7%">
                                        <asp:Label ID="Label5" Text="２" runat="server" CssClass="defaultLabel"></asp:Label>
                                    </td>
                                    <td style="width: 7%">
                                        <asp:Label ID="Label6" Text="３" runat="server" CssClass="defaultLabel"></asp:Label>
                                    </td>
                                    <td style="width: 7%">
                                        <asp:Label ID="Label7" Text="４" runat="server" CssClass="defaultLabel"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td rowspan="2" style="width: 15%;">
                                        <asp:Label ID="Label8" runat="server" CssClass="defaultLabel"></asp:Label>
                                    </td>
                                    <td>
                                        <asp:Label ID="Label9" runat="server" CssClass="defaultLabel"></asp:Label>
                                    </td>
                                    <td style="width: 7%;">
                                        <asp:RadioButton ID="rdoRK1" GroupName="rdoGroupRK" runat="server" Text="" />
                                    </td>
                                    <td style="width: 7%;">
                                        <asp:RadioButton ID="rdoRK2" GroupName="rdoGroupRK" runat="server" Text="" />
                                    </td>
                                    <td style="width: 7%;">
                                        <asp:RadioButton ID="rdoRK3" GroupName="rdoGroupRK" runat="server" Text="" />
                                    </td>
                                    <td style="width: 7%;">
                                        <asp:RadioButton ID="rdoRK4" GroupName="rdoGroupRK" runat="server" Text="" />
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Label ID="Label10" runat="server" CssClass="defaultLabel"></asp:Label>
                                    </td>
                                    <td style="width: 7%;">
                                        <asp:RadioButton ID="rdoRC1" GroupName="rdoGroupRC" runat="server" />
                                    </td>
                                    <td style="width: 7%;">
                                        <asp:RadioButton ID="rdoRC2" GroupName="rdoGroupRC" runat="server" Text="" />
                                    </td>
                                    <td style="width: 7%;">
                                        <asp:RadioButton ID="rdoRC3" GroupName="rdoGroupRC" runat="server" Text="" />
                                    </td>
                                    <td style="width: 7%;">
                                        <asp:RadioButton ID="rdoRC4" GroupName="rdoGroupRC" runat="server" Text="" />
                                    </td>
                                </tr>
                                <tr>
                                    <td rowspan="2" style="width: 15%;">
                                        <asp:Label ID="Label11" runat="server" CssClass="defaultLabel"></asp:Label>
                                    </td>
                                    <td>
                                        <asp:Label ID="Label12" runat="server" CssClass="defaultLabel"></asp:Label>
                                    </td>
                                    <td style="width: 7%;">
                                        <asp:RadioButton ID="rdoHK1" GroupName="rdoGroupHK" runat="server" Text="" />
                                    </td>
                                    <td style="width: 7%;">
                                        <asp:RadioButton ID="rdoHK2" GroupName="rdoGroupHK" runat="server" Text="" />
                                    </td>
                                    <td style="width: 7%;">
                                        <asp:RadioButton ID="rdoHK3" GroupName="rdoGroupHK" runat="server" Text="" />
                                    </td>
                                    <td style="width: 7%;">
                                        <asp:RadioButton ID="rdoHK4" GroupName="rdoGroupHK" runat="server" Text="" />
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Label ID="Label13" runat="server" CssClass="defaultLabel"></asp:Label>
                                    </td>
                                    <td style="width: 7%;">
                                        <asp:RadioButton ID="rdoHC1" GroupName="rdoGroupCK" runat="server" Text="" />
                                    </td>
                                    <td style="width: 7%;">
                                        <asp:RadioButton ID="rdoHC2" GroupName="rdoGroupCK" runat="server" Text="" />
                                    </td>
                                    <td style="width: 7%;">
                                        <asp:RadioButton ID="rdoHC3" GroupName="rdoGroupCK" runat="server" Text="" />
                                    </td>
                                    <td style="width: 7%;">
                                        <asp:RadioButton ID="rdoHC4" GroupName="rdoGroupCK" runat="server" Text="" />
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="6" style="width: 100%">
                                        <asp:Label ID="Label14" runat="server" CssClass="defaultLabel"></asp:Label>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            <asp:Label ID="Label19" runat="server" CssClass="defaultLabel"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            <asp:Image ID="Image2" runat="server" ImageUrl="~/KHImage/outOp2.gif" />
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Panel ID="pnlPort" runat="server" Width="98%" HorizontalAlign="Left">
                <table style="width: 100%;">
                    <tr>
                        <td style="width: 40%;">
                            <asp:Label ID="Label20" runat="server" CssClass="titleLabel" Width="90%"></asp:Label>
                        </td>
                        <td style="width: 60%;">
                            <asp:DropDownList ID="cmbPort" runat="server" CssClass="dropDown">
                            </asp:DropDownList>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Panel ID="pnlPortSize" runat="server" Width="98%" HorizontalAlign="Left">
                <table style="width: 100%;">
                    <tr>
                        <td style="width: 40%;">
                            <asp:Label ID="Label21" runat="server" CssClass="titleLabel" Width="90%"></asp:Label>
                        </td>
                        <td style="width: 60%;">
                            <asp:DropDownList ID="cmbPortSize" runat="server" CssClass="dropDown">
                            </asp:DropDownList>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Panel ID="pnlMounting" runat="server" Width="98%" HorizontalAlign="Left">
                <table style="width: 100%;">
                    <tr>
                        <td style="width: 40%;">
                            <asp:Label ID="Label22" runat="server" CssClass="titleLabel" Width="90%"></asp:Label>
                        </td>
                        <td style="width: 60%;">
                            <asp:DropDownList ID="cmbMounting" runat="server" CssClass="dropDown">
                            </asp:DropDownList>
                        </td>
                    </tr>
                </table>
                <table style="width: 100%;">
                    <tr>
                        <td>
                            <asp:Label ID="Label23" runat="server" CssClass="defaultLabel"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:Image ID="Image1" runat="server" ImageUrl="../KHImage/outOp3.gif" />
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Panel ID="pnlTrunnion" runat="server" Width="98%" HorizontalAlign="Left">
                <table style="width: 100%;">
                    <tr>
                        <td style="width: 40%;">
                            <asp:Label ID="Label24" Text="◆トラニオン位置指定" runat="server" CssClass="titleLabel" BorderStyle="None"
                                Width="90%"></asp:Label>
                        </td>
                        <td style="width: 5%;">
                            <asp:Label ID="Label25" Text="AQ" runat="server" Font-Size="14pt" CssClass="titleLabel"
                                Width="100%" BorderStyle="None"></asp:Label>
                        </td>
                        <td>
                            <asp:TextBox ID="txtTrunnion" runat="server" Width="35%" CssClass="textBox"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            <asp:Image ID="Image3" runat="server" ImageUrl="~/KHImage/outOp4.gif" />
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Panel ID="pnlClevis" runat="server" Width="98%" HorizontalAlign="Left">
                <table style="width: 100%;">
                    <tr>
                        <td style="width: 40%;">
                            <asp:Label ID="Label26" Text="◆二山ナックル・二山クレビス" runat="server" CssClass="titleLabel"
                                Width="90%" BorderStyle="None"></asp:Label>
                        </td>
                        <td>
                            <asp:DropDownList ID="cmbClevis" runat="server" CssClass="dropDown">
                            </asp:DropDownList>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Panel ID="pnlTieRod" runat="server" Width="98%" HorizontalAlign="Left">
                <table style="width: 100%;">
                    <tr>
                        <td style="width: 40%;">
                            <asp:Label ID="Label27" Text="◆タイロッド延長寸法" runat="server" CssClass="titleLabel" BorderStyle="None"
                                Width="90%"></asp:Label>
                        </td>
                        <td>
                        </td>
                    </tr>
                </table>
                <asp:Table ID="tblTieRod" runat="server" CssClass="dispTable" Width="100%">
                    <asp:TableHeaderRow ID="TableHeaderRow1" runat="server">
                        <asp:TableCell ID="TableCell1" Width="5%" CssClass="DispTable" runat="server"></asp:TableCell>
                        <asp:TableCell ID="TableCell2" Width="15%" CssClass="DispTable" runat="server">
                            <asp:Label ID="Label28" Text="記号" runat="server" CssClass="defaultLabel"></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell ID="TableCell3" Width="40%" CssClass="DispTable" runat="server">
                            <asp:Label ID="Label29" Text="内容" runat="server" CssClass="defaultLabel"></asp:Label>
                        </asp:TableCell>
                        <%--<asp:TableCell ID="TableCell4" Width="2%" runat="server"></asp:TableCell>--%>
                        <asp:TableCell ID="TableCell5" Width="38%" RowSpan="8" runat="server" HorizontalAlign="Center">
                            <asp:Image ID="Image5" runat="server" ImageUrl="~/KHImage/outOp5_SCS.gif" />
                        </asp:TableCell>
                    </asp:TableHeaderRow>
                    <asp:TableRow ID="TableRow1" Visible="False" runat="server">
                        <asp:TableCell ID="TableCell6" CssClass="DispTable" runat="server">
                            <asp:RadioButton ID="rdoTie" runat="server" GroupName="A" Width="85%" />
                        </asp:TableCell>
                        <asp:TableCell ID="TableCell7" CssClass="DispTable" runat="server">
                            <asp:Label ID="Label30" Text="MX(寸法)" runat="server" CssClass="defaultLabel"></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell ID="TableCell8" CssClass="DispTable" runat="server">
                            <asp:Label ID="Label31" Text="ロッド・ヘッド側８本延長" runat="server" CssClass="defaultLabel"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow ID="TableRow2" Visible="False" runat="server">
                        <asp:TableCell ID="TableCell10" CssClass="DispTable" runat="server">
                            <asp:RadioButton ID="rdoTieR" runat="server" GroupName="A" Width="85%" />
                        </asp:TableCell>
                        <asp:TableCell ID="TableCell11" CssClass="DispTable" runat="server">
                            <asp:Label ID="Label32" Text="MX(寸法)R" runat="server" CssClass="defaultLabel"></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell ID="TableCell12" CssClass="DispTable" runat="server">
                            <asp:Label ID="Label33" Text="ロッド側４本延長" runat="server" CssClass="defaultLabel"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow ID="TableRow3" Visible="False" runat="server">
                        <asp:TableCell ID="TableCell14" CssClass="DispTable" runat="server">
                            <asp:RadioButton ID="rdoTieR1" runat="server" GroupName="A" Width="85%" />
                        </asp:TableCell>
                        <asp:TableCell ID="TableCell15" CssClass="DispTable" runat="server">
                            <asp:Label ID="Label34" Text="MX(寸法)R1" runat="server" CssClass="defaultLabel"></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell ID="TableCell16" CssClass="DispTable" runat="server">
                            <asp:Label ID="Label35" Text="ロッド／ポート有り側２本延長" runat="server" CssClass="defaultLabel"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow ID="TableRow4" Visible="False" runat="server">
                        <asp:TableCell ID="TableCell18" CssClass="DispTable" runat="server">
                            <asp:RadioButton ID="rdoTieR2" runat="server" GroupName="A" Width="85%" />
                        </asp:TableCell>
                        <asp:TableCell ID="TableCell19" CssClass="DispTable" runat="server">
                            <asp:Label ID="Label36" Text="MX(寸法)R2" runat="server" CssClass="defaultLabel"></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell ID="TableCell20" CssClass="DispTable" runat="server">
                            <asp:Label ID="Label37" Text="ロッド／ポート無し側２本延長" runat="server" CssClass="defaultLabel"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow ID="TableRow5" Visible="False" runat="server">
                        <asp:TableCell ID="TableCell22" CssClass="DispTable" runat="server">
                            <asp:RadioButton ID="rdoTieH" runat="server" GroupName="A" Width="85%" />
                        </asp:TableCell>
                        <asp:TableCell ID="TableCell23" CssClass="DispTable" runat="server">
                            <asp:Label ID="Label38" Text="MX(寸法)H" runat="server" CssClass="defaultLabel"></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell ID="TableCell24" CssClass="DispTable" runat="server">
                            <asp:Label ID="Label39" Text="ヘッド側４本延長" runat="server" CssClass="defaultLabel"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow ID="TableRow6" Visible="False" runat="server">
                        <asp:TableCell ID="TableCell26" CssClass="DispTable" runat="server">
                            <asp:RadioButton ID="rdoTieH1" runat="server" GroupName="A" Width="85%" />
                        </asp:TableCell>
                        <asp:TableCell ID="TableCell27" CssClass="DispTable" runat="server">
                            <asp:Label ID="Label40" Text="MX(寸法)H1" runat="server" CssClass="defaultLabel"></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell ID="TableCell28" CssClass="DispTable" runat="server">
                            <asp:Label ID="Label41" Text="ヘッド／ポート有り側２本延長" runat="server" CssClass="defaultLabel"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow ID="TableRow7" Visible="False" runat="server">
                        <asp:TableCell ID="TableCell30" CssClass="DispTable" runat="server">
                            <asp:RadioButton ID="rdoTieH2" runat="server" GroupName="A" Width="85%" />
                        </asp:TableCell>
                        <asp:TableCell ID="TableCell31" CssClass="DispTable" runat="server">
                            <asp:Label ID="Label42" Text="MX(寸法)H2" runat="server" CssClass="defaultLabel"></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell ID="TableCell32" CssClass="DispTable" runat="server">
                            <asp:Label ID="Label43" Text="ヘッド／ポート無し側２本延長" runat="server" CssClass="defaultLabel"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow ID="TableRow8" runat="server">
                        <asp:TableCell ID="TableCell34" ColumnSpan="2" runat="server"></asp:TableCell>
                        <asp:TableCell ID="TableCell35" runat="server">
                            <table id="tblData2" border="1" width="90%" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td style="width: 40%; text-align: center;">
                                        <asp:Label ID="Label44" Text="標準寸法" runat="server" Width="100%" BorderWidth="0"></asp:Label>
                                    </td>
                                    <td style="width: 60%; text-align: center;">
                                        <asp:Label ID="Label45" Text="特注寸法" runat="server" Width="100%" BorderWidth="0"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Label ID="lblDefault" Text="20" runat="server" Width="100%" CssClass="right"></asp:Label>
                                    </td>
                                    <td>
                                        <asp:TextBox ID="txtTieRodCstm" runat="server" Width="90%" CssClass="textBox"></asp:TextBox>
                                        <asp:DropDownList ID="cmbTieRodCstm" runat="server" Visible="False" CssClass="dropDown">
                                        </asp:DropDownList>
                                    </td>
                                </tr>
                            </table>
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </asp:Panel>
            <asp:Panel ID="pnlSUS" runat="server" Width="98%" HorizontalAlign="Left">
                <table style="width: 100%;">
                    <tr>
                        <td style="width: 50%;">
                            <asp:Label ID="Label46" Text="◆タイロッド材質ＳＵＳ" runat="server" CssClass="titleLabel" Width="90%"></asp:Label>
                        </td>
                        <td style="width: 50%;">
                            <asp:DropDownList ID="cmbSUS" runat="server" CssClass="dropDown">
                            </asp:DropDownList>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Panel ID="pnlJM" runat="server" Width="98%" HorizontalAlign="Left">
                <table style="width: 100%;">
                    <tr>
                        <td style="width: 50%;">
                            <asp:Label ID="Label47" Text="◆ピストンロッドはジャバラ付寸法でジャバラなし" runat="server" CssClass="titleLabel"
                                Width="90%"></asp:Label>
                        </td>
                        <td style="width: 50%;">
                            <asp:DropDownList ID="cmbJM" runat="server" CssClass="dropDown">
                            </asp:DropDownList>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Panel ID="pnlFluoroRub" runat="server" Width="98%" HorizontalAlign="Left">
                <table style="width: 100%;">
                    <tr>
                        <td style="width: 50%;">
                            <asp:Label ID="Label48" Text="◆スクレーバー、ロッドパッキンのみフッ素ゴム" runat="server" CssClass="titleLabel"
                                Width="90%"></asp:Label>
                        </td>
                        <td style="width: 50%;">
                            <asp:DropDownList ID="cmbFluoroRub" runat="server" CssClass="dropDown">
                            </asp:DropDownList>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
        </div>
    </asp:Panel>
    <asp:Panel ID="Panel4" runat="server" BackColor="#C7EDCC">
        <table style="width: 98%">
            <tr>
                <td style="width: 60%">
                </td>
                <td style="width: 20%">
                    <asp:Button ID="btnOK" runat="server" Width="100px" Text="OK" OnClientClick="f_OutOfOption_OK('ctl00_ContentDetail_WebUC_OutOfOption')"
                        onfocus="this.style.color = 'white';" onblur="this.style.color = 'black';" CssClass="button" />
                </td>
                <td style="width: 20%">
                    <asp:Button ID="btnCancel" runat="server" Width="100px" Text="Cancel" onfocus="this.style.color = 'white';"
                        onblur="this.style.color = 'black';" CssClass="button" />
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:HiddenField ID="HdnPortPlace1" runat="server" />
    <asp:HiddenField ID="HdnPortPlace2" runat="server" />
    <asp:HiddenField ID="HdnPortPlace3" runat="server" />
    <asp:HiddenField ID="HdnPortPlace4" runat="server" />
    <asp:HiddenField ID="HdnTieRodRdio" runat="server" />
    <asp:HiddenField ID="HdnPtnCnt" runat="server" />
    <asp:HiddenField ID="HdnSelPortCushon" runat="server" />
    <asp:HiddenField ID="HdnValPortCushon" runat="server" />
    <asp:HiddenField ID="HdnSelPortPlace" runat="server" />
    <asp:HiddenField ID="HdnSelPort" runat="server" />
    <asp:HiddenField ID="HdnValPort" runat="server" />
    <asp:HiddenField ID="HdnSelPortSize" runat="server" />
    <asp:HiddenField ID="HdnValPortSize" runat="server" />
    <asp:HiddenField ID="HdnSelMounting" runat="server" />
    <asp:HiddenField ID="HdnValMounting" runat="server" />
    <asp:HiddenField ID="HdnSelTrunnion" runat="server" />
    <asp:HiddenField ID="HdnSelClevis" runat="server" />
    <asp:HiddenField ID="HdnValClevis" runat="server" />
    <asp:HiddenField ID="HdnSelTieRod" runat="server" />
    <asp:HiddenField ID="HdnSelTieRodDefault" runat="server" />
    <asp:HiddenField ID="HdnSeltxtTieRodCstm" runat="server" />
    <asp:HiddenField ID="HdnSelcmbTieRodCstm" runat="server" />
    <asp:HiddenField ID="HdnSelSUS" runat="server" />
    <asp:HiddenField ID="HdnValSUS" runat="server" />
    <asp:HiddenField ID="HdnSelJM" runat="server" />
    <asp:HiddenField ID="HdnValJM" runat="server" />
    <asp:HiddenField ID="HdnSelFluoroRub" runat="server" />
    <asp:HiddenField ID="HdnValFluoroRub" runat="server" />
    <asp:HiddenField ID="HdnActionType" runat="server" />
</div>
