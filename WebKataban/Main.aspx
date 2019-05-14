<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/WebUserControl/Master_Main.master"
    CodeBehind="Main.aspx.vb" Inherits="WebKataban._Main" %>

<%@ Register Src="WebUserControl/WebUC_Type.ascx" TagName="WebUC_Type" TagPrefix="uc1" %>
<%@ Register Src="WebUserControl/WebUC_Youso.ascx" TagName="WebUC_Youso" TagPrefix="uc2" %>
<%@ Register Src="WebUserControl/WebUC_Login.ascx" TagName="WebUC_Login" TagPrefix="uc3" %>
<%@ Register Src="WebUserControl/WebUC_Menu.ascx" TagName="WebUC_Menu" TagPrefix="uc4" %>
<%@ Register Src="WebUserControl/WebUC_Tanka.ascx" TagName="WebUC_Tanka" TagPrefix="uc5" %>
<%@ Register Src="WebUserControl/WebUC_ISOTanka.ascx" TagName="WebUC_ISOTanka" TagPrefix="uc6" %>
<%@ Register Src="WebUserControl/WebUC_RodEndOrder.ascx" TagName="WebUC_RodEndOrder"
    TagPrefix="uc7" %>
<%@ Register Src="WebUserControl/WebUC_RodEnd.ascx" TagName="WebUC_RodEnd" TagPrefix="uc8" %>
<%@ Register Src="WebUserControl/WebUC_OutOfOption.ascx" TagName="WebUC_OutOfOption"
    TagPrefix="uc9" %>
<%@ Register Src="WebUserControl/WebUC_Stopper.ascx" TagName="WebUC_Stopper" TagPrefix="uc10" %>
<%@ Register Src="WebUserControl/WebUC_PriceCopy.ascx" TagName="WebUC_PriceCopy"
    TagPrefix="uc11" %>
<%@ Register Src="WebUserControl/WebUC_Siyou.ascx" TagName="WebUC_Siyou" TagPrefix="uc12" %>
<%@ Register Src="WebUserControl/WebUC_Error.ascx" TagName="WebUC_Error" TagPrefix="uc13" %>
<%@ Register Src="WebUserControl/WebUC_Master.ascx" TagName="WebUC_Master" TagPrefix="uc14" %>
<%@ Register Src="WebUserControl/WebUC_TEST/WebUC_KatOut.ascx" TagName="WebUC_KatOut"
    TagPrefix="uc15" %>
<%@ Register Src="WebUserControl/WebUC_TEST/WebUC_KatSep.ascx" TagName="WebUC_KatSep"
    TagPrefix="uc16" %>
<%@ Register Src="WebUserControl/WebUC_TEST/WebUC_100Test.ascx" TagName="WebUC_100Test"
    TagPrefix="uc17" %>
<%@ Register Src="WebUserControl/WebUC_Motor.ascx" TagName="WebUC_Motor" TagPrefix="uc18" %>
<%@ Register Src="WebUserControl/WebUC_PriceDetail.ascx" TagName="WebUC_PriceDetail"
    TagPrefix="uc19" %>
<%@ Register Src="WebUserControl/WebUC_TypeAnonymous.ascx" tagName="WebUC_TypeAnonymous" tagPrefix="uc20" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentTitle" runat="server">
    <script type="text/javascript" language="javascript">
        Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);
        function EndRequestHandler(sender, args) {
            if (args.get_error() != undefined) {
                args.set_errorHandled(true);
            }
        }
    </script>
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <asp:UpdatePanel ID="UpdatePanelMenu" runat="server">
        <ContentTemplate>
            <table class="menuBar" cellspacing="0px" cellpadding="0px">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td>
                        <img alt="CKD" src="KHImage/logo4.jpg" />
                    </td>
                    <td style="width: 2%">
                    </td>
                    <td style="width: 67%; height: 25px">
                        <asp:Panel ID="PanMainMenu" runat="server" Height="25px" HorizontalAlign="Left">
                            <asp:Button ID="Button1" OnClientClick="SetButtonID('1')" runat="server" CssClass="button" />
                            <asp:Button ID="Button3" OnClientClick="SetButtonID('3')" runat="server" CssClass="button" />
                            <asp:Button ID="Button4" OnClientClick="SetButtonID('4')" runat="server" CssClass="button" />
                            <!--匿名ユーザー機種選択画面ボタン-->
                            <asp:Button ID="Button7" OnClientClick="SetButtonID('7')" runat="server" CssClass="button" />
                            <asp:Button ID="Button5" OnClientClick="SetButtonID('5')" runat="server" CssClass="button" />
                            <asp:Button ID="Button6" OnClientClick="SetButtonID('6')" runat="server" CssClass="button" />
                            <asp:HiddenField ID="HidRunForm" runat="server" />
                        </asp:Panel>
                    </td>
                    <td style="width: 15%">
                            <!-- 言語選択 -->
                            <asp:DropDownList ID="selLang" runat="server" AutoPostBack="True" DataTextField="language_nm"
                                DataValueField="language_cd" CssClass="DropDownList">
                            </asp:DropDownList>
                    </td>
                    <td style="width: 10%">
                        <asp:Button ID="Button2" runat="server" Visible="False" CssClass="button" />
                    </td>
                </tr>
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td>
                    </td>
                    <td style="width: 2%">
                    </td>
                    <td colspan="3">
                        <asp:Panel ID="pnlMaster" runat="server" HorizontalAlign="Left">
                            <asp:Button ID="Button10" OnClientClick="SetButtonID('10')" runat="server" CssClass="button" />
                            <asp:Button ID="Button11" OnClientClick="SetButtonID('11')" runat="server" CssClass="button" />
                            <asp:Button ID="Button12" OnClientClick="SetButtonID('12')" runat="server" CssClass="button" />
                            <asp:Button ID="Button13" OnClientClick="SetButtonID('13')" runat="server" CssClass="button" />
                            <asp:Button ID="Button14" OnClientClick="SetButtonID('14')" runat="server" CssClass="button" />
                            <asp:Button ID="Button15" OnClientClick="SetButtonID('15')" runat="server" CssClass="button" />
                        </asp:Panel>
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
    <asp:HiddenField ID="AppFlg" runat="server" Value="false" />
    <asp:HiddenField ID="DownloadMode" runat="server" />
    <asp:LinkButton ID="btnDownload" runat="server"></asp:LinkButton>
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="ContentDetail" runat="server">
    <script type="text/javascript" language="javascript">
        Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);
        function EndRequestHandler(sender, args) {
            if (args.get_error() != undefined) {
                args.set_errorHandled(true);
            }
        }
    </script>
    <asp:UpdatePanel ID="UpdatePanelPage" runat="server">
        <ContentTemplate>
            <uc3:WebUC_Login ID="WebUC_Login" runat="server" />
            <uc4:WebUC_Menu ID="WebUC_Menu" runat="server" Visible="False" />
            <uc1:WebUC_Type ID="WebUC_Type" runat="server" Visible="False" />
            <uc20:WebUC_TypeAnonymous ID="WebUC_TypeAnonymous" runat="server" Visible="False" />
            <uc2:WebUC_Youso ID="WebUC_Youso" runat="server" Visible="False" />
            <uc5:WebUC_Tanka ID="WebUC_Tanka" runat="server" Visible="False" />
            <uc6:WebUC_ISOTanka ID="WebUC_ISOTanka" runat="server" Visible="False" />
            <uc7:WebUC_RodEndOrder ID="WebUC_RodEndOrder" runat="server" Visible="False" />
            <uc8:WebUC_RodEnd ID="WebUC_RodEnd" runat="server" Visible="False" />
            <uc9:WebUC_OutOfOption ID="WebUC_OutOfOption" runat="server" Visible="False" />
            <uc10:WebUC_Stopper ID="WebUC_Stopper" runat="server" Visible="False" />
            <uc18:WebUC_Motor ID="WebUC_Motor" runat="server" Visible="False" />
            <uc11:WebUC_PriceCopy ID="WebUC_PriceCopy" runat="server" Visible="False" />
            <uc13:WebUC_Error ID="WebUC_Error" runat="server" Visible="False" />
            <uc14:WebUC_Master ID="WebUC_Master" runat="server" Visible="False" />
            <uc15:WebUC_KatOut ID="WebUC_KatOut" runat="server" Visible="False" />
            <uc16:WebUC_KatSep ID="WebUC_KatSep" runat="server" Visible="False" />
            <uc17:WebUC_100Test ID="WebUC_100Test" runat="server" Visible="False" />
            <uc12:WebUC_Siyou ID="WebUC_Siyou" runat="server" Visible="False" />
            <uc19:WebUC_PriceDetail ID="WebUC_PriceDetail" runat="server" Visible="False" />
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
<asp:Content ID="Content4" ContentPlaceHolderID="footer" runat="server">
    <div style="text-align: right; font-size: smaller; color: Gray; background: #C7EDCC">
        <asp:Label ID="pageFooter" runat="server" Text="主体备案号：沪ICP备17029745号　　网站备案号：沪ICP备17029745号-1"/>
    </div>
</asp:Content>
