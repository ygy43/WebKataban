<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="WebUC_TypeAnonymous.ascx.vb"
Inherits="WebKataban.WebUC_TypeAnonymous" %>

<div class="typeanonymous left">
    <asp:Panel ID="PnlTypeAnonymous" runat="server" CssClass="mainContainer">
        <asp:TreeView ID="TreeViewSeries" runat="server" CssClass="marginleft50top10" NodeIndent="15" ExpandDepth="0" ShowExpandCollapse="False" OnSelectedNodeChanged="TreeViewSeries_OnSelectedNodeChanged">
            <RootNodeStyle></RootNodeStyle>
            <HoverNodeStyle Font-Underline="True" ForeColor="#6666AA"/>
            <NodeStyle Font-Size="12pt" ForeColor="Black" HorizontalPadding="2px" NodeSpacing="2px" VerticalPadding="2px" />
            <ParentNodeStyle Font-Bold="True" Font-Size="14pt" />
        </asp:TreeView>
    </asp:Panel>
</div>