<%@ Page Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="Default.aspx.vb" Inherits="Atom.Web._Default" %>
<%@ Import Namespace="System.Web" %>

<asp:Content ID="cntHead" ContentPlaceHolderID="cphHead" runat="server">
    <title>ATOM メインメニュー</title>
</asp:Content>

<asp:Content ID="cntContent" ContentPlaceHolderID="cphContent" runat="server">

    <div class="menu-root">
        <asp:Repeater ID="rptBlocks" runat="server" EnableViewState="False">
            <ItemTemplate>
                <section class="menu-block">
                    <h2 class="menu-block-title"><%# HttpUtility.HtmlEncode(CStr(Eval("Title"))) %></h2>
                    <div class="menu-block-body">
                        <asp:Repeater ID="rptEntries" runat="server" EnableViewState="False">
                            <ItemTemplate>
                                <asp:HyperLink ID="lnkMenu" runat="server"
                                    CssClass="menu-button"
                                    NavigateUrl='<%# CStr(Eval("NavigateUrl")) %>'
                                    Enabled='<%# CBool(Eval("IsEnabled")) %>'
                                    ToolTip='<%# "メニュー番号 " & CStr(Eval("MenuNo")) %>'>
                                    <span class="menu-button-title"><%# HttpUtility.HtmlEncode(CStr(Eval("Title"))) %></span>
                                    <span class="menu-button-status"><%# HttpUtility.HtmlEncode(CStr(Eval("StatusText"))) %></span>
                                </asp:HyperLink>
                            </ItemTemplate>
                        </asp:Repeater>
                    </div>
                </section>
            </ItemTemplate>
        </asp:Repeater>
    </div>

</asp:Content>
