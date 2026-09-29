<%@ Page Title="ATOM メインメニュー" Language="vb" MasterPageFile="~/Site.Master" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="Atom.Web.DefaultPage" %>

<asp:Content ID="cntHead" ContentPlaceHolderID="cphHead" runat="server">
</asp:Content>

<asp:Content ID="cntContent" ContentPlaceHolderID="cphContent" runat="server">

    <div class="atom-menu">
        <asp:Repeater ID="rptBlocks" runat="server">
            <ItemTemplate>
                <section class='<%# Server.HtmlEncode(CStr(Eval("CssClass"))) %>'>
                    <h2 class="atom-block-title"><%# Server.HtmlEncode(CStr(Eval("Title"))) %></h2>
                    <ul class="atom-block-items">
                        <asp:Repeater ID="rptScreens" runat="server" DataSource='<%# Eval("Screens") %>'>
                            <ItemTemplate>
                                <li>
                                    <asp:LinkButton ID="lnkScreen" runat="server"
                                        CssClass="atom-menu-button"
                                        CausesValidation="false"
                                        CommandName="Open"
                                        CommandArgument='<%# CStr(Eval("MenuNo")) %>'
                                        Enabled='<%# CBool(Eval("IsEnabled")) %>'
                                        Text='<%# Server.HtmlEncode(CStr(Eval("Title"))) %>' />
                                    <asp:Label ID="lblStatus" runat="server"
                                        CssClass="atom-menu-status"
                                        Text='<%# Server.HtmlEncode(CStr(Eval("StatusText"))) %>'
                                        Visible='<%# CStr(Eval("StatusText")).Length > 0 %>' />
                                </li>
                            </ItemTemplate>
                        </asp:Repeater>
                    </ul>
                </section>
            </ItemTemplate>
        </asp:Repeater>
    </div>

    <div class="atom-menu-footer">
        <asp:Button ID="btnLogout" runat="server" Text="ログアウト"
            CssClass="atom-logout-button"
            CausesValidation="false"
            OnClientClick="return AtomPage.confirmLogout();" />
    </div>

</asp:Content>
