<%@ Page Title="ATOM メインメニュー" Language="vb" MasterPageFile="~/Site.Master" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="Atom.Web.DefaultPage" %>

<asp:Content ID="cntHead" ContentPlaceHolderID="cphHead" runat="server">
    <%-- この画面だけで使うスタイルです。共通のスタイルは Content/site.css にあります。
         ここへ書いているのは Visual Studio のデザイナでプレビューを確認できるようにするためです。
         色は site.css の共通変数を参照し、この画面で HEX 値を直接書かないようにします。 --%>
    <style type="text/css">
        /* 3 列に並べます。どのブロックをどの列へ置くかは画面定義の Column で決めます。
           列ごとに独立して積み上げるため、隣の列の高さに引きずられた余白ができません。 */
        .atom-menu {
            display: flex;
            flex-wrap: wrap;
            align-items: flex-start;
            gap: 16px;
        }

        .atom-column {
            display: flex;
            flex: 0 0 auto;
            flex-direction: column;
            gap: 16px;
            width: var(--atom-block-width);
        }

        .atom-block {
            display: flex;
            flex-direction: column;
            gap: 8px;
            width: 100%;
            padding: 8px;
            border: 1px solid var(--atom-menu-border);
        }

        .atom-block-blue {
            background-color: var(--atom-menu-blue-bg);
        }

        .atom-block-pink {
            background-color: var(--atom-menu-pink-bg);
        }

        .atom-block-title {
            margin: 0;
            padding: 6px 4px;
            overflow: hidden;
            border: 1px solid var(--atom-menu-border);
            background-color: var(--atom-menu-pink-bg);
            color: var(--atom-menu-blue-text);
            font-size: 13px;
            font-weight: bold;
            text-align: center;
            /* 「上海輸入定期便＆AIR便」のような長い見出しが途中で折り返さないようにする */
            white-space: nowrap;
        }

        .atom-block-items {
            display: flex;
            flex-direction: column;
            gap: 8px;
            margin: 0;
            padding: 0;
            list-style: none;
        }

        .atom-menu-button {
            display: block;
            width: 100%;
            padding: 8px 4px;
            border: 1px solid var(--atom-button-border);
            background-color: var(--atom-menu-button-bg);
            color: var(--atom-menu-blue-text);
            cursor: pointer;
            font-size: 14px;
            text-align: center;
            text-decoration: none;
        }

        .atom-menu-button:hover {
            filter: brightness(0.95);
        }

        .atom-menu-button.aspNetDisabled {
            color: var(--atom-disabled-text);
            cursor: default;
        }

        .atom-menu-status {
            display: block;
            color: var(--atom-disabled-text);
            font-size: 12px;
            text-align: center;
        }

        /* ログアウトは 3 列目（輸入経費入力の下）に置きます */
        .atom-menu-footer {
            display: flex;
        }

        .atom-logout-button {
            width: 100%;
            padding: 12px 4px;
            border: 1px solid var(--atom-button-border);
            background-color: var(--atom-menu-button-bg);
            color: var(--atom-text);
            cursor: pointer;
            font-size: 16px;
            text-align: center;
        }

        .atom-logout-button:hover {
            filter: brightness(0.95);
        }
    </style>
</asp:Content>

<asp:Content ID="cntContent" ContentPlaceHolderID="cphContent" runat="server">

    <div class="atom-menu">

        <div class="atom-column">
            <asp:Repeater ID="rptColumn1" runat="server">
                <ItemTemplate>
                    <section class='<%# Server.HtmlEncode(CStr(Eval("CssClass"))) %>'>
                        <h2 class="atom-block-title"><%# Server.HtmlEncode(CStr(Eval("Title"))) %></h2>
                        <ul class="atom-block-items">
                            <asp:Repeater ID="rptScreens" runat="server" DataSource='<%# Eval("VisibleScreens") %>'>
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

        <div class="atom-column">
            <asp:Repeater ID="rptColumn2" runat="server">
                <ItemTemplate>
                    <section class='<%# Server.HtmlEncode(CStr(Eval("CssClass"))) %>'>
                        <h2 class="atom-block-title"><%# Server.HtmlEncode(CStr(Eval("Title"))) %></h2>
                        <ul class="atom-block-items">
                            <asp:Repeater ID="rptScreens" runat="server" DataSource='<%# Eval("VisibleScreens") %>'>
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

        <div class="atom-column">
            <asp:Repeater ID="rptColumn3" runat="server">
                <ItemTemplate>
                    <section class='<%# Server.HtmlEncode(CStr(Eval("CssClass"))) %>'>
                        <h2 class="atom-block-title"><%# Server.HtmlEncode(CStr(Eval("Title"))) %></h2>
                        <ul class="atom-block-items">
                            <asp:Repeater ID="rptScreens" runat="server" DataSource='<%# Eval("VisibleScreens") %>'>
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

            <div class="atom-menu-footer">
                <asp:Button ID="btnLogout" runat="server" Text="ログアウト"
                    CssClass="atom-logout-button"
                    CausesValidation="false"
                    OnClientClick="return AtomPage.confirmLogout();" />
            </div>
        </div>

    </div>

</asp:Content>
