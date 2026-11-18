<%@ Page Title="ATOM メインメニュー" Language="vb" MasterPageFile="~/Site.Master" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="Atom.Web.DefaultPage" %>

<asp:Content ID="cntHead" ContentPlaceHolderID="cphHead" runat="server">
    <%-- この画面だけで使うスタイルです。共通のスタイルは Content/site.css にあります。
         ここへ書いているのは Visual Studio のデザイナでプレビューを確認できるようにするためです。
         色は site.css の共通変数を参照し、この画面で HEX 値を直接書かないようにします。 --%>
    <style type="text/css">
        /* 3 列に並べます。どのブロックをどの列へ置くかは画面定義の Column で決めます。
           列ごとに独立して積み上げるため、隣の列の高さに引きずられた余白ができません。
           列は幅が固定のため、放っておくと左へ寄ります。大きなモニターでも
           真ん中に見えるよう、列ごと中央へ寄せます。 */
        .atom-menu {
            display: flex;
            flex-wrap: wrap;
            align-items: flex-start;
            justify-content: center;
            gap: 18px;
        }

        .atom-column {
            display: flex;
            flex: 0 0 auto;
            flex-direction: column;
            gap: 18px;
            width: var(--atom-block-width);
        }

        /* ブロックは白いカード。枠線ではなく淡い影で区切ります。 */
        .atom-block {
            display: flex;
            flex-direction: column;
            gap: 0;
            width: 100%;
            padding: 16px 12px 14px;
            border: 0;
            border-radius: var(--atom-radius-l);
            background-color: var(--atom-surface);
            box-shadow: var(--atom-shadow);
        }

        /* 見出しの色だけで系統を分けます。面は白のまま揃えます。 */
        .atom-block-blue .atom-block-title {
            color: var(--atom-accent);
        }

        .atom-block-pink .atom-block-title {
            color: var(--atom-menu-pink-text);
        }

        .atom-block-title {
            margin: 0 0 12px;
            padding: 0 6px;
            overflow: hidden;
            border: 0;
            background-color: transparent;
            font-size: 13px;
            font-weight: bold;
            text-align: left;
            text-overflow: ellipsis;
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
            padding: 11px 14px;
            border: 0;
            border-radius: var(--atom-radius);
            background-color: var(--atom-surface-soft);
            color: var(--atom-text);
            cursor: pointer;
            font-family: inherit;
            font-size: 13px;
            text-align: left;
            text-decoration: none;
        }

        .atom-menu-button:hover {
            background-color: var(--atom-accent-weak);
            color: var(--atom-accent);
            filter: none;
            text-decoration: none;
        }

        /* 未実装の画面。押す前に分かるよう薄くします。
           押すこと自体は今までどおりできます（押すと「現在未実装です。」が出ます）。 */
        .atom-menu-button-pending {
            background-color: #FBFCFD;
            color: var(--atom-disabled-text);
        }

        .atom-menu-button-pending:hover {
            background-color: var(--atom-surface-soft);
            color: var(--atom-text-muted);
        }

        /* 権限が無い画面。こちらは押せません。 */
        .atom-menu-button.aspNetDisabled {
            background-color: #FBFCFD;
            color: var(--atom-disabled-text);
            cursor: default;
        }

        .atom-menu-button.aspNetDisabled:hover {
            background-color: #FBFCFD;
            color: var(--atom-disabled-text);
        }

        .atom-menu-status {
            display: block;
            margin-top: 3px;
            padding: 0 14px;
            color: var(--atom-disabled-text);
            font-size: 11.5px;
            text-align: left;
        }

        /* ログアウトは 3 列目（輸入経費入力の下）に置きます */
        .atom-menu-footer {
            display: flex;
        }

        .atom-logout-button {
            width: 100%;
            padding: 12px 14px;
            border: 1px solid var(--atom-line-strong);
            border-radius: var(--atom-radius);
            background-color: var(--atom-surface);
            color: #344054;
            cursor: pointer;
            font-family: inherit;
            font-size: 13px;
            text-align: center;
        }

        .atom-logout-button:hover {
            border-color: var(--atom-accent);
            background-color: var(--atom-accent-weak);
            color: var(--atom-accent);
            filter: none;
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
                                            CssClass='<%# Server.HtmlEncode(CStr(Eval("MenuButtonCssClass"))) %>'
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
                                            CssClass='<%# Server.HtmlEncode(CStr(Eval("MenuButtonCssClass"))) %>'
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
                                            CssClass='<%# Server.HtmlEncode(CStr(Eval("MenuButtonCssClass"))) %>'
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
