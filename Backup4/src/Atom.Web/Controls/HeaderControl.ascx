<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="HeaderControl.ascx.vb" Inherits="Atom.Web.HeaderControl" %>
<div id="divHeader" runat="server" class="atom-header">

    <div class="atom-header-icons">
        <button type="button" class="atom-header-icon" title="再読み込み" onclick="return AtomPage.reload();">
            <svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" focusable="false">
                <circle cx="12" cy="12" r="8" fill="none" stroke="#333333" stroke-width="2" />
                <path d="M12 4 A8 8 0 0 1 20 12" fill="none" stroke="#333333" stroke-width="2" />
                <path d="M4 12 A8 8 0 0 0 12 20" fill="none" stroke="#333333" stroke-width="2" />
                <path d="M2 12 H22" fill="none" stroke="#333333" stroke-width="2" />
            </svg>
        </button>
        <asp:LinkButton ID="lnkManual" runat="server" CssClass="atom-header-icon" ToolTip="マニュアル" CausesValidation="false">
            <svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" focusable="false">
                <rect x="4" y="3" width="16" height="18" fill="none" stroke="#333333" stroke-width="2" />
                <path d="M4 7 H20" fill="none" stroke="#333333" stroke-width="2" />
                <path d="M8 12 H16" fill="none" stroke="#333333" stroke-width="2" />
                <path d="M8 16 H16" fill="none" stroke="#333333" stroke-width="2" />
            </svg>
        </asp:LinkButton>
    </div>

    <div class="atom-header-title">
        <asp:Panel ID="pnlLogo" runat="server" CssClass="atom-header-logo">
            <span class="atom-header-logo-main">ATOM</span>
            <span class="atom-header-logo-sub">（ Arrangement of Tool Orders &amp; Materials System ）</span>
            <span class="atom-header-logo-version"><asp:Literal ID="litVersion" runat="server" Mode="Encode" /></span>
        </asp:Panel>
        <asp:Panel ID="pnlScreenTitle" runat="server" CssClass="atom-header-screen" Visible="false">
            <asp:Literal ID="litScreenTitle" runat="server" Mode="Encode" />
        </asp:Panel>
    </div>

    <div class="atom-header-actions">
        <asp:Repeater ID="rptActions" runat="server">
            <ItemTemplate>
                <%-- OnClientClick は遷移しないボタン（別タブの「閉じる」等）でのみ値が入ります。
                     false を返すスクリプトを与えるとポストバックせずブラウザ側だけで処理されます。 --%>
                <asp:LinkButton ID="lnkAction" runat="server"
                    CssClass="atom-header-action"
                    CausesValidation="false"
                    CommandName="Navigate"
                    CommandArgument='<%# CStr(Eval("MenuNo")) %>'
                    OnClientClick='<%# CStr(Eval("ClientScript")) %>'
                    Text='<%# Server.HtmlEncode(CStr(Eval("Text"))) %>' />
            </ItemTemplate>
        </asp:Repeater>
    </div>

    <div class="atom-header-user">
        <asp:Panel ID="pnlUser" runat="server" CssClass="atom-header-user-box" Visible="false">
            <asp:Literal ID="litLoginName" runat="server" Mode="Encode" />
        </asp:Panel>
    </div>

</div>
