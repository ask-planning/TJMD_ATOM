<%@ Page Title="ATOM ログイン" Language="vb" MasterPageFile="~/Site.Master" AutoEventWireup="false" CodeBehind="Login.aspx.vb" Inherits="Atom.Web.Pages.Common.Login" %>

<asp:Content ID="cntHead" ContentPlaceHolderID="cphHead" runat="server">
    <%-- この画面だけで使うスタイルです。共通のスタイルは Content/site.css にあります。
         ここへ書いているのは Visual Studio のデザイナでプレビューを確認できるようにするためです。
         色は site.css の共通変数を参照し、この画面で HEX 値を直接書かないようにします。 --%>
    <style type="text/css">
        .atom-login {
            max-width: 380px;
            margin: 48px auto 0;
            padding: 24px;
            border: 1px solid var(--atom-menu-border);
        }

        .atom-login-title {
            margin: 0 0 16px;
            font-size: 18px;
            font-weight: bold;
        }

        .atom-login-row {
            display: flex;
            flex-direction: column;
            gap: 4px;
            margin-bottom: 16px;
        }

        .atom-login-input {
            width: 100%;
            padding: 6px 8px;
            border: 1px solid var(--atom-button-border);
            font-size: 16px;
            ime-mode: disabled;
        }

        .atom-login-button {
            width: 100%;
            padding: 10px 4px;
            border: 1px solid var(--atom-button-border);
            background-color: var(--atom-menu-button-bg);
            cursor: pointer;
            font-size: 16px;
        }
    </style>
</asp:Content>

<asp:Content ID="cntContent" ContentPlaceHolderID="cphContent" runat="server">

    <%-- DefaultButton を指定すると、この中で Enter を押したときにログインボタンが押されます。
         ログインボタンは UseSubmitBehavior="false"（二重送信を防ぐため押下時に無効化する）ため、
         指定が無いと Enter で送信されません。 --%>
    <asp:Panel ID="pnlLogin" runat="server" CssClass="atom-login" DefaultButton="btnLogin">
        <h1 class="atom-login-title">ログイン</h1>

        <div class="atom-login-row">
            <asp:Label ID="lblLoginCode" runat="server" AssociatedControlID="txtLoginCode" Text="社員コード（6桁）" />
            <asp:TextBox ID="txtLoginCode" runat="server" CssClass="atom-login-input" MaxLength="6" autocomplete="off" />
            <asp:RequiredFieldValidator ID="rfvLoginCode" runat="server"
                ControlToValidate="txtLoginCode"
                CssClass="atom-validator"
                Display="Dynamic"
                ErrorMessage="社員コードを入力してください。"
                Text="社員コードを入力してください。" />
            <asp:RegularExpressionValidator ID="revLoginCode" runat="server"
                ControlToValidate="txtLoginCode"
                CssClass="atom-validator"
                Display="Dynamic"
                ValidationExpression="^.{6}$"
                ErrorMessage="入力桁数が不正です（６桁）。"
                Text="入力桁数が不正です（６桁）。" />
        </div>

        <asp:Button ID="btnLogin" runat="server" Text="ログイン"
            CssClass="atom-login-button"
            UseSubmitBehavior="false"
            OnClientClick="if (typeof Page_ClientValidate === 'function' &amp;&amp; !Page_ClientValidate('')) { return false; } this.disabled = true;" />
    </asp:Panel>

</asp:Content>
