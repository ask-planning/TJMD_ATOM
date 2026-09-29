<%@ Page Title="ATOM ログイン" Language="vb" MasterPageFile="~/Site.Master" AutoEventWireup="false" CodeBehind="Login.aspx.vb" Inherits="Atom.Web.Pages.Common.Login" %>

<asp:Content ID="cntHead" ContentPlaceHolderID="cphHead" runat="server">
</asp:Content>

<asp:Content ID="cntContent" ContentPlaceHolderID="cphContent" runat="server">

    <div class="atom-login">
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
    </div>

</asp:Content>
