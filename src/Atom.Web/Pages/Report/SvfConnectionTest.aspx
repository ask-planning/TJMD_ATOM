<%@ Page Title="SVF 連携確認" Language="vb" MasterPageFile="~/Site.Master" AutoEventWireup="false" CodeBehind="SvfConnectionTest.aspx.vb" Inherits="Atom.Web.Pages.Report.SvfConnectionTest" %>

<asp:Content ID="cntHead" ContentPlaceHolderID="cphHead" runat="server">
    <%-- この画面だけで使うスタイルです。共通のスタイルは Content/site.css にあります。 --%>
    <style type="text/css">
        .atom-check-table {
            border-collapse: collapse;
            margin-bottom: 16px;
        }

        .atom-check-table th,
        .atom-check-table td {
            padding: 6px 12px;
            border: 1px solid var(--atom-menu-border);
            text-align: left;
            vertical-align: top;
            font-size: 13px;
        }

        .atom-check-table th {
            background-color: var(--atom-grid-head-bg);
            white-space: nowrap;
        }

        .atom-check-result {
            margin-top: 16px;
            padding: 12px;
            border: 1px solid var(--atom-menu-border);
            background-color: var(--atom-grid-area-bg);
            font-size: 13px;
            white-space: pre-wrap;
        }
    </style>
</asp:Content>

<asp:Content ID="cntContent" ContentPlaceHolderID="cphContent" runat="server">

    <h1 class="atom-page-title">SVF 連携確認</h1>

    <p class="atom-form-note">
        Web サーバーから SVF の監視フォルダへ連携ファイルを置けるかを確かめます。<br />
        配置に成功すると SVF がファイルを拾い、PDF を出力します。<strong>動作確認用の画面です。</strong>
    </p>

    <table class="atom-check-table">
        <tr>
            <th>実行アカウント</th>
            <td><asp:Literal ID="litAccount" runat="server" Mode="Encode" /></td>
        </tr>
        <tr>
            <th>監視フォルダ</th>
            <td><asp:Literal ID="litWatchFolder" runat="server" Mode="Encode" /></td>
        </tr>
        <tr>
            <th>フォルダの状態</th>
            <td><asp:Literal ID="litFolderState" runat="server" Mode="Encode" /></td>
        </tr>
        <tr>
            <th>ファイル名</th>
            <td><asp:Literal ID="litFileName" runat="server" Mode="Encode" /></td>
        </tr>
    </table>

    <asp:Button ID="btnWrite" runat="server" Text="連携ファイルを置く" CssClass="atom-action-button"
        UseSubmitBehavior="false" OnClientClick="AtomPage.beginProcessing(this);" />

    <asp:Panel ID="pnlResult" runat="server" CssClass="atom-check-result" Visible="false">
        <asp:Literal ID="litResult" runat="server" Mode="Encode" />
    </asp:Panel>

</asp:Content>
