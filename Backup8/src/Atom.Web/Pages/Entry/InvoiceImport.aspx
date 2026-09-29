<%@ Page Title="INVOICE/PACKING取込み" Language="vb" MasterPageFile="~/Site.Master" AutoEventWireup="false" CodeBehind="InvoiceImport.aspx.vb" Inherits="Atom.Web.Pages.Entry.InvoiceImport" %>

<asp:Content ID="cntHead" ContentPlaceHolderID="cphHead" runat="server">
</asp:Content>

<asp:Content ID="cntContent" ContentPlaceHolderID="cphContent" runat="server">

    <div class="atom-form">

        <div class="atom-form-row">
            <asp:Label ID="lblChumonsaki" runat="server" CssClass="atom-form-label" Text="発注先" />
            <div class="atom-form-field">
                <asp:RadioButtonList ID="rblChumonsaki" runat="server"
                    CssClass="atom-radio-list"
                    RepeatDirection="Horizontal"
                    RepeatLayout="Flow" />
            </div>
        </div>

        <div class="atom-form-row">
            <asp:Label ID="lblTsuka" runat="server" CssClass="atom-form-label" Text="通貨" />
            <div class="atom-form-field">
                <asp:RadioButtonList ID="rblTsuka" runat="server"
                    CssClass="atom-radio-list"
                    RepeatDirection="Horizontal"
                    RepeatLayout="Flow">
                    <asp:ListItem Value="JPY" Text="JPY" />
                    <asp:ListItem Value="RMB" Text="RMB" />
                </asp:RadioButtonList>
            </div>
        </div>

        <div class="atom-form-row">
            <asp:Label ID="lblShanghaiChotatsu" runat="server" CssClass="atom-form-label"
                AssociatedControlID="ddlShanghaiChotatsu" Text="担当" />
            <div class="atom-form-field">
                <asp:DropDownList ID="ddlShanghaiChotatsu" runat="server" CssClass="atom-select">
                    <asp:ListItem Value="1" Text="上海" />
                    <asp:ListItem Value="2" Text="調達" />
                    <asp:ListItem Value="3" Text="貿易" />
                </asp:DropDownList>
            </div>
        </div>

        <div class="atom-form-row">
            <asp:Label ID="lblInvoiceFile" runat="server" CssClass="atom-form-label"
                AssociatedControlID="fuInvoice" Text="INVOICEファイル" />
            <div class="atom-form-field">
                <asp:FileUpload ID="fuInvoice" runat="server" CssClass="atom-file" accept=".xls,.xlsx" />
                <%-- 取り込んだファイル名です。ファイル選択欄はポストバックで必ず空になるため、
                     何を取り込んだのかが分かるように名前だけを表示します（ファイルは残しません）。 --%>
                <asp:Label ID="lblInvoiceFileName" runat="server" CssClass="atom-file-name" Visible="false" />
            </div>
        </div>

        <div class="atom-form-row">
            <asp:Label ID="lblPackingFile" runat="server" CssClass="atom-form-label"
                AssociatedControlID="fuPacking" Text="PACKINGファイル" />
            <div class="atom-form-field">
                <asp:FileUpload ID="fuPacking" runat="server" CssClass="atom-file" accept=".xls,.xlsx" />
                <asp:Label ID="lblPackingFileName" runat="server" CssClass="atom-file-name" Visible="false" />
            </div>
        </div>

        <div class="atom-form-row">
            <asp:Label ID="lblRefNo" runat="server" CssClass="atom-form-label"
                AssociatedControlID="txtRefNo" Text="REF_NO" />
            <div class="atom-form-field">
                <asp:TextBox ID="txtRefNo" runat="server" CssClass="atom-readonly" ReadOnly="true" />
            </div>
        </div>

        <div class="atom-form-buttons">
            <asp:Button ID="btnImport" runat="server" Text="取り込み"
                CssClass="atom-action-button atom-action-primary"
                ToolTip="Excelを読み直して取込データを作り直します。"
                CausesValidation="false"
                UseSubmitBehavior="false"
                OnClientClick="AtomPage.beginProcessing(this);" />
            <asp:Button ID="btnRerun" runat="server" Text="再実行"
                CssClass="atom-action-button"
                ToolTip="Excelは読み直しません。取込データのまま、チェックと登録を行います。"
                CausesValidation="false"
                UseSubmitBehavior="false"
                OnClientClick="AtomPage.beginProcessing(this);" />
            <%-- 取込内容確認は別のタブで開きます。ポストバックしないため、
                 この画面の入力内容と REF_NO はそのまま残ります。
                 開くのをスクリプト経由にしているのは、確認画面の「閉じる」で
                 タブを閉じられるようにするためです（window.close はスクリプトで
                 開いたタブしか閉じられません）。開けなかったときは Target の動作にまかせます。 --%>
            <asp:HyperLink ID="lnkKakunin" runat="server" Text="取込内容確認"
                CssClass="atom-action-button atom-action-link"
                Target="_blank"
                onclick="return AtomPage.openScreen(this.href);" />
        </div>

        <p class="atom-form-note">
            「取込内容確認」は<strong>別のタブ</strong>で開きます。この画面の入力内容は残るため、
            確認して直したあとは、このタブに戻って<strong>「再実行」</strong>を押してください。<br />
            「取り込み」を押すと Excel を読み直すため、<strong>確認画面で直した内容は消えます。</strong><br />
            ファイル欄の横に出るのは<strong>取り込んだファイルの名前</strong>です。
            続けてチェックと登録をやり直すだけなら、ファイルを選び直す必要はありません。
        </p>

    </div>

    <%-- 処理中の表示です。Visible="false" にすると HTML が出力されず
         スクリプトから見つけられないため、隠すのは CSS（.atom-processing の display:none）に任せます。 --%>
    <asp:Panel ID="pnlProcessing" runat="server" CssClass="atom-processing">
        処理しています。しばらくお待ちください。
    </asp:Panel>

    <asp:Panel ID="pnlConfirm" runat="server" CssClass="atom-modal" Visible="false">
        <div class="atom-modal-box">
            <p class="atom-modal-message">
                <asp:Literal ID="litConfirmMessage" runat="server" Mode="Encode" />
            </p>
            <%-- 確認ダイアログのボタンからも本登録が走るため、取り込みボタンと同じく
                 連打を止めます。UseSubmitBehavior="false" にしているのは、
                 押下時に無効化しても送信が取り消されないようにするためです。 --%>
            <div class="atom-modal-buttons">
                <asp:Button ID="btnConfirm1" runat="server" CssClass="atom-action-button atom-action-primary"
                    CausesValidation="false"
                    UseSubmitBehavior="false"
                    OnClientClick="AtomPage.beginProcessing(this);" />
                <asp:Button ID="btnConfirm2" runat="server" CssClass="atom-action-button"
                    CausesValidation="false"
                    UseSubmitBehavior="false"
                    OnClientClick="AtomPage.beginProcessing(this);" />
                <asp:Button ID="btnConfirm3" runat="server" CssClass="atom-action-button"
                    CausesValidation="false" Visible="false"
                    UseSubmitBehavior="false"
                    OnClientClick="AtomPage.beginProcessing(this);" />
            </div>
        </div>
    </asp:Panel>

</asp:Content>
