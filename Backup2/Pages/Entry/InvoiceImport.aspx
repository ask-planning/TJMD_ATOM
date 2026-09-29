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
            </div>
        </div>

        <div class="atom-form-row">
            <asp:Label ID="lblPackingFile" runat="server" CssClass="atom-form-label"
                AssociatedControlID="fuPacking" Text="PACKINGファイル" />
            <div class="atom-form-field">
                <asp:FileUpload ID="fuPacking" runat="server" CssClass="atom-file" accept=".xls,.xlsx" />
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
                CausesValidation="false"
                UseSubmitBehavior="false"
                OnClientClick="return AtomPage.beginProcessing(this);" />
            <asp:Button ID="btnRerun" runat="server" Text="再実行"
                CssClass="atom-action-button"
                CausesValidation="false"
                UseSubmitBehavior="false"
                OnClientClick="return AtomPage.beginProcessing(this);" />
            <asp:Button ID="btnKakunin" runat="server" Text="取込内容確認"
                CssClass="atom-action-button"
                CausesValidation="false" />
        </div>

    </div>

    <asp:Panel ID="pnlProcessing" runat="server" CssClass="atom-processing" Visible="false">
        INVOICE/PACKINGファイルを取り込んでいます。しばらくお待ちください。
    </asp:Panel>

    <asp:Panel ID="pnlConfirm" runat="server" CssClass="atom-modal" Visible="false">
        <div class="atom-modal-box">
            <p class="atom-modal-message">
                <asp:Literal ID="litConfirmMessage" runat="server" Mode="Encode" />
            </p>
            <div class="atom-modal-buttons">
                <asp:Button ID="btnConfirm1" runat="server" CssClass="atom-action-button atom-action-primary"
                    CausesValidation="false" />
                <asp:Button ID="btnConfirm2" runat="server" CssClass="atom-action-button"
                    CausesValidation="false" />
                <asp:Button ID="btnConfirm3" runat="server" CssClass="atom-action-button"
                    CausesValidation="false" Visible="false" />
            </div>
        </div>
    </asp:Panel>

</asp:Content>
