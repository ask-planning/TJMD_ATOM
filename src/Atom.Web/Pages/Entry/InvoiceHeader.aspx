<%@ Page Title="INVOICEヘッダー入力" Language="vb" MasterPageFile="~/Site.Master" AutoEventWireup="false" CodeBehind="InvoiceHeader.aspx.vb" Inherits="Atom.Web.Pages.Entry.InvoiceHeader" %>

<asp:Content ID="cntHead" ContentPlaceHolderID="cphHead" runat="server">
    <%-- この画面だけで使うスタイルです。共通のスタイルは Content/site.css にあります。
         色は site.css の共通変数を参照し、この画面で HEX 値を直接書かないようにします。 --%>
    <style type="text/css">
        .atom-header-area {
            display: flex;
            flex-wrap: wrap;
            align-items: flex-start;
            gap: 24px;
            margin-bottom: 16px;
        }

        .atom-header-inputs {
            flex: 1 1 520px;
            min-width: 420px;
        }

        .atom-header-side {
            flex: 0 0 300px;
        }

        .atom-header-buttons {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-bottom: 12px;
        }

        .atom-guide {
            padding: 8px 12px;
            border: 1px solid var(--atom-menu-border);
            background-color: var(--atom-grid-area-bg);
            font-size: 12px;
            line-height: 1.6;
        }

        .atom-guide p {
            margin: 0 0 4px;
        }

        .atom-guide p:last-child {
            margin-bottom: 0;
        }

        .atom-total-row {
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 16px;
            padding: 8px 0;
        }

        .atom-total {
            display: flex;
            align-items: center;
            gap: 4px;
        }

        .atom-total-unit {
            font-size: 12px;
        }

        .atom-detail-head {
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 16px;
            margin-bottom: 8px;
        }

        .atom-detail-note {
            margin: 0;
            font-size: 12px;
        }

        .atom-detail-actions {
            margin-left: auto;
        }

        .atom-grid-area {
            max-height: 50vh;
            overflow: auto;
            padding: 2px;
            border: 1px solid var(--atom-menu-border);
            background-color: var(--atom-grid-area-bg);
        }
    </style>
</asp:Content>

<asp:Content ID="cntContent" ContentPlaceHolderID="cphContent" runat="server">

    <h1 class="atom-page-title">INVOICEヘッダー入力</h1>

    <div class="atom-header-area">

        <%-- 左：ヘッダー情報の入力 --%>
        <div class="atom-header-inputs">

            <div class="atom-form">

                <div class="atom-form-row">
                    <asp:Label ID="lblInvoiceNoMain" runat="server" CssClass="atom-form-label"
                        Text="元INV_NO" AssociatedControlID="ddlInvoiceNoMain" />
                    <div class="atom-form-field">
                        <asp:DropDownList ID="ddlInvoiceNoMain" runat="server"
                            CssClass="atom-select" AutoPostBack="true" />
                    </div>
                </div>

                <div class="atom-form-row">
                    <asp:Label ID="lblBlDate" runat="server" CssClass="atom-form-label"
                        Text="B/L DATE" AssociatedControlID="txtBlDate" />
                    <div class="atom-form-field">
                        <asp:TextBox ID="txtBlDate" runat="server" CssClass="atom-cell atom-cell-date"
                            TextMode="Date" />
                    </div>
                </div>

                <div class="atom-form-row">
                    <asp:Label ID="lblVessel" runat="server" CssClass="atom-form-label"
                        Text="船名/AIR便名" AssociatedControlID="ddlVessel" />
                    <div class="atom-form-field">
                        <asp:DropDownList ID="ddlVessel" runat="server" CssClass="atom-select" />
                    </div>
                </div>

                <div class="atom-form-row">
                    <asp:Label ID="lblSyukkaHoho" runat="server" CssClass="atom-form-label"
                        Text="出荷方法" AssociatedControlID="ddlSyukkaHoho" />
                    <div class="atom-form-field">
                        <asp:DropDownList ID="ddlSyukkaHoho" runat="server" CssClass="atom-select" />
                    </div>
                </div>

                <div class="atom-form-row">
                    <asp:Label ID="lblOtsunaka" runat="server" CssClass="atom-form-label"
                        Text="乙仲" AssociatedControlID="ddlOtsunaka" />
                    <div class="atom-form-field">
                        <asp:DropDownList ID="ddlOtsunaka" runat="server" CssClass="atom-select" />
                    </div>
                </div>

            </div>

            <%-- 明細から集計する項目。入力はできません。 --%>
            <div class="atom-total-row">
                <div class="atom-total">
                    <asp:Label ID="lblCartonQty" runat="server" Text="カートン数" AssociatedControlID="txtCartonQty" />
                    <asp:TextBox ID="txtCartonQty" runat="server"
                        CssClass="atom-readonly atom-cell-number" ReadOnly="true" />
                    <span class="atom-total-unit">CTNS</span>
                </div>
                <div class="atom-total">
                    <asp:Label ID="lblTotalSuryo" runat="server" Text="数量" AssociatedControlID="txtTotalSuryo" />
                    <asp:TextBox ID="txtTotalSuryo" runat="server"
                        CssClass="atom-readonly atom-cell-number" ReadOnly="true" />
                    <span class="atom-total-unit">PCS</span>
                </div>
                <div class="atom-total">
                    <asp:Label ID="lblTotalAmount" runat="server" Text="金額" AssociatedControlID="txtTotalAmount" />
                    <asp:TextBox ID="txtTotalAmount" runat="server"
                        CssClass="atom-readonly atom-cell-number" ReadOnly="true" />
                </div>
                <div class="atom-total">
                    <asp:Label ID="lblReturnPallet" runat="server" Text="返却パレット" AssociatedControlID="txtReturnPallet" />
                    <asp:TextBox ID="txtReturnPallet" runat="server"
                        CssClass="atom-cell atom-cell-number" />
                </div>
            </div>

        </div>

        <%-- 右：操作ボタンと注意書き --%>
        <div class="atom-header-side">

            <div class="atom-header-buttons">
                <asp:Button ID="btnRegister" runat="server" Text="登 録"
                    CssClass="atom-action-button atom-action-primary"
                    UseSubmitBehavior="false" OnClientClick="AtomPage.beginProcessing(this);" />
                <%-- 設計書 §10.2 の確認ダイアログ。「いいえ」のときは送信しません。 --%>
                <asp:Button ID="btnHaisoTehai" runat="server" Text="配送手配書"
                    CssClass="atom-action-button"
                    UseSubmitBehavior="false"
                    OnClientClick="if (!confirm('配送タイプ、コンテナ区分、コンテナNOの入力は済んでいますか。')) { return false; } AtomPage.beginProcessing(this);" />
                <%-- ポップアップブロックを避けるため、ボタンではなくリンクで別タブに開きます。
                     ボタンのポストバック後に window.open を呼ぶと、利用者の操作と見なされず塞がれます。
                     PDF が出ているかの判定は遷移先（ReportView）で行います。 --%>
                <asp:HyperLink ID="lnkHaisoTehaiView" runat="server" Text="表示"
                    CssClass="atom-action-button" Target="_blank"
                    NavigateUrl="~/Pages/Report/ReportView.aspx" />
            </div>

            <div class="atom-header-buttons">
                <asp:Button ID="btnNyukaMeisai" runat="server" Text="入荷明細書"
                    CssClass="atom-action-button"
                    UseSubmitBehavior="false" OnClientClick="AtomPage.beginProcessing(this);" />
                <asp:Button ID="btnNyukaMeisaiView" runat="server" Text="表示"
                    CssClass="atom-action-button" />
            </div>

            <%-- 旧 Access 画面の文言をそのまま踏襲しています。 --%>
            <div class="atom-guide">
                <p>●フィールドがグレーの場合、ヘッダーはまだ未登録です。登録ボタンで登録してください。</p>
                <p>●明細更新でコンテナ番号を更新できるのは 1 コンテナサイズ、1 コンテナNO の場合のみです。同サイズのコンテナが複数ある場合は更新できません。</p>
            </div>

        </div>

    </div>

    <%-- INVOICE詳細 --%>
    <div class="atom-detail-head">
        <strong>INVOICE詳細</strong>
        <p class="atom-detail-note">印刷用INVOICE_NOは、登録後自動編集します。</p>
        <div class="atom-detail-actions">
            <asp:Button ID="btnMeisaiUpdate" runat="server" Text="明細更新"
                CssClass="atom-action-button"
                UseSubmitBehavior="false" OnClientClick="AtomPage.beginProcessing(this);" />
        </div>
    </div>

    <div class="atom-grid-area">
        <asp:GridView ID="gvDetail" runat="server" CssClass="atom-grid"
            AutoGenerateColumns="false" GridLines="None" UseAccessibleHeader="true"
            EmptyDataText="元INV_NO を選択してください。">
            <Columns>
                <asp:TemplateField HeaderText="INVOICE_NO">
                    <ItemTemplate>
                        <asp:Literal ID="litInvoiceNo" runat="server" Mode="Encode" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="配送予定日">
                    <ItemTemplate>
                        <asp:TextBox ID="txtHaisoYoteiYmd" runat="server"
                            CssClass="atom-cell atom-cell-date" TextMode="Date" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="配送タイプ">
                    <ItemTemplate>
                        <asp:DropDownList ID="ddlHaisoType" runat="server" CssClass="atom-select" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="コンテナ（AIR=AIR）">
                    <ItemTemplate>
                        <asp:DropDownList ID="ddlContainerType" runat="server" CssClass="atom-select" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="コンテナNO（AIR=商品群）">
                    <ItemTemplate>
                        <asp:TextBox ID="txtContainerNo" runat="server"
                            CssClass="atom-cell" MaxLength="20" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="シールNO">
                    <ItemTemplate>
                        <asp:TextBox ID="txtSealNo" runat="server"
                            CssClass="atom-cell" MaxLength="20" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="カートン数">
                    <ItemTemplate>
                        <asp:TextBox ID="txtCartonQtySub" runat="server"
                            CssClass="atom-cell atom-cell-number" MaxLength="7" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="印刷用INVOICE_NO">
                    <ItemTemplate>
                        <asp:Literal ID="litInvoiceNoPrint" runat="server" Mode="Encode" />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
    </div>

</asp:Content>
