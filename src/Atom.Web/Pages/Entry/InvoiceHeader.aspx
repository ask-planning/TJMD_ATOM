<%@ Page Title="INVOICEヘッダー入力" Language="vb" MasterPageFile="~/Site.Master" AutoEventWireup="false" CodeBehind="InvoiceHeader.aspx.vb" Inherits="Atom.Web.Pages.Entry.InvoiceHeader" %>

<asp:Content ID="cntHead" ContentPlaceHolderID="cphHead" runat="server">
    <%-- この画面だけで使うスタイルです。共通のスタイルは Content/site.css にあります。
         色は site.css の共通変数を参照し、この画面で HEX 値を直接書かないようにします。 --%>
    <style type="text/css">
        /* 左に入力、右に操作。画面が狭くなると下へ回り込みます。
           広い画面では入力の列が伸びて右の操作列と離れてしまうため、
           2 列ぶんの幅で止めて中央へ寄せます。 */
        .atom-header-area {
            display: flex;
            flex-wrap: wrap;
            align-items: flex-start;
            justify-content: center;
            gap: 20px;
            max-width: 1072px;
            margin-right: auto;
            margin-bottom: 24px;
            margin-left: auto;
        }

        .atom-header-inputs {
            flex: 1 1 520px;
            min-width: 420px;
            max-width: 780px;
        }

        /* 操作の列もカードに載せて、入力のカードと高さを揃えます。 */
        .atom-header-side {
            flex: 0 0 272px;
            padding: 22px;
            border-radius: var(--atom-radius-l);
            background-color: var(--atom-surface);
            box-shadow: var(--atom-shadow);
        }

        /* 「操作」「帳票」の見出し。用途が違うボタンを混ぜないための仕切りです。 */
        .atom-side-title {
            margin: 0 0 10px;
            color: var(--atom-text-muted);
            font-size: 12px;
            font-weight: bold;
        }

        .atom-side-title + * + .atom-side-title,
        .atom-report-row + .atom-side-title {
            margin-top: 20px;
        }

        .atom-header-buttons {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-bottom: 0;
        }

        /* 帳票は「出す」と「見る」で 1 組。横に並べて組だと分かるようにします。 */
        .atom-report-row {
            display: flex;
            gap: 6px;
            margin-bottom: 8px;
        }

        .atom-report-main {
            flex: 1 1 auto;
            padding-right: 8px;
            padding-left: 8px;
        }

        .atom-report-view {
            flex: 0 0 72px;
            padding-right: 0;
            padding-left: 0;
        }

        .atom-action-wide {
            width: 100%;
        }

        /* 入荷明細書の確認（§11.2）。押した結果ここに選択肢が出るため、
           帳票ボタンのすぐ下に、囲みを付けて別のまとまりだと分かるようにします。 */
        .atom-confirm-panel {
            margin-top: 8px;
            padding: 10px;
            border: 1px solid var(--atom-line-strong);
            border-radius: 4px;
            background-color: var(--atom-surface-soft);
            font-size: 12px;
            line-height: 1.6;
        }

        .atom-confirm-actions {
            display: flex;
            flex-direction: column;
            gap: 6px;
            margin-top: 8px;
        }

        /* 未実装であることを伝える一言。ボタンの真下に小さく出します。 */
        .atom-side-note {
            margin: 6px 0 0;
            color: var(--atom-text-muted);
            font-size: 11.5px;
        }

        /* 画面の使い方の説明。本文と混ざらないよう薄い面に載せます。 */
        .atom-guide {
            margin-top: 20px;
            padding: 14px 16px;
            border: 0;
            border-radius: var(--atom-radius);
            background-color: var(--atom-surface-soft);
            color: #475467;
            font-size: 12.5px;
            line-height: 1.8;
        }

        .atom-guide p {
            margin: 0 0 4px;
        }

        .atom-guide p:last-child {
            margin-bottom: 0;
        }

        /* 明細から集計した値。入力欄と区別できるよう 1 本の帯にまとめます。 */
        .atom-total-row {
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 20px;
            margin-top: 4px;
            padding: 14px 16px;
            border-radius: var(--atom-radius);
            background-color: var(--atom-surface-soft);
        }

        .atom-total {
            display: flex;
            align-items: center;
            gap: 6px;
            color: var(--atom-text-muted);
            font-size: 12px;
        }

        /* 帯の中の数値。読み取り専用の欄は枠を消して、文字だけに見せます。 */
        .atom-total .atom-readonly {
            width: 96px;
            padding: 4px 8px;
            border-color: transparent;
            background-color: transparent;
            color: var(--atom-text);
            font-size: 14px;
            text-align: right;
        }

        .atom-total-unit {
            font-size: 11.5px;
        }

        .atom-detail-head {
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 14px;
            margin-bottom: 12px;
        }

        .atom-detail-note {
            margin: 0;
            color: var(--atom-text-muted);
            font-size: 12px;
        }

        .atom-detail-actions {
            margin-left: auto;
        }

        /* 一覧は白い面に載せます。列が多いため、縮めずに横へスクロールさせます。 */
        .atom-grid-area {
            max-height: 50vh;
            overflow: auto;
            padding: 0;
            border: 0;
            border-radius: var(--atom-radius-l);
            background-color: var(--atom-surface);
            box-shadow: var(--atom-shadow);
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

                <%-- 明細から集計する項目。返却パレット以外は入力できません。
                     入力欄と区別できるよう、カードの中で帯にまとめています。 --%>
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

        </div>

        <%-- 右：操作ボタン。用途ごとに分けて並べます。 --%>
        <div class="atom-header-side">

            <p class="atom-side-title">操作</p>
            <div class="atom-header-buttons">
                <asp:Button ID="btnRegister" runat="server" Text="登 録"
                    CssClass="atom-action-button atom-action-primary atom-action-wide"
                    UseSubmitBehavior="false" OnClientClick="AtomPage.beginProcessing(this);" />
            </div>

            <p class="atom-side-title">帳票</p>
            <%-- 入荷明細書は、押した時点では件数を調べるだけです。その 1 往復で画面全体が
                 描き直されると、ちらついて落ち着きません。ここだけを差し替えます。
                 メッセージ欄は Site.Master 側の UpdatePanel が受け持つため、ここには含めません。 --%>
            <asp:UpdatePanel ID="upnReport" runat="server" UpdateMode="Conditional">
                <ContentTemplate>
            <%-- 帳票は「出す」と「見る」で 1 組です。横に並べて組だと分かるようにします。 --%>
            <div class="atom-report-row">
                <%-- 設計書 §10.2 の確認ダイアログ。「いいえ」のときは送信しません。 --%>
                <asp:Button ID="btnHaisoTehai" runat="server" Text="配送手配書"
                    CssClass="atom-action-button atom-report-main"
                    UseSubmitBehavior="false"
                    OnClientClick="if (!confirm('配送タイプ、コンテナ区分、コンテナNOの入力は済んでいますか。')) { return false; } AtomPage.beginProcessing(this);" />
                <%-- ポップアップブロックを避けるため、ボタンではなくリンクで別タブに開きます。
                     ボタンのポストバック後に window.open を呼ぶと、利用者の操作と見なされず塞がれます。
                     PDF が出ているかの判定は遷移先（ReportView）で行います。 --%>
                <asp:HyperLink ID="lnkHaisoTehaiView" runat="server" Text="表示"
                    CssClass="atom-action-button atom-action-link atom-report-view" Target="_blank"
                    NavigateUrl="~/Pages/Report/ReportView.aspx" />
            </div>

            <div class="atom-report-row">
                <asp:Button ID="btnNyukaMeisai" runat="server" Text="入荷明細書"
                    CssClass="atom-action-button atom-report-main"
                    UseSubmitBehavior="false" OnClientClick="AtomPage.beginProcessing(this);" />
                <%-- 配送手配書と同じ理由でリンクにしています（ポップアップブロック対策）。
                     **帳票区分を必ず付けます。**付けないと、先に見つかった別の帳票の PDF が開きます。
                     値は HaisoTehaiForm の入荷明細書の帳票区分（NYUKA）と揃えること。 --%>
                <asp:HyperLink ID="lnkNyukaMeisaiView" runat="server" Text="表示"
                    CssClass="atom-action-button atom-action-link atom-report-view" Target="_blank"
                    NavigateUrl="~/Pages/Report/ReportView.aspx?kubun=NYUKA" />
            </div>

            <%-- 設計書 §11.2 の確認です。作成済みがあるかで選択肢が変わるため、
                 両方のパターンをここに書いておき、Visible で出し分けます。
                 旧 Access は 1 / 2 / 3 を打たせるダイアログでしたが、
                 Web ではボタンのほうが誤入力が起きないため、選択肢ごとのボタンにしています。 --%>
            <asp:Panel ID="pnlNyukaConfirm" runat="server" Visible="false" CssClass="atom-confirm-panel">
                <asp:Literal ID="litNyukaConfirmMessage" runat="server" Mode="Encode" />
                <div class="atom-confirm-actions">
                    <asp:Button ID="btnNyukaRecreate" runat="server" Text="作り直す"
                        CssClass="atom-action-button"
                        UseSubmitBehavior="false"
                        OnClientClick="if (!confirm('作成済みの入荷明細を削除して作り直します。よろしいですか。')) { return false; } AtomPage.beginProcessing(this);" />
                    <asp:Button ID="btnNyukaGoConfirm" runat="server" Text="内容を確認する"
                        CssClass="atom-action-button"
                        UseSubmitBehavior="false" OnClientClick="AtomPage.beginProcessing(this);" />
                    <asp:Button ID="btnNyukaPrint" runat="server" Text="印刷する"
                        CssClass="atom-action-button atom-action-primary"
                        UseSubmitBehavior="false" OnClientClick="AtomPage.beginProcessing(this);" />
                    <asp:Button ID="btnNyukaCreateYes" runat="server" Text="はい"
                        CssClass="atom-action-button atom-action-primary"
                        UseSubmitBehavior="false" OnClientClick="AtomPage.beginProcessing(this);" />
                    <asp:Button ID="btnNyukaCancel" runat="server" Text="いいえ"
                        CssClass="atom-action-button" CausesValidation="false"
                        UseSubmitBehavior="false" />
                </div>
            </asp:Panel>
                </ContentTemplate>
            </asp:UpdatePanel>

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

    <%-- 旧 Access 画面の文言をそのまま踏襲しています。
         操作ボタンの近くではなく画面の最後へ置き、作業の邪魔にならないようにしました。 --%>
    <div class="atom-guide">
        <p>フィールドがグレーの場合、ヘッダーはまだ<strong>未登録</strong>です。登録ボタンで登録してください。</p>
        <p>明細更新でコンテナ番号を更新できるのは<strong>1 コンテナサイズ、1 コンテナNO の場合のみ</strong>です。同サイズのコンテナが複数ある場合は更新できません。</p>
    </div>

</asp:Content>
