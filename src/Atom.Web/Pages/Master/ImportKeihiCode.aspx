<%@ Page Title="輸入経費コード登録" Language="vb" MasterPageFile="~/Site.Master" AutoEventWireup="false" CodeBehind="ImportKeihiCode.aspx.vb" Inherits="Atom.Web.Pages.Master.ImportKeihiCode" %>

<asp:Content ID="cntHead" ContentPlaceHolderID="cphHead" runat="server">
    <%-- この画面だけで使うスタイルです。共通のスタイルは Content/site.css にあります。
         色は site.css の共通変数を参照し、この画面で HEX 値を直接書かないようにします。 --%>
    <style type="text/css">
        /* ページの幅いっぱいを使います。`.atom-content` 側が中央寄せなので、
           ここで幅を絞ると「操作」列がスクロールの外へ押し出されます。 */
        .atom-master-area {
            width: 100%;
        }

        .atom-master-note {
            margin: 0 0 12px;
            color: var(--atom-text-muted);
            font-size: 12px;
        }

        /* 列幅を下の .atom-col-* で決め打ちします。
           内容に合わせる既定（table-layout: auto）だと、経費名称の長い行に
           引きずられて表が画面より広くなり、右端の「操作」列が隠れるためです。 */
        .atom-master-area .atom-grid {
            width: 100%;
            table-layout: fixed;
        }

        /* 列が 7 つあります。左右の余白を詰めます。 */
        .atom-master-area .atom-grid th,
        .atom-master-area .atom-grid td {
            padding-right: 8px;
            padding-left: 8px;
        }

        /* 経費名称と REF は 50 文字まで入ります。1 行に収まらないときは
           切らずに折り返します（「…」で省略すると末尾が読めなくなるため）。
           site.css の .atom-grid td は nowrap なので、この 2 列だけ戻します。 */
        .atom-master-area .atom-grid td.atom-col-name,
        .atom-master-area .atom-grid td.atom-col-ref {
            white-space: normal;
            word-break: break-word;
        }

        /* ---- 列幅。合計で 100% になるようにしています ---- */
        .atom-col-code {
            width: 90px;
        }

        /* 経費名称。残りを受け持つため幅を指定しません。 */
        .atom-col-ref {
            width: 240px;
        }

        .atom-col-kbn {
            width: 150px;
        }

        .atom-col-flg {
            width: 80px;
            text-align: right;
        }

        .atom-col-date {
            width: 140px;
        }

        /* 「操作」列。ボタンが折り返したり切れたりしないよう場所を確保します。 */
        .atom-col-action {
            width: 150px;
            white-space: nowrap;
        }

        /* ---- 入力欄。列の幅に収めます ---- */
        .atom-cell-code {
            width: 100%;
        }

        .atom-cell-text {
            width: 100%;
        }

        .atom-cell-kbn {
            width: 100%;
        }

        .atom-cell-flg {
            width: 100%;
            text-align: right;
        }

        /* 一覧の中に置くボタン。行の高さを押し上げないよう小さくします。 */
        .atom-action-small {
            padding: 3px 10px;
            font-size: 11.5px;
        }

        /* 一覧の縦スクロール。見出し行は site.css の .atom-grid th が
           position: sticky なので、この箱を付けるだけで固定されます。 */
        .atom-grid-scroll {
            max-height: 460px;
            overflow-y: auto;
            border: 1px solid var(--atom-line);
            border-radius: var(--atom-radius-s);
            background-color: var(--atom-surface);
        }

        /* 新規登録の行。一覧と区別が付くよう、薄い面に載せます。 */
        .atom-new-row {
            margin-bottom: 16px;
            padding: 12px;
            border: 1px solid var(--atom-line-strong);
            border-radius: 4px;
            background-color: var(--atom-surface-soft);
        }

        .atom-new-row-title {
            margin: 0 0 10px;
            color: var(--atom-text-muted);
            font-size: 12px;
            font-weight: bold;
        }

        .atom-new-row-fields {
            display: flex;
            flex-wrap: wrap;
            align-items: flex-end;
            gap: 10px;
        }

        /* 新規登録の欄は表の列幅に縛られないため、ここで幅を決めます。 */
        .atom-new-field {
            display: flex;
            flex-direction: column;
            gap: 5px;
        }

        .atom-new-field label {
            color: var(--atom-text-muted);
            font-size: 11.5px;
        }

        .atom-new-field .atom-cell-code {
            width: 80px;
        }

        .atom-new-field .atom-cell-flg {
            width: 70px;
        }

        .atom-new-field .atom-cell {
            width: 220px;
        }
    </style>
</asp:Content>

<asp:Content ID="cntBody" ContentPlaceHolderID="cphContent" runat="server">

    <div class="atom-master-area">

        <%-- 編集・追加・削除のたびに画面全体が描き直されると見づらいため、
             この画面の操作はすべて UpdatePanel の中で完結させます。
             メッセージ欄は Site.Master 側の UpdatePanel が受け持ちます。 --%>
        <asp:UpdatePanel ID="upnKeihiCode" runat="server" UpdateMode="Conditional">
            <ContentTemplate>

                <%-- 新規登録。旧は一覧の最終行で入力していましたが、
                     Web では一覧の上に置いて入力欄を分けています。 --%>
                <div class="atom-new-row">
                    <p class="atom-new-row-title">新規登録</p>
                    <div class="atom-new-row-fields">
                        <div class="atom-new-field">
                            <label for="<%= Me.txtNewKeihiCode.ClientID %>">伝票コード</label>
                            <asp:TextBox ID="txtNewKeihiCode" runat="server"
                                CssClass="atom-cell atom-cell-code" MaxLength="4" />
                        </div>
                        <div class="atom-new-field">
                            <label for="<%= Me.txtNewKeihiCodeText.ClientID %>">経費名称</label>
                            <asp:TextBox ID="txtNewKeihiCodeText" runat="server"
                                CssClass="atom-cell" MaxLength="50" />
                        </div>
                        <div class="atom-new-field">
                            <label for="<%= Me.txtNewKeihiCodeRef.ClientID %>">REF</label>
                            <asp:TextBox ID="txtNewKeihiCodeRef" runat="server"
                                CssClass="atom-cell" MaxLength="50" />
                        </div>
                        <div class="atom-new-field">
                            <label for="<%= Me.ddlNewKeihiKbn.ClientID %>">区分</label>
                            <asp:DropDownList ID="ddlNewKeihiKbn" runat="server" CssClass="atom-select" />
                        </div>
                        <div class="atom-new-field">
                            <label for="<%= Me.txtNewHokenTaisyoFlg.ClientID %>">保険対象</label>
                            <asp:TextBox ID="txtNewHokenTaisyoFlg" runat="server"
                                CssClass="atom-cell atom-cell-flg" MaxLength="10" />
                        </div>
                        <asp:Button ID="btnAdd" runat="server" Text="追 加"
                            CssClass="atom-action-button atom-action-primary"
                            UseSubmitBehavior="false" OnClientClick="AtomPage.beginProcessing(this);" />
                    </div>
                </div>

                <p class="atom-master-note">
                    輸入諸経費で使う経費コードのマスタです。行ごとに編集・削除できます。
                </p>

                <%-- 一覧。行の編集は GridView の編集モードで行います。
                     旧 Access は行を移動した時点で自動保存でしたが、Web には同じ仕組みが無いため、
                     「編集」→「更新」の形にしています（設計書 C-15）。 --%>
                <div class="atom-grid-scroll">
                    <asp:GridView ID="gvKeihiCode" runat="server" CssClass="atom-grid"
                        AutoGenerateColumns="false" GridLines="None" UseAccessibleHeader="true"
                        DataKeyNames="KeihiCode" EmptyDataText="対象データがありません。">
                        <Columns>
                            <asp:TemplateField HeaderText="伝票コード" ItemStyle-CssClass="atom-col-code" HeaderStyle-CssClass="atom-col-code">
                                <ItemTemplate>
                                    <asp:Literal ID="litKeihiCode" runat="server" Mode="Encode" />
                                </ItemTemplate>
                                <EditItemTemplate>
                                    <%-- 旧は主キー列も編集できました。同じ挙動にしています（設計書 C-8）。 --%>
                                    <asp:TextBox ID="txtKeihiCode" runat="server"
                                        CssClass="atom-cell atom-cell-code" MaxLength="4" />
                                </EditItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="経費名称" ItemStyle-CssClass="atom-col-name" HeaderStyle-CssClass="atom-col-name">
                                <ItemTemplate>
                                    <asp:Literal ID="litKeihiCodeText" runat="server" Mode="Encode" />
                                </ItemTemplate>
                                <EditItemTemplate>
                                    <asp:TextBox ID="txtKeihiCodeText" runat="server"
                                        CssClass="atom-cell atom-cell-text" MaxLength="50" />
                                </EditItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="REF" ItemStyle-CssClass="atom-col-ref" HeaderStyle-CssClass="atom-col-ref">
                                <ItemTemplate>
                                    <asp:Literal ID="litKeihiCodeRef" runat="server" Mode="Encode" />
                                </ItemTemplate>
                                <EditItemTemplate>
                                    <asp:TextBox ID="txtKeihiCodeRef" runat="server"
                                        CssClass="atom-cell atom-cell-text" MaxLength="50" />
                                </EditItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="区分" ItemStyle-CssClass="atom-col-kbn" HeaderStyle-CssClass="atom-col-kbn">
                                <ItemTemplate>
                                    <asp:Literal ID="litKeihiKbn" runat="server" Mode="Encode" />
                                </ItemTemplate>
                                <EditItemTemplate>
                                    <asp:DropDownList ID="ddlKeihiKbn" runat="server" CssClass="atom-select atom-cell-kbn" />
                                </EditItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="保険対象" ItemStyle-CssClass="atom-col-flg" HeaderStyle-CssClass="atom-col-flg">
                                <ItemTemplate>
                                    <asp:Literal ID="litHokenTaisyoFlg" runat="server" Mode="Encode" />
                                </ItemTemplate>
                                <EditItemTemplate>
                                    <%-- 旧はテキストボックスで、0 / 1 以外も入れられました（設計書 C-10）。 --%>
                                    <asp:TextBox ID="txtHokenTaisyoFlg" runat="server"
                                        CssClass="atom-cell atom-cell-flg" MaxLength="10" />
                                </EditItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="更新日" ItemStyle-CssClass="atom-col-date" HeaderStyle-CssClass="atom-col-date">
                                <ItemTemplate>
                                    <asp:Literal ID="litUpdateYmd" runat="server" Mode="Encode" />
                                </ItemTemplate>
                                <EditItemTemplate>
                                    <asp:Literal ID="litUpdateYmdEdit" runat="server" Mode="Encode" />
                                </EditItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="操作" ItemStyle-CssClass="atom-col-action"
                                HeaderStyle-CssClass="atom-col-action">
                                <ItemTemplate>
                                    <asp:Button ID="btnEdit" runat="server" Text="編集"
                                        CssClass="atom-action-button atom-action-small"
                                        CommandName="Edit" CausesValidation="false"
                                        UseSubmitBehavior="false" />
                                    <%-- 旧に ATOM 独自の確認文言が無いため、Web 版で新設しています（設計書 C-29）。 --%>
                                    <asp:Button ID="btnDelete" runat="server" Text="削除"
                                        CssClass="atom-action-button atom-action-small"
                                        CommandName="Delete" CausesValidation="false"
                                        UseSubmitBehavior="false"
                                        OnClientClick="if (!confirm('選択した経費コードを削除します。よろしいですか？')) { return false; } AtomPage.beginProcessing(this);" />
                                </ItemTemplate>
                                <EditItemTemplate>
                                    <asp:Button ID="btnUpdate" runat="server" Text="更新"
                                        CssClass="atom-action-button atom-action-small atom-action-primary"
                                        CommandName="Update"
                                        UseSubmitBehavior="false" OnClientClick="AtomPage.beginProcessing(this);" />
                                    <asp:Button ID="btnCancel" runat="server" Text="取消"
                                        CssClass="atom-action-button atom-action-small"
                                        CommandName="Cancel" CausesValidation="false"
                                        UseSubmitBehavior="false" />
                                </EditItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </div>

            </ContentTemplate>
        </asp:UpdatePanel>

        <%-- 部分更新のたびに一覧が作り直されるため、そのままだと縦スクロールが先頭に戻ります。
             この画面だけの都合なので、共通の atom.js ではなくここに置いています。 --%>
        <script type="text/javascript">
            (function () {
                if (typeof Sys === "undefined" || !Sys.WebForms || !Sys.WebForms.PageRequestManager) {
                    return;
                }

                var position = 0;
                var manager = Sys.WebForms.PageRequestManager.getInstance();

                manager.add_beginRequest(function () {
                    var box = document.querySelector(".atom-grid-scroll");
                    position = box ? box.scrollTop : 0;
                });

                manager.add_endRequest(function () {
                    var box = document.querySelector(".atom-grid-scroll");
                    if (box) {
                        box.scrollTop = position;
                    }
                });
            })();
        </script>

    </div>

</asp:Content>
