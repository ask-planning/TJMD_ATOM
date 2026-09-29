<%@ Page Title="テキストデータ取込結果" Language="vb" MasterPageFile="~/Site.Master" AutoEventWireup="false" CodeBehind="ImportResult.aspx.vb" Inherits="Atom.Web.Pages.Entry.ImportResult" %>

<asp:Content ID="cntHead" ContentPlaceHolderID="cphHead" runat="server">
    <%-- この画面だけで使うスタイルです。共通のスタイルは Content/site.css にあります。
         色は site.css の共通変数を参照し、この画面で HEX 値を直接書かないようにします。 --%>
    <style type="text/css">
        .atom-result-toolbar {
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 16px;
            margin-bottom: 12px;
        }

        .atom-result-filter {
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .atom-result-tabs {
            display: flex;
            gap: 4px;
            margin-left: auto;
        }

        .atom-result-count {
            color: var(--atom-disabled-text);
            font-size: 13px;
        }

        .atom-grid-area {
            max-height: 60vh;
            overflow: auto;
            padding: 2px;
            border: 1px solid var(--atom-menu-border);
            background-color: var(--atom-grid-area-bg);
        }

        .atom-result-note {
            margin: 0 0 12px;
            font-size: 13px;
        }
    </style>
</asp:Content>

<asp:Content ID="cntContent" ContentPlaceHolderID="cphContent" runat="server">

    <h1 class="atom-page-title">テキストデータ取込結果</h1>

    <p class="atom-result-note">
        入力欄を直すと、その場で取込データ（一時データ）へ反映されます。ERR_CONTENTS は変更できません。<br />
        品目コード・INVOICE_NO は INVOICE と PACKING の突合キーです。片方だけ直すと突合できなくなるため、
        両方の一覧で同じ値に直してください。<br />
        直し終えたら<strong>取込画面のタブに戻り、「チェック→登録」</strong>を押してください。
        取込画面で「取り込み」を押すと Excel を読み直すため、<strong>ここで直した内容は消えます。</strong>
    </p>

    <asp:UpdatePanel ID="upnResult" runat="server" UpdateMode="Conditional" ChildrenAsTriggers="true">
        <ContentTemplate>

            <div class="atom-result-toolbar">
                <div class="atom-result-filter">
                    <asp:Label ID="lblInvoiceNo" runat="server" AssociatedControlID="ddlInvoiceNo" Text="元INV_NO" />
                    <asp:DropDownList ID="ddlInvoiceNo" runat="server" CssClass="atom-select" AutoPostBack="true" />
                    <asp:Label ID="lblCount" runat="server" CssClass="atom-result-count" />
                </div>

                <div class="atom-result-tabs">
                    <asp:LinkButton ID="lnkTabInvoice" runat="server" CssClass="atom-tab atom-tab-active"
                        CausesValidation="false" Text="INVOICE" />
                    <asp:LinkButton ID="lnkTabPacking" runat="server" CssClass="atom-tab"
                        CausesValidation="false" Text="PACKING" />
                </div>
            </div>

            <asp:Panel ID="pnlInvoice" runat="server" CssClass="atom-grid-area">
                <asp:GridView ID="gvInvoice" runat="server"
                    AutoGenerateColumns="False"
                    DataKeyNames="SeqNo"
                    CssClass="atom-grid"
                    GridLines="Both"
                    UseAccessibleHeader="true"
                    EmptyDataText="対象データがありません。">
                    <Columns>
                        <asp:TemplateField HeaderText="ERR_CONTENTS" ItemStyle-CssClass="atom-grid-err">
                            <ItemTemplate>
                                <asp:Literal ID="litErrContents" runat="server" Mode="Encode"
                                    Text='<%# Eval("ErrContents") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="数量" ItemStyle-CssClass="atom-grid-number">
                            <ItemTemplate>
                                <asp:TextBox ID="txtSuryo" runat="server" CssClass="atom-cell atom-cell-number"
                                    AutoPostBack="true" Text='<%# Eval("Suryo", "{0:#,##0}") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="単位">
                            <ItemTemplate>
                                <asp:TextBox ID="txtSuryoTani" runat="server" CssClass="atom-cell atom-cell-short"
                                    MaxLength="4" AutoPostBack="true" Text='<%# Eval("SuryoTani") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="通貨">
                            <ItemTemplate>
                                <asp:TextBox ID="txtTsuka" runat="server" CssClass="atom-cell atom-cell-short"
                                    MaxLength="3" AutoPostBack="true" Text='<%# Eval("Tsuka") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="単価" ItemStyle-CssClass="atom-grid-number">
                            <ItemTemplate>
                                <asp:TextBox ID="txtUnitPrice" runat="server" CssClass="atom-cell atom-cell-number"
                                    AutoPostBack="true" Text='<%# Eval("UnitPrice", "{0:#,##0.000}") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="金額" ItemStyle-CssClass="atom-grid-number">
                            <ItemTemplate>
                                <asp:TextBox ID="txtAmount" runat="server" CssClass="atom-cell atom-cell-number"
                                    AutoPostBack="true" Text='<%# Eval("Amount", "{0:#,##0.00}") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="購買伝票#">
                            <ItemTemplate>
                                <asp:TextBox ID="txtKoubaiDenpyoNo" runat="server" CssClass="atom-cell"
                                    MaxLength="50" AutoPostBack="true" Text='<%# Eval("R3KoubaiDenpyoNo") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="INVOICE#">
                            <ItemTemplate>
                                <asp:TextBox ID="txtInvoiceNo" runat="server" CssClass="atom-cell"
                                    MaxLength="50" AutoPostBack="true" Text='<%# Eval("InvoiceNo") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="B/L DATE">
                            <ItemTemplate>
                                <asp:TextBox ID="txtBlDate" runat="server" CssClass="atom-cell"
                                    MaxLength="10" AutoPostBack="true" Text='<%# Eval("BlDate", "{0:yyyy/MM/dd}") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="出荷方法">
                            <ItemTemplate>
                                <asp:TextBox ID="txtSyukkaHoho" runat="server" CssClass="atom-cell atom-cell-short"
                                    MaxLength="50" AutoPostBack="true" Text='<%# Eval("SyukkaHoho") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <%-- 品目コードは PACKING との突合キーです。片方だけ直すと突合が外れるため、
                             両方の一覧で編集できるようにしています。 --%>
                        <asp:TemplateField HeaderText="品目コード">
                            <ItemTemplate>
                                <asp:TextBox ID="txtShanghaiCode" runat="server" CssClass="atom-cell"
                                    MaxLength="50" AutoPostBack="true" Text='<%# Eval("ShanghaiCode") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="品名">
                            <ItemTemplate>
                                <asp:TextBox ID="txtShanghaiCodeText" runat="server" CssClass="atom-cell atom-cell-wide"
                                    MaxLength="128" AutoPostBack="true" Text='<%# Eval("ShanghaiCodeText") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </asp:Panel>

            <asp:Panel ID="pnlPacking" runat="server" CssClass="atom-grid-area" Visible="false">
                <%-- atom-grid-sub を付けると入力欄が薄い黄色になり、INVOICE 側と見分けられます。 --%>
                <asp:GridView ID="gvPacking" runat="server"
                    AutoGenerateColumns="False"
                    DataKeyNames="SeqNo"
                    CssClass="atom-grid atom-grid-sub"
                    GridLines="Both"
                    UseAccessibleHeader="true"
                    EmptyDataText="対象データがありません。">
                    <Columns>
                        <asp:TemplateField HeaderText="ERR_CONTENTS" ItemStyle-CssClass="atom-grid-err">
                            <ItemTemplate>
                                <asp:Literal ID="litErrContents" runat="server" Mode="Encode"
                                    Text='<%# Eval("ErrContents") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="INVOICE_NO">
                            <ItemTemplate>
                                <asp:TextBox ID="txtInvoiceNo" runat="server" CssClass="atom-cell"
                                    MaxLength="50" AutoPostBack="true" Text='<%# Eval("InvoiceNo") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="PO">
                            <ItemTemplate>
                                <asp:TextBox ID="txtTajimaPoNo" runat="server" CssClass="atom-cell"
                                    MaxLength="50" AutoPostBack="true" Text='<%# Eval("TajimaPoNo") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="T_F">
                            <ItemTemplate>
                                <asp:TextBox ID="txtTrueOrFalse" runat="server" CssClass="atom-cell atom-cell-short"
                                    MaxLength="10" AutoPostBack="true" Text='<%# Eval("TrueOrFalse") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="品目コード">
                            <ItemTemplate>
                                <asp:TextBox ID="txtShanghaiCode" runat="server" CssClass="atom-cell"
                                    MaxLength="50" AutoPostBack="true" Text='<%# Eval("ShanghaiCode") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="出荷方法">
                            <ItemTemplate>
                                <asp:TextBox ID="txtSyukkaHoho" runat="server" CssClass="atom-cell atom-cell-short"
                                    MaxLength="50" AutoPostBack="true" Text='<%# Eval("SyukkaHoho") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="BL_DATE">
                            <ItemTemplate>
                                <asp:TextBox ID="txtBlDate" runat="server" CssClass="atom-cell"
                                    MaxLength="10" AutoPostBack="true" Text='<%# Eval("BlDate", "{0:yyyy/MM/dd}") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="数量" ItemStyle-CssClass="atom-grid-number">
                            <ItemTemplate>
                                <asp:TextBox ID="txtSuryo" runat="server" CssClass="atom-cell atom-cell-number"
                                    AutoPostBack="true" Text='<%# Eval("Suryo", "{0:#,##0}") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="NET" ItemStyle-CssClass="atom-grid-number">
                            <ItemTemplate>
                                <asp:TextBox ID="txtNetWeight" runat="server" CssClass="atom-cell atom-cell-number"
                                    AutoPostBack="true" Text='<%# Eval("NetWeight", "{0:#,##0.###}") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="GRS" ItemStyle-CssClass="atom-grid-number">
                            <ItemTemplate>
                                <asp:TextBox ID="txtGrossWeight" runat="server" CssClass="atom-cell atom-cell-number"
                                    AutoPostBack="true" Text='<%# Eval("GrossWeight", "{0:#,##0.###}") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="M3" ItemStyle-CssClass="atom-grid-number">
                            <ItemTemplate>
                                <asp:TextBox ID="txtM3" runat="server" CssClass="atom-cell atom-cell-number"
                                    AutoPostBack="true" Text='<%# Eval("M3", "{0:#,##0.###}") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="C/NO 開始">
                            <ItemTemplate>
                                <asp:TextBox ID="txtCartonNoFrom" runat="server" CssClass="atom-cell atom-cell-short"
                                    MaxLength="50" AutoPostBack="true" Text='<%# Eval("CartonNoFrom") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="C/NO 終了">
                            <ItemTemplate>
                                <asp:TextBox ID="txtCartonNoTo" runat="server" CssClass="atom-cell atom-cell-short"
                                    MaxLength="50" AutoPostBack="true" Text='<%# Eval("CartonNoTo") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="単価" ItemStyle-CssClass="atom-grid-number">
                            <ItemTemplate>
                                <asp:TextBox ID="txtUnitPrice" runat="server" CssClass="atom-cell atom-cell-number"
                                    AutoPostBack="true" Text='<%# Eval("UnitPrice", "{0:#,##0.000}") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="合計" ItemStyle-CssClass="atom-grid-number">
                            <ItemTemplate>
                                <asp:TextBox ID="txtAmount" runat="server" CssClass="atom-cell atom-cell-number"
                                    AutoPostBack="true" Text='<%# Eval("Amount", "{0:#,##0.00}") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </asp:Panel>

        </ContentTemplate>
    </asp:UpdatePanel>

    <div class="atom-form-buttons atom-result-buttons">
        <asp:Button ID="btnExcel" runat="server" Text="EXCELへ"
            CssClass="atom-action-button" CausesValidation="false" />
    </div>

</asp:Content>
