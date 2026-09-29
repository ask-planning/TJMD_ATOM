<%@ Page Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="ExpenseCode.aspx.vb" Inherits="Atom.Web.ExpenseCodeMaintenance" %>

<asp:Content ID="cntHead" ContentPlaceHolderID="cphHead" runat="server">
    <title>経費コード登録</title>
</asp:Content>

<asp:Content ID="cntContent" ContentPlaceHolderID="cphContent" runat="server">

    <asp:GridView ID="gvExpenseCode" runat="server"
        AutoGenerateColumns="False"
        DataKeyNames="KeihiCode"
        ShowFooter="True"
        CssClass="atom-grid"
        GridLines="Both"
        EmptyDataText="登録されている経費コードはありません。">

        <Columns>

            <asp:TemplateField HeaderText="経費コード" ItemStyle-CssClass="cell-code" FooterStyle-CssClass="cell-code">
                <ItemTemplate>
                    <asp:TextBox ID="txtKeihiCode" runat="server" MaxLength="10" CssClass="input-code" />
                    <asp:HiddenField ID="hdnOriginalKeihiCode" runat="server" />
                    <asp:HiddenField ID="hdnUpdateYmd" runat="server" />
                </ItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="txtNewKeihiCode" runat="server" MaxLength="10" CssClass="input-code" />
                </FooterTemplate>
            </asp:TemplateField>

            <asp:TemplateField HeaderText="経費名称">
                <ItemTemplate>
                    <asp:TextBox ID="txtKeihiCodeText" runat="server" MaxLength="100" CssClass="input-name" />
                </ItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="txtNewKeihiCodeText" runat="server" MaxLength="100" CssClass="input-name" />
                </FooterTemplate>
            </asp:TemplateField>

            <asp:TemplateField HeaderText="経費区分">
                <ItemTemplate>
                    <asp:TextBox ID="txtKeihiKbn" runat="server" MaxLength="10" CssClass="input-kubun" />
                </ItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="txtNewKeihiKbn" runat="server" MaxLength="10" CssClass="input-kubun" />
                </FooterTemplate>
            </asp:TemplateField>

            <asp:TemplateField HeaderText="更新日時" ItemStyle-CssClass="cell-readonly">
                <ItemTemplate>
                    <asp:Literal ID="litUpdateYmd" runat="server" Mode="Encode" />
                </ItemTemplate>
            </asp:TemplateField>

            <asp:TemplateField HeaderText="更新者" ItemStyle-CssClass="cell-readonly">
                <ItemTemplate>
                    <asp:Literal ID="litUpdateLogin" runat="server" Mode="Encode" />
                </ItemTemplate>
            </asp:TemplateField>

            <asp:TemplateField HeaderText="" ItemStyle-HorizontalAlign="Center" FooterStyle-HorizontalAlign="Center">
                <ItemTemplate>
                    <asp:LinkButton ID="lnkDelete" runat="server"
                        CssClass="atom-grid-link"
                        CommandName="DeleteRow"
                        CommandArgument='<%# Container.DataItemIndex %>'
                        CausesValidation="false"
                        Text="削除"
                        OnClientClick="return confirm('この行を削除します。よろしいですか。');" />
                </ItemTemplate>
                <FooterTemplate>
                    <asp:LinkButton ID="lnkAdd" runat="server"
                        CssClass="atom-grid-link"
                        CommandName="AddRow"
                        CausesValidation="false"
                        Text="追加" />
                </FooterTemplate>
            </asp:TemplateField>

        </Columns>
    </asp:GridView>

    <div class="atom-actions">
        <asp:Button ID="btnSave" runat="server" Text="保存" CssClass="atom-button-primary" />
        <asp:Button ID="btnReload" runat="server" Text="再表示" CssClass="atom-button" CausesValidation="false" />
    </div>

</asp:Content>
