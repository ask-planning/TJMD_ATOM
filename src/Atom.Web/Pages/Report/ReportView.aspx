<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ReportView.aspx.vb" Inherits="Atom.Web.Pages.Report.ReportView" %>
<!DOCTYPE html>
<%--
    SVF が出力した PDF をそのまま返す画面です。
    帳票が 1 つだけのときは中身を返すため、ここは表示されません。
    PDF が無いときと、複数あって選んでもらうときだけ画面になります。
--%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>帳票の表示</title>
    <link rel="stylesheet" href="../../Content/atom.css" />
</head>
<body>
    <form id="frmReportView" runat="server">
        <div class="atom-report-view">
            <p><asp:Literal ID="litMessage" runat="server" Mode="Encode" /></p>

            <%-- 出力されている帳票が複数あるときだけ出します。 --%>
            <asp:Repeater ID="rptReports" runat="server" Visible="false">
                <ItemTemplate>
                    <p>
                        <asp:HyperLink ID="lnkReport" runat="server" CssClass="atom-action-button"
                            Text='<%# Eval("DisplayName") %>'
                            NavigateUrl='<%# CreateReportUrl(CStr(Eval("ReportKubun"))) %>' />
                    </p>
                </ItemTemplate>
            </asp:Repeater>
        </div>
    </form>
</body>
</html>
