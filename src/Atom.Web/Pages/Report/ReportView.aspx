<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ReportView.aspx.vb" Inherits="Atom.Web.Pages.Report.ReportView" %>
<!DOCTYPE html>
<%--
    SVF が出力した PDF をそのまま返す画面です。
    帳票が 1 つだけのときは中身を返すため、ここは表示されません。
    PDF が無いときと、複数あって選んでもらうときだけ画面になります。

    別のタブで開くため Site.Master は使いません。共通ヘッダーも出しません。
    そのぶん、この画面だけのスタイルをここに書いています。
--%>
<html lang="ja">
<head runat="server">
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>帳票の表示</title>
    <link rel="stylesheet" href="~/Content/site.css" runat="server" />
    <%-- 色・角丸・影は site.css の共通変数を参照し、ここで直接書きません。 --%>
    <style type="text/css">
        .atom-report-view {
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 100vh;
            padding: 24px 16px;
        }

        .atom-report-box {
            width: 100%;
            max-width: 420px;
            padding: 32px 28px;
            border-radius: var(--atom-radius-l);
            background-color: var(--atom-surface);
            box-shadow: var(--atom-shadow);
            text-align: center;
        }

        .atom-report-title {
            margin: 0 0 6px;
            color: var(--atom-accent);
            font-size: 18px;
            font-weight: bold;
        }

        .atom-report-message {
            margin: 0;
            color: var(--atom-text-muted);
            font-size: 13px;
            line-height: 1.8;
        }

        /* 帳票を選ぶリンク。縦に積んで、どれも同じ幅に見えるようにします。 */
        .atom-report-list {
            display: flex;
            flex-direction: column;
            gap: 8px;
            margin: 22px 0 0;
            padding: 0;
            list-style: none;
        }

        .atom-report-list .atom-action-button {
            width: 100%;
        }
    </style>
</head>
<body>
    <form id="frmReportView" runat="server">
        <div class="atom-report-view">
            <div class="atom-report-box">
                <p class="atom-report-title">帳票</p>
                <p class="atom-report-message">
                    <asp:Literal ID="litMessage" runat="server" Mode="Encode" />
                </p>

                <%-- 出力されている帳票が複数あるときだけ出します。 --%>
                <asp:Repeater ID="rptReports" runat="server" Visible="false">
                    <HeaderTemplate>
                        <ul class="atom-report-list">
                    </HeaderTemplate>
                    <ItemTemplate>
                        <li>
                            <asp:HyperLink ID="lnkReport" runat="server"
                                CssClass="atom-action-button atom-action-link"
                                Text='<%# Eval("DisplayName") %>'
                                NavigateUrl='<%# CreateReportUrl(CStr(Eval("ReportKubun"))) %>' />
                        </li>
                    </ItemTemplate>
                    <FooterTemplate>
                        </ul>
                    </FooterTemplate>
                </asp:Repeater>
            </div>
        </div>
    </form>
</body>
</html>
