Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Configuration
Imports System.IO
Imports System.Web
Imports Atom.Core.Logging
Imports Atom.Report
Imports Atom.Web.Common

Namespace Pages.Report

    ''' <summary>
    ''' SVF が出力した PDF を表示する画面です。
    ''' </summary>
    ''' <remarks>
    ''' 帳票の「表示」ボタンから別のタブで開かれます。
    '''
    ''' **任意のファイルを読み出せないようにしています。**
    ''' 受け取るのは帳票区分だけで、定義済みの値かを確かめてから使います。
    ''' ファイル名のもう一方であるログインコードは**セッションの値**を使い、
    ''' 画面から渡された値は使いません。フォルダのパスも Web.config から取ります。
    ''' </remarks>
    Public Class ReportView
        Inherits BasePage

#Region "定数"
        ''' <summary>
        ''' この画面のメニュー番号です。
        ''' </summary>
        ''' <remarks>
        ''' **配送手配書を出せる画面（103）と同じ番号にします。**
        ''' 0 にすると権限判定が素通りし、103 の権限を外された利用者でも
        ''' 自分の PDF を開けてしまいます。
        ''' </remarks>
        Private Const ThisMenuNo As Integer = 103

        ''' <summary>ログの発生元名。</summary>
        Private Const LogSource As String = "ReportView"

        ''' <summary>帳票区分を受け取るクエリ名。</summary>
        Private Const ReportKubunKey As String = "kubun"

        ''' <summary>PDF の出力先フォルダのパスを保持する設定キー。</summary>
        Private Const PdfFolderKey As String = "Svf.PdfFolderPath"

        ''' <summary>PDF の MIME タイプ。</summary>
        Private Const PdfContentType As String = "application/pdf"

        ''' <summary>帳票区分が正しくないときの文言。</summary>
        Private Const InvalidKubunMessage As String = "表示できない帳票です。"

        ''' <summary>PDF がまだ無いときの文言。</summary>
        Private Const NotReadyMessage As String = "まだ印刷中です。少し待ってから開き直してください。"

        ''' <summary>設定が足りないときの文言。</summary>
        Private Const SettingMissingMessage As String = "帳票の設定が不足しています。管理者へ連絡してください。"

        ''' <summary>複数の帳票が出ているときの文言。</summary>
        Private Const SelectReportMessage As String = "開く帳票を選んでください。"
#End Region

#Region "フィールド"
        ''' <summary>SVF が出力した PDF の扱い。</summary>
        Private ReadOnly _pdfStore As New SvfPdfStore()
#End Region

#Region "イベント"
        ''' <summary>
        ''' 初期化時の処理です。メニュー番号を設定します。
        ''' </summary>
        ''' <param name="e">イベント引数。</param>
        ''' <remarks>
        ''' メニュー番号は権限判定より前に決める必要があるため、Page_Load ではなくここで設定します。
        ''' </remarks>
        Protected Overrides Sub OnInit(e As EventArgs)
            Me.MenuNo = ThisMenuNo
            MyBase.OnInit(e)
        End Sub

        ''' <summary>
        ''' 画面の読み込み時の処理です。PDF をそのまま返します。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
            ' 権限が無いとき
            If Me.IsAccessDenied Then
                Return
            End If

            Dim pdfFolderPath As String = GetSetting(PdfFolderKey)
            ' 設定が足りないとき
            If String.IsNullOrWhiteSpace(pdfFolderPath) Then
                Me.litMessage.Text = SettingMissingMessage
                Return
            End If

            Dim loginCode As String = SessionContext.Current.LoginCode
            Dim requestedKubun As String = Me.Request.QueryString(ReportKubunKey)

            ' 帳票区分の指定があるとき（一覧から選ばれた場合）
            If Not String.IsNullOrEmpty(requestedKubun) Then
                Me.ShowOne(requestedKubun, pdfFolderPath, loginCode)
                Return
            End If

            Me.ShowAvailable(pdfFolderPath, loginCode)
        End Sub
#End Region

#Region "内部処理（表示の振り分け）"
        ''' <summary>
        ''' 指定された帳票区分の PDF を表示します。
        ''' </summary>
        ''' <param name="reportKubun">帳票区分。</param>
        ''' <param name="pdfFolderPath">PDF の出力先フォルダ。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        Private Sub ShowOne(reportKubun As String, pdfFolderPath As String, loginCode As String)
            Dim form As HaisoTehaiForm = HaisoTehaiForm.FindByReportKubun(reportKubun)
            ' 定義に無い帳票区分のとき
            If form Is Nothing Then
                Me.litMessage.Text = InvalidKubunMessage
                Return
            End If

            Dim filePath As String = Me._pdfStore.CreatePath(pdfFolderPath, form.ReportKubun, loginCode)
            ' まだ出力されていないとき
            If Not File.Exists(filePath) Then
                Me.litMessage.Text = NotReadyMessage
                Return
            End If

            Me.TransmitPdf(filePath, form, loginCode)
        End Sub

        ''' <summary>
        ''' 出力されている配送手配書を探して表示します。
        ''' </summary>
        ''' <param name="pdfFolderPath">PDF の出力先フォルダ。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <remarks>
        ''' 配送先の箇所数によって様式が分かれるため、**1 回の出力で 2 つ以上できることがあります。**
        ''' 1 つだけならそのまま開き、複数あるときは選べるようにリンクを並べます。
        ''' </remarks>
        Private Sub ShowAvailable(pdfFolderPath As String, loginCode As String)
            Dim available As New List(Of HaisoTehaiForm)()
            For Each form As HaisoTehaiForm In HaisoTehaiForm.All()
                ' 出力されていないとき
                If Not Me._pdfStore.Exists(pdfFolderPath, form.ReportKubun, loginCode) Then
                    Continue For
                End If
                available.Add(form)
            Next

            ' 1 つも出力されていないとき
            If available.Count = 0 Then
                Me.litMessage.Text = NotReadyMessage
                Return
            End If

            ' 1 つだけのとき（選ばせずにそのまま開く）
            If available.Count = 1 Then
                Me.ShowOne(available(0).ReportKubun, pdfFolderPath, loginCode)
                Return
            End If

            Me.ShowLinks(available)
        End Sub

        ''' <summary>
        ''' 開ける帳票のリンクを並べます。
        ''' </summary>
        ''' <param name="forms">出力されている様式。</param>
        ''' <remarks>
        ''' 帳票区分も表示名も様式の定義から取るため、利用者の入力は混ざりません。
        ''' </remarks>
        Private Sub ShowLinks(forms As IList(Of HaisoTehaiForm))
            Me.litMessage.Text = SelectReportMessage
            Me.rptReports.DataSource = forms
            Me.rptReports.DataBind()
            Me.rptReports.Visible = True
        End Sub

        ''' <summary>
        ''' 帳票を開くリンクの URL を組み立てます。一覧の各行から呼び出します。
        ''' </summary>
        ''' <param name="reportKubun">帳票区分。</param>
        ''' <returns>この画面への URL。</returns>
        Protected Function CreateReportUrl(reportKubun As String) As String
            Return Me.ResolveUrl("~/Pages/Report/ReportView.aspx?" & ReportKubunKey & "=" &
                                 HttpUtility.UrlEncode(reportKubun))
        End Function
#End Region

#Region "内部処理"
        ''' <summary>
        ''' PDF をそのまま返します。
        ''' </summary>
        ''' <param name="filePath">PDF のフルパス。</param>
        ''' <param name="form">様式。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <remarks>
        ''' `inline` を指定し、ブラウザ内で開かせます。保存は利用者の操作に任せます。
        ''' </remarks>
        Private Sub TransmitPdf(filePath As String, form As HaisoTehaiForm, loginCode As String)
            Try
                Me.Response.Clear()
                Me.Response.ContentType = PdfContentType
                Me.Response.AppendHeader("Content-Disposition",
                                         "inline; filename=" & form.ReportKubun & ".pdf")
                Me.Response.TransmitFile(filePath)
                Me.Response.End()
            Catch ex As Threading.ThreadAbortException
                ' Response.End が投げるもの。処理は正常に終わっているため何もしません。
                Throw
            Catch ex As IOException
                Me.ResetResponse()
                Me.litMessage.Text = NotReadyMessage
                Logger.WriteWarning(LogSource, "PDF を読み出せませんでした。" & filePath & " " & ex.Message, loginCode)
            Catch ex As UnauthorizedAccessException
                Me.ResetResponse()
                Me.litMessage.Text = SettingMissingMessage
                Logger.WriteError(LogSource, "PDF の読み出しが拒否されました。" & filePath, ex, loginCode)
            End Try
        End Sub

        ''' <summary>
        ''' 応答を画面として返せる状態に戻します。
        ''' </summary>
        ''' <remarks>
        ''' PDF として返す準備をした後に失敗すると、**文言が PDF として扱われてしまいます。**
        ''' 付けたヘッダーと内容を捨ててから文言を出します。
        ''' </remarks>
        Private Sub ResetResponse()
            Me.Response.Clear()
            Me.Response.ClearHeaders()
            Me.Response.ContentType = "text/html"
        End Sub

        ''' <summary>
        ''' 設定値を取得します。
        ''' </summary>
        ''' <param name="key">設定キー。</param>
        ''' <returns>設定値。未設定のときは空文字。</returns>
        Private Shared Function GetSetting(key As String) As String
            Dim value As String = ConfigurationManager.AppSettings(key)
            Return If(value, String.Empty)
        End Function
#End Region

    End Class
End Namespace
