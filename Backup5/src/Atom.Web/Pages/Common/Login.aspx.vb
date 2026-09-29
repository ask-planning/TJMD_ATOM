Option Strict On
Option Explicit On
Option Infer Off

Imports System.Data.SqlClient
Imports Atom.Core.Entities
Imports Atom.Data.Repositories
Imports Atom.Web.Common

Namespace Pages.Common

    ''' <summary>
    ''' ログイン画面です。社員コードだけで認証します。
    ''' 旧 Access 版の ATOM_Login（InputBox による社員コード入力）に相当します。
    ''' </summary>
    Public Class Login
        Inherits BasePage

#Region "定数"
        ''' <summary>社員コードの桁数。</summary>
        Private Const LoginCodeLength As Integer = 6

        ''' <summary>ログイン履歴に記録するメニュー番号。ログイン時は 0 とします。</summary>
        Private Const LoginLogMenuNo As Integer = 0

        ''' <summary>桁数が不正なときの文言。</summary>
        Private Const InvalidLengthMessage As String = "入力桁数が不正です（６桁）。"

        ''' <summary>社員コードが存在しないときの文言。</summary>
        Private Const NotFoundMessage As String = "存在しない社員コードが入力されました。"

        ''' <summary>データベースへ接続できないときの文言。</summary>
        Private Const DatabaseErrorMessage As String = "ログイン処理に失敗しました。時間をおいて再度お試しください。"
#End Region

#Region "プロパティ"
        ''' <summary>
        ''' ログイン画面自体はログインを必要としません。
        ''' </summary>
        Protected Overrides ReadOnly Property RequiresLogin As Boolean
            Get
                Return False
            End Get
        End Property
#End Region

#Region "イベントハンドラ"
        ''' <summary>
        ''' 読み込み時の処理です。ログイン済みのときはメニューへ戻します。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
            ' すでに遷移を指示しているとき
            If Me.IsRedirecting Then
                Return
            End If

            ' 初回表示のとき
            If Not Me.IsPostBack Then
                ' すでにログインしているとき
                If SessionContext.Current.IsLoggedIn() Then
                    Me.RedirectTo(MenuUrl)
                    Return
                End If
                Me.txtLoginCode.Focus()
            End If
        End Sub

        ''' <summary>
        ''' ログインボタン押下時の処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
            ' クライアント側の検証だけに頼らず、サーバー側でも必ず検証する
            Me.Validate()
            If Not Me.IsValid Then
                Return
            End If

            Dim loginCode As String = Me.txtLoginCode.Text.Trim()
            ' 桁数が 6 桁でないとき
            If loginCode.Length <> LoginCodeLength Then
                Me.ShowMessage(InvalidLengthMessage, True)
                Return
            End If

            Dim repository As New LoginRepository()
            Dim user As LoginUser = Nothing

            Try
                user = repository.SelectLoginUser(loginCode)
            Catch ex As SqlException
                ' 例外の詳細は Repository 側でログへ記録済みです。画面には定型文だけを出します。
                Me.ShowMessage(DatabaseErrorMessage, True)
                Return
            End Try

            ' 社員コードが存在しないとき
            If user Is Nothing Then
                Me.ShowMessage(NotFoundMessage, True)
                Return
            End If

            SessionContext.SignIn(user)
            repository.InsertLoginLog(user.LoginCode, AppSettings.SystemVersion, LoginLogMenuNo)

            Me.RedirectTo(Me.ResolveReturnUrl())
        End Sub
#End Region

#Region "内部処理"
        ''' <summary>
        ''' ログイン後の遷移先を求めます。
        ''' </summary>
        ''' <returns>遷移先の URL。指定が無いときや外部サイトを指すときはメニュー画面。</returns>
        ''' <remarks>
        ''' 遷移先はクエリ文字列から渡されるため改ざんされ得ます。
        ''' 外部サイトへ誘導されないよう、同一サイト内の相対パスだけを許可します。
        ''' </remarks>
        Private Function ResolveReturnUrl() As String
            Dim returnUrl As String = Me.Request.QueryString(ReturnUrlKey)
            ' 指定が無いとき
            If String.IsNullOrEmpty(returnUrl) Then
                Return MenuUrl
            End If
            ' 絶対 URL や別サイトを指しているとき
            If Not returnUrl.StartsWith("/", StringComparison.Ordinal) Then
                Return MenuUrl
            End If
            ' 「//host」「/\host」の形はスキーム相対として外部サイトへ飛び得るため除外する
            If returnUrl.StartsWith("//", StringComparison.Ordinal) Then
                Return MenuUrl
            End If
            If returnUrl.StartsWith("/\", StringComparison.Ordinal) Then
                Return MenuUrl
            End If
            If returnUrl.Contains(":") OrElse returnUrl.Contains("\") Then
                Return MenuUrl
            End If
            Return returnUrl
        End Function
#End Region

    End Class
End Namespace
