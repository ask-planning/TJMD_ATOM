Option Strict On
Option Explicit On
Option Infer Off

Imports System.Web
Imports System.Web.UI

Namespace Common

    ''' <summary>
    ''' 全画面の基底ページです。ログイン判定・権限判定・画面遷移・メッセージ表示を共通化します。
    ''' 各業務画面は認証処理を意識せず業務ロジックだけを実装できます。
    ''' </summary>
    Public Class BasePage
        Inherits Page

#Region "定数"
        ''' <summary>ログイン画面の URL。</summary>
        Public Const LoginUrl As String = "~/Pages/Common/Login.aspx"

        ''' <summary>メニュー画面の URL。</summary>
        Public Const MenuUrl As String = "~/Default.aspx"

        ''' <summary>ログイン後に戻る URL を渡すクエリ名。</summary>
        Public Const ReturnUrlKey As String = "ReturnUrl"

        ''' <summary>未実装の画面を開こうとしたときの文言。</summary>
        Private Const NotImplementedMessage As String = "現在未実装です。"

        ''' <summary>権限が無い画面を開こうとしたときの文言。</summary>
        Private Const NoPermissionMessage As String = "処理権限がありません。"
#End Region

#Region "プロパティ"
        ''' <summary>
        ''' 権限判定に使うメニュー番号です。0 のときは判定しません。
        ''' 各画面の OnInit で設定します。
        ''' </summary>
        Public Property MenuNo As Integer = 0

        ''' <summary>
        ''' ログインを必須とするかどうかです。ログイン画面だけが False を返します。
        ''' </summary>
        Protected Overridable ReadOnly Property RequiresLogin As Boolean
            Get
                Return True
            End Get
        End Property

        ''' <summary>
        ''' 画面遷移の指示済みかどうかです。
        ''' Response.Redirect の第 2 引数を False にするとページの処理は続くため、
        ''' 遷移を指示したあとの処理を各画面で止められるようにします。
        ''' </summary>
        Public ReadOnly Property IsRedirecting As Boolean
            Get
                Return Me.Response.IsRequestBeingRedirected
            End Get
        End Property
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' 画面上のメッセージ領域に文言を表示します。
        ''' </summary>
        ''' <param name="message">表示する文言。</param>
        ''' <param name="isError">エラーとして表示するとき True。</param>
        Public Sub ShowMessage(message As String, isError As Boolean)
            Dim master As SiteMaster = TryCast(Me.Master, SiteMaster)
            ' マスターページが取得できないときは何もしない
            If master Is Nothing Then
                Return
            End If
            master.ShowMessage(message, isError)
        End Sub

        ''' <summary>
        ''' メニュー番号を指定して画面を遷移します。
        ''' 権限が無いとき、または未実装のときは遷移せず、現在画面にメッセージを表示します。
        ''' </summary>
        ''' <param name="menuNo">遷移先のメニュー番号。</param>
        Public Sub NavigateTo(menuNo As Integer)
            Dim screen As ScreenDefinition = ScreenCatalog.Current.FindScreen(menuNo)
            ' 定義が無いとき
            If screen Is Nothing Then
                Me.ShowMessage(NotImplementedMessage, False)
                Return
            End If
            ' 権限が無いとき（ボタンの非表示だけに頼らずここでも判定する）
            If Not screen.IsAllowed Then
                Me.ShowMessage(NoPermissionMessage, True)
                Return
            End If
            ' 画面が未実装のとき
            If Not screen.IsImplemented Then
                Me.ShowMessage(NotImplementedMessage, False)
                Return
            End If

            Me.RedirectTo(screen.Url)
        End Sub

        ''' <summary>
        ''' 画面を遷移します。
        ''' </summary>
        ''' <param name="url">遷移先のアプリケーション相対 URL。</param>
        Public Sub RedirectTo(url As String)
            Me.Response.Redirect(url, False)
            Me.Context.ApplicationInstance.CompleteRequest()
        End Sub

        ''' <summary>
        ''' ログアウトし、ログイン画面へ遷移します。
        ''' </summary>
        Public Sub SignOutAndRedirect()
            SessionContext.SignOut()
            Me.RedirectTo(LoginUrl)
        End Sub
#End Region

#Region "イベントハンドラ"
        ''' <summary>
        ''' 事前初期化時の処理です。未ログインのときはログイン画面へ戻します。
        ''' </summary>
        ''' <param name="e">イベント引数。</param>
        Protected Overrides Sub OnPreInit(e As EventArgs)
            MyBase.OnPreInit(e)

            ' ログインが必要な画面に未ログインで来たとき
            If Me.RequiresLogin AndAlso Not SessionContext.Current.IsLoggedIn() Then
                Me.RedirectToLogin()
                Return
            End If
        End Sub

        ''' <summary>
        ''' 初期化時の処理です。CSRF 対策のキーを設定します。
        ''' </summary>
        ''' <param name="e">イベント引数。</param>
        Protected Overrides Sub OnInit(e As EventArgs)
            If Me.Session IsNot Nothing Then
                ' セッションが空のままだと保存されず、次のリクエストで ID が変わってしまう。
                ' 鍵に使う ID を安定させるため、先にセッションへ内容を持たせる。
                SessionContext.EnsureCreated()

                ' ViewState の改ざんと CSRF を防ぐためセッション ID を鍵にする
                Me.ViewStateUserKey = Me.Session.SessionID
            End If
            MyBase.OnInit(e)
        End Sub

        ''' <summary>
        ''' 読み込み時の処理です。権限が無い場合はメニューへ戻します。
        ''' </summary>
        ''' <param name="e">イベント引数。</param>
        Protected Overrides Sub OnLoad(e As EventArgs)
            MyBase.OnLoad(e)

            ' すでに別の画面へ遷移を指示しているとき
            If Me.IsRedirecting Then
                Return
            End If

            ' 権限が無いとき
            If Not PermissionChecker.IsAllowed(Me.MenuNo) Then
                Me.RedirectTo(MenuUrl & "?denied=" & Me.MenuNo.ToString())
                Return
            End If
        End Sub
#End Region

#Region "内部処理"
        ''' <summary>
        ''' ログイン画面へ遷移します。ログイン後に元の画面へ戻れるよう、現在の URL を渡します。
        ''' </summary>
        Private Sub RedirectToLogin()
            Dim currentUrl As String = Me.Request.Url.PathAndQuery
            Dim url As String = LoginUrl & "?" & ReturnUrlKey & "=" & HttpUtility.UrlEncode(currentUrl)
            Me.RedirectTo(url)
        End Sub
#End Region

    End Class
End Namespace
