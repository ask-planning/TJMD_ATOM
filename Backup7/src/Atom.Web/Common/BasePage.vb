Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
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
        ''' 共通ヘッダーの右側に表示する画面固有ボタンです。
        ''' 各画面の OnInit で追加します。空のときはボタン領域を表示しません。
        ''' </summary>
        Public ReadOnly Property HeaderActions As IList([Of] HeaderAction) = New List([Of] HeaderAction)()

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
            Dim master As SiteMaster = [TryCast](Me.Master, SiteMaster)
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
        ''' 未保存の入力があるという状態を解除します。登録が終わったときに呼び出します。
        ''' </summary>
        ''' <remarks>
        ''' ログアウト時の確認は、入力欄を触ったかどうかでブラウザ側が判定しています。
        ''' 登録が済んだあとも確認が出続けないよう、ここで解除します。
        ''' </remarks>
        Public Sub ClearUnsavedState()
            Me.ClientScript.RegisterStartupScript(Me.GetType(), "AtomClearUnsaved",
                                                  "AtomPage.clearUnsaved();", True)
        End Sub

        ''' <summary>
        ''' メニュー番号から画面の URL を求めます。別のタブで開くリンクを作るときに使います。
        ''' </summary>
        ''' <param name="menuNo">対象のメニュー番号。</param>
        ''' <returns>アプリケーション相対 URL。定義が無い・権限が無い・未実装のときは空文字。</returns>
        ''' <remarks>
        ''' 判定の条件は NavigateTo と同じです。空文字が返ったときは、
        ''' 呼び出し側でリンクを非表示にしてください。
        ''' </remarks>
        Public Function TryResolveScreenUrl(menuNo As Integer) As String
            Dim screen As ScreenDefinition = ScreenCatalog.Current.FindScreen(menuNo)
            ' 定義が無いとき
            If screen Is Nothing Then
                Return String.Empty
            End If
            ' 権限が無いとき
            If Not screen.IsAllowed Then
                Return String.Empty
            End If
            ' 画面が未実装のとき
            If Not screen.IsImplemented Then
                Return String.Empty
            End If
            Return screen.Url
        End Function

        ''' <summary>
        ''' 画面を遷移します。
        ''' </summary>
        ''' <param name="url">遷移先のアプリケーション相対 URL。</param>
        Public Sub RedirectTo(url As String)
            Me.Response.Redirect(url, False)
            Me.Context.ApplicationInstance.CompleteRequest()
        End Sub

        ''' <summary>
        ''' メニュー画面へ遷移します。
        ''' </summary>
        Public Sub NavigateToMenu()
            Me.RedirectTo(MenuUrl)
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
            If Me.Session [IsNot] Nothing Then
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
        ''' <remarks>
        ''' 判定は MyBase.OnLoad より前に行います。MyBase.OnLoad は各画面の Page_Load を
        ''' 呼び出すため、後で判定すると権限が無い利用者でも画面の初期化（検索や一覧の割り当て）が
        ''' 一度走ってしまいます。
        ''' </remarks>
        Protected Overrides Sub OnLoad(e As EventArgs)
            ' すでに別の画面へ遷移を指示しているとき
            If Me.IsRedirecting Then
                Return
            End If

            ' 権限が無いとき
            If Not PermissionChecker.IsAllowed(Me.MenuNo) Then
                Me.RedirectTo(MenuUrl & "?denied=" & Me.MenuNo.ToString())
                Return
            End If

            MyBase.OnLoad(e)
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
