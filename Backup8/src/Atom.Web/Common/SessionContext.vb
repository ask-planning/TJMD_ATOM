Option Strict On
Option Explicit On
Option Infer Off

Imports System.Web
Imports Atom.Core.Entities

Namespace Common

    ''' <summary>
    ''' ログイン利用者の情報を保持します。セッションに格納され、リクエストをまたいで引き継がれます。
    ''' 旧 Access 版のグローバル変数に相当します。
    ''' </summary>
    ''' <remarks>
    ''' ModeAtom / ShanghaiChotatsu / BumonCode / MailAddress は用途が未確定です。
    ''' 旧 Access 版の後続処理で参照されている可能性があるため値の保持のみ行い、
    ''' これらによる処理の分岐は実装しません。
    ''' </remarks>
    <Serializable()>
    Public NotInheritable Class SessionContext

#Region "定数"
        ''' <summary>セッションに格納するときのキー。</summary>
        Private Const SessionKey As String = "Atom.SessionContext"
#End Region

#Region "プロパティ"
        ''' <summary>社員コード。</summary>
        Public Property LoginCode As String = String.Empty

        ''' <summary>社員名。ヘッダーに表示します。</summary>
        Public Property LoginName As String = String.Empty

        ''' <summary>部門コード。</summary>
        Public Property BumonCode As String = String.Empty

        ''' <summary>ATOM の利用権限（login_mode_atom）。値の保持のみ行います。</summary>
        Public Property ModeAtom As Integer?

        ''' <summary>上海調達の区分。値の保持のみ行います。</summary>
        Public Property ShanghaiChotatsu As Integer?

        ''' <summary>メールアドレス。値の保持のみ行います。</summary>
        Public Property MailAddress As String = String.Empty

        ''' <summary>
        ''' 現在のセッションに紐づく情報を取得します。
        ''' 未生成のときは新しく作成してセッションに格納します。
        ''' </summary>
        Public Shared ReadOnly Property Current As SessionContext
            Get
                Dim context As HttpContext = HttpContext.Current
                ' セッションが使えないとき（アプリ起動直後など）は一時的な入れ物を返す
                If context Is Nothing OrElse context.Session Is Nothing Then
                    Return New SessionContext()
                End If

                Dim stored As SessionContext = [TryCast](context.Session(SessionKey), SessionContext)
                ' 未生成のとき
                If stored Is Nothing Then
                    stored = New SessionContext()
                    context.Session(SessionKey) = stored
                End If
                Return stored
            End Get
        End Property
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' セッションの入れ物を用意します。
        ''' 空のセッションは保存されずリクエストごとに ID が変わるため、
        ''' ViewState の鍵に使う ID を安定させる目的で呼び出します。
        ''' </summary>
        Public Shared Sub EnsureCreated()
            Dim created As SessionContext = SessionContext.Current
            ' 生成できたかどうかは利用側で判断しないため、戻り値は返しません。
            If created Is Nothing Then
                Return
            End If
        End Sub

        ''' <summary>
        ''' ログインした社員の情報をセッションへ設定します。
        ''' </summary>
        ''' <param name="user">認証済みの社員情報。</param>
        Public Shared Sub SignIn(user As LoginUser)
            ' 社員情報が無いとき
            If user Is Nothing Then
                Throw New ArgumentNullException(NameOf(user))
            End If

            Dim context As HttpContext = HttpContext.Current
            ' セッションが使えないとき
            If context Is Nothing OrElse context.Session Is Nothing Then
                Return
            End If

            Dim stored As New SessionContext()
            stored.LoginCode = user.LoginCode
            stored.LoginName = user.LoginName
            stored.BumonCode = user.BumonCode
            stored.ModeAtom = user.ModeAtom
            stored.ShanghaiChotatsu = user.ShanghaiChotatsu
            stored.MailAddress = user.MailAddress
            context.Session(SessionKey) = stored
        End Sub

        ''' <summary>
        ''' ログイン状態を解除し、セッションを破棄します。
        ''' </summary>
        Public Shared Sub SignOut()
            Dim context As HttpContext = HttpContext.Current
            ' セッションが使えないとき
            If context Is Nothing OrElse context.Session Is Nothing Then
                Return
            End If

            context.Session.Remove(SessionKey)
            context.Session.Abandon()
        End Sub

        ''' <summary>ログイン済みかどうかを返します。</summary>
        ''' <returns>社員コードが設定されているとき True。</returns>
        Public Function IsLoggedIn() As Boolean
            Return Not String.IsNullOrEmpty(Me.LoginCode)
        End Function
#End Region

    End Class
End Namespace
