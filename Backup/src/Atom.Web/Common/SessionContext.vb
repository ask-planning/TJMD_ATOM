Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Web

Namespace Common

    ''' <summary>
    ''' ログイン利用者の情報と、許可されたメニュー番号を保持します。
    ''' セッションに格納され、リクエストをまたいで引き継がれます。
    ''' </summary>
    <Serializable()>
    Public NotInheritable Class SessionContext

#Region "定数"
        ''' <summary>セッションに格納するときのキー。</summary>
        Private Const SessionKey As String = "Atom.SessionContext"
#End Region

#Region "フィールド"
        Private ReadOnly _allowedMenus As New HashSet(Of Integer)()
#End Region

#Region "コンストラクタ"
        ''' <summary>既定の状態で初期化します。</summary>
        Public Sub New()
        End Sub
#End Region

#Region "プロパティ"
        ''' <summary>ログインコード。</summary>
        Public Property LoginCode As String = String.Empty

        ''' <summary>利用者名。</summary>
        Public Property LoginName As String = String.Empty

        ''' <summary>所属部門。</summary>
        Public Property Department As String = String.Empty

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

                Dim stored As SessionContext = TryCast(context.Session(SessionKey), SessionContext)
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
        ''' 許可メニュー番号を設定します。ログイン時に権限取得結果を反映するために使います。
        ''' </summary>
        ''' <param name="menuNos">許可するメニュー番号の一覧。</param>
        Public Sub SetPermissions(menuNos As IEnumerable(Of Integer))
            Me._allowedMenus.Clear()
            ' 指定が無いときは何も許可しない
            If menuNos Is Nothing Then
                Return
            End If
            For Each menuNo As Integer In menuNos
                Me._allowedMenus.Add(menuNo)
            Next
        End Sub

        ''' <summary>
        ''' 指定したメニューが許可されているかを返します。
        ''' </summary>
        ''' <param name="menuNo">メニュー番号。</param>
        ''' <returns>許可されているとき True。</returns>
        Public Function IsAllowed(menuNo As Integer) As Boolean
            Return Me._allowedMenus.Contains(menuNo)
        End Function

        ''' <summary>ログイン済みかどうかを返します。</summary>
        ''' <returns>ログインコードが設定されているとき True。</returns>
        Public Function IsLoggedIn() As Boolean
            Return Not String.IsNullOrEmpty(Me.LoginCode)
        End Function
#End Region

    End Class
End Namespace
