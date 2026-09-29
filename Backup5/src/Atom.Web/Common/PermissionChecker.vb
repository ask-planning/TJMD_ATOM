Option Strict On
Option Explicit On
Option Infer Off

Namespace Common

    ''' <summary>
    ''' メニュー番号に対する権限判定を行います。
    ''' </summary>
    ''' <remarks>
    ''' 権限値（login_mode_atom）は 0 / 1 / 2 / 9 を取りますが、値ごとの意味と
    ''' 画面ごとの可否が未確定です。そのため現時点ではログイン済みかどうかだけで判定し、
    ''' 権限値による分岐は実装していません。仕様が確定したらこのクラスだけを変更します。
    ''' </remarks>
    Public NotInheritable Class PermissionChecker

#Region "コンストラクタ"
        ''' <summary>インスタンス化を禁止します。</summary>
        Private Sub New()
        End Sub
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' 指定したメニューが許可されているかを返します。
        ''' </summary>
        ''' <param name="menuNo">メニュー番号。0 以下のときは判定せず許可扱いとします。</param>
        ''' <returns>許可されているとき True。</returns>
        Public Shared Function IsAllowed(menuNo As Integer) As Boolean
            ' 判定不要のとき（メニュー画面・ログイン画面）
            If menuNo <= 0 Then
                Return True
            End If
            Return SessionContext.Current.IsLoggedIn()
        End Function
#End Region

    End Class
End Namespace
