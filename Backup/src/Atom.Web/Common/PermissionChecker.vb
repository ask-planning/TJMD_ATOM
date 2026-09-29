Option Strict On
Option Explicit On
Option Infer Off

Namespace Common

    ''' <summary>
    ''' メニュー番号に対する権限判定を行います。現行の管理者フラグ判定に相当します。
    ''' </summary>
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
            ' 判定不要のとき
            If menuNo <= 0 Then
                Return True
            End If
            Return SessionContext.Current.IsAllowed(menuNo)
        End Function
#End Region

    End Class
End Namespace
