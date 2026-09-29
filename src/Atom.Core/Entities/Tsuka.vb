Option Strict On
Option Explicit On
Option Infer Off

Namespace Entities

    ''' <summary>
    ''' 通貨の区分です。データベースには文字列で保持します。
    ''' </summary>
    ''' <remarks>
    ''' 旧 Access 版は画面上で 1 = RMB / 2 = JPY という数値で持っていましたが、
    ''' 一時テーブルおよび明細テーブルは 'JPY' / 'RMB' の文字列で保持しており、
    ''' 各チェック処理もこの文字列で分岐するため、文字列で扱います。
    ''' </remarks>
    Public NotInheritable Class Tsuka

#Region "定数"
        ''' <summary>日本円。</summary>
        Public Const Jpy As String = "JPY"

        ''' <summary>人民元。</summary>
        Public Const Rmb As String = "RMB"
#End Region

#Region "コンストラクタ"
        ''' <summary>インスタンス化を禁止します。</summary>
        Private Sub New()
        End Sub
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' 通貨として扱える値かどうかを返します。
        ''' </summary>
        ''' <param name="value">判定する値。</param>
        ''' <returns>JPY または RMB のとき True。</returns>
        Public Shared Function IsValid(value As String) As Boolean
            Return String.Equals(value, Jpy, StringComparison.Ordinal) OrElse
                   String.Equals(value, Rmb, StringComparison.Ordinal)
        End Function
#End Region

    End Class
End Namespace
