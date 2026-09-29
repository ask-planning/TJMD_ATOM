Option Strict On
Option Explicit On
Option Infer Off

Imports System.Runtime.Serialization

Namespace Exceptions

    ''' <summary>
    ''' 他の利用者による更新と競合したときに送出します。
    ''' Web では画面を開いている間ロックを保持できないため、更新日時による楽観ロックで検知します。
    ''' </summary>
    <Serializable()>
    Public Class ConcurrencyException
        Inherits Exception

#Region "コンストラクタ"
        ''' <summary>既定の文言で初期化します。</summary>
        Public Sub New()
            MyBase.New("他の利用者によって更新されています。最新の内容を読み込み直してください。")
        End Sub

        ''' <summary>指定した文言で初期化します。</summary>
        ''' <param name="message">例外の説明。</param>
        Public Sub New(message As String)
            MyBase.New(message)
        End Sub

        ''' <summary>指定した文言と内側の例外で初期化します。</summary>
        ''' <param name="message">例外の説明。</param>
        ''' <param name="innerException">原因となった例外。</param>
        Public Sub New(message As String, innerException As Exception)
            MyBase.New(message, innerException)
        End Sub

        ''' <summary>直列化された情報から復元します。</summary>
        ''' <param name="info">直列化情報。</param>
        ''' <param name="context">直列化の文脈。</param>
        Protected Sub New(info As SerializationInfo, context As StreamingContext)
            MyBase.New(info, context)
        End Sub
#End Region

    End Class
End Namespace
