Option Strict On
Option Explicit On
Option Infer Off

Namespace Exceptions

    ''' <summary>
    ''' 取込対象の Excel の形式が想定と異なるときに送出します。
    ''' </summary>
    <Serializable()>
    Public Class ImportFormatException
        Inherits Exception

#Region "コンストラクタ"
        ''' <summary>文言を指定して初期化します。</summary>
        ''' <param name="message">利用者へ表示する文言。</param>
        Public Sub New(message As String)
            MyBase.New(message)
        End Sub

        ''' <summary>文言と原因の例外を指定して初期化します。</summary>
        ''' <param name="message">利用者へ表示する文言。</param>
        ''' <param name="innerException">原因となった例外。</param>
        Public Sub New(message As String, innerException As Exception)
            MyBase.New(message, innerException)
        End Sub
#End Region

    End Class
End Namespace
