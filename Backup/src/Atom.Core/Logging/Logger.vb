Option Strict On
Option Explicit On
Option Infer Off

Namespace Logging

    ''' <summary>
    ''' アプリケーション全体のログ出力窓口です。
    ''' 起動時に Register で書き出し先を登録してから使います。
    ''' </summary>
    Public NotInheritable Class Logger

#Region "フィールド"
        ''' <summary>排他制御用のロックオブジェクト。</summary>
        Private Shared ReadOnly _lockObject As New Object()

        ''' <summary>登録された書き出し先。</summary>
        Private Shared _writer As ILogWriter
#End Region

#Region "コンストラクタ"
        ''' <summary>インスタンス化を禁止します。</summary>
        Private Sub New()
        End Sub
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' ログの書き出し先を登録します。アプリケーション起動時に 1 回だけ呼び出します。
        ''' </summary>
        ''' <param name="writer">書き出し先。</param>
        Public Shared Sub Register(writer As ILogWriter)
            SyncLock _lockObject
                _writer = writer
            End SyncLock
        End Sub

        ''' <summary>
        ''' 動作情報を記録します。
        ''' </summary>
        ''' <param name="source">発生元。画面名やクラス名を渡します。</param>
        ''' <param name="message">記録する文言。</param>
        ''' <param name="loginCode">操作した利用者のログインコード。</param>
        Public Shared Sub WriteInformation(source As String, message As String, loginCode As String)
            Write(LogLevel.Information, source, message, String.Empty, loginCode)
        End Sub

        ''' <summary>
        ''' 警告を記録します。
        ''' </summary>
        ''' <param name="source">発生元。</param>
        ''' <param name="message">記録する文言。</param>
        ''' <param name="loginCode">操作した利用者のログインコード。</param>
        Public Shared Sub WriteWarning(source As String, message As String, loginCode As String)
            Write(LogLevel.Warning, source, message, String.Empty, loginCode)
        End Sub

        ''' <summary>
        ''' 例外を記録します。
        ''' </summary>
        ''' <param name="source">発生元。</param>
        ''' <param name="message">記録する文言。</param>
        ''' <param name="ex">発生した例外。</param>
        ''' <param name="loginCode">操作した利用者のログインコード。</param>
        Public Shared Sub WriteError(source As String, message As String, ex As Exception, loginCode As String)
            Dim detail As String = String.Empty
            ' 例外が渡されたとき
            If ex IsNot Nothing Then
                detail = ex.ToString()
            End If
            Write(LogLevel.[Error], source, message, detail, loginCode)
        End Sub
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 登録された書き出し先へログを渡します。
        ''' ログ出力の失敗が業務処理を止めないよう、ここで例外を吸収します。
        ''' </summary>
        ''' <param name="level">重要度。</param>
        ''' <param name="source">発生元。</param>
        ''' <param name="message">記録する文言。</param>
        ''' <param name="detail">詳細。</param>
        ''' <param name="loginCode">ログインコード。</param>
        Private Shared Sub Write(level As LogLevel, source As String, message As String, detail As String, loginCode As String)
            Dim writer As ILogWriter = Nothing
            SyncLock _lockObject
                writer = _writer
            End SyncLock

            ' 書き出し先が未登録のとき
            If writer Is Nothing Then
                System.Diagnostics.Trace.WriteLine(source & " : " & message)
                Return
            End If

            Try
                writer.Write(level, source, message, detail, loginCode)
            Catch ex As Exception
                ' ログ出力の失敗は業務処理を止めない
                System.Diagnostics.Trace.WriteLine("ログ出力に失敗しました。" & ex.Message)
            End Try
        End Sub
#End Region

    End Class
End Namespace
