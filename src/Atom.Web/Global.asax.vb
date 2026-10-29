Option Strict On
Option Explicit On
Option Infer Off

Imports System.Configuration
Imports System.Web
Imports Atom.Core.Logging

''' <summary>
''' アプリケーション全体の起動処理と共通のエラー処理を担います。
''' </summary>
''' <remarks>
''' ログの書き出し先は Web.config の `Log.Writer` で決まります。
''' 現時点で登録できるのは `Trace` だけです。
''' `Sql` / `File` の書き出し先が決まったら、Atom.Core.Logging.ILogWriter を実装したクラスを作り、
''' <see cref="RegisterLogWriter"/> に分岐を足してください。
''' </remarks>
Public Class Global_asax
    Inherits HttpApplication

#Region "定数"
    ''' <summary>ログの書き出し先を指定する設定キー。</summary>
    Private Const LogWriterKey As String = "Log.Writer"

    ''' <summary>Trace へ書き出すときの設定値。</summary>
    Private Const TraceWriterName As String = "Trace"
#End Region

#Region "イベントハンドラ"
    ''' <summary>
    ''' アプリケーション開始時の処理です。
    ''' </summary>
    ''' <param name="sender">イベントの発生元。</param>
    ''' <param name="e">イベント引数。</param>
    Private Sub Application_Start(sender As Object, e As EventArgs)
        RegisterLogWriter()
        Logger.WriteInformation("Application", "ATOM を起動しました。", String.Empty)
    End Sub

    ''' <summary>
    ''' ログの書き出し先を登録します。
    ''' </summary>
    ''' <remarks>
    ''' **登録しないと例外の詳細が記録されません。**
    ''' 書き出し先が未登録のときの Logger の代替出力は文言だけで、
    ''' 例外の内容が落ちてしまい、原因を追えませんでした。
    ''' </remarks>
    Private Shared Sub RegisterLogWriter()
        Dim writerName As String = ConfigurationManager.AppSettings(LogWriterKey)

        ' Trace のとき
        If String.Equals(writerName, TraceWriterName, StringComparison.OrdinalIgnoreCase) Then
            Logger.Register(New TraceLogWriter())
            Return
        End If

        ' TODO: Sql / File の書き出し先が決まったらここに足す
    End Sub

    ''' <summary>
    ''' 未処理の例外が発生したときの処理です。ログへ記録します。
    ''' </summary>
    ''' <param name="sender">イベントの発生元。</param>
    ''' <param name="e">イベント引数。</param>
    Private Sub Application_Error(sender As Object, e As EventArgs)
        Dim ex As Exception = Me.Server.GetLastError()
        ' 例外を取得できたとき
        If ex IsNot Nothing Then
            Logger.WriteError("Application", "未処理の例外が発生しました。", ex, String.Empty)
        End If
    End Sub
#End Region

End Class
