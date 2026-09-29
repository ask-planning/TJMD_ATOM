Option Strict On
Option Explicit On
Option Infer Off

Imports System.Web
Imports Atom.Core.Logging

''' <summary>
''' アプリケーション全体の起動処理と共通のエラー処理を担います。
''' </summary>
''' <remarks>
''' ログの書き出し先は未確定です。書き出し先を登録していないあいだ、
''' Logger は出力を System.Diagnostics.Trace へ流します（データベースへは書き込みません）。
''' 既存のログ出力先が判明したら、Atom.Core.Logging.ILogWriter を実装したクラスを作成し、
''' Application_Start で Logger.Register を呼び出してください。
''' </remarks>
Public Class Global_asax
    Inherits HttpApplication

#Region "イベントハンドラ"
    ''' <summary>
    ''' アプリケーション開始時の処理です。
    ''' </summary>
    ''' <param name="sender">イベントの発生元。</param>
    ''' <param name="e">イベント引数。</param>
    Private Sub Application_Start(sender As Object, e As EventArgs)
        ' TODO: 既存のログ出力先が判明したら ILogWriter の実装を登録する
        Logger.WriteInformation("Application", "ATOM を起動しました。", String.Empty)
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
