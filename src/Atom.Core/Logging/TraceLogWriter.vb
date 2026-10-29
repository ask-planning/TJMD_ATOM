Option Strict On
Option Explicit On
Option Infer Off

Imports System.Globalization

Namespace Logging

    ''' <summary>
    ''' ログを System.Diagnostics.Trace へ書き出します。
    ''' </summary>
    ''' <remarks>
    ''' Web.config の `Log.Writer` が `Trace` のときに使います。
    ''' Visual Studio の「出力」ウィンドウ（デバッグ）に出るため、開発中の調査に使えます。
    '''
    ''' **例外の詳細も必ず書き出します。**
    ''' 書き出し先が未登録のときの Logger 側の代替出力は文言だけで、
    ''' 例外の内容が落ちてしまい原因が追えませんでした。
    '''
    ''' データベースに依存しないため Atom.Core に置いています。
    ''' </remarks>
    Public Class TraceLogWriter
        Implements ILogWriter

#Region "公開メソッド"
        ''' <summary>
        ''' ログを 1 件書き出します。
        ''' </summary>
        ''' <param name="level">重要度。</param>
        ''' <param name="source">発生元。</param>
        ''' <param name="message">記録する文言。</param>
        ''' <param name="detail">例外情報などの詳細。無いときは空文字。</param>
        ''' <param name="loginCode">操作した利用者のログインコード。不明なときは空文字。</param>
        Public Sub Write(level As LogLevel, source As String, message As String,
                         detail As String, loginCode As String) Implements ILogWriter.Write

            Dim timestamp As String = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture)
            System.Diagnostics.Trace.WriteLine(
                timestamp & " [" & level.ToString() & "] " & source & " : " & message &
                If(String.IsNullOrEmpty(loginCode), String.Empty, "（" & loginCode & "）"))

            ' 詳細があるとき
            If Not String.IsNullOrEmpty(detail) Then
                System.Diagnostics.Trace.WriteLine(detail)
            End If
        End Sub
#End Region

    End Class
End Namespace
