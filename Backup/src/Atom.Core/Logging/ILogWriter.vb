Option Strict On
Option Explicit On
Option Infer Off

Namespace Logging

    ''' <summary>
    ''' ログの書き出し先を表します。実装は Atom.Data 側に置きます。
    ''' Atom.Core がデータアクセスに依存しないよう、書き出し先を抽象化しています。
    ''' </summary>
    Public Interface ILogWriter

        ''' <summary>
        ''' ログを 1 件書き出します。
        ''' </summary>
        ''' <param name="level">重要度。</param>
        ''' <param name="source">発生元。画面名やクラス名を渡します。</param>
        ''' <param name="message">記録する文言。</param>
        ''' <param name="detail">例外情報などの詳細。無いときは空文字。</param>
        ''' <param name="loginCode">操作した利用者のログインコード。不明なときは空文字。</param>
        Sub Write(level As LogLevel, source As String, message As String, detail As String, loginCode As String)

    End Interface
End Namespace
