Option Strict On
Option Explicit On
Option Infer Off

Namespace Logging

    ''' <summary>
    ''' ログの重要度を表します。
    ''' </summary>
    Public Enum LogLevel

        ''' <summary>動作確認用の詳細情報。</summary>
        Debug = 0

        ''' <summary>通常の動作情報。</summary>
        Information = 1

        ''' <summary>処理は継続できるが注意が必要な状態。</summary>
        Warning = 2

        ''' <summary>処理を継続できない異常。</summary>
        [Error] = 3

    End Enum
End Namespace
