Imports System.Collections.Generic
Imports System.Drawing

Namespace Common
    ''' <summary>ヘッダーの画面別設定（タイトル・背景色・ボタン構成）を保持する。</summary>
    Public Class HeaderConfig

        ''' <summary>ヘッダーに表示するタイトル。</summary>
        Public Property Title As String
        ''' <summary>ヘッダーの背景色。</summary>
        Public Property HeaderBackColor As Color
        ''' <summary>ヘッダー右側に並べるボタンの一覧（先頭から左→右の順で配置）。</summary>
        Public Property Buttons As List(Of HeaderButton)

        ''' <summary>既定値（タイトル空・標準背景色・ボタンなし）で初期化する。</summary>
        Public Sub New()
            Me.Title = String.Empty
            Me.HeaderBackColor = SystemColors.Control
            Me.Buttons = New List(Of HeaderButton)()
        End Sub
    End Class

    ''' <summary>ヘッダーボタン1個分の設定（表示文字と押下時の処理）。</summary>
    Public Class HeaderButton

        ''' <summary>ボタンに表示する文字。</summary>
        Public Property Text As String
        ''' <summary>押下時に実行する処理。</summary>
        Public Property OnClick As Action

        ''' <summary>表示文字と押下時処理を指定して初期化する。</summary>
        ''' <param name="text">ボタンに表示する文字。</param>
        ''' <param name="onClick">押下時に実行する処理。</param>
        Public Sub New(text As String, onClick As Action)
            Me.Text = text
            Me.OnClick = onClick
        End Sub
    End Class
End Namespace
