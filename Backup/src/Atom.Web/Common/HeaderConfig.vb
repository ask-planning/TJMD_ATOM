Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic

Namespace Common

    ''' <summary>
    ''' ヘッダーに配置するボタン 1 個分の設定です。
    ''' </summary>
    Public NotInheritable Class HeaderButton

#Region "コンストラクタ"
        ''' <summary>ヘッダーボタンを初期化します。</summary>
        ''' <param name="text">ボタンに表示する文言。</param>
        ''' <param name="navigateUrl">押下時の遷移先 URL。</param>
        Public Sub New(text As String, navigateUrl As String)
            Me.Text = text
            Me.NavigateUrl = navigateUrl
        End Sub
#End Region

#Region "プロパティ"
        ''' <summary>ボタンに表示する文言。</summary>
        Public ReadOnly Property Text As String

        ''' <summary>押下時の遷移先 URL。</summary>
        Public ReadOnly Property NavigateUrl As String
#End Region

    End Class

    ''' <summary>
    ''' 画面ごとのヘッダー表示設定です。各画面が BasePage.CreateHeaderConfig で返します。
    ''' </summary>
    Public NotInheritable Class HeaderConfig

#Region "定数"
        ''' <summary>ヘッダーの既定の背景色。</summary>
        Public Const DefaultBackColor As String = "#1F4E79"
#End Region

#Region "コンストラクタ"
        ''' <summary>既定の設定で初期化します。</summary>
        Public Sub New()
            Me.Buttons = New List(Of HeaderButton)()
        End Sub
#End Region

#Region "プロパティ"
        ''' <summary>ヘッダーに表示する画面名。</summary>
        Public Property Title As String = String.Empty

        ''' <summary>ヘッダーの背景色。CSS の色指定を文字列で持ちます。</summary>
        Public Property BackColor As String = DefaultBackColor

        ''' <summary>ヘッダー右側に並べるボタン。</summary>
        Public ReadOnly Property Buttons As IList(Of HeaderButton)
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' メニューへ戻る「閉じる」ボタンを追加します。
        ''' </summary>
        ''' <returns>続けて設定できるようにするため自身を返します。</returns>
        Public Function AddCloseButton() As HeaderConfig
            Return Me.AddCloseButton("~/Default.aspx")
        End Function

        ''' <summary>
        ''' 遷移先を指定して「閉じる」ボタンを追加します。
        ''' 呼び出し元の画面へ戻す必要がある画面で使います。
        ''' </summary>
        ''' <param name="navigateUrl">閉じたあとの遷移先 URL。</param>
        ''' <returns>続けて設定できるようにするため自身を返します。</returns>
        Public Function AddCloseButton(navigateUrl As String) As HeaderConfig
            Dim url As String = navigateUrl
            ' 遷移先が指定されていないとき
            If String.IsNullOrEmpty(url) Then
                url = "~/Default.aspx"
            End If
            Me.Buttons.Add(New HeaderButton("閉じる", url))
            Return Me
        End Function
#End Region

    End Class
End Namespace
