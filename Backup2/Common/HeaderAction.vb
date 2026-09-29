Option Strict On
Option Explicit On
Option Infer Off

Namespace Common

    ''' <summary>
    ''' 共通ヘッダーの右側に置く画面固有ボタン 1 個分の設定です。
    ''' </summary>
    ''' <remarks>
    ''' 各画面が BasePage.HeaderActions へ追加します。押下時は共通の画面遷移処理を通るため、
    ''' 権限が無いときや未実装のときはメッセージを表示して現在画面に留まります。
    ''' </remarks>
    Public NotInheritable Class HeaderAction

#Region "定数"
        ''' <summary>メニュー画面へ戻ることを表すメニュー番号。</summary>
        Public Const MenuScreenNo As Integer = 0
#End Region

#Region "コンストラクタ"
        ''' <summary>
        ''' ボタンを初期化します。
        ''' </summary>
        ''' <param name="text">ボタンに表示する文言。</param>
        ''' <param name="menuNo">遷移先のメニュー番号。0 のときはメニュー画面へ戻ります。</param>
        Public Sub New(text As String, menuNo As Integer)
            Me.Text = text
            Me.MenuNo = menuNo
        End Sub
#End Region

#Region "プロパティ"
        ''' <summary>ボタンに表示する文言。</summary>
        Public ReadOnly Property Text As String

        ''' <summary>遷移先のメニュー番号。0 のときはメニュー画面。</summary>
        Public ReadOnly Property MenuNo As Integer
#End Region

    End Class
End Namespace
