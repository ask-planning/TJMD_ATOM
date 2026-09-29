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
    ''' 別のタブで開いた画面の「閉じる」のように、遷移せずブラウザ側だけで処理したいときは
    ''' CreateClientAction で作ります。
    ''' </remarks>
    Public NotInheritable Class HeaderAction

#Region "定数"
        ''' <summary>メニュー画面へ戻ることを表すメニュー番号。</summary>
        Public Const MenuScreenNo As Integer = 0

        ''' <summary>画面遷移を行わないことを表すメニュー番号。</summary>
        Private Const NoNavigationMenuNo As Integer = -1
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
            Me.ClientScript = String.Empty
        End Sub

        ''' <summary>
        ''' 画面遷移せず、ブラウザ側の処理だけを行うボタンを初期化します。
        ''' </summary>
        ''' <param name="text">ボタンに表示する文言。</param>
        ''' <param name="clientScript">押下時に実行するスクリプト。</param>
        Private Sub New(text As String, clientScript As String)
            Me.Text = text
            Me.MenuNo = NoNavigationMenuNo
            Me.ClientScript = clientScript
        End Sub
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' 画面遷移せず、ブラウザ側の処理だけを行うボタンを作ります。
        ''' </summary>
        ''' <param name="text">ボタンに表示する文言。</param>
        ''' <param name="clientScript">押下時に実行するスクリプト。ポストバックを止めるため false を返すこと。</param>
        ''' <returns>ボタンの設定。</returns>
        ''' <remarks>
        ''' スクリプトは画面側でリテラルとして与えます。利用者の入力を混ぜてはいけません。
        ''' </remarks>
        Public Shared Function CreateClientAction(text As String, clientScript As String) As HeaderAction
            Return New HeaderAction(text, clientScript)
        End Function
#End Region

#Region "プロパティ"
        ''' <summary>ボタンに表示する文言。</summary>
        Public ReadOnly Property Text As String

        ''' <summary>遷移先のメニュー番号。0 のときはメニュー画面。遷移しないときは -1。</summary>
        Public ReadOnly Property MenuNo As Integer

        ''' <summary>押下時にブラウザ側で実行するスクリプト。遷移するボタンでは空文字。</summary>
        Public ReadOnly Property ClientScript As String
#End Region

    End Class
End Namespace
