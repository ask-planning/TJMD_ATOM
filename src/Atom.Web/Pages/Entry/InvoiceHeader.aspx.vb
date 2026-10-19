Option Strict On
Option Explicit On
Option Infer Off

Imports System.Web.UI
Imports Atom.Web.Common

Namespace Pages.Entry

    ''' <summary>
    ''' INVOICEヘッダー入力画面（menuNo=103）です。
    ''' 旧 Access 版の F_1_上海輸入_INVOICE_HEADER に相当します。
    ''' </summary>
    ''' <remarks>
    ''' 仕様は docs/specs/103_INVOICEヘッダー入力.md を正とします。
    '''
    ''' **現時点は画面のレイアウトのみです。**
    ''' 元INV_NO の選択肢の取得、明細の表示、登録、帳票出力はいずれも未実装で、
    ''' ボタンを押すと「未実装です」とだけ表示します。
    ''' 実装は設計書の確認事項に回答を得てから行います。
    ''' </remarks>
    Public Class InvoiceHeader
        Inherits BasePage

#Region "定数"
        ''' <summary>この画面のメニュー番号。</summary>
        Private Const ThisMenuNo As Integer = 103

        ''' <summary>遷移先。R3在庫転送作成。</summary>
        Private Const StockTransferMenuNo As Integer = 106

        ''' <summary>遷移先。R3購買発注確認。</summary>
        Private Const PurchaseConfirmMenuNo As Integer = 105

        ''' <summary>遷移先。R3購買発注取込。</summary>
        Private Const PurchaseImportMenuNo As Integer = 104

        ''' <summary>遷移先。INVOICE確認。</summary>
        Private Const InvoiceConfirmMenuNo As Integer = 102

        ''' <summary>未実装であることを伝える文言。</summary>
        Private Const NotImplementedMessage As String = "この処理はまだ実装していません。"
#End Region

#Region "イベント"
        ''' <summary>
        ''' 初期化時の処理です。メニュー番号とヘッダーのボタンを設定します。
        ''' </summary>
        ''' <param name="e">イベント引数。</param>
        ''' <remarks>
        ''' メニュー番号は権限判定より前に決める必要があるため、Page_Load ではなくここで設定します。
        ''' </remarks>
        Protected Overrides Sub OnInit(e As EventArgs)
            Me.MenuNo = ThisMenuNo

            Me.HeaderActions.Add(New HeaderAction("在庫転送", StockTransferMenuNo))
            Me.HeaderActions.Add(New HeaderAction("発注確認", PurchaseConfirmMenuNo))
            Me.HeaderActions.Add(New HeaderAction("発注取込", PurchaseImportMenuNo))
            Me.HeaderActions.Add(New HeaderAction("INV確認", InvoiceConfirmMenuNo))
            Me.HeaderActions.Add(New HeaderAction("MENUへ", HeaderAction.MenuScreenNo))

            MyBase.OnInit(e)
        End Sub

        ''' <summary>
        ''' 画面の読み込み時の処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
            ' 初回表示のとき
            If Not Me.IsPostBack Then
                Me.ShowMessage("画面のレイアウトのみの状態です。データの取得・登録・帳票出力は未実装です。", False)
            End If
        End Sub

        ''' <summary>
        ''' 元INV_NO を選んだときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub DdlInvoiceNoMain_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlInvoiceNoMain.SelectedIndexChanged
            Me.ShowMessage(NotImplementedMessage, False)
        End Sub

        ''' <summary>
        ''' 「登 録」を押したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnRegister_Click(sender As Object, e As EventArgs) Handles btnRegister.Click
            Me.ShowMessage(NotImplementedMessage, False)
        End Sub

        ''' <summary>
        ''' 「明細更新」を押したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnMeisaiUpdate_Click(sender As Object, e As EventArgs) Handles btnMeisaiUpdate.Click
            Me.ShowMessage(NotImplementedMessage, False)
        End Sub

        ''' <summary>
        ''' 「配送手配書」を押したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnHaisoTehai_Click(sender As Object, e As EventArgs) Handles btnHaisoTehai.Click
            Me.ShowMessage(NotImplementedMessage, False)
        End Sub

        ''' <summary>
        ''' 配送手配書の「表示」を押したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnHaisoTehaiView_Click(sender As Object, e As EventArgs) Handles btnHaisoTehaiView.Click
            Me.ShowMessage(NotImplementedMessage, False)
        End Sub

        ''' <summary>
        ''' 「入荷明細書」を押したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnNyukaMeisai_Click(sender As Object, e As EventArgs) Handles btnNyukaMeisai.Click
            Me.ShowMessage(NotImplementedMessage, False)
        End Sub

        ''' <summary>
        ''' 入荷明細書の「表示」を押したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnNyukaMeisaiView_Click(sender As Object, e As EventArgs) Handles btnNyukaMeisaiView.Click
            Me.ShowMessage(NotImplementedMessage, False)
        End Sub
#End Region

    End Class
End Namespace
