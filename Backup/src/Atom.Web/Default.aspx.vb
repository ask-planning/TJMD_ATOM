Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports Atom.Web.Common

''' <summary>
''' メインメニュー画面です。現行の F_0_START_MENU に相当します。
''' ブロックごとにメニューを表示し、各項目は画面へのリンクとして描画します。
''' </summary>
''' <remarks>
''' メニューはポストバックを伴わないリンクで描画します。
''' 入れ子の Repeater ではボタンの ItemCommand が外側まで伝わらないため、
''' ボタン方式は採用していません。
''' </remarks>
Public Class _Default
    Inherits BasePage

#Region "公開メソッド"
    ''' <summary>
    ''' メニュー画面のヘッダー設定を返します。閉じるボタンは表示しません。
    ''' </summary>
    ''' <returns>ヘッダー設定。</returns>
    Public Overrides Function CreateHeaderConfig() As HeaderConfig
        Dim config As New HeaderConfig()
        config.Title = "ATOM メインメニュー"
        config.BackColor = HeaderConfig.DefaultBackColor
        Return config
    End Function
#End Region

#Region "イベントハンドラ"
    ''' <summary>
    ''' 読み込み時の処理です。仮ログインを設定し、メニューを表示します。
    ''' </summary>
    ''' <param name="sender">イベントの発生元。</param>
    ''' <param name="e">イベント引数。</param>
    Private Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.SetTemporaryLogin()
        Me.BindMenu()
        Me.ShowDeniedMessage()
    End Sub

    ''' <summary>
    ''' ブロックのデータ割り当て時に、そのブロック内のメニュー項目を割り当てます。
    ''' </summary>
    ''' <param name="sender">イベントの発生元。</param>
    ''' <param name="e">イベント引数。</param>
    Private Sub RptBlocks_ItemDataBound(sender As Object, e As RepeaterItemEventArgs) Handles rptBlocks.ItemDataBound
        ' 見出し行やフッター行のときは何もしない
        If e.Item.ItemType <> ListItemType.Item AndAlso e.Item.ItemType <> ListItemType.AlternatingItem Then
            Return
        End If

        Dim block As MenuBlock = TryCast(e.Item.DataItem, MenuBlock)
        If block Is Nothing Then
            Return
        End If

        Dim entries As Repeater = TryCast(e.Item.FindControl("rptEntries"), Repeater)
        If entries Is Nothing Then
            Return
        End If

        entries.DataSource = block.Entries
        entries.DataBind()
    End Sub
#End Region

#Region "内部処理"
    ''' <summary>
    ''' メニューをブロック単位で画面に割り当てます。
    ''' </summary>
    Private Sub BindMenu()
        Dim blocks As IList(Of MenuBlock) = MenuDefinition.CreateBlocks()
        Me.rptBlocks.DataSource = blocks
        Me.rptBlocks.DataBind()
    End Sub

    ''' <summary>
    ''' 権限不足で戻された場合に、その旨を表示します。
    ''' </summary>
    Private Sub ShowDeniedMessage()
        Dim denied As String = Me.Request.QueryString("denied")
        ' 権限不足の戻りでないとき
        If String.IsNullOrEmpty(denied) Then
            Return
        End If

        Dim menuNo As Integer = 0
        If Integer.TryParse(denied, menuNo) Then
            Me.ShowMessage("処理権限がありません。（メニュー番号 " & menuNo.ToString() & "）", True)
        End If
    End Sub

    ''' <summary>
    ''' 仮のログイン情報を設定します。
    ''' TODO: ログイン画面と権限取得処理を実装したら削除します。
    ''' </summary>
    Private Sub SetTemporaryLogin()
        Dim session As SessionContext = SessionContext.Current
        ' すでにログイン済みのとき
        If session.IsLoggedIn() Then
            Return
        End If

        session.LoginCode = "000001"
        session.LoginName = "山田 太郎"
        session.Department = "輸入部"
        session.SetPermissions(New Integer() {101, 102, 103, 104, 105, 106, 107,
                                              201, 202, 203,
                                              301,
                                              902, 903, 904, 905, 906, 907})
    End Sub
#End Region

End Class
