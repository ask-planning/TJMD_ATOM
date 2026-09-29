Option Strict On
Option Explicit On
Option Infer Off

Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports Atom.Web.Common

''' <summary>
''' メインメニュー画面です。旧 Access 版の F_0_START_MENU に相当します。
''' ブロックごとにメニューを表示し、押下されたメニュー番号で共通の画面遷移処理を呼び出します。
''' </summary>
Public Class DefaultPage
    Inherits BasePage

#Region "定数"
    ''' <summary>メニュー押下時のコマンド名。</summary>
    Private Const OpenCommandName As String = "Open"

    ''' <summary>内側のリピーターの ID。</summary>
    Private Const ScreenRepeaterId As String = "rptScreens"

    ''' <summary>権限不足で戻されたことを表すクエリ名。</summary>
    Private Const DeniedKey As String = "denied"
#End Region

#Region "イベントハンドラ"
    ''' <summary>
    ''' 読み込み時の処理です。メニューを表示します。
    ''' </summary>
    ''' <param name="sender">イベントの発生元。</param>
    ''' <param name="e">イベント引数。</param>
    Private Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' ログイン画面へ遷移を指示しているとき
        If Me.IsRedirecting Then
            Return
        End If

        ' 初回表示のとき
        If Not Me.IsPostBack Then
            Me.BindMenu()
            Me.ShowDeniedMessage()
        End If
    End Sub

    ''' <summary>
    ''' ブロックの生成時に、内側のメニュー押下イベントを受け取れるようにします。
    ''' </summary>
    ''' <param name="sender">イベントの発生元。</param>
    ''' <param name="e">イベント引数。</param>
    ''' <remarks>
    ''' 内側のリピーターはマークアップに静的に置いているため、受け取り先だけを結び付けます。
    ''' データ割り当て時にしか発生しない ItemDataBound ではポストバック時に結び付かないため、
    ''' 項目が再生成されるたびに発生する ItemCreated を使います。
    ''' </remarks>
    Private Sub RptBlocks_ItemCreated(sender As Object, e As RepeaterItemEventArgs) Handles rptBlocks.ItemCreated
        ' 見出し行やフッター行のときは何もしない
        If e.Item.ItemType <> ListItemType.Item AndAlso e.Item.ItemType <> ListItemType.AlternatingItem Then
            Return
        End If

        Dim screens As Repeater = TryCast(e.Item.FindControl(ScreenRepeaterId), Repeater)
        ' 内側のリピーターが取得できないとき
        If screens Is Nothing Then
            Return
        End If

        AddHandler screens.ItemCommand, AddressOf Me.RptScreens_ItemCommand
    End Sub

    ''' <summary>
    ''' メニュー押下時の処理です。共通の画面遷移処理へ渡します。
    ''' </summary>
    ''' <param name="source">イベントの発生元。</param>
    ''' <param name="e">イベント引数。</param>
    Private Sub RptScreens_ItemCommand(source As Object, e As RepeaterCommandEventArgs)
        ' 想定外のコマンドのとき
        If Not String.Equals(e.CommandName, OpenCommandName, StringComparison.Ordinal) Then
            Return
        End If

        Dim menuNo As Integer = 0
        ' メニュー番号を読み取れないとき
        If Not Integer.TryParse(CStr(e.CommandArgument), menuNo) Then
            Return
        End If

        Me.NavigateTo(menuNo)
    End Sub

    ''' <summary>
    ''' ログアウトボタン押下時の処理です。
    ''' </summary>
    ''' <param name="sender">イベントの発生元。</param>
    ''' <param name="e">イベント引数。</param>
    Private Sub BtnLogout_Click(sender As Object, e As EventArgs) Handles btnLogout.Click
        Me.SignOutAndRedirect()
    End Sub
#End Region

#Region "内部処理"
    ''' <summary>
    ''' メニューをブロック単位で画面に割り当てます。
    ''' </summary>
    Private Sub BindMenu()
        Me.rptBlocks.DataSource = ScreenCatalog.Current.Blocks
        Me.rptBlocks.DataBind()
    End Sub

    ''' <summary>
    ''' 権限不足で戻された場合に、その旨を表示します。
    ''' </summary>
    Private Sub ShowDeniedMessage()
        Dim denied As String = Me.Request.QueryString(DeniedKey)
        ' 権限不足の戻りでないとき
        If String.IsNullOrEmpty(denied) Then
            Return
        End If

        Dim menuNo As Integer = 0
        ' メニュー番号を読み取れたとき
        If Integer.TryParse(denied, menuNo) Then
            Me.ShowMessage("処理権限がありません。（メニュー番号 " & menuNo.ToString() & "）", True)
        End If
    End Sub
#End Region

End Class
