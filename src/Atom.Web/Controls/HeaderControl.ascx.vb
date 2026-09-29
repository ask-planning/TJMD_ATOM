Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports Atom.Web.Common

''' <summary>
''' 全画面共通のヘッダーです。
''' メニュー画面では ATOM のタイトルを、それ以外の画面では画面名を表示します。
''' </summary>
Public Class HeaderControl
    Inherits UserControl

#Region "定数"
    ''' <summary>バージョン表示の接頭辞。</summary>
    Private Const VersionPrefix As String = "Ver."

    ''' <summary>メニュー画面のメニュー番号。</summary>
    Private Const MenuScreenNo As Integer = 0

    ''' <summary>マニュアルの URL が未設定のときの文言。</summary>
    Private Const NotImplementedMessage As String = "現在未実装です。"

    ''' <summary>色の指定として認める書式。</summary>
    Private Const ColorPattern As String = "^#[0-9A-Fa-f]{6}$"

    ''' <summary>色の指定が想定外だったときに使う背景色。</summary>
    Private Const DefaultHeaderBackColor As String = "#400080"

    ''' <summary>画面固有ボタン押下時のコマンド名。</summary>
    Private Const NavigateCommandName As String = "Navigate"
#End Region

#Region "公開メソッド"
    ''' <summary>
    ''' 表示中の画面に合わせてヘッダーの内容を設定します。Site.Master から呼び出します。
    ''' </summary>
    ''' <param name="page">表示中の画面。</param>
    Public Sub ApplyFor(page As BasePage)
        Me.litVersion.Text = VersionPrefix & AppSettings.SystemVersion

        Dim menuNo As Integer = MenuScreenNo
        ' 基底ページを継承した画面のとき
        If page IsNot Nothing Then
            menuNo = page.MenuNo
        End If

        Dim catalog As ScreenCatalog = ScreenCatalog.Current

        ' メニュー画面・ログイン画面のとき（ATOM のタイトルを表示する）
        If menuNo = MenuScreenNo Then
            Me.pnlLogo.Visible = True
            Me.pnlScreenTitle.Visible = False
        ' 業務画面のとき（画面名を表示する）
        Else
            Me.pnlLogo.Visible = False
            Me.pnlScreenTitle.Visible = True
            Me.litScreenTitle.Text = catalog.GetTitle(menuNo)
        End If

        ' 画面ごとの背景色は外部定義で切り替えるため、CSS クラスではなく直接指定します。
        ' 外部ファイルの値をそのまま style 属性へ出さないよう、色の書式を検証します。
        Me.divHeader.Style("background-color") = Me.ToSafeColor(catalog.GetHeaderBackColor(menuNo))

        ' 画面固有ボタンは画面ごとに件数が変わるためデータバインドで描画します。
        Dim actions As IList(Of HeaderAction) = New List(Of HeaderAction)()
        If page IsNot Nothing Then
            actions = page.HeaderActions
        End If
        Me.rptActions.DataSource = actions
        Me.rptActions.DataBind()

        Dim session As SessionContext = SessionContext.Current
        Me.pnlUser.Visible = session.IsLoggedIn()
        Me.litLoginName.Text = session.LoginName
    End Sub
#End Region

#Region "イベントハンドラ"
    ''' <summary>
    ''' 画面固有ボタン押下時の処理です。共通の画面遷移処理へ渡します。
    ''' </summary>
    ''' <param name="source">イベントの発生元。</param>
    ''' <param name="e">イベント引数。</param>
    Private Sub RptActions_ItemCommand(source As Object, e As RepeaterCommandEventArgs) Handles rptActions.ItemCommand
        ' 想定外のコマンドのとき
        If Not String.Equals(e.CommandName, NavigateCommandName, StringComparison.Ordinal) Then
            Return
        End If

        Dim page As BasePage = TryCast(Me.Page, BasePage)
        ' 基底ページを継承していないとき
        If page Is Nothing Then
            Return
        End If

        Dim menuNo As Integer = 0
        ' メニュー番号を読み取れないとき
        If Not Integer.TryParse(CStr(e.CommandArgument), menuNo) Then
            Return
        End If

        ' 遷移しないボタンのとき（ブラウザ側だけで処理するため、ここでは何もしない）
        If menuNo < HeaderAction.MenuScreenNo Then
            Return
        End If

        ' メニュー画面へ戻るとき
        If menuNo = HeaderAction.MenuScreenNo Then
            page.NavigateToMenu()
            Return
        End If
        page.NavigateTo(menuNo)
    End Sub
#End Region

#Region "イベントハンドラ"
    ''' <summary>
    ''' マニュアルボタン押下時の処理です。設定された URL を新しいタブで開きます。
    ''' </summary>
    ''' <param name="sender">イベントの発生元。</param>
    ''' <param name="e">イベント引数。</param>
    Private Sub LnkManual_Click(sender As Object, e As EventArgs) Handles lnkManual.Click
        Dim url As String = AppSettings.ManualUrl
        Dim page As BasePage = TryCast(Me.Page, BasePage)

        ' URL が未設定のとき
        If String.IsNullOrEmpty(url) Then
            ' メッセージ表示のため、基底ページを継承した画面でのみ通知します。
            If page IsNot Nothing Then
                page.ShowMessage(NotImplementedMessage, False)
            End If
            Return
        End If

        ' URL が設定されているとき（新しいタブで開く）
        Me.Page.ClientScript.RegisterStartupScript(Me.GetType(), "AtomOpenManual",
            "AtomPage.openManual(" & Me.CreateJsonUrl(url) & ");", True)
    End Sub
#End Region

#Region "内部処理"
    ''' <summary>
    ''' 色の指定を検証します。#RRGGBB 以外のときは既定色を返します。
    ''' </summary>
    ''' <param name="color">画面定義から読み込んだ色の指定。</param>
    ''' <returns>style 属性へ出力してよい色の指定。</returns>
    Private Function ToSafeColor(color As String) As String
        ' 書式が #RRGGBB のとき
        If Not String.IsNullOrEmpty(color) AndAlso Regex.IsMatch(color, ColorPattern) Then
            Return color
        End If
        ' 書式が想定外のとき
        Return DefaultHeaderBackColor
    End Function

    ''' <summary>
    ''' URL をスクリプトへ渡せる形式の文字列リテラルにします。
    ''' </summary>
    ''' <param name="url">対象の URL。</param>
    ''' <returns>引用符で囲んだ文字列リテラル。</returns>
    Private Function CreateJsonUrl(url As String) As String
        Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
        Return serializer.Serialize(url)
    End Function
#End Region

End Class
