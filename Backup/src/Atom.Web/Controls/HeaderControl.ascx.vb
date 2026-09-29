Option Strict On
Option Explicit On
Option Infer Off

Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports Atom.Web.Common

''' <summary>
''' 全画面共通のヘッダーです。画面名・利用者名・画面固有ボタンを表示します。
''' </summary>
Public Class HeaderControl
    Inherits UserControl

#Region "公開メソッド"
    ''' <summary>
    ''' ヘッダー設定を反映します。マスターページの初期化時に呼び出します。
    ''' </summary>
    ''' <param name="config">画面ごとのヘッダー設定。</param>
    Public Sub Apply(config As HeaderConfig)
        ' 設定が無いときは既定値で表示する
        If config Is Nothing Then
            config = New HeaderConfig()
        End If

        Me.lblTitle.Text = Me.Server.HtmlEncode(config.Title)
        Me.divHeader.Style("background-color") = config.BackColor

        ' 利用者名を表示する
        Dim session As SessionContext = SessionContext.Current
        Me.lblName.Text = Me.Server.HtmlEncode(session.LoginName)

        Me.BuildButtons(config)
    End Sub
#End Region

#Region "内部処理"
    ''' <summary>
    ''' ヘッダーのボタンを組み立てます。
    ''' </summary>
    ''' <param name="config">画面ごとのヘッダー設定。</param>
    Private Sub BuildButtons(config As HeaderConfig)
        Me.phButtons.Controls.Clear()

        For Each button As HeaderButton In config.Buttons
            Dim link As New HyperLink()
            link.Text = Me.Server.HtmlEncode(button.Text)
            link.NavigateUrl = button.NavigateUrl
            link.CssClass = "atom-header-button"
            Me.phButtons.Controls.Add(link)
        Next
    End Sub
#End Region

End Class
