Option Strict On
Option Explicit On
Option Infer Off

Imports System.Reflection
Imports System.Web.UI
Imports Atom.Core.Logging
Imports Atom.Data
Imports Atom.Web.Common

''' <summary>
''' 全画面共通のレイアウトです。WinForms 版の BaseForm の外枠に相当します。
''' ヘッダー、メッセージ領域、フッター（再接続）を提供します。
''' </summary>
Public Class SiteMaster
    Inherits MasterPage

#Region "公開メソッド"
    ''' <summary>
    ''' メッセージ領域に文言を表示します。
    ''' </summary>
    ''' <param name="message">表示する文言。空のときは領域を隠します。</param>
    ''' <param name="isError">エラーとして表示するとき True。</param>
    Public Sub ShowMessage(message As String, isError As Boolean)
        ' 文言が無いとき
        If String.IsNullOrEmpty(message) Then
            Me.pnlMessage.Visible = False
            Return
        End If

        Me.litMessage.Text = message
        ' エラーのとき
        If isError Then
            Me.pnlMessage.CssClass = "atom-message atom-message-error"
        Else
            Me.pnlMessage.CssClass = "atom-message atom-message-info"
        End If
        Me.pnlMessage.Visible = True
    End Sub
#End Region

#Region "イベントハンドラ"
    ''' <summary>
    ''' 初期化時の処理です。画面ごとのヘッダー設定を反映します。
    ''' </summary>
    ''' <param name="e">イベント引数。</param>
    Protected Overrides Sub OnInit(e As EventArgs)
        MyBase.OnInit(e)

        Dim basePage As BasePage = TryCast(Me.Page, BasePage)
        ' BasePage を継承した画面のとき
        If basePage IsNot Nothing Then
            Me.hdrHeader.Apply(basePage.CreateHeaderConfig())
        Else
            Me.hdrHeader.Apply(New HeaderConfig())
        End If

        Me.litVersion.Text = "ATOM  ver." & Assembly.GetExecutingAssembly().GetName().Version.ToString()
    End Sub

    ''' <summary>
    ''' 再接続ボタン押下時の処理です。データベースへの接続を確認します。
    ''' </summary>
    ''' <param name="sender">イベントの発生元。</param>
    ''' <param name="e">イベント引数。</param>
    Private Sub BtnReconnect_Click(sender As Object, e As EventArgs) Handles btnReconnect.Click
        Try
            ' 接続できたとき
            If Database.TestConnection() Then
                Me.ShowMessage("再接続しました。", False)
            Else
                Me.ShowMessage("再接続に失敗しました。", True)
            End If
        Catch ex As Exception
            Logger.WriteError("SiteMaster", "再接続に失敗しました。", ex, SessionContext.Current.LoginCode)
            Me.ShowMessage("再接続に失敗しました。", True)
        End Try
    End Sub
#End Region

End Class
