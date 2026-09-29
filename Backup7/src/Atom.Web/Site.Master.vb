Option Strict On
Option Explicit On
Option Infer Off

Imports System.Web.UI
Imports Atom.Web.Common

''' <summary>
''' 全画面共通のレイアウトです。共通ヘッダーとメッセージ領域を持ちます。
''' </summary>
Public Class SiteMaster
    Inherits MasterPage

#Region "定数"
    ''' <summary>メッセージ領域の通常表示に使う CSS クラス。</summary>
    Private Const InformationCssClass As String = "atom-message atom-message-information"

    ''' <summary>メッセージ領域のエラー表示に使う CSS クラス。</summary>
    Private Const ErrorCssClass As String = "atom-message atom-message-error"
#End Region

#Region "公開メソッド"
    ''' <summary>
    ''' メッセージ領域に文言を表示します。
    ''' </summary>
    ''' <param name="message">表示する文言。</param>
    ''' <param name="isError">エラーとして表示するとき True。</param>
    Public Sub ShowMessage(message As String, isError As Boolean)
        ' 文言が無いときは表示しない
        If String.IsNullOrEmpty(message) Then
            Me.pnlMessage.Visible = False
            Return
        End If

        Me.litMessage.Text = message
        Me.pnlMessage.Visible = True

        ' エラーのとき
        If isError Then
            Me.pnlMessage.CssClass = ErrorCssClass
        ' エラーでないとき
        Else
            Me.pnlMessage.CssClass = InformationCssClass
        End If
    End Sub
#End Region

#Region "イベントハンドラ"
    ''' <summary>
    ''' 描画前の処理です。表示中の画面に合わせてヘッダーを設定します。
    ''' </summary>
    ''' <param name="e">イベント引数。</param>
    Protected Overrides Sub OnPreRender(e As EventArgs)
        MyBase.OnPreRender(e)
        Me.hdrHeader.ApplyFor(TryCast(Me.Page, BasePage))
    End Sub
#End Region

End Class
