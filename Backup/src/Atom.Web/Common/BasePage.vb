Option Strict On
Option Explicit On
Option Infer Off

Imports System.Web
Imports System.Web.UI

Namespace Common

    ''' <summary>
    ''' 全画面の基底ページです。WinForms 版の BaseForm に相当します。
    ''' ヘッダー設定、権限チェック、メッセージ表示の共通処理を提供します。
    ''' </summary>
    Public Class BasePage
        Inherits Page

#Region "プロパティ"
        ''' <summary>
        ''' 権限判定に使うメニュー番号です。0 のときは判定しません。
        ''' 各画面のコンストラクタまたは OnInit で設定します。
        ''' </summary>
        Public Property MenuNo As Integer = 0
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' この画面のヘッダー設定を作成します。画面ごとに上書きしてください。
        ''' </summary>
        ''' <returns>ヘッダー設定。</returns>
        Public Overridable Function CreateHeaderConfig() As HeaderConfig
            Dim config As New HeaderConfig()
            config.Title = String.Empty
            config.BackColor = HeaderConfig.DefaultBackColor
            ' メニュー以外の画面には「閉じる」ボタンを置く
            If Me.MenuNo > 0 Then
                config.AddCloseButton()
            End If
            Return config
        End Function

        ''' <summary>
        ''' 画面上のメッセージ領域に文言を表示します。
        ''' </summary>
        ''' <param name="message">表示する文言。</param>
        ''' <param name="isError">エラーとして表示するとき True。</param>
        Public Sub ShowMessage(message As String, isError As Boolean)
            Dim master As SiteMaster = TryCast(Me.Master, SiteMaster)
            ' マスターページが取得できないときは何もしない
            If master Is Nothing Then
                Return
            End If
            master.ShowMessage(message, isError)
        End Sub
#End Region

#Region "イベントハンドラ"
        ''' <summary>
        ''' 初期化時の処理です。CSRF 対策のキーを設定します。
        ''' </summary>
        ''' <param name="e">イベント引数。</param>
        Protected Overrides Sub OnInit(e As EventArgs)
            ' ViewState の改ざんと CSRF を防ぐためセッション ID を鍵にする
            If Me.Session IsNot Nothing Then
                Me.ViewStateUserKey = Me.Session.SessionID
            End If
            MyBase.OnInit(e)
        End Sub

        ''' <summary>
        ''' 読み込み時の処理です。権限が無い場合はメニューへ戻します。
        ''' </summary>
        ''' <param name="e">イベント引数。</param>
        Protected Overrides Sub OnLoad(e As EventArgs)
            MyBase.OnLoad(e)

            ' 権限が無いとき
            If Not PermissionChecker.IsAllowed(Me.MenuNo) Then
                Me.Response.Redirect("~/Default.aspx?denied=" & Me.MenuNo.ToString(), False)
                Me.Context.ApplicationInstance.CompleteRequest()
                Return
            End If
        End Sub
#End Region

    End Class
End Namespace
