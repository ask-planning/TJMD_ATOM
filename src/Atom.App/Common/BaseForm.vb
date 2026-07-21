Imports System.Drawing
Imports System.Windows.Forms

Namespace Common
    ''' <summary>
    ''' 全画面の基底フォーム。ヘッダー・フッター（閉じる/再接続）・起動時権限チェック・最大化を提供する。
    ''' 各画面はこれを継承し、ContentPanel に固有UIを載せる。
    ''' </summary>
    Public Class BaseForm
        Inherits Form

#Region "フィールド"
        Private ReadOnly hdrHeader As New HeaderControl()
        Private ReadOnly pnlFooter As New Panel()
        Private ReadOnly btnClose As New Button()
        Private ReadOnly btnReconnect As New Button()
        Private ReadOnly pnlContent As New Panel()
#End Region

#Region "プロパティ"
        ''' <summary>各画面が固有UIを配置する領域。</summary>
        Protected ReadOnly Property ContentPanel As Panel
            Get
                Return Me.pnlContent
            End Get
        End Property

        ''' <summary>ヘッダーに表示する画面名。</summary>
        Public Property ScreenTitle As String = ""
        ''' <summary>権限判定に使うメニュー番号（0なら判定しない）。</summary>
        Public Property MenuNo As Integer = 0
#End Region

#Region "コンストラクター"
        ''' <summary>基底フォームを初期化する（ヘッダー・フッター・本体領域を配置）。</summary>
        Public Sub New()
            Me.Font = New Font("Meiryo UI", 9.0!)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.WindowState = FormWindowState.Maximized

            ' フッター（下端）に 再接続 / 閉じる を右寄せで並べる
            Me.pnlFooter.Dock = DockStyle.Bottom
            Me.pnlFooter.Height = 46
            Me.pnlFooter.Padding = New Padding(8)
            Me.btnReconnect.Text = "再接続"
            Me.btnReconnect.Width = 100
            Me.btnReconnect.Dock = DockStyle.Right
            Me.btnClose.Text = "閉じる"
            Me.btnClose.Width = 100
            Me.btnClose.Dock = DockStyle.Right
            AddHandler Me.btnClose.Click, AddressOf Me.OnCloseClick
            AddHandler Me.btnReconnect.Click, AddressOf Me.OnReconnectClick
            Dim pnlSpacer As New Panel() With {.Width = 6, .Dock = DockStyle.Right}
            Me.pnlFooter.Controls.Add(Me.btnClose)
            Me.pnlFooter.Controls.Add(pnlSpacer)
            Me.pnlFooter.Controls.Add(Me.btnReconnect)

            ' 本体領域は残りを埋める
            Me.pnlContent.Dock = DockStyle.Fill

            ' Fill を先、次に端（Bottom/Top）を追加すると正しく段組みされる
            Me.Controls.Add(Me.pnlContent)
            Me.Controls.Add(Me.pnlFooter)
            Me.Controls.Add(Me.hdrHeader)
        End Sub
#End Region

#Region "イベントハンドラー"
        ''' <summary>読み込み時：タイトル解決・ヘッダー反映・権限チェックを行う。</summary>
        Protected Overrides Sub OnLoad(e As EventArgs)
            MyBase.OnLoad(e)
            ' 画面タイトルを menuNo からリソース解決する（未設定のときだけ）
            If String.IsNullOrEmpty(Me.ScreenTitle) AndAlso Me.MenuNo > 0 Then
                Me.ScreenTitle = ScreenTitles.GetTitle(Me.MenuNo)
            End If
            ' ウィンドウ見出しとヘッダーへ反映する
            Me.Text = Me.ScreenTitle
            Me.hdrHeader.Bind(SessionContext.Current, Me.ScreenTitle)
            ' 権限が無いとき：メッセージを出して閉じる
            If Me.MenuNo > 0 AndAlso Not SessionContext.Current.IsAllowed(Me.MenuNo) Then
                MessageBox.Show("処理権限がありません。", "権限", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Me.BeginInvoke(New Action(Sub() Me.Close()))
            End If
        End Sub

        ''' <summary>「閉じる」動作。既定はフォームを閉じ、呼び出し元へ戻る。</summary>
        Protected Overridable Sub OnCloseClick(sender As Object, e As EventArgs)
            Me.Close()
        End Sub

        ''' <summary>「再接続」動作。</summary>
        Private Sub OnReconnectClick(sender As Object, e As EventArgs)
            ReconnectService.Reconnect(Me)
        End Sub
#End Region

    End Class
End Namespace
