Imports System.Drawing
Imports System.Windows.Forms

Namespace Common
    ''' <summary>
    ''' 全画面の基底フォーム。ヘッダー・フッター（再接続）・起動時権限チェック・最大化を提供する。
    ''' ヘッダーのタイトル・背景色・ボタン構成は CreateHeaderConfig() を画面ごとに上書きして指定する。
    ''' レイアウトはデザイナー(BaseForm.Designer.vb)。各画面はこれを継承し pnlContent に固有UIを載せる。
    ''' </summary>
    Public Class BaseForm

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
        ''' <summary>基底フォームを初期化する。</summary>
        Public Sub New()
            InitializeComponent()
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
            ' ウィンドウ見出しへ反映する
            Me.Text = Me.ScreenTitle
            ' ヘッダーへ、ログイン情報→画面別設定の順で反映する（画面別設定が最終的に優先される）
            Me.hdrHeader.Bind(SessionContext.Current, Me.ScreenTitle)
            Me.hdrHeader.Apply(Me.CreateHeaderConfig())
            ' 権限が無いとき：メッセージを出して閉じる
            If Me.MenuNo > 0 AndAlso Not SessionContext.Current.IsAllowed(Me.MenuNo) Then
                MessageBox.Show("処理権限がありません。", "権限", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Me.BeginInvoke(New Action(Sub() Me.Close()))
            End If
        End Sub

        ''' <summary>「再接続」動作。</summary>
        Private Sub OnReconnectClick(sender As Object, e As EventArgs) Handles btnReconnect.Click
            ReconnectService.Reconnect(Me)
        End Sub
#End Region

#Region "拡張ポイント"
        ''' <summary>
        ''' ヘッダー設定（タイトル・背景色・ボタン構成）を返す。
        ''' 既定はタイトル＝ScreenTitle・標準背景色・「閉じる」ボタンのみ。画面ごとに上書きする。
        ''' </summary>
        Protected Overridable Function CreateHeaderConfig() As HeaderConfig
            Dim config As New HeaderConfig()
            config.Title = Me.ScreenTitle
            config.HeaderBackColor = SystemColors.Control
            config.Buttons.Add(New HeaderButton("閉じる", AddressOf Me.OnCloseRequested))
            Return config
        End Function

        ''' <summary>「閉じる」ボタン押下時の既定動作。既定はフォームを閉じる。画面ごとに上書き可能。</summary>
        Protected Overridable Sub OnCloseRequested()
            Me.Close()
        End Sub
#End Region

    End Class
End Namespace
