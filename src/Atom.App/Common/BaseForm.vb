Imports System.Windows.Forms

Namespace Common
    ''' <summary>
    ''' 全画面の基底フォーム。ヘッダー・フッター（閉じる/再接続）・起動時権限チェック・最大化を提供する。
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
        Protected Overridable Sub OnCloseClick(sender As Object, e As EventArgs) Handles btnClose.Click
            Me.Close()
        End Sub

        ''' <summary>「再接続」動作。</summary>
        Private Sub OnReconnectClick(sender As Object, e As EventArgs) Handles btnReconnect.Click
            ReconnectService.Reconnect(Me)
        End Sub
#End Region

    End Class
End Namespace
