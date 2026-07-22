Imports System.Windows.Forms

Namespace Common
    ''' <summary>各画面共通のヘッダー部品（氏名/画面名/ボタン群）。レイアウトはデザイナー(HeaderControl.Designer.vb)。</summary>
    Public Class HeaderControl

        ''' <summary>コンストラクター（デザイナー初期化）。</summary>
        Public Sub New()
            InitializeComponent()
        End Sub

        ''' <summary>セッション情報を表示へ反映する（氏名のみ）。</summary>
        ''' <param name="session">ログイン中の利用者情報。</param>
        ''' <param name="screenTitle">画面タイトル（Apply未使用時の既定値として反映）。</param>
        Public Sub Bind(session As SessionContext, screenTitle As String)
            ' ヘッダー各項目へ反映する（氏名と画面名のみ）
            Me.lblName.Text = session.LoginName
            Me.lblTitle.Text = screenTitle
        End Sub

        ''' <summary>画面別のヘッダー設定（タイトル・背景色・ボタン構成）を反映する。</summary>
        ''' <param name="config">反映するヘッダー設定。</param>
        Public Sub Apply(config As HeaderConfig)
            ' タイトルと背景色を反映する（Bindでの初期値を上書きする）
            Me.lblTitle.Text = config.Title
            Me.BackColor = config.HeaderBackColor
            ' 既存のボタンを消してから、設定分を作り直す
            Me.pnlButtons.Controls.Clear()
            For Each buttonConfig As HeaderButton In config.Buttons
                Dim button As New Button()
                button.Text = buttonConfig.Text
                button.AutoSize = True
                button.Margin = New Padding(4, 2, 0, 2)
                button.UseVisualStyleBackColor = True
                AddHandler button.Click, Sub(sender As Object, e As EventArgs) buttonConfig.OnClick.Invoke()
                Me.pnlButtons.Controls.Add(button)
            Next
        End Sub
    End Class
End Namespace