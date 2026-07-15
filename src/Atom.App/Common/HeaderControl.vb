\xEF\xBB\xBFNamespace Common
    ''' <summary>各画面共通のヘッダー部品（ログイン/氏名/部門/画面名）。レイアウトはデザイナー(HeaderControl.Designer.vb)。</summary>
    Public Class HeaderControl

        ''' <summary>コンストラクター（デザイナー初期化）。</summary>
        Public Sub New()
            InitializeComponent()
        End Sub

        ''' <summary>セッション情報と画面名を表示へ反映する。</summary>
        ''' <param name="session">ログイン中の利用者情報。</param>
        ''' <param name="screenTitle">画面タイトル。</param>
        Public Sub Bind(session As SessionContext, screenTitle As String)
            ' ヘッダー各項目へ反映する
            Me.lblLogin.Text = session.LoginCode
            Me.lblName.Text = session.LoginName
            Me.lblDept.Text = session.Department
            Me.lblTitle.Text = screenTitle
        End Sub
    End Class
End Namespace
