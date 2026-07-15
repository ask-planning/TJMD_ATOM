\xEF\xBB\xBFImports System.Drawing
Imports System.Windows.Forms
Imports Atom.App.Common

''' <summary>メインメニュー画面（現行 F_0_START_MENU 相当）。赤枠4ブロックのメニューを表示し各画面へ遷移する。</summary>
Public Class MainForm

#Region "フィールド"
    Private ReadOnly hdrHeader As New HeaderControl()
#End Region

#Region "イベントハンドラー"
    ''' <summary>読み込み時：仮ログイン設定・メニュー生成・ヘッダー表示を行う。</summary>
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 仮ログイン（TODO: 実ログイン画面 + MenuPermissionRepository で権限取得）
        Dim session As SessionContext = SessionContext.Current
        session.LoginCode = "000001"
        session.LoginName = "山田 太郎"
        session.Department = "輸入部"
        session.SetPermissions(New Integer() {101, 102, 103, 104, 105, 106, 107,
                                              201, 202, 203, 301,
                                              902, 903, 904, 905, 906, 907})

        Me.Text = "ATOM メインメニュー"
        Me.Font = New Font("Meiryo UI", 9.0!)

        ' メニュー（Fill）を先に、ヘッダー（Top）を後に追加する
        Me.BuildMenu()
        Me.Controls.Add(Me.hdrHeader)
        Me.hdrHeader.Bind(session, ScreenTitles.GetTitle(0))
    End Sub

    ''' <summary>メニューボタン押下：権限確認→画面生成→表示。</summary>
    Private Sub OnMenuClick(sender As Object, e As EventArgs)
        ' 押されたボタンから menuNo を取り出す
        Dim button As Button = DirectCast(sender, Button)
        Dim menuNo As Integer = CInt(button.Tag)
        ' 権限が無いとき：メッセージを出して終了
        If Not SessionContext.Current.IsAllowed(menuNo) Then
            MessageBox.Show("処理権限がありません。", "権限", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        ' 対応画面を生成する（未実装なら Nothing）
        Dim screen As Form = ScreenFactory.Create(menuNo)
        If screen Is Nothing Then
            MessageBox.Show("この画面は未実装です。（menuNo=" & menuNo.ToString() & "）",
                            "未実装", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        ' 画面を開く
        NavigationManager.OpenForm(screen, Me)
    End Sub
#End Region

#Region "内部メソッド"
    ''' <summary>赤枠4ブロックのメニューを組み立てる。</summary>
    Private Sub BuildMenu()
        ' 横並び・折り返しのレイアウト領域
        Dim flow As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill, .Padding = New Padding(14), .AutoScroll = True,
            .FlowDirection = FlowDirection.LeftToRight, .WrapContents = True}

        ' 上海輸入定期便＆AIR便
        Dim block1 As GroupBox = Me.CreateBlock("上海輸入定期便＆AIR便")
        Me.AddButton(block1, "INVOICE取込", 101)
        Me.AddButton(block1, "INVOICE確認", 102)
        Me.AddButton(block1, "INVOICEヘッダー入力", 103)
        Me.AddButton(block1, "R3購買発注取込", 104)
        Me.AddButton(block1, "R3購買発注確認", 105)
        Me.AddButton(block1, "R3在庫転送作成", 106)
        Me.AddButton(block1, "SAP伝票金額入力", 107)
        flow.Controls.Add(block1)

        ' その他輸入関連（調達）
        Dim block2 As GroupBox = Me.CreateBlock("その他輸入関連")
        Me.AddButton(block2, "調達R3購買発注取込", 201)
        Me.AddButton(block2, "調達R3購買発注確認", 202)
        Me.AddButton(block2, "調達配送手配書作成", 203)
        flow.Controls.Add(block2)

        ' マスタ関連
        Dim block3 As GroupBox = Me.CreateBlock("マスタ関連")
        Me.AddButton(block3, "社内参照番号発番", 902)
        Me.AddButton(block3, "項目マスタ", 903)
        Me.AddButton(block3, "カレンダー", 904)
        Me.AddButton(block3, "仕入先マスタ登録", 905)
        Me.AddButton(block3, "商品名称入力", 906)
        Me.AddButton(block3, "経費コード登録", 907)
        flow.Controls.Add(block3)

        ' 輸入経費入力
        Dim block4 As GroupBox = Me.CreateBlock("輸入経費入力")
        Me.AddButton(block4, "諸経費入力", 301)
        flow.Controls.Add(block4)

        ' OTHERS（終了）
        Dim block5 As GroupBox = Me.CreateBlock("OTHERS")
        Dim btnExit As New Button() With {.Text = "終  了", .Width = 210, .Height = 40, .Margin = New Padding(4)}
        AddHandler btnExit.Click, Sub() Me.Close()
        DirectCast(block5.Tag, FlowLayoutPanel).Controls.Add(btnExit)
        flow.Controls.Add(block5)

        Me.Controls.Add(flow)
    End Sub

    ''' <summary>ブロック（GroupBox＋縦並び領域）を作る。</summary>
    ''' <param name="title">ブロック見出し。</param>
    Private Function CreateBlock(title As String) As GroupBox
        Dim box As New GroupBox() With {.Text = title, .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink, .Margin = New Padding(10),
            .Padding = New Padding(8), .MinimumSize = New Size(236, 0)}
        Dim inner As New FlowLayoutPanel() With {
            .FlowDirection = FlowDirection.TopDown, .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink, .WrapContents = False,
            .Dock = DockStyle.Top, .Padding = New Padding(4, 6, 4, 4)}
        box.Controls.Add(inner)
        box.Tag = inner
        Return box
    End Function

    ''' <summary>ブロックへメニューボタンを1つ追加する。</summary>
    ''' <param name="block">追加先ブロック。</param>
    ''' <param name="label">ボタン表示名。</param>
    ''' <param name="menuNo">メニュー番号。</param>
    Private Sub AddButton(block As GroupBox, label As String, menuNo As Integer)
        Dim inner As FlowLayoutPanel = DirectCast(block.Tag, FlowLayoutPanel)
        Dim button As New Button() With {.Text = label, .Width = 210, .Height = 34, .Margin = New Padding(4)}
        button.Tag = menuNo
        AddHandler button.Click, AddressOf Me.OnMenuClick
        inner.Controls.Add(button)
    End Sub
#End Region

End Class
