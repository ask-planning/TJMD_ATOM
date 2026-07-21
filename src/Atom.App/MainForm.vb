Imports System.Windows.Forms
Imports Atom.App.Common

''' <summary>メインメニュー画面（現行 F_0_START_MENU 相当）。レイアウトはデザイナー(MainForm.Designer.vb)。</summary>
Public Class MainForm

#Region "イベントハンドラー"
    ''' <summary>読み込み時：仮ログイン設定とヘッダー表示を行う。</summary>
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 仮ログイン（TODO: 実ログイン画面 + MenuPermissionRepository で権限取得）
        Dim session As SessionContext = SessionContext.Current
        session.LoginCode = "000001"
        session.LoginName = "山田 太郎"
        session.Department = "輸入部"
        session.SetPermissions(New Integer() {101, 102, 103, 104, 105, 106, 107, 201, 202, 203, 902, 903, 904, 905, 906, 907, 301})

        ' ヘッダーへ反映する
        Me.hdrHeader.Bind(session, ScreenTitles.GetTitle(0))
    End Sub

    ''' <summary>メニューボタン押下：権限確認→画面生成→表示。</summary>
    Private Sub OnMenuClick(sender As Object, e As EventArgs) Handles btnInvoiceImport.Click, btnInvoiceCheck.Click, btnInvoiceHeader.Click, btnR3PoImport.Click, btnR3PoCheck.Click, btnR3StockTransfer.Click, btnSapAmount.Click, btnProcPoImport.Click, btnProcPoCheck.Click, btnProcDelivery.Click, btnRefNumber.Click, btnItemMaster.Click, btnCalendar.Click, btnSupplierMaster.Click, btnProductName.Click, btnExpenseCode.Click, btnMiscExpense.Click
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

    ''' <summary>終了ボタン：画面を閉じる。</summary>
    Private Sub OnExitClick(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
#End Region

End Class
