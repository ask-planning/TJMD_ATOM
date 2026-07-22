Imports System.Collections.Generic
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms
Imports Atom.App.Common
Imports Atom.Data.Repositories

Namespace Forms.Expense
    ''' <summary>経費コード登録（現行 F_M_経費コード登録_MAIN + _SUB を1フォームに統合）。tbl_m_import_keihi_code を直接一覧表示する。</summary>
    Public Class ExpenseCodeMaintenanceForm

#Region "フィールド"
        Private ReadOnly _repo As New ExpenseCodeRepository()
        Private _table As DataTable
#End Region

#Region "コンストラクター"
        ''' <summary>画面を初期化する（タイトルは menuNo からリソース解決）。</summary>
        Public Sub New()
            InitializeComponent()
            Me.MenuNo = 907
        End Sub
#End Region

#Region "拡張ポイント"
        ''' <summary>この画面のヘッダー設定：タイトル「輸入経費コード登録」・背景色オレンジ・「閉じる」のみ。</summary>
        Protected Overrides Function CreateHeaderConfig() As HeaderConfig
            Dim config As New HeaderConfig()
            config.Title = "輸入経費コード登録"
            config.HeaderBackColor = ColorTranslator.FromHtml("#FF8040")
            config.Buttons.Add(New HeaderButton("閉じる", AddressOf Me.OnCloseRequested))
            Return config
        End Function
#End Region

#Region "イベントハンドラー"
        ''' <summary>読み込み時：コンボ候補の設定と初期表示（一覧の取得・表示）を行う。</summary>
        Protected Overrides Sub OnLoad(e As EventArgs)
            MyBase.OnLoad(e)
            Me.SetupCombos()
            Me.LoadData()
        End Sub

        ''' <summary>グリッドの表示エラーは無視して画面を保つ（防御的）。</summary>
        Private Sub OnGridDataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles dgvList.DataError
            e.ThrowException = False
        End Sub

        Private Sub OnGridCellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles dgvList.CellFormatting
            ' 伝票コード列：セル確定表示はコードのみにする（生値が「コード　名称」なので先頭のコードを取り出す）
            If e.ColumnIndex >= 0 AndAlso Me.dgvList.Columns(e.ColumnIndex).Name = "colDenpyo" AndAlso e.Value IsNot Nothing Then
                Dim raw As String = Convert.ToString(e.Value)
                ' 全角スペース／半角スペースどちらの区切りでも先頭部分を取り出す
                Dim head As String = raw.Split(New Char() {"　"c, " "c})(0).Trim()
                Dim codeValue As Integer
                ' 数値化できれば4桁ゼロ埋め、できなければ取り出した文字をそのまま
                If Integer.TryParse(head, codeValue) Then
                    e.Value = codeValue.ToString("D4")
                Else
                    e.Value = head
                End If
                e.FormattingApplied = True
            End If
        End Sub
#End Region

#Region "内部メソッド"

        ''' <summary>伝票コード・区分のコンボ候補をマスタから設定する（取得失敗時は空のまま）。</summary>
        Private Sub SetupCombos()
            ' 伝票コード：コード＋名称を「コード　名称」で表示（値はコード）
            Try
                Dim codes As DataTable = Me._repo.SelectCodeList()
                If Not codes.Columns.Contains("display") Then
                    codes.Columns.Add("display", GetType(String))
                End If
                For Each codeRow As DataRow In codes.Rows
                    codeRow("display") = Convert.ToString(codeRow("pk_keihi_code")) & "　" & Convert.ToString(codeRow("keihi_code_text"))
                Next
                Me.colDenpyo.DataSource = codes
                Me.colDenpyo.ValueMember = "pk_keihi_code"
                Me.colDenpyo.DisplayMember = "display"
            Catch ex As Exception
                ' 取得失敗時：候補なし（表示は素の値）
            End Try
            ' 区分：View__keihi_kbn から（値=コード / 表示=名称）
            Try
                Dim kbn As DataTable = Me._repo.SelectKbnList()
                Me.colKbn.DataSource = kbn
                Me.colKbn.ValueMember = "code"
                Me.colKbn.DisplayMember = "name"
            Catch ex As Exception
                ' 取得失敗時：候補なし
            End Try
        End Sub

        ''' <summary>初期表示：経費コードマスタを取得してグリッドへ直接バインドする（現行の直接バインドを再現）。</summary>
        Private Sub LoadData()
            Try
                ' 現行同様、tbl_m_import_keihi_code を取得
                Me._table = Me._repo.SelectAll()
                ' 取得したテーブルを直接バインド（型付きDSを介さない）
                Me.bsExpense.DataMember = ""
                Me.bsExpense.DataSource = Me._table
            Catch ex As Exception
                ' 取得失敗時：原因を通知し、一覧は空にする（列はデザイナー定義のまま残る）
                Me.bsExpense.DataMember = ""
                Me.bsExpense.DataSource = Nothing
                MessageBox.Show("一覧の取得に失敗しました。" & vbCrLf & ex.Message, "読み込み",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub
#End Region

    End Class
End Namespace