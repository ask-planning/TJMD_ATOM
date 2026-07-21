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
            Me.SetupAppearance()
        End Sub
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
#End Region

#Region "内部メソッド"
        ''' <summary>入力列（伝票コード/区分/保険対象）のセルに背景色を付ける。</summary>
        Private Sub SetupAppearance()
            ' 入力欄の背景色（うすい緑）
            Dim inputBack As Color = Color.FromArgb(226, 240, 217)
            Me.colDenpyo.DefaultCellStyle.BackColor = inputBack
            Me.colKbn.DefaultCellStyle.BackColor = inputBack
            Me.colHoken.DefaultCellStyle.BackColor = inputBack
        End Sub

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
                ' 取得失敗時：原因を通知し、空スキーマ（デザインの列）を表示
                Me._table = Me.expenseDataSet.ExpenseCode
                Me.bsExpense.DataMember = ""
                Me.bsExpense.DataSource = Me._table
                MessageBox.Show("一覧の取得に失敗しました。" & vbCrLf & ex.Message, "読み込み",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub
#End Region

    End Class
End Namespace
