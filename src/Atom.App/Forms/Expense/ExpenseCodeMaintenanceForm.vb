\xEF\xBB\xBFImports System.Collections.Generic
Imports System.Data
Imports System.Windows.Forms
Imports Atom.App.Common
Imports Atom.Data.Repositories

Namespace Forms.Expense
    ''' <summary>経費コード登録（現行 F_M_経費コード登録_MAIN + _SUB を1フォームに統合）。</summary>
    Public Class ExpenseCodeMaintenanceForm
        Inherits BaseForm

#Region "定数"
        ' 経費区分の選択肢（TODO: View__keihi_kbn から取得）
        Private Shared ReadOnly KbnOptions As String() = {"運賃", "諸掛", "保険", "関税", "その他"}
#End Region

#Region "フィールド"
        Private ReadOnly _repo As New ExpenseCodeRepository()
        Private ReadOnly bsExpense As New BindingSource()
        Private ReadOnly dgvList As New DataGridView()
        Private _table As DataTable
#End Region

#Region "コンストラクター"
        ''' <summary>画面を初期化する（タイトルは menuNo からリソース解決）。</summary>
        Public Sub New()
            Me.MenuNo = 907
            Me.BuildUi()
        End Sub
#End Region

#Region "イベントハンドラー"
        ''' <summary>読み込み時：権限チェック（基底）後に一覧を取得・表示する。</summary>
        Protected Overrides Sub OnLoad(e As EventArgs)
            MyBase.OnLoad(e)
            Me.LoadData()
        End Sub

        ''' <summary>編集時に更新日時・更新者を自動セットする（現行 AfterUpdate 相当）。</summary>
        Private Sub OnCellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
            ' 見出し行・無効セルは対象外
            If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then
                Return
            End If
            ' 更新日時・更新者の列自体が変わったときは何もしない
            Dim propName As String = Me.dgvList.Columns(e.ColumnIndex).DataPropertyName
            If propName = "update_ymd" OrElse propName = "update_login" Then
                Return
            End If
            ' 新規入力中の行は対象外
            Dim row As DataGridViewRow = Me.dgvList.Rows(e.RowIndex)
            If row.IsNewRow Then
                Return
            End If
            ' 更新日時・更新者をセットする
            row.Cells("colYmd").Value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
            row.Cells("colUsr").Value = SessionContext.Current.LoginCode
        End Sub

        ''' <summary>行削除ボタン：選択行を削除する。</summary>
        Private Sub OnDeleteClick(sender As Object, e As EventArgs)
            ' 未選択・新規行のときは何もしない
            If Me.dgvList.CurrentRow Is Nothing OrElse Me.dgvList.CurrentRow.IsNewRow Then
                Return
            End If
            ' 確認のうえ削除する
            If MessageBox.Show("選択中の行を削除します。よろしいですか？", "確認",
                               MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Me.dgvList.Rows.Remove(Me.dgvList.CurrentRow)
            End If
        End Sub

        ''' <summary>保存ボタン：未入力・重複を確認して保存する。</summary>
        Private Sub OnSaveClick(sender As Object, e As EventArgs)
            Me.bsExpense.EndEdit()
            Dim seen As New HashSet(Of String)()
            ' 全行の経費コードを検査する
            For Each row As DataRow In Me._table.Rows
                If row.RowState = DataRowState.Deleted Then
                    Continue For
                End If
                Dim code As String = Convert.ToString(row("pk_keihi_code"))
                ' 未入力のとき：中止
                If String.IsNullOrWhiteSpace(code) Then
                    MessageBox.Show("経費コードが未入力の行があります。", "入力エラー",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
                ' 重複のとき：中止
                If Not seen.Add(code) Then
                    MessageBox.Show("経費コードが重複しています: " & code, "入力エラー",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
            Next
            ' 保存する（永続化は未実装）
            Try
                ' TODO: _repo.SaveExpenseCode(_table) を実装
                MessageBox.Show("保存しました。（永続化処理は未実装）", "保存",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("保存に失敗しました。" & vbCrLf & ex.Message, "保存",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
#End Region

#Region "内部メソッド"
        ''' <summary>画面UI（操作バー＋グリッド）を組み立てる。</summary>
        Private Sub BuildUi()
            ' 操作バー（新規行追加 / 行削除 / 保存）
            Dim bar As New FlowLayoutPanel() With {
                .Dock = DockStyle.Top, .Height = 40, .Padding = New Padding(6, 6, 6, 4),
                .FlowDirection = FlowDirection.LeftToRight}
            Dim btnAdd As New Button() With {.Text = "新規行追加", .Width = 100, .Height = 26, .Margin = New Padding(0, 0, 6, 0)}
            Dim btnDel As New Button() With {.Text = "行削除", .Width = 90, .Height = 26, .Margin = New Padding(0, 0, 6, 0)}
            Dim btnSave As New Button() With {.Text = "保存", .Width = 90, .Height = 26}
            AddHandler btnAdd.Click, Sub(s As Object, ev As EventArgs) Me.bsExpense.AddNew()
            AddHandler btnDel.Click, AddressOf Me.OnDeleteClick
            AddHandler btnSave.Click, AddressOf Me.OnSaveClick
            bar.Controls.Add(btnAdd)
            bar.Controls.Add(btnDel)
            bar.Controls.Add(btnSave)

            ' 一覧グリッド
            Me.dgvList.Dock = DockStyle.Fill
            Me.dgvList.AutoGenerateColumns = False
            Me.dgvList.AllowUserToAddRows = True
            Me.dgvList.AllowUserToDeleteRows = True
            Me.dgvList.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            Me.dgvList.RowHeadersWidth = 24
            Me.BuildColumns()
            AddHandler Me.dgvList.CellValueChanged, AddressOf Me.OnCellValueChanged
            Me.dgvList.DataSource = Me.bsExpense

            ' Fill を先、Top を後に追加する
            Me.ContentPanel.Controls.Add(Me.dgvList)
            Me.ContentPanel.Controls.Add(bar)
        End Sub

        ''' <summary>グリッドの列を定義する。</summary>
        Private Sub BuildColumns()
            Me.dgvList.Columns.Add(Me.TextColumn("colCode", "経費コード", "pk_keihi_code", 120, False))
            Me.dgvList.Columns.Add(Me.TextColumn("colText", "経費コード名称", "keihi_code_text", 220, False))

            ' 経費区分（コンボ）
            Dim kbn As New DataGridViewComboBoxColumn() With {
                .Name = "colKbn", .HeaderText = "経費区分", .DataPropertyName = "keihi_kbn", .Width = 110}
            kbn.Items.AddRange(DirectCast(KbnOptions, Object()))
            kbn.FlatStyle = FlatStyle.Flat
            Me.dgvList.Columns.Add(kbn)

            Me.dgvList.Columns.Add(Me.TextColumn("colRef", "参照コード", "keihi_code_ref", 100, False))
            Me.dgvList.Columns.Add(Me.TextColumn("colYmd", "更新日時", "update_ymd", 150, True))
            Me.dgvList.Columns.Add(Me.TextColumn("colUsr", "更新者", "update_login", 90, True))
        End Sub

        ''' <summary>テキスト列を作る。</summary>
        Private Function TextColumn(name As String, header As String, prop As String, width As Integer, readOnlyColumn As Boolean) As DataGridViewTextBoxColumn
            Return New DataGridViewTextBoxColumn() With {
                .Name = name, .HeaderText = header, .DataPropertyName = prop,
                .Width = width, .[ReadOnly] = readOnlyColumn}
        End Function

        ''' <summary>一覧を取得して表示する（DB未接続時は空一覧）。</summary>
        Private Sub LoadData()
            Try
                ' 経費コード一覧の取得・表示
                Me._table = Me._repo.SelectAll()
            Catch ex As Exception
                ' DB未接続などのとき：空スキーマで表示する
                Me._table = Me.BuildEmptyTable()
            End Try
            Me.bsExpense.DataSource = Me._table
        End Sub

        ''' <summary>DB未接続時に使う空の一覧スキーマを作る。</summary>
        Private Function BuildEmptyTable() As DataTable
            Dim table As New DataTable()
            table.Columns.Add("pk_keihi_code", GetType(String))
            table.Columns.Add("keihi_code_text", GetType(String))
            table.Columns.Add("keihi_kbn", GetType(String))
            table.Columns.Add("keihi_code_ref", GetType(String))
            table.Columns.Add("update_ymd", GetType(String))
            table.Columns.Add("update_login", GetType(String))
            Return table
        End Function
#End Region

    End Class
End Namespace
