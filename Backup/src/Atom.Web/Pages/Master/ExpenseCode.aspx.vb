Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports Atom.Core.Entities
Imports Atom.Core.Exceptions
Imports Atom.Core.Logging
Imports Atom.Data.Repositories
Imports Atom.Web.Common

''' <summary>
''' 経費コード登録画面（menuNo=907）です。
''' 旧 ATOM の F_M_経費コード登録_MAIN / _SUB に対応します。
''' </summary>
''' <remarks>
''' 旧画面の挙動との対応
''' ・btn_close_Click     … ヘッダーの「閉じる」ボタン。呼び出し元が指定されていればそこへ、
''' 　　　　　　　　　　　　 指定が無ければ輸入諸経費入力（menuNo=301）へ戻ります。
''' ・btn_setsuzoku_Click … フッターの「再接続」ボタン（Site.Master 側で実装）。
''' ・各項目の AfterUpdate … 保存時に更新日時と更新者を書き込みます。
''' ・Form_Open の Maximize … Web では画面幅いっぱいに表示されるため対応不要です。
''' ・SUB_Enter の GoToRecord acFirst … Web にはレコード位置の概念が無いため対応不要です。
''' </remarks>
Public Class ExpenseCodeMaintenance
    Inherits BasePage

#Region "定数"
    ''' <summary>この画面のメニュー番号。</summary>
    Private Const ScreenMenuNo As Integer = 907

    ''' <summary>ヘッダーの背景色。</summary>
    Private Const HeaderBackColor As String = "#FF8040"

    ''' <summary>呼び出し元を受け取るクエリ文字列の名前。旧画面の return_form_name に対応します。</summary>
    Private Const ReturnUrlKey As String = "return"

    ''' <summary>呼び出し元の指定が無いときに戻る画面のメニュー番号（輸入諸経費入力）。</summary>
    Private Const DefaultReturnMenuNo As Integer = 301

    ''' <summary>ログの発生元名。</summary>
    Private Const LogSource As String = "ExpenseCodeMaintenance"
#End Region

#Region "コンストラクタ"
    ''' <summary>画面を初期化します。</summary>
    Public Sub New()
        Me.MenuNo = ScreenMenuNo
    End Sub
#End Region

#Region "公開メソッド"
    ''' <summary>
    ''' この画面のヘッダー設定を返します。
    ''' </summary>
    ''' <returns>ヘッダー設定。</returns>
    Public Overrides Function CreateHeaderConfig() As HeaderConfig
        Dim config As New HeaderConfig()
        config.Title = "経費コード登録"
        config.BackColor = HeaderBackColor
        config.AddCloseButton(Me.GetReturnUrl())
        Return config
    End Function
#End Region

#Region "イベントハンドラ"
    ''' <summary>
    ''' 読み込み時の処理です。初回のみ一覧を表示します。
    ''' </summary>
    ''' <param name="sender">イベントの発生元。</param>
    ''' <param name="e">イベント引数。</param>
    Private Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' 初回表示のとき
        If Not Me.IsPostBack Then
            Me.BindGrid()
        End If
    End Sub

    ''' <summary>
    ''' 行のデータ割り当て時に、各入力欄へ値を設定します。
    ''' </summary>
    ''' <param name="sender">イベントの発生元。</param>
    ''' <param name="e">イベント引数。</param>
    Private Sub GvExpenseCode_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gvExpenseCode.RowDataBound
        ' 明細行でないとき
        If e.Row.RowType <> DataControlRowType.DataRow Then
            Return
        End If

        Dim entity As ExpenseCode = TryCast(e.Row.DataItem, ExpenseCode)
        If entity Is Nothing Then
            Return
        End If

        Me.FindTextBox(e.Row, "txtKeihiCode").Text = entity.KeihiCode
        Me.FindTextBox(e.Row, "txtKeihiCodeText").Text = entity.KeihiCodeText
        Me.FindTextBox(e.Row, "txtKeihiKbn").Text = entity.KeihiKbn

        Me.FindHiddenField(e.Row, "hdnOriginalKeihiCode").Value = entity.KeihiCode
        Me.FindHiddenField(e.Row, "hdnUpdateYmd").Value = Me.ToHiddenValue(entity.UpdateYmd)

        Me.FindLiteral(e.Row, "litUpdateYmd").Text = Me.ToDisplayValue(entity.UpdateYmd)
        Me.FindLiteral(e.Row, "litUpdateLogin").Text = entity.UpdateLogin
    End Sub

    ''' <summary>
    ''' 一覧内のボタン押下時の処理です。行の追加と削除を受け付けます。
    ''' </summary>
    ''' <param name="sender">イベントの発生元。</param>
    ''' <param name="e">イベント引数。</param>
    Private Sub GvExpenseCode_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles gvExpenseCode.RowCommand
        Select Case e.CommandName
            ' 追加のとき
            Case "AddRow"
                Me.AddNewRow()

            ' 削除のとき
            Case "DeleteRow"
                Dim rowIndex As Integer = 0
                ' 行番号として解釈できたとき
                If Integer.TryParse(Convert.ToString(e.CommandArgument), rowIndex) Then
                    Me.DeleteRow(rowIndex)
                Else
                    Me.ShowMessage("削除対象の行を特定できませんでした。", True)
                End If
        End Select
    End Sub

    ''' <summary>
    ''' 保存ボタン押下時の処理です。変更された行だけを更新します。
    ''' </summary>
    ''' <param name="sender">イベントの発生元。</param>
    ''' <param name="e">イベント引数。</param>
    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Me.SaveChanges()
    End Sub

    ''' <summary>
    ''' 再表示ボタン押下時の処理です。入力内容を破棄して読み直します。
    ''' </summary>
    ''' <param name="sender">イベントの発生元。</param>
    ''' <param name="e">イベント引数。</param>
    Private Sub BtnReload_Click(sender As Object, e As EventArgs) Handles btnReload.Click
        Me.BindGrid()
        Me.ShowMessage("最新の内容を読み込みました。", False)
    End Sub
#End Region

#Region "内部処理（表示）"
    ''' <summary>
    ''' 一覧を読み込んで表示します。
    ''' </summary>
    Private Sub BindGrid()
        Me.gvExpenseCode.DataSource = Me.CreateRepository().SelectAll()
        Me.gvExpenseCode.DataBind()
    End Sub

    ''' <summary>
    ''' 「閉じる」ボタンの遷移先を求めます。
    ''' 呼び出し元の指定があればそこへ、無ければ輸入諸経費入力へ戻します。
    ''' </summary>
    ''' <returns>遷移先の URL。</returns>
    Private Function GetReturnUrl() As String
        Dim requested As String = Me.Request.QueryString(ReturnUrlKey)

        ' 呼び出し元が指定されているとき（改ざん対策として登録済み URL だけを許可する）
        If Not String.IsNullOrEmpty(requested) Then
            Dim menuNo As Integer = 0
            If Integer.TryParse(requested, menuNo) Then
                Dim url As String = PageUrlMap.GetUrl(menuNo)
                If Not String.IsNullOrEmpty(url) Then
                    Return url
                End If
            End If
        End If

        ' 呼び出し元の指定が無いとき
        Dim defaultUrl As String = PageUrlMap.GetUrl(DefaultReturnMenuNo)
        If Not String.IsNullOrEmpty(defaultUrl) Then
            Return defaultUrl
        End If

        ' 輸入諸経費入力が未実装のとき
        Return "~/Default.aspx"
    End Function
#End Region

#Region "内部処理（更新）"
    ''' <summary>
    ''' 画面の入力内容のうち、変更された行だけを更新します。
    ''' </summary>
    Private Sub SaveChanges()
        Dim repository As ExpenseCodeRepository = Me.CreateRepository()

        Try
            ' 現在の登録内容を読み込み、変更の有無を判定する材料にする
            Dim stored As Dictionary(Of String, ExpenseCode) = Me.CreateStoredMap(repository.SelectAll())
            Dim updatedCount As Integer = 0

            For Each row As GridViewRow In Me.gvExpenseCode.Rows
                ' 明細行でないとき
                If row.RowType <> DataControlRowType.DataRow Then
                    Continue For
                End If

                Dim originalKeihiCode As String = Me.FindHiddenField(row, "hdnOriginalKeihiCode").Value
                Dim entity As ExpenseCode = Me.CreateEntityFromRow(row)

                ' 経費コードが未入力のとき
                If String.IsNullOrEmpty(entity.KeihiCode) Then
                    Me.ShowMessage("経費コードは必須です。", True)
                    Return
                End If

                Dim original As ExpenseCode = Nothing
                ' 登録が見つからないとき（他の利用者が削除した）
                If Not stored.TryGetValue(originalKeihiCode, original) Then
                    Me.ShowMessage("他の利用者によって削除されています。最新の内容を読み込み直してください。", True)
                    Me.BindGrid()
                    Return
                End If

                ' 変更が無いときは更新しない（不要な更新日時の書き換えを避ける）
                If Not Me.HasChanged(original, entity) Then
                    Continue For
                End If

                repository.Update(entity, originalKeihiCode)
                updatedCount += 1
            Next

            Me.BindGrid()

            ' 更新があったとき
            If updatedCount > 0 Then
                Me.ShowMessage(updatedCount.ToString() & " 件を保存しました。", False)
            Else
                Me.ShowMessage("変更はありませんでした。", False)
            End If

        Catch ex As ConcurrencyException
            Me.BindGrid()
            Me.ShowMessage(ex.Message, True)

        Catch ex As Exception
            Logger.WriteError(LogSource, "経費コードの保存に失敗しました。", ex, Me.GetLoginCode())
            Me.ShowMessage("保存時にエラーが発生しました。", True)
        End Try
    End Sub

    ''' <summary>
    ''' 追加行の入力内容を新規登録します。
    ''' </summary>
    Private Sub AddNewRow()
        Dim footer As GridViewRow = Me.gvExpenseCode.FooterRow
        ' 追加行が無いとき
        If footer Is Nothing Then
            Return
        End If

        Dim entity As New ExpenseCode()
        entity.KeihiCode = Me.FindTextBox(footer, "txtNewKeihiCode").Text.Trim()
        entity.KeihiCodeText = Me.FindTextBox(footer, "txtNewKeihiCodeText").Text.Trim()
        entity.KeihiKbn = Me.FindTextBox(footer, "txtNewKeihiKbn").Text.Trim()

        ' 経費コードが未入力のとき
        If String.IsNullOrEmpty(entity.KeihiCode) Then
            Me.ShowMessage("経費コードを入力してください。", True)
            Return
        End If

        Dim repository As ExpenseCodeRepository = Me.CreateRepository()

        Try
            ' すでに同じ経費コードが登録されているとき
            If repository.SelectByKey(entity.KeihiCode) IsNot Nothing Then
                Me.ShowMessage("経費コード " & entity.KeihiCode & " はすでに登録されています。", True)
                Return
            End If

            repository.Insert(entity)
            Me.BindGrid()
            Me.ShowMessage("経費コード " & entity.KeihiCode & " を追加しました。", False)

        Catch ex As Exception
            Logger.WriteError(LogSource, "経費コードの追加に失敗しました。", ex, Me.GetLoginCode())
            Me.ShowMessage("追加時にエラーが発生しました。", True)
        End Try
    End Sub

    ''' <summary>
    ''' 指定した行を削除します。
    ''' </summary>
    ''' <param name="rowIndex">一覧上の行番号。</param>
    Private Sub DeleteRow(rowIndex As Integer)
        ' 行番号が範囲外のとき
        If rowIndex < 0 OrElse rowIndex >= Me.gvExpenseCode.Rows.Count Then
            Me.ShowMessage("削除対象の行を特定できませんでした。", True)
            Return
        End If

        Dim row As GridViewRow = Me.gvExpenseCode.Rows(rowIndex)
        Dim keihiCode As String = Me.FindHiddenField(row, "hdnOriginalKeihiCode").Value
        Dim updateYmd As Date? = Me.ToNullableDate(Me.FindHiddenField(row, "hdnUpdateYmd").Value)

        Try
            Me.CreateRepository().Delete(keihiCode, updateYmd)
            Me.BindGrid()
            Me.ShowMessage("経費コード " & keihiCode & " を削除しました。", False)

        Catch ex As ConcurrencyException
            Me.BindGrid()
            Me.ShowMessage(ex.Message, True)

        Catch ex As Exception
            Logger.WriteError(LogSource, "経費コードの削除に失敗しました。", ex, Me.GetLoginCode())
            Me.ShowMessage("削除時にエラーが発生しました。", True)
        End Try
    End Sub

    ''' <summary>
    ''' 明細行の入力内容からエンティティを作成します。
    ''' </summary>
    ''' <param name="row">対象の行。</param>
    ''' <returns>作成したエンティティ。</returns>
    Private Function CreateEntityFromRow(row As GridViewRow) As ExpenseCode
        Dim entity As New ExpenseCode()
        entity.KeihiCode = Me.FindTextBox(row, "txtKeihiCode").Text.Trim()
        entity.KeihiCodeText = Me.FindTextBox(row, "txtKeihiCodeText").Text.Trim()
        entity.KeihiKbn = Me.FindTextBox(row, "txtKeihiKbn").Text.Trim()
        entity.UpdateYmd = Me.ToNullableDate(Me.FindHiddenField(row, "hdnUpdateYmd").Value)
        Return entity
    End Function

    ''' <summary>
    ''' 経費コードを鍵にした対応表を作成します。
    ''' </summary>
    ''' <param name="list">対象の一覧。</param>
    ''' <returns>経費コードと内容の対応表。</returns>
    Private Function CreateStoredMap(list As IList(Of ExpenseCode)) As Dictionary(Of String, ExpenseCode)
        Dim map As New Dictionary(Of String, ExpenseCode)()
        For Each entity As ExpenseCode In list
            map(entity.KeihiCode) = entity
        Next
        Return map
    End Function

    ''' <summary>
    ''' 入力内容が登録内容から変更されているかを判定します。
    ''' </summary>
    ''' <param name="original">登録されている内容。</param>
    ''' <param name="input">画面で入力された内容。</param>
    ''' <returns>変更されているとき True。</returns>
    Private Function HasChanged(original As ExpenseCode, input As ExpenseCode) As Boolean
        If original.KeihiCode <> input.KeihiCode Then
            Return True
        End If
        If original.KeihiCodeText <> input.KeihiCodeText Then
            Return True
        End If
        If original.KeihiKbn <> input.KeihiKbn Then
            Return True
        End If
        Return False
    End Function
#End Region

#Region "内部処理（部品の取得と変換）"
    ''' <summary>データアクセスの窓口を作成します。</summary>
    ''' <returns>経費コードのデータアクセス。</returns>
    Private Function CreateRepository() As ExpenseCodeRepository
        Return New ExpenseCodeRepository(Me.GetLoginCode())
    End Function

    ''' <summary>操作している利用者のログインコードを取得します。</summary>
    ''' <returns>ログインコード。</returns>
    Private Function GetLoginCode() As String
        Return SessionContext.Current.LoginCode
    End Function

    ''' <summary>行の中から入力欄を取得します。</summary>
    ''' <param name="row">対象の行。</param>
    ''' <param name="controlId">コントロールの ID。</param>
    ''' <returns>見つかった入力欄。</returns>
    Private Function FindTextBox(row As GridViewRow, controlId As String) As TextBox
        Return DirectCast(row.FindControl(controlId), TextBox)
    End Function

    ''' <summary>行の中から隠し項目を取得します。</summary>
    ''' <param name="row">対象の行。</param>
    ''' <param name="controlId">コントロールの ID。</param>
    ''' <returns>見つかった隠し項目。</returns>
    Private Function FindHiddenField(row As GridViewRow, controlId As String) As HiddenField
        Return DirectCast(row.FindControl(controlId), HiddenField)
    End Function

    ''' <summary>行の中から表示欄を取得します。</summary>
    ''' <param name="row">対象の行。</param>
    ''' <param name="controlId">コントロールの ID。</param>
    ''' <returns>見つかった表示欄。</returns>
    Private Function FindLiteral(row As GridViewRow, controlId As String) As Literal
        Return DirectCast(row.FindControl(controlId), Literal)
    End Function

    ''' <summary>日時を隠し項目へ格納する文字列に変換します。</summary>
    ''' <param name="value">変換する日時。</param>
    ''' <returns>格納する文字列。未設定のときは空文字。</returns>
    Private Function ToHiddenValue(value As Date?) As String
        ' 未設定のとき
        If Not value.HasValue Then
            Return String.Empty
        End If
        ' ミリ秒まで保持して楽観ロックの判定精度を保つ
        Return value.Value.ToString("yyyy/MM/dd HH:mm:ss.fff")
    End Function

    ''' <summary>日時を画面表示用の文字列に変換します。</summary>
    ''' <param name="value">変換する日時。</param>
    ''' <returns>表示する文字列。未設定のときは空文字。</returns>
    Private Function ToDisplayValue(value As Date?) As String
        ' 未設定のとき
        If Not value.HasValue Then
            Return String.Empty
        End If
        Return value.Value.ToString("yyyy/MM/dd HH:mm")
    End Function

    ''' <summary>隠し項目の文字列を日時に戻します。</summary>
    ''' <param name="value">隠し項目の値。</param>
    ''' <returns>変換した日時。変換できないときは Nothing。</returns>
    Private Function ToNullableDate(value As String) As Date?
        Dim parsed As Date = Nothing
        ' 日時として解釈できたとき
        If Date.TryParse(value, parsed) Then
            Return parsed
        End If
        Return Nothing
    End Function
#End Region

End Class
