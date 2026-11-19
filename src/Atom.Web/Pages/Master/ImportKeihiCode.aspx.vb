Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports Atom.Core.Entities
Imports Atom.Core.Logging
Imports Atom.Data.Services
Imports Atom.Web.Common

Namespace Pages.Master

    ''' <summary>
    ''' 輸入経費コード登録画面（menuNo=907）です。
    ''' 旧 Access 版の F_M_経費コード登録_MAIN ＋ _SUB に相当します。
    ''' </summary>
    ''' <remarks>
    ''' 仕様は docs/specs/907_輸入経費コード登録.md を正とします。
    '''
    ''' **旧と同じ挙動にする方針です。**旧は入場制限も更新制限も無く、
    ''' すべての利用者が追加・更新・削除できました（設計書 C-28 / 既知の不具合 K-3）。
    ''' 入力チェックも持たず、DB の制約だけで弾いていました。
    '''
    ''' 旧は行を移動した時点で自動保存でしたが、Web に同じ仕組みが無いため
    ''' 「編集」→「更新」の形にしています（設計書 C-15）。
    ''' </remarks>
    Public Class ImportKeihiCode
        Inherits BasePage

#Region "定数"
        ''' <summary>この画面のメニュー番号。</summary>
        Private Const ThisMenuNo As Integer = 907

        ''' <summary>ログの発生元名。</summary>
        Private Const LogSource As String = "ImportKeihiCode"

        ''' <summary>更新日の表示書式。</summary>
        Private Const UpdateYmdFormat As String = "yyyy/MM/dd HH:mm"

        ''' <summary>区分が未選択のときの表示。</summary>
        Private Const NoSelectionText As String = "（未選択）"

        ''' <summary>区分が未選択のときの値。</summary>
        Private Const NoSelectionValue As String = ""

        ''' <summary>登録できたときの文言。</summary>
        Private Const AddedMessage As String = "追加しました。"

        ''' <summary>更新できたときの文言。</summary>
        Private Const UpdatedMessage As String = "更新しました。"

        ''' <summary>削除できたときの文言。</summary>
        Private Const DeletedMessage As String = "削除しました。"

        ''' <summary>取得に失敗したときの文言。</summary>
        Private Const LoadFailedMessage As String = "データの取得に失敗しました。"

        ''' <summary>保存に失敗したときの文言。</summary>
        Private Const SaveFailedMessage As String = "保存に失敗しました。管理者へ連絡してください。"
#End Region

#Region "フィールド"
        ''' <summary>業務処理。</summary>
        Private ReadOnly _service As New ImportKeihiCodeService()

        ''' <summary>
        ''' 区分の選択肢。
        ''' </summary>
        ''' <remarks>
        ''' 1 回の要求の中で、編集行のプルダウンと新規登録のプルダウンの両方から使います。
        ''' 毎回読み直さないよう、最初に読んだものを保持します。
        ''' </remarks>
        Private _kbnList As IList(Of ImportKeihiKbn)
#End Region

#Region "ページのイベント"
        ''' <summary>
        ''' 画面の初期化です。
        ''' </summary>
        ''' <param name="e">イベント引数。</param>
        ''' <remarks>
        ''' **入場制限は設けません。**項目マスタの `contents_num` が 0 で、
        ''' 旧でも管理者限定ではありませんでした（設計書 C-27）。
        ''' </remarks>
        Protected Overrides Sub OnInit(e As EventArgs)
            Me.MenuNo = ThisMenuNo
            Me.HeaderActions.Add(New HeaderAction("MENUへ", HeaderAction.MenuScreenNo))
            MyBase.OnInit(e)
        End Sub

        ''' <summary>
        ''' 画面を開いたときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
            ' 画面を開いたときだけ読み込みます
            If Not Me.IsPostBack Then
                Me.LoadKbnOptions(Me.ddlNewKeihiKbn)
                Me.LoadList()
            End If
        End Sub
#End Region

#Region "一覧のイベント"
        ''' <summary>
        ''' 行を組み立てるときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        ''' <remarks>
        ''' **`RowDataBound` ではなく `RowCreated` で値を入れます。**
        ''' ポストバックのときは `RowDataBound` が起きないため、
        ''' テンプレート内のコントロールに値が入らなくなります。
        ''' </remarks>
        Protected Sub GvKeihiCode_RowCreated(sender As Object, e As GridViewRowEventArgs) _
            Handles gvKeihiCode.RowCreated

            ' データの行ではないとき
            If e.Row.RowType <> DataControlRowType.DataRow Then
                Return
            End If

            ' 編集中の行のとき（プルダウンの選択肢はここで入れておく必要があります）
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Dim ddl As DropDownList = TryCast(e.Row.FindControl("ddlKeihiKbn"), DropDownList)
                If ddl IsNot Nothing Then
                    Me.LoadKbnOptions(ddl)
                End If
            End If
        End Sub

        ''' <summary>
        ''' 行に値を流し込むときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Protected Sub GvKeihiCode_RowDataBound(sender As Object, e As GridViewRowEventArgs) _
            Handles gvKeihiCode.RowDataBound

            ' データの行ではないとき
            If e.Row.RowType <> DataControlRowType.DataRow Then
                Return
            End If

            Dim entity As Atom.Core.Entities.ImportKeihiCode =
                TryCast(e.Row.DataItem, Atom.Core.Entities.ImportKeihiCode)
            ' 取り出せないとき
            If entity Is Nothing Then
                Return
            End If

            ' 表示した時点の更新日を控えます。楽観ロックの比較に使います
            Me.ViewState(ViewStateKey(entity.KeihiCode)) = entity.UpdateYmd

            ' 編集中の行のとき
            If (e.Row.RowState And DataControlRowState.Edit) = DataControlRowState.Edit Then
                Me.SetEditRow(e.Row, entity)
                Return
            End If

            Me.SetViewRow(e.Row, entity)
        End Sub

        ''' <summary>
        ''' 「編集」を押したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Protected Sub GvKeihiCode_RowEditing(sender As Object, e As GridViewEditEventArgs) _
            Handles gvKeihiCode.RowEditing

            Me.gvKeihiCode.EditIndex = e.NewEditIndex
            Me.LoadList()
        End Sub

        ''' <summary>
        ''' 「取消」を押したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Protected Sub GvKeihiCode_RowCancelingEdit(sender As Object, e As GridViewCancelEditEventArgs) _
            Handles gvKeihiCode.RowCancelingEdit

            Me.gvKeihiCode.EditIndex = -1
            Me.LoadList()
        End Sub

        ''' <summary>
        ''' 「更新」を押したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Protected Sub GvKeihiCode_RowUpdating(sender As Object, e As GridViewUpdateEventArgs) _
            Handles gvKeihiCode.RowUpdating

            e.Cancel = True

            Dim row As GridViewRow = Me.gvKeihiCode.Rows(e.RowIndex)
            Dim originalKeihiCode As String =
                Convert.ToString(Me.gvKeihiCode.DataKeys(e.RowIndex).Value, CultureInfo.InvariantCulture)

            Dim entity As New Atom.Core.Entities.ImportKeihiCode()
            entity.KeihiCode = GetTextBoxValue(row, "txtKeihiCode")
            entity.KeihiCodeText = GetTextBoxValue(row, "txtKeihiCodeText")
            entity.KeihiCodeRef = GetTextBoxValue(row, "txtKeihiCodeRef")
            entity.KeihiKbn = GetDropDownValue(row, "ddlKeihiKbn")

            Dim flg As Integer
            ' 保険対象が数値でないとき
            If Not TryParseFlg(GetTextBoxValue(row, "txtHokenTaisyoFlg"), flg) Then
                Me.ShowMessage(ImportKeihiCodeService.HokenTaisyoFlgInvalidMessage, True)
                Return
            End If
            entity.HokenTaisyoFlg = flg

            Dim loginCode As String = SessionContext.Current.LoginCode

            Try
                Dim message As String = Me._service.Update(entity, originalKeihiCode,
                                                           Me.GetOriginalUpdateYmd(originalKeihiCode), loginCode)
                ' 保存できなかったとき
                If Not String.IsNullOrEmpty(message) Then
                    Me.ShowMessage(message, True)
                    Me.LoadList()
                    Return
                End If

                Me.gvKeihiCode.EditIndex = -1
                Me.LoadList()
                Me.ShowMessage(UpdatedMessage, False)
            Catch ex As Exception
                Logger.WriteError(LogSource, "経費コードの更新に失敗しました。", ex, loginCode)
                Me.ShowMessage(SaveFailedMessage, True)
            End Try
        End Sub

        ''' <summary>
        ''' 「削除」を押したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        ''' <remarks>
        ''' **物理削除です。**明細で使用中かどうかは見ません（設計書 C-7 / 既知の不具合 K-4）。
        ''' </remarks>
        Protected Sub GvKeihiCode_RowDeleting(sender As Object, e As GridViewDeleteEventArgs) _
            Handles gvKeihiCode.RowDeleting

            e.Cancel = True

            Dim keihiCode As String =
                Convert.ToString(Me.gvKeihiCode.DataKeys(e.RowIndex).Value, CultureInfo.InvariantCulture)
            Dim loginCode As String = SessionContext.Current.LoginCode

            Try
                Dim message As String = Me._service.Delete(keihiCode,
                                                           Me.GetOriginalUpdateYmd(keihiCode), loginCode)
                ' 削除できなかったとき
                If Not String.IsNullOrEmpty(message) Then
                    Me.ShowMessage(message, True)
                    Me.LoadList()
                    Return
                End If

                Me.gvKeihiCode.EditIndex = -1
                Me.LoadList()
                Me.ShowMessage(DeletedMessage, False)
            Catch ex As Exception
                Logger.WriteError(LogSource, "経費コードの削除に失敗しました。", ex, loginCode)
                Me.ShowMessage(SaveFailedMessage, True)
            End Try
        End Sub
#End Region

#Region "ボタンのイベント"
        ''' <summary>
        ''' 「追加」を押したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Protected Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
            Dim entity As New Atom.Core.Entities.ImportKeihiCode()
            entity.KeihiCode = Me.txtNewKeihiCode.Text.Trim()
            entity.KeihiCodeText = Me.txtNewKeihiCodeText.Text.Trim()
            entity.KeihiCodeRef = Me.txtNewKeihiCodeRef.Text.Trim()
            entity.KeihiKbn = Me.ddlNewKeihiKbn.SelectedValue

            Dim flg As Integer
            ' 保険対象が数値でないとき
            If Not TryParseFlg(Me.txtNewHokenTaisyoFlg.Text, flg) Then
                Me.ShowMessage(ImportKeihiCodeService.HokenTaisyoFlgInvalidMessage, True)
                Return
            End If
            entity.HokenTaisyoFlg = flg

            Dim loginCode As String = SessionContext.Current.LoginCode

            Try
                Dim message As String = Me._service.Insert(entity, loginCode)
                ' 登録できなかったとき
                If Not String.IsNullOrEmpty(message) Then
                    Me.ShowMessage(message, True)
                    Return
                End If

                Me.ClearNewRow()
                Me.gvKeihiCode.EditIndex = -1
                Me.LoadList()
                Me.ShowMessage(AddedMessage, False)
            Catch ex As Exception
                Logger.WriteError(LogSource, "経費コードの登録に失敗しました。", ex, loginCode)
                Me.ShowMessage(SaveFailedMessage, True)
            End Try
        End Sub
#End Region

#Region "内部処理（表示）"
        ''' <summary>
        ''' 一覧を読み込んで表示します。
        ''' </summary>
        Private Sub LoadList()
            Try
                Me.gvKeihiCode.DataSource = Me._service.SelectAll()
                Me.gvKeihiCode.DataBind()
            Catch ex As Exception
                Logger.WriteError(LogSource, "経費コードの取得に失敗しました。", ex,
                                  SessionContext.Current.LoginCode)
                Me.ShowMessage(LoadFailedMessage, True)
            End Try
        End Sub

        ''' <summary>
        ''' 区分の選択肢をプルダウンへ設定します。
        ''' </summary>
        ''' <param name="target">設定先のプルダウン。</param>
        ''' <remarks>
        ''' 保存するのは区分の値です。並びはビューの定義と同じです（設計書 §13.1 ②）。
        '''
        ''' 選択肢には**値と名称の両方**を並べます。値だけでは何を選んでいるか分からず、
        ''' 名称だけでは明細側に入る値が見えないためです。
        ''' 一覧の表示列は設計書どおり名称だけです。
        ''' </remarks>
        Private Sub LoadKbnOptions(target As DropDownList)
            target.Items.Clear()
            target.Items.Add(New ListItem(NoSelectionText, NoSelectionValue))

            For Each kbn As ImportKeihiKbn In Me.GetKbnList()
                target.Items.Add(New ListItem(ToKbnItemText(kbn), kbn.Code))
            Next
        End Sub

        ''' <summary>
        ''' 区分の選択肢に出す文言を作ります。
        ''' </summary>
        ''' <param name="kbn">区分。</param>
        ''' <returns>「値　名称」の形の文言。名称が無いときは値だけ。</returns>
        Private Shared Function ToKbnItemText(kbn As ImportKeihiKbn) As String
            ' 名称が無いとき
            If String.IsNullOrEmpty(kbn.Name) Then
                Return kbn.Code
            End If
            Return kbn.Code & "　" & kbn.Name
        End Function

        ''' <summary>
        ''' 区分の選択肢を取得します。1 回の要求の中では読み直しません。
        ''' </summary>
        ''' <returns>選択肢の一覧。</returns>
        Private Function GetKbnList() As IList(Of ImportKeihiKbn)
            ' まだ読んでいないとき
            If Me._kbnList Is Nothing Then
                Try
                    Me._kbnList = Me._service.SelectKbnList()
                Catch ex As Exception
                    Logger.WriteError(LogSource, "経費区分の取得に失敗しました。", ex,
                                      SessionContext.Current.LoginCode)
                    Me._kbnList = New List(Of ImportKeihiKbn)()
                    Me.ShowMessage(LoadFailedMessage, True)
                End Try
            End If
            Return Me._kbnList
        End Function

        ''' <summary>
        ''' 表示中の行へ値を設定します。
        ''' </summary>
        ''' <param name="row">対象の行。</param>
        ''' <param name="entity">表示する内容。</param>
        Private Sub SetViewRow(row As GridViewRow, entity As Atom.Core.Entities.ImportKeihiCode)
            SetLiteral(row, "litKeihiCode", entity.KeihiCode)
            SetLiteral(row, "litKeihiCodeText", entity.KeihiCodeText)
            SetLiteral(row, "litKeihiCodeRef", entity.KeihiCodeRef)
            SetLiteral(row, "litKeihiKbn", Me.ResolveKbnName(entity.KeihiKbn))
            SetLiteral(row, "litHokenTaisyoFlg",
                       entity.HokenTaisyoFlg.ToString(CultureInfo.InvariantCulture))
            SetLiteral(row, "litUpdateYmd", ToUpdateYmdText(entity.UpdateYmd))
        End Sub

        ''' <summary>
        ''' 編集中の行へ値を設定します。
        ''' </summary>
        ''' <param name="row">対象の行。</param>
        ''' <param name="entity">表示する内容。</param>
        Private Sub SetEditRow(row As GridViewRow, entity As Atom.Core.Entities.ImportKeihiCode)
            SetTextBox(row, "txtKeihiCode", entity.KeihiCode)
            SetTextBox(row, "txtKeihiCodeText", entity.KeihiCodeText)
            SetTextBox(row, "txtKeihiCodeRef", entity.KeihiCodeRef)
            SetTextBox(row, "txtHokenTaisyoFlg",
                       entity.HokenTaisyoFlg.ToString(CultureInfo.InvariantCulture))
            SetLiteral(row, "litUpdateYmdEdit", ToUpdateYmdText(entity.UpdateYmd))

            Dim ddl As DropDownList = TryCast(row.FindControl("ddlKeihiKbn"), DropDownList)
            ' プルダウンが無いとき
            If ddl Is Nothing Then
                Return
            End If

            Dim item As ListItem = ddl.Items.FindByValue(If(entity.KeihiKbn, String.Empty))
            ' 選択肢に無い値のとき（マスタから消えた区分など）
            If item Is Nothing Then
                ddl.SelectedValue = NoSelectionValue
                Return
            End If
            ddl.SelectedValue = item.Value
        End Sub

        ''' <summary>
        ''' 区分の値から表示名を求めます。
        ''' </summary>
        ''' <param name="code">区分の値。</param>
        ''' <returns>表示名。選択肢に無いときは値をそのまま返します。</returns>
        Private Function ResolveKbnName(code As String) As String
            ' 未設定のとき
            If String.IsNullOrEmpty(code) Then
                Return String.Empty
            End If

            For Each kbn As ImportKeihiKbn In Me.GetKbnList()
                ' 一致したとき
                If String.Equals(kbn.Code, code, StringComparison.Ordinal) Then
                    Return kbn.Name
                End If
            Next

            ' 選択肢に無い値は、消さずにそのまま見せます（データの異常に気付けるようにするため）
            Return code
        End Function

        ''' <summary>
        ''' 新規登録の入力欄を空にします。
        ''' </summary>
        Private Sub ClearNewRow()
            Me.txtNewKeihiCode.Text = String.Empty
            Me.txtNewKeihiCodeText.Text = String.Empty
            Me.txtNewKeihiCodeRef.Text = String.Empty
            Me.txtNewHokenTaisyoFlg.Text = String.Empty
            Me.ddlNewKeihiKbn.SelectedValue = NoSelectionValue
        End Sub
#End Region

#Region "内部処理（値の取り出し）"
        ''' <summary>
        ''' 画面に表示した時点の更新日を取得します。
        ''' </summary>
        ''' <param name="keihiCode">伝票コード。</param>
        ''' <returns>更新日。取れないときは Nothing。</returns>
        ''' <remarks>
        ''' 楽観ロックの比較に使います。**表示した時点の値**でなければ意味がないため、
        ''' 一覧を読み込んだときにビューステートへ入れた値を使います。
        ''' </remarks>
        Private Function GetOriginalUpdateYmd(keihiCode As String) As Date?
            Dim stored As Object = Me.ViewState(ViewStateKey(keihiCode))
            ' 保持していないとき
            If stored Is Nothing Then
                Return Nothing
            End If
            Return CType(stored, Date?)
        End Function

        ''' <summary>
        ''' 更新日を保持するためのビューステートのキーを作ります。
        ''' </summary>
        ''' <param name="keihiCode">伝票コード。</param>
        ''' <returns>キー。</returns>
        Private Shared Function ViewStateKey(keihiCode As String) As String
            Return "UpdateYmd_" & keihiCode
        End Function

        ''' <summary>
        ''' テキストボックスの値を取り出します。
        ''' </summary>
        ''' <param name="row">対象の行。</param>
        ''' <param name="id">コントロールの ID。</param>
        ''' <returns>前後の空白を落とした値。見つからないときは空文字。</returns>
        Private Shared Function GetTextBoxValue(row As GridViewRow, id As String) As String
            Dim textBox As TextBox = TryCast(row.FindControl(id), TextBox)
            ' 見つからないとき
            If textBox Is Nothing Then
                Return String.Empty
            End If
            Return textBox.Text.Trim()
        End Function

        ''' <summary>
        ''' プルダウンの選択値を取り出します。
        ''' </summary>
        ''' <param name="row">対象の行。</param>
        ''' <param name="id">コントロールの ID。</param>
        ''' <returns>選択値。見つからないときは空文字。</returns>
        Private Shared Function GetDropDownValue(row As GridViewRow, id As String) As String
            Dim ddl As DropDownList = TryCast(row.FindControl(id), DropDownList)
            ' 見つからないとき
            If ddl Is Nothing Then
                Return String.Empty
            End If
            Return ddl.SelectedValue
        End Function

        ''' <summary>
        ''' 保険対象の入力を数値にします。
        ''' </summary>
        ''' <param name="value">入力された文字列。</param>
        ''' <param name="result">変換した値。</param>
        ''' <returns>変換できたとき True。</returns>
        ''' <remarks>
        ''' 未入力は 0 とします。旧はテキストボックスで 0 / 1 以外も入れられたため、
        ''' ここでも範囲は見ません（設計書 C-10）。
        ''' </remarks>
        Private Shared Function TryParseFlg(value As String, ByRef result As Integer) As Boolean
            ' 未入力のとき
            If String.IsNullOrWhiteSpace(value) Then
                result = 0
                Return True
            End If
            Return Integer.TryParse(value.Trim(), NumberStyles.Integer,
                                    CultureInfo.InvariantCulture, result)
        End Function

        ''' <summary>
        ''' リテラルへ値を設定します。
        ''' </summary>
        ''' <param name="row">対象の行。</param>
        ''' <param name="id">コントロールの ID。</param>
        ''' <param name="value">設定する値。</param>
        Private Shared Sub SetLiteral(row As GridViewRow, id As String, value As String)
            Dim literal As Literal = TryCast(row.FindControl(id), Literal)
            ' 見つからないとき
            If literal Is Nothing Then
                Return
            End If
            literal.Text = If(value, String.Empty)
        End Sub

        ''' <summary>
        ''' テキストボックスへ値を設定します。
        ''' </summary>
        ''' <param name="row">対象の行。</param>
        ''' <param name="id">コントロールの ID。</param>
        ''' <param name="value">設定する値。</param>
        Private Shared Sub SetTextBox(row As GridViewRow, id As String, value As String)
            Dim textBox As TextBox = TryCast(row.FindControl(id), TextBox)
            ' 見つからないとき
            If textBox Is Nothing Then
                Return
            End If
            textBox.Text = If(value, String.Empty)
        End Sub

        ''' <summary>
        ''' 更新日を画面に出す書式にします。
        ''' </summary>
        ''' <param name="value">更新日。</param>
        ''' <returns>文字列。値が無いときは空文字。</returns>
        Private Shared Function ToUpdateYmdText(value As Date?) As String
            ' 値が無いとき（既存行の多くが NULL）
            If Not value.HasValue Then
                Return String.Empty
            End If
            Return value.Value.ToString(UpdateYmdFormat, CultureInfo.InvariantCulture)
        End Function
#End Region

    End Class
End Namespace
