Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports Atom.Core.Entities
Imports Atom.Data.Services
Imports Atom.Web.Common

Namespace Pages.Entry

    ''' <summary>
    ''' 取込結果確認画面（menuNo=108）です。
    ''' 旧 Access 版の F_1_取込内容確認_MAIN に相当します。
    ''' </summary>
    ''' <remarks>
    ''' 一時テーブル（tbl_t_shanghai_invoice_import_check_temp /
    ''' tbl_t_shanghai_packing_import_check_temp）の内容を、操作している利用者の分だけ表示します。
    ''' ERR_CONTENTS 以外は編集でき、入力欄から離れた時点で一時データへ反映します。
    ''' チェック処理はやり直さないため、直しても ERR_CONTENTS は変わりません（旧 Access 版と同じ挙動）。
    ''' 品目コードを直すと R3 品目コードも Service 側で併せて更新されます。
    ''' </remarks>
    Public Class ImportResult
        Inherits BasePage

#Region "定数"
        ''' <summary>この画面のメニュー番号。</summary>
        Private Const ThisMenuNo As Integer = 108

        ''' <summary>
        ''' 「閉じる」で実行するスクリプトです。このタブを閉じます。
        ''' </summary>
        ''' <remarks>
        ''' 利用者の入力は混ぜず、固定の文字列だけを渡します。
        ''' ポストバックさせないため false を返します。
        ''' </remarks>
        Private Const CloseTabScript As String = "return AtomPage.closeTab();"

        ''' <summary>元INV_NO の「すべて」を表す値。</summary>
        Private Const AllInvoiceNoValue As String = ""

        ''' <summary>元INV_NO の「すべて」の表示文言。</summary>
        Private Const AllInvoiceNoText As String = "（すべて）"

        ''' <summary>表示モードを保持する ViewState のキー。</summary>
        Private Const DisplayModeKey As String = "DisplayMode"

        ''' <summary>タブの通常時の CSS クラス。</summary>
        Private Const TabCssClass As String = "atom-tab"

        ''' <summary>タブの選択時の CSS クラス。</summary>
        Private Const ActiveTabCssClass As String = "atom-tab atom-tab-active"

        ''' <summary>チェック結果の区分。エラー。</summary>
        Private Const ErrorStatus As Integer = 1

        ''' <summary>チェック結果の区分。警告。</summary>
        Private Const WarningStatus As Integer = 2

        ''' <summary>エラー行の CSS クラス。</summary>
        Private Const ErrorRowCssClass As String = "atom-grid-row-error"

        ''' <summary>警告行の CSS クラス。</summary>
        Private Const WarningRowCssClass As String = "atom-grid-row-warning"

        ''' <summary>メッセージ。</summary>
        Private Const SelectFailedMessage As String = "データの取得に失敗しました。"
        Private Const UpdateFailedMessage As String = "データの更新に失敗しました。"
        Private Const NotImplementedMessage As String = "現在未実装です。"
        Private Const UpdatedMessage As String = "更新しました。"

        ''' <summary>日付として認める書式。</summary>
        Private Shared ReadOnly DateFormats As String() = {
            "yyyy/MM/dd", "yyyy/M/d", "yyyy-MM-dd", "yyyyMMdd", "yy/MM/dd", "yy/M/d"
        }
#End Region

#Region "列挙型"
        ''' <summary>一覧に表示する対象です。</summary>
        Private Enum ResultDisplayMode
            Invoice = 0
            Packing = 1
        End Enum
#End Region

#Region "フィールド"
        ''' <summary>取込処理。</summary>
        Private ReadOnly _service As New ShanghaiImportService()
#End Region

#Region "イベントハンドラ"
        ''' <summary>
        ''' 初期化時の処理です。メニュー番号とヘッダーのボタンを設定します。
        ''' </summary>
        ''' <param name="e">イベント引数。</param>
        Protected Overrides Sub OnInit(e As EventArgs)
            Me.MenuNo = ThisMenuNo
            ' この画面は取込画面から別のタブで開きます。取込画面へ遷移させると
            ' 同じ画面が 2 つ開いた状態になるため、「閉じる」はこのタブを閉じるだけにします。
            Me.HeaderActions.Add(HeaderAction.CreateClientAction("閉じる", CloseTabScript))
            Me.HeaderActions.Add(New HeaderAction("MENUへ", HeaderAction.MenuScreenNo))
            MyBase.OnInit(e)
        End Sub

        ''' <summary>
        ''' 読み込み時の処理です。初回表示のみ選択肢と一覧を用意します。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
            ' すでに遷移を指示しているとき
            If Me.IsRedirecting Then
                Return
            End If

            ' 初回表示のとき
            If Not Me.IsPostBack Then
                Me.CurrentDisplayMode = ResultDisplayMode.Invoice
                Me.BindInvoiceNoList()
                Me.BindResults()
            End If
        End Sub

        ''' <summary>
        ''' 元INV_NO を変更したときの処理です。一覧を取り直します。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub DdlInvoiceNo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlInvoiceNo.SelectedIndexChanged
            Me.BindResults()
        End Sub

        ''' <summary>
        ''' INVOICE タブ押下時の処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub LnkTabInvoice_Click(sender As Object, e As EventArgs) Handles lnkTabInvoice.Click
            Me.CurrentDisplayMode = ResultDisplayMode.Invoice
            Me.BindResults()
        End Sub

        ''' <summary>
        ''' PACKING タブ押下時の処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub LnkTabPacking_Click(sender As Object, e As EventArgs) Handles lnkTabPacking.Click
            Me.CurrentDisplayMode = ResultDisplayMode.Packing
            Me.BindResults()
        End Sub

        ''' <summary>
        ''' EXCEL へボタン押下時の処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        ''' <remarks>
        ''' 旧 Access 版では Excel を添付したメール送信画面へ遷移します。
        ''' メール送信画面が未実装のため、現時点では案内のみ表示します。
        ''' </remarks>
        Private Sub BtnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
            Me.ShowMessage(NotImplementedMessage, False)
        End Sub

        ''' <summary>
        ''' INVOICE 一覧の行を割り当てるときの処理です。チェック結果に応じて行の色を変えます。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub GvInvoice_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gvInvoice.RowDataBound
            Dim entity As ShanghaiInvoiceImportRow = TryCast(e.Row.DataItem, ShanghaiInvoiceImportRow)
            ' 明細行でないとき
            If entity Is Nothing Then
                Return
            End If
            ' バインド時の状態変更のため、行の CSS クラスをここで設定します。
            Me.ApplyRowStyle(e.Row, entity.ErrStatus)
        End Sub

        ''' <summary>
        ''' INVOICE 一覧の行を作るときの処理です。入力欄の変更を受け取れるようにします。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        ''' <remarks>
        ''' RowDataBound はデータを割り当てるときだけ発生するため、ポストバックでは呼ばれません。
        ''' ポストバック時は ViewState から行が組み立て直される RowCreated で結び付ける必要があります。
        ''' ここで結び付けないと TextChanged が発生せず、編集内容が保存されません。
        ''' </remarks>
        Private Sub GvInvoice_RowCreated(sender As Object, e As GridViewRowEventArgs) Handles gvInvoice.RowCreated
            Me.AttachCellHandlers(e.Row, AddressOf Me.InvoiceCell_TextChanged)
        End Sub

        ''' <summary>
        ''' PACKING 一覧の行を割り当てるときの処理です。チェック結果に応じて行の色を変えます。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub GvPacking_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gvPacking.RowDataBound
            Dim entity As ShanghaiPackingImportRow = TryCast(e.Row.DataItem, ShanghaiPackingImportRow)
            ' 明細行でないとき
            If entity Is Nothing Then
                Return
            End If
            ' バインド時の状態変更のため、行の CSS クラスをここで設定します。
            Me.ApplyRowStyle(e.Row, entity.ErrStatus)
        End Sub

        ''' <summary>
        ''' PACKING 一覧の行を作るときの処理です。入力欄の変更を受け取れるようにします。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        ''' <remarks>
        ''' 結び付ける場所を RowCreated にする理由は GvInvoice_RowCreated と同じです。
        ''' </remarks>
        Private Sub GvPacking_RowCreated(sender As Object, e As GridViewRowEventArgs) Handles gvPacking.RowCreated
            Me.AttachCellHandlers(e.Row, AddressOf Me.PackingCell_TextChanged)
        End Sub

        ''' <summary>
        ''' INVOICE の入力欄を変更したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub InvoiceCell_TextChanged(sender As Object, e As EventArgs)
            Me.SaveCell(sender, True)
        End Sub

        ''' <summary>
        ''' PACKING の入力欄を変更したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub PackingCell_TextChanged(sender As Object, e As EventArgs)
            Me.SaveCell(sender, False)
        End Sub
#End Region

#Region "内部処理（編集）"
        ''' <summary>
        ''' 行にある入力欄へ、変更時の受け取り先を結び付けます。
        ''' </summary>
        ''' <param name="row">対象の行。</param>
        ''' <param name="handler">変更時に呼び出す処理。</param>
        ''' <remarks>
        ''' 入力欄はテンプレート内にあり designer からは参照できないため、ここで結び付けます。
        ''' </remarks>
        Private Sub AttachCellHandlers(row As GridViewRow, handler As EventHandler)
            ' 明細行でないとき（見出し・フッターには入力欄が無い）
            If row.RowType <> DataControlRowType.DataRow Then
                Return
            End If

            For Each cell As TableCell In row.Cells
                For Each control As Control In cell.Controls
                    Dim textBox As TextBox = TryCast(control, TextBox)
                    ' 入力欄でないとき
                    If textBox Is Nothing Then
                        Continue For
                    End If
                    AddHandler textBox.TextChanged, handler
                Next
            Next
        End Sub

        ''' <summary>
        ''' 変更された 1 項目を一時データへ反映します。
        ''' </summary>
        ''' <param name="sender">変更された入力欄。</param>
        ''' <param name="isInvoice">INVOICE 側のとき True。</param>
        ''' <remarks>
        ''' 入力が項目の型に合わないときは更新せず、一覧を読み直して元の値へ戻します。
        ''' チェック処理はやり直さないため ERR_CONTENTS は変わりません（旧 Access 版と同じ挙動）。
        ''' </remarks>
        Private Sub SaveCell(sender As Object, isInvoice As Boolean)
            Dim textBox As TextBox = TryCast(sender, TextBox)
            ' 入力欄が取得できないとき
            If textBox Is Nothing Then
                Return
            End If

            Dim row As GridViewRow = TryCast(textBox.NamingContainer, GridViewRow)
            ' 行が取得できないとき
            If row Is Nothing Then
                Return
            End If

            ' 画面から渡された ID は必ず定義一覧と照合し、想定外の項目は更新しません。
            Dim field As ShanghaiImportField
            ' INVOICE のとき
            If isInvoice Then
                field = ShanghaiImportFieldCatalog.FindInvoiceField(textBox.ID)
            ' PACKING のとき
            Else
                field = ShanghaiImportFieldCatalog.FindPackingField(textBox.ID)
            End If
            ' 定義に無い項目のとき
            If field Is Nothing Then
                Return
            End If

            Dim seqNo As Integer = Me.GetSeqNo(row, isInvoice)
            ' 行を特定できないとき
            If seqNo <= 0 Then
                Me.ShowMessage(UpdateFailedMessage, True)
                Me.BindResults()
                Return
            End If

            Dim value As Object = Nothing
            ' 入力が項目の型に合わないとき
            If Not Me.TryConvertValue(field, textBox.Text, value) Then
                Me.ShowMessage(field.DisplayName & "の入力が正しくありません。元の値に戻しました。", True)
                Me.BindResults()
                Return
            End If

            Try
                ' INVOICE のとき
                If isInvoice Then
                    Me._service.UpdateInvoiceResultValue(Me.CurrentLoginCode, seqNo, field, value)
                ' PACKING のとき
                Else
                    Me._service.UpdatePackingResultValue(Me.CurrentLoginCode, seqNo, field, value)
                End If
            Catch ex As SqlException
                ' 詳細は Service 側でログへ記録済みです。
                Me.ShowMessage(UpdateFailedMessage, True)
                Me.BindResults()
                Return
            End Try

            Me.ShowMessage(UpdatedMessage, False)
            ' 書式をそろえるため、保存後は読み直します。
            Me.BindResults()
        End Sub

        ''' <summary>
        ''' 行に対応する一時データの連番を取得します。
        ''' </summary>
        ''' <param name="row">対象の行。</param>
        ''' <param name="isInvoice">INVOICE 側のとき True。</param>
        ''' <returns>連番。取得できないときは 0。</returns>
        Private Function GetSeqNo(row As GridViewRow, isInvoice As Boolean) As Integer
            Dim grid As GridView = Me.gvPacking
            ' INVOICE のとき
            If isInvoice Then
                grid = Me.gvInvoice
            End If

            ' 行の位置が範囲外のとき
            If row.RowIndex < 0 OrElse row.RowIndex >= grid.DataKeys.Count Then
                Return 0
            End If

            Dim stored As Object = grid.DataKeys(row.RowIndex).Value
            ' キーが無いとき
            If stored Is Nothing Then
                Return 0
            End If

            Dim seqNo As Integer = 0
            ' 数値として読めないとき
            If Not Integer.TryParse(stored.ToString(), seqNo) Then
                Return 0
            End If
            Return seqNo
        End Function

        ''' <summary>
        ''' 入力された文字列を、項目の型に合わせた値へ変換します。
        ''' </summary>
        ''' <param name="field">項目の定義。</param>
        ''' <param name="text">入力された文字列。</param>
        ''' <param name="value">変換した値。空のときは Nothing。</param>
        ''' <returns>変換できたとき True。</returns>
        Private Function TryConvertValue(field As ShanghaiImportField, text As String, ByRef value As Object) As Boolean
            Dim trimmed As String = String.Empty
            ' 入力があるとき
            If text IsNot Nothing Then
                trimmed = text.Trim()
            End If

            ' 空欄のとき（NULL を設定する）
            If trimmed.Length = 0 Then
                value = Nothing
                Return True
            End If

            Select Case field.Kind
                ' 整数のとき
                Case ShanghaiImportValueKind.WholeNumber
                    Dim parsedWhole As Decimal = 0
                    If Not Decimal.TryParse(trimmed, NumberStyles.Number, CultureInfo.InvariantCulture, parsedWhole) Then
                        Return False
                    End If
                    ' 小数を保持できない列のため、四捨五入して整数にします。
                    value = Decimal.Round(parsedWhole, 0, MidpointRounding.AwayFromZero)
                    Return True

                ' 小数を含む数値のとき
                Case ShanghaiImportValueKind.DecimalNumber
                    Dim parsedNumber As Double = 0
                    If Not Double.TryParse(trimmed, NumberStyles.Number, CultureInfo.InvariantCulture, parsedNumber) Then
                        Return False
                    End If
                    value = parsedNumber
                    Return True

                ' 日付のとき
                Case ShanghaiImportValueKind.DateValue
                    Dim parsedDate As Date = Date.MinValue
                    If Date.TryParseExact(trimmed, DateFormats, CultureInfo.InvariantCulture,
                                          DateTimeStyles.None, parsedDate) Then
                        value = parsedDate
                        Return True
                    End If
                    Return False

                ' 文字列のとき
                Case Else
                    ' 列の桁数を超えているとき
                    If trimmed.Length > field.MaxLength Then
                        Return False
                    End If
                    value = trimmed
                    Return True
            End Select
        End Function
#End Region

#Region "内部処理（画面の状態）"
        ''' <summary>現在の表示対象。</summary>
        Private Property CurrentDisplayMode As ResultDisplayMode
            Get
                Dim stored As Object = Me.ViewState(DisplayModeKey)
                ' 保持していないとき
                If stored Is Nothing Then
                    Return ResultDisplayMode.Invoice
                End If
                Return CType(stored, ResultDisplayMode)
            End Get
            Set(value As ResultDisplayMode)
                Me.ViewState(DisplayModeKey) = value
            End Set
        End Property

        ''' <summary>選択されている元INV_NO。すべてのときは空文字。</summary>
        Private ReadOnly Property SelectedInvoiceNo As String
            Get
                ' 選択されていないとき
                If Me.ddlInvoiceNo.SelectedIndex < 0 Then
                    Return AllInvoiceNoValue
                End If
                Return Me.ddlInvoiceNo.SelectedValue
            End Get
        End Property

        ''' <summary>操作している利用者のログインコード。</summary>
        Private ReadOnly Property CurrentLoginCode As String
            Get
                Return SessionContext.Current.LoginCode
            End Get
        End Property
#End Region

#Region "内部処理（表示）"
        ''' <summary>
        ''' 元INV_NO の選択肢を割り当てます。
        ''' </summary>
        Private Sub BindInvoiceNoList()
            Dim invoiceNos As IList(Of String)
            Try
                invoiceNos = Me._service.SelectInvoiceNoList(Me.CurrentLoginCode)
            Catch ex As SqlException
                ' 詳細は Service 側でログへ記録済みです。
                Me.ShowMessage(SelectFailedMessage, True)
                Return
            End Try

            Me.ddlInvoiceNo.Items.Clear()
            Me.ddlInvoiceNo.Items.Add(New ListItem(AllInvoiceNoText, AllInvoiceNoValue))
            For Each invoiceNo As String In invoiceNos
                Me.ddlInvoiceNo.Items.Add(New ListItem(invoiceNo, invoiceNo))
            Next
        End Sub

        ''' <summary>
        ''' 表示対象に応じて一覧を割り当てます。
        ''' </summary>
        Private Sub BindResults()
            Me.ApplyTabState()

            ' PACKING を表示するとき
            If Me.CurrentDisplayMode = ResultDisplayMode.Packing Then
                Me.BindPackingResults()
                Return
            End If
            ' INVOICE を表示するとき
            Me.BindInvoiceResults()
        End Sub

        ''' <summary>
        ''' INVOICE の一覧を割り当てます。
        ''' </summary>
        Private Sub BindInvoiceResults()
            Try
                Dim rows As IList(Of ShanghaiInvoiceImportRow) =
                    Me._service.SelectInvoiceResults(Me.CurrentLoginCode, Me.SelectedInvoiceNo)
                Me.gvInvoice.DataSource = rows
                Me.gvInvoice.DataBind()
                Me.ShowCount(rows.Count)
            Catch ex As SqlException
                ' 詳細は Service 側でログへ記録済みです。表示中の一覧は維持します。
                Me.ShowMessage(SelectFailedMessage, True)
            End Try
        End Sub

        ''' <summary>
        ''' PACKING の一覧を割り当てます。
        ''' </summary>
        Private Sub BindPackingResults()
            Try
                Dim rows As IList(Of ShanghaiPackingImportRow) =
                    Me._service.SelectPackingResults(Me.CurrentLoginCode, Me.SelectedInvoiceNo)
                Me.gvPacking.DataSource = rows
                Me.gvPacking.DataBind()
                Me.ShowCount(rows.Count)
            Catch ex As SqlException
                ' 詳細は Service 側でログへ記録済みです。表示中の一覧は維持します。
                Me.ShowMessage(SelectFailedMessage, True)
            End Try
        End Sub

        ''' <summary>
        ''' タブと一覧の表示状態を、現在の表示対象に合わせます。
        ''' </summary>
        Private Sub ApplyTabState()
            Dim isPacking As Boolean = (Me.CurrentDisplayMode = ResultDisplayMode.Packing)

            Me.pnlInvoice.Visible = Not isPacking
            Me.pnlPacking.Visible = isPacking

            ' PACKING を表示しているとき
            If isPacking Then
                Me.lnkTabInvoice.CssClass = TabCssClass
                Me.lnkTabPacking.CssClass = ActiveTabCssClass
                Return
            End If
            ' INVOICE を表示しているとき
            Me.lnkTabInvoice.CssClass = ActiveTabCssClass
            Me.lnkTabPacking.CssClass = TabCssClass
        End Sub

        ''' <summary>
        ''' 件数を表示します。
        ''' </summary>
        ''' <param name="count">表示している件数。</param>
        Private Sub ShowCount(count As Integer)
            Me.lblCount.Text = count.ToString() & " 件"
        End Sub

        ''' <summary>
        ''' チェック結果に応じて行の CSS クラスを設定します。
        ''' </summary>
        ''' <param name="row">対象の行。</param>
        ''' <param name="errStatus">チェック結果の区分。</param>
        Private Sub ApplyRowStyle(row As GridViewRow, errStatus As Integer)
            Select Case errStatus
                ' エラーのとき
                Case ErrorStatus
                    row.CssClass = ErrorRowCssClass
                ' 警告のとき
                Case WarningStatus
                    row.CssClass = WarningRowCssClass
            End Select
        End Sub
#End Region

    End Class
End Namespace
