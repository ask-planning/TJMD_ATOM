Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports Atom.Core.Entities
Imports Atom.Data.Services
Imports Atom.Web.Common

Namespace Pages.Entry

    ''' <summary>
    ''' INVOICEヘッダー入力画面（menuNo=103）です。
    ''' 旧 Access 版の F_1_上海輸入_INVOICE_HEADER に相当します。
    ''' </summary>
    ''' <remarks>
    ''' 仕様は docs/specs/103_INVOICEヘッダー入力.md を正とします。
    '''
    ''' **現時点は表示までです。**登録・明細更新・帳票出力は未実装で、
    ''' ボタンを押すと「未実装です」とだけ表示します。
    '''
    ''' 明細は作業用の一時テーブル（dbo.tbl_t_shanghai_invoice_header_sub_temp）に
    ''' 利用者単位で保持します。画面での編集はその一時データへ反映し、
    ''' 登録時に本体へ移します（設計書 §8）。
    ''' </remarks>
    Public Class InvoiceHeader
        Inherits BasePage

#Region "定数"
        ''' <summary>この画面のメニュー番号。</summary>
        Private Const ThisMenuNo As Integer = 103

        ''' <summary>遷移先。R3在庫転送作成。</summary>
        Private Const StockTransferMenuNo As Integer = 106

        ''' <summary>遷移先。R3購買発注確認。</summary>
        Private Const PurchaseConfirmMenuNo As Integer = 105

        ''' <summary>遷移先。R3購買発注取込。</summary>
        Private Const PurchaseImportMenuNo As Integer = 104

        ''' <summary>遷移先。INVOICE確認。</summary>
        Private Const InvoiceConfirmMenuNo As Integer = 102

        ''' <summary>未実装であることを伝える文言。</summary>
        Private Const NotImplementedMessage As String = "この処理はまだ実装していません。"

        ''' <summary>新規登録であることを伝える文言。旧 Access 版を踏襲します。</summary>
        Private Const NewHeaderMessage As String = "新規登録です。"

        ''' <summary>
        ''' 外部システムへの通知が未実装であることを伝える文言です。
        ''' </summary>
        ''' <remarks>
        ''' 旧 Access は配送予定日が入っているとき DataSpider へ通知していました（設計書 §8.5 / C-5）。
        ''' 扱いが決まるまで、利用者に分かるよう画面へ出しておきます。
        ''' </remarks>
        Private Const PendingNotifyMessage As String =
            "（配送予定日が入力されています。外部システムへの連携は未実装です。）"

        ''' <summary>元INV_NO の未選択を表す値。</summary>
        Private Const NoSelectionValue As String = ""

        ''' <summary>元INV_NO の未選択の表示文言。</summary>
        Private Const NoSelectionText As String = "（選択してください）"

        ''' <summary>選択肢の未設定を表す値。</summary>
        Private Const EmptyOptionValue As String = ""

        ''' <summary>選択肢の未設定の表示文言。</summary>
        Private Const EmptyOptionText As String = ""

        ''' <summary>3 桁区切りの表示書式。</summary>
        Private Const NumberFormat As String = "#,##0"

        ''' <summary>日付欄の書式。TextMode="Date" が受け取れる形にします。</summary>
        Private Const DateFormat As String = "yyyy-MM-dd"

        ''' <summary>ヘッダが未登録のときに入力欄へ付ける CSS クラス。</summary>
        Private Const DisabledCssClass As String = "atom-readonly"

        ''' <summary>ヘッダが登録済みのときに入力欄へ付ける CSS クラス。</summary>
        Private Const EnabledCssClass As String = "atom-select"

        ''' <summary>選択中の元INV_NO を保持する ViewState のキー。</summary>
        Private Const SelectedInvoiceNoKey As String = "SelectedInvoiceNoMain"
#End Region

#Region "フィールド"
        ''' <summary>業務処理。</summary>
        Private ReadOnly _service As New ShanghaiInvoiceHeaderService()

        ''' <summary>明細の配送タイプの選択肢。</summary>
        ''' <remarks>
        ''' ViewState には持たせません。一覧の行は ViewState の読み込み時に作り直されるため、
        ''' それより前の Page_Init で毎回読み込みます。
        ''' </remarks>
        Private _haisoTypeOptions As IList(Of CodeOption) = New List(Of CodeOption)()

        ''' <summary>明細のコンテナタイプの選択肢。</summary>
        ''' <remarks>配送タイプと同じ理由で ViewState には持たせません。</remarks>
        Private _containerTypeOptions As IList(Of CodeOption) = New List(Of CodeOption)()
#End Region

#Region "プロパティ"
        ''' <summary>選択中の元INV_NO。未選択のときは空文字。</summary>
        Private Property SelectedInvoiceNoMain As String
            Get
                Dim stored As Object = Me.ViewState(SelectedInvoiceNoKey)
                ' 保持していないとき
                If stored Is Nothing Then
                    Return String.Empty
                End If
                Return CStr(stored)
            End Get
            Set(value As String)
                Me.ViewState(SelectedInvoiceNoKey) = value
            End Set
        End Property
#End Region

#Region "イベント"
        ''' <summary>
        ''' 初期化時の処理です。メニュー番号とヘッダーのボタンを設定します。
        ''' </summary>
        ''' <param name="e">イベント引数。</param>
        ''' <remarks>
        ''' メニュー番号は権限判定より前に決める必要があるため、Page_Load ではなくここで設定します。
        ''' </remarks>
        Protected Overrides Sub OnInit(e As EventArgs)
            Me.MenuNo = ThisMenuNo

            Me.HeaderActions.Add(New HeaderAction("在庫転送", StockTransferMenuNo))
            Me.HeaderActions.Add(New HeaderAction("発注確認", PurchaseConfirmMenuNo))
            Me.HeaderActions.Add(New HeaderAction("発注取込", PurchaseImportMenuNo))
            Me.HeaderActions.Add(New HeaderAction("INV確認", InvoiceConfirmMenuNo))
            Me.HeaderActions.Add(New HeaderAction("MENUへ", HeaderAction.MenuScreenNo))

            MyBase.OnInit(e)
        End Sub

        ''' <summary>
        ''' 初期化時の処理です。選択肢を読み込みます。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        ''' <remarks>
        ''' 一覧の行は ViewState の読み込み時に作り直され、そのとき配送タイプの選択肢が要ります。
        ''' それより前に用意する必要があるため、Page_Load ではなくここで読み込みます。
        ''' 選択肢は毎回読み込みます。ViewState に載せないので、その分だけ軽くなります。
        ''' </remarks>
        Private Sub Page_Init(sender As Object, e As EventArgs) Handles Me.Init
            Me.LoadOptions()
        End Sub

        ''' <summary>
        ''' 画面の読み込み時の処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
            ' 初回表示のとき
            If Not Me.IsPostBack Then
                Me.ClearHeaderInputs()
                Me.ApplyHeaderEditable(False)
            End If
        End Sub

        ''' <summary>
        ''' 元INV_NO を選んだときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub DdlInvoiceNoMain_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlInvoiceNoMain.SelectedIndexChanged
            Dim invoiceNoMain As String = Me.ddlInvoiceNoMain.SelectedValue
            Me.SelectedInvoiceNoMain = invoiceNoMain

            ' 未選択に戻したとき
            If String.IsNullOrEmpty(invoiceNoMain) Then
                Me.ClearHeaderInputs()
                Me.ApplyHeaderEditable(False)
                Me.BindDetails(New List(Of ShanghaiInvoiceHeaderSub)())
                Return
            End If

            Me.ShowSelectedInvoice(invoiceNoMain)
        End Sub

        ''' <summary>
        ''' 明細の行を作るときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        ''' <remarks>
        ''' 行内のコントロールへの値の設定は RowDataBound で行いますが、
        ''' 配送タイプの選択肢はポストバック後も必要になるため、ここで組み立てます。
        ''' </remarks>
        Private Sub GvDetail_RowCreated(sender As Object, e As GridViewRowEventArgs) Handles gvDetail.RowCreated
            ' データ行以外のとき
            If e.Row.RowType <> DataControlRowType.DataRow Then
                Return
            End If

            Dim ddlHaisoType As DropDownList = TryCast(e.Row.FindControl("ddlHaisoType"), DropDownList)
            ' 配送タイプがあるとき
            If ddlHaisoType IsNot Nothing Then
                Me.FillCodeOptions(ddlHaisoType, Me._haisoTypeOptions, True)
            End If

            Dim ddlContainerType As DropDownList = TryCast(e.Row.FindControl("ddlContainerType"), DropDownList)
            ' コンテナタイプがあるとき
            If ddlContainerType IsNot Nothing Then
                Me.FillCodeOptions(ddlContainerType, Me._containerTypeOptions, True)
            End If
        End Sub

        ''' <summary>
        ''' 明細の行にデータを割り当てるときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub GvDetail_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gvDetail.RowDataBound
            ' データ行以外のとき
            If e.Row.RowType <> DataControlRowType.DataRow Then
                Return
            End If

            Dim detail As ShanghaiInvoiceHeaderSub = TryCast(e.Row.DataItem, ShanghaiInvoiceHeaderSub)
            ' 想定した型でないとき
            If detail Is Nothing Then
                Return
            End If

            Me.SetLiteral(e.Row, "litInvoiceNo", detail.InvoiceNo)
            Me.SetLiteral(e.Row, "litInvoiceNoPrint", detail.InvoiceNoPrint)
            Me.SetTextBox(e.Row, "txtHaisoYoteiYmd", Me.ToDateText(detail.HaisoYoteiYmd))
            Me.SetDropDown(e.Row, "ddlContainerType", Me.ToIntegerText(detail.ContainerType))
            Me.SetTextBox(e.Row, "txtContainerNo", detail.ContainerNo)
            Me.SetTextBox(e.Row, "txtSealNo", detail.SealNo)
            Me.SetTextBox(e.Row, "txtCartonQtySub", Me.ToIntegerText(detail.CartonQty))
            Me.SetDropDown(e.Row, "ddlHaisoType", Me.ToIntegerText(detail.HaisoType))
        End Sub

        ''' <summary>
        ''' 「登 録」を押したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnRegister_Click(sender As Object, e As EventArgs) Handles btnRegister.Click
            Dim invoiceNoMain As String = Me.SelectedInvoiceNoMain
            Dim loginCode As String = SessionContext.Current.LoginCode

            Dim result As ShanghaiInvoiceRegisterResult =
                Me._service.Register(loginCode, Me.CollectHeader(invoiceNoMain), Me.CollectDetails(invoiceNoMain))

            ' 入力に問題があったとき
            If Not result.IsSuccess Then
                Me.ShowMessage(result.ValidationMessage, True)
                Return
            End If

            ' 登録後の内容で画面を作り直す（設計書 §8.7）
            Me.ShowSelectedInvoice(invoiceNoMain)
            Me.ClearUnsavedState()

            Dim message As String = ShanghaiInvoiceHeaderService.RegisteredMessage
            ' 配送予定日が入っているとき（旧 Access は外部システムへ通知していた。§8.5）
            If result.HasHaisoYotei Then
                message &= PendingNotifyMessage
            End If
            ' POEM の更新を行わなかったとき
            If result.IsPoemSkipped Then
                message &= ShanghaiInvoiceHeaderService.PoemSkippedMessage
            End If
            Me.ShowMessage(message, False)
        End Sub

        ''' <summary>
        ''' 「明細更新」を押したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnMeisaiUpdate_Click(sender As Object, e As EventArgs) Handles btnMeisaiUpdate.Click
            Dim invoiceNoMain As String = Me.SelectedInvoiceNoMain
            Dim loginCode As String = SessionContext.Current.LoginCode

            Dim message As String =
                Me._service.UpdateContainerNo(loginCode, invoiceNoMain, Me.CollectDetails(invoiceNoMain))

            Dim isError As Boolean = Not String.Equals(message, ShanghaiInvoiceHeaderService.ContainerUpdatedMessage,
                                                       StringComparison.Ordinal)
            ' 更新できたとき（明細の表示を作り直す）
            If Not isError Then
                Me.ShowSelectedInvoice(invoiceNoMain)
            End If

            Me.ShowMessage(message, isError)
        End Sub

        ''' <summary>
        ''' 「配送手配書」を押したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnHaisoTehai_Click(sender As Object, e As EventArgs) Handles btnHaisoTehai.Click
            Me.ShowMessage(NotImplementedMessage, False)
        End Sub

        ''' <summary>
        ''' 配送手配書の「表示」を押したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnHaisoTehaiView_Click(sender As Object, e As EventArgs) Handles btnHaisoTehaiView.Click
            Me.ShowMessage(NotImplementedMessage, False)
        End Sub

        ''' <summary>
        ''' 「入荷明細書」を押したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnNyukaMeisai_Click(sender As Object, e As EventArgs) Handles btnNyukaMeisai.Click
            Me.ShowMessage(NotImplementedMessage, False)
        End Sub

        ''' <summary>
        ''' 入荷明細書の「表示」を押したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnNyukaMeisaiView_Click(sender As Object, e As EventArgs) Handles btnNyukaMeisaiView.Click
            Me.ShowMessage(NotImplementedMessage, False)
        End Sub
#End Region

#Region "内部処理（表示）"
        ''' <summary>
        ''' 選択肢を読み込んで画面へ設定します。
        ''' </summary>
        Private Sub LoadOptions()
            Dim options As ShanghaiInvoiceHeaderOptions = Me._service.LoadOptions()

            Me.ddlInvoiceNoMain.Items.Clear()
            Me.ddlInvoiceNoMain.Items.Add(New ListItem(NoSelectionText, NoSelectionValue))
            For Each invoiceNoMain As String In options.InvoiceNoMains
                Me.ddlInvoiceNoMain.Items.Add(New ListItem(invoiceNoMain, invoiceNoMain))
            Next

            Me.ddlVessel.Items.Clear()
            Me.ddlVessel.Items.Add(New ListItem(EmptyOptionText, EmptyOptionValue))
            For Each vessel As String In options.Vessels
                Me.ddlVessel.Items.Add(New ListItem(vessel, vessel))
            Next

            Me.FillCodeOptions(Me.ddlSyukkaHoho, options.SyukkaHohos, True)
            Me.FillCodeOptions(Me.ddlOtsunaka, options.Otsunakas, True)

            ' 明細の行ごとに使うため保持する
            Me._haisoTypeOptions = options.HaisoTypes
            Me._containerTypeOptions = options.ContainerTypes
        End Sub

        ''' <summary>
        ''' 選択した INVOICE の内容を画面へ表示します。
        ''' </summary>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        Private Sub ShowSelectedInvoice(invoiceNoMain As String)
            Dim loginCode As String = SessionContext.Current.LoginCode
            Dim chotatsu As Integer = Me.CurrentShanghaiChotatsu()

            Dim view As ShanghaiInvoiceHeaderView = Me._service.LoadView(loginCode, chotatsu, invoiceNoMain)

            ' BL DATE と出荷方法が食い違っているとき
            If view.Existence.HasConflict Then
                Me.ClearHeaderInputs()
                Me.ApplyHeaderEditable(False)
                Me.BindDetails(New List(Of ShanghaiInvoiceHeaderSub)())
                Me.ShowMessage(ShanghaiInvoiceHeaderService.ConflictMessage, True)
                Return
            End If

            Me.ClearHeaderInputs()
            Me.ShowHeader(view)
            Me.ShowTotal(view)
            Me.ApplyHeaderEditable(Not view.Existence.IsNewHeader)
            Me.BindDetails(view.Details)

            ' ヘッダが未登録のとき
            If view.Existence.IsNewHeader Then
                Me.ShowMessage(NewHeaderMessage, False)
            End If
        End Sub

        ''' <summary>
        ''' ヘッダの内容を画面へ設定します。
        ''' </summary>
        ''' <param name="view">表示内容。</param>
        Private Sub ShowHeader(view As ShanghaiInvoiceHeaderView)
            ' ヘッダが未登録のとき
            If view.Header Is Nothing Then
                Return
            End If

            Me.txtBlDate.Text = Me.ToDateText(view.Header.BlDate)
            Me.SelectValue(Me.ddlVessel, view.Header.Vessel)
            Me.SelectValue(Me.ddlSyukkaHoho, Me.ToIntegerText(view.Header.SyukkaHoho))
            Me.SelectValue(Me.ddlOtsunaka, Me.ToIntegerText(view.Header.OtsunakaNo))
            Me.txtReturnPallet.Text = Me.ToIntegerText(view.Header.ReturnPallet)
        End Sub

        ''' <summary>
        ''' 明細から集計した合計を画面へ設定します。
        ''' </summary>
        ''' <param name="view">表示内容。</param>
        ''' <remarks>
        ''' ヘッダが未登録のときは、BL DATE と出荷方法も集計値で埋めます（設計書 §6.2-6）。
        ''' </remarks>
        Private Sub ShowTotal(view As ShanghaiInvoiceHeaderView)
            ' 合計が取れなかったとき
            If view.Total Is Nothing Then
                Return
            End If

            Me.txtCartonQty.Text = Me.ToNumberText(view.Total.TotalCarton)
            Me.txtTotalSuryo.Text = Me.ToNumberText(view.Total.TotalSuryo)
            Me.txtTotalAmount.Text = Me.ToNumberText(view.Total.TotalAmount)

            ' ヘッダが登録済みのときは、ヘッダ側の値を優先する
            If Not view.Existence.IsNewHeader Then
                Return
            End If

            Me.txtBlDate.Text = Me.ToDateText(view.Total.BlDate)
            Me.SelectValue(Me.ddlSyukkaHoho, Me.ToIntegerText(view.Total.SyukkaHoho))
        End Sub

        ''' <summary>
        ''' 明細を一覧へ割り当てます。
        ''' </summary>
        ''' <param name="details">明細の一覧。</param>
        Private Sub BindDetails(details As IList(Of ShanghaiInvoiceHeaderSub))
            Me.gvDetail.DataSource = details
            Me.gvDetail.DataBind()
        End Sub

        ''' <summary>
        ''' ヘッダの入力欄を空にします。
        ''' </summary>
        Private Sub ClearHeaderInputs()
            Me.txtBlDate.Text = String.Empty
            Me.SelectValue(Me.ddlVessel, String.Empty)
            Me.SelectValue(Me.ddlSyukkaHoho, String.Empty)
            Me.SelectValue(Me.ddlOtsunaka, String.Empty)
            Me.txtCartonQty.Text = String.Empty
            Me.txtTotalSuryo.Text = String.Empty
            Me.txtTotalAmount.Text = String.Empty
            Me.txtReturnPallet.Text = String.Empty
        End Sub

        ''' <summary>
        ''' ヘッダの登録状況に応じて入力欄の見た目を切り替えます。
        ''' </summary>
        ''' <param name="isRegistered">ヘッダが登録済みのときは True。</param>
        ''' <remarks>
        ''' 旧 Access 版は未登録のとき BL DATE と出荷方法をグレーにしています（設計書 §3.2）。
        ''' 値そのものは明細の集計で埋まるため、入力は禁止しません。
        ''' </remarks>
        Private Sub ApplyHeaderEditable(isRegistered As Boolean)
            Dim textCssClass As String = If(isRegistered, "atom-cell atom-cell-date", "atom-readonly atom-cell-date")
            Dim selectCssClass As String = If(isRegistered, EnabledCssClass, DisabledCssClass)

            Me.txtBlDate.CssClass = textCssClass
            Me.ddlSyukkaHoho.CssClass = selectCssClass
        End Sub
#End Region

#Region "内部処理（部品）"
        ''' <summary>
        ''' 選択肢をドロップダウンへ設定します。
        ''' </summary>
        ''' <param name="target">設定先。</param>
        ''' <param name="options">選択肢。</param>
        ''' <param name="addEmpty">先頭に空の選択肢を入れるときは True。</param>
        Private Sub FillCodeOptions(target As DropDownList, options As IList(Of CodeOption), addEmpty As Boolean)
            target.Items.Clear()

            ' 空の選択肢を入れるとき
            If addEmpty Then
                target.Items.Add(New ListItem(EmptyOptionText, EmptyOptionValue))
            End If

            For Each item As CodeOption In options
                target.Items.Add(New ListItem(item.Text, item.Value.ToString(CultureInfo.InvariantCulture)))
            Next
        End Sub

        ''' <summary>
        ''' ドロップダウンの選択値を設定します。選択肢に無い値は無視します。
        ''' </summary>
        ''' <param name="target">設定先。</param>
        ''' <param name="value">選択する値。</param>
        Private Sub SelectValue(target As DropDownList, value As String)
            Dim item As ListItem = target.Items.FindByValue(If(value, String.Empty))
            target.ClearSelection()

            ' 該当が無いとき
            If item Is Nothing Then
                Return
            End If

            item.Selected = True
        End Sub

        ''' <summary>
        ''' 行内の表示用コントロールへ値を設定します。
        ''' </summary>
        ''' <param name="row">対象の行。</param>
        ''' <param name="controlId">コントロールの ID。</param>
        ''' <param name="value">設定する値。</param>
        Private Sub SetLiteral(row As GridViewRow, controlId As String, value As String)
            Dim target As Literal = TryCast(row.FindControl(controlId), Literal)
            ' 見つからないとき
            If target Is Nothing Then
                Return
            End If
            target.Text = value
        End Sub

        ''' <summary>
        ''' 行内の入力欄へ値を設定します。
        ''' </summary>
        ''' <param name="row">対象の行。</param>
        ''' <param name="controlId">コントロールの ID。</param>
        ''' <param name="value">設定する値。</param>
        Private Sub SetTextBox(row As GridViewRow, controlId As String, value As String)
            Dim target As TextBox = TryCast(row.FindControl(controlId), TextBox)
            ' 見つからないとき
            If target Is Nothing Then
                Return
            End If
            target.Text = value
        End Sub

        ''' <summary>
        ''' 行内のドロップダウンの選択値を設定します。
        ''' </summary>
        ''' <param name="row">対象の行。</param>
        ''' <param name="controlId">コントロールの ID。</param>
        ''' <param name="value">選択する値。</param>
        Private Sub SetDropDown(row As GridViewRow, controlId As String, value As String)
            Dim target As DropDownList = TryCast(row.FindControl(controlId), DropDownList)
            ' 見つからないとき
            If target Is Nothing Then
                Return
            End If
            Me.SelectValue(target, value)
        End Sub

        ''' <summary>
        ''' 画面のヘッダ内容を集めます。
        ''' </summary>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <returns>ヘッダ内容。</returns>
        Private Function CollectHeader(invoiceNoMain As String) As ShanghaiInvoiceHeader
            Dim header As New ShanghaiInvoiceHeader()
            header.InvoiceNoMain = invoiceNoMain
            header.BlDate = Me.ParseDate(Me.txtBlDate.Text)
            header.Vessel = Me.ddlVessel.SelectedValue
            header.SyukkaHoho = Me.ParseInteger(Me.ddlSyukkaHoho.SelectedValue)
            header.OtsunakaNo = Me.ParseInteger(Me.ddlOtsunaka.SelectedValue)
            header.CartonQty = Me.ParseInteger(Me.txtCartonQty.Text)
            header.ReturnPallet = Me.ParseInteger(Me.txtReturnPallet.Text)
            Return header
        End Function

        ''' <summary>
        ''' 画面の明細内容を集めます。
        ''' </summary>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <returns>明細の一覧。</returns>
        ''' <remarks>
        ''' 行を識別する INVOICE 番号は、画面に出している表示専用の値から取ります。
        ''' 入力欄として出していないため、利用者が書き換えることはできません。
        ''' </remarks>
        Private Function CollectDetails(invoiceNoMain As String) As IList(Of ShanghaiInvoiceHeaderSub)
            Dim details As New List(Of ShanghaiInvoiceHeaderSub)()

            For Each row As GridViewRow In Me.gvDetail.Rows
                ' データ行以外のとき
                If row.RowType <> DataControlRowType.DataRow Then
                    Continue For
                End If

                Dim detail As New ShanghaiInvoiceHeaderSub()
                detail.InvoiceNoMain = invoiceNoMain
                detail.InvoiceNo = Me.GetLiteral(row, "litInvoiceNo")
                detail.HaisoType = Me.ParseInteger(Me.GetDropDown(row, "ddlHaisoType"))
                detail.ContainerType = Me.ParseInteger(Me.GetDropDown(row, "ddlContainerType"))
                detail.HaisoYoteiYmd = Me.ParseDate(Me.GetTextBox(row, "txtHaisoYoteiYmd"))
                detail.ContainerNo = Me.GetTextBox(row, "txtContainerNo")
                detail.SealNo = Me.GetTextBox(row, "txtSealNo")
                detail.CartonQty = Me.ParseInteger(Me.GetTextBox(row, "txtCartonQtySub"))

                ' 行を識別できないとき
                If detail.InvoiceNo.Length = 0 Then
                    Continue For
                End If

                details.Add(detail)
            Next

            Return details
        End Function

        ''' <summary>
        ''' 行内の表示用コントロールの値を取得します。
        ''' </summary>
        ''' <param name="row">対象の行。</param>
        ''' <param name="controlId">コントロールの ID。</param>
        ''' <returns>値。見つからないときは空文字。</returns>
        Private Function GetLiteral(row As GridViewRow, controlId As String) As String
            Dim target As Literal = TryCast(row.FindControl(controlId), Literal)
            ' 見つからないとき
            If target Is Nothing Then
                Return String.Empty
            End If
            Return target.Text
        End Function

        ''' <summary>
        ''' 行内の入力欄の値を取得します。
        ''' </summary>
        ''' <param name="row">対象の行。</param>
        ''' <param name="controlId">コントロールの ID。</param>
        ''' <returns>値。見つからないときは空文字。</returns>
        Private Function GetTextBox(row As GridViewRow, controlId As String) As String
            Dim target As TextBox = TryCast(row.FindControl(controlId), TextBox)
            ' 見つからないとき
            If target Is Nothing Then
                Return String.Empty
            End If
            Return target.Text.Trim()
        End Function

        ''' <summary>
        ''' 行内のドロップダウンの選択値を取得します。
        ''' </summary>
        ''' <param name="row">対象の行。</param>
        ''' <param name="controlId">コントロールの ID。</param>
        ''' <returns>値。見つからないときは空文字。</returns>
        Private Function GetDropDown(row As GridViewRow, controlId As String) As String
            Dim target As DropDownList = TryCast(row.FindControl(controlId), DropDownList)
            ' 見つからないとき
            If target Is Nothing Then
                Return String.Empty
            End If
            Return target.SelectedValue
        End Function

        ''' <summary>
        ''' 文字列を日付に変換します。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>日付。変換できないときは Nothing。</returns>
        Private Function ParseDate(value As String) As Date?
            Dim parsed As Date
            ' 変換できないとき
            If Not Date.TryParse(value, parsed) Then
                Return Nothing
            End If
            Return parsed
        End Function

        ''' <summary>
        ''' 文字列を整数に変換します。
        ''' </summary>
        ''' <param name="value">元の値。3 桁区切りを含んでいても変換します。</param>
        ''' <returns>整数。変換できないときは Nothing。</returns>
        Private Function ParseInteger(value As String) As Integer?
            Dim parsed As Integer
            ' 変換できないとき
            If Not Integer.TryParse(value, NumberStyles.Number, CultureInfo.CurrentCulture, parsed) Then
                Return Nothing
            End If
            Return parsed
        End Function

        ''' <summary>
        ''' 操作している利用者の担当区分を取得します。
        ''' </summary>
        ''' <returns>担当区分。未設定のときは 0。</returns>
        Private Function CurrentShanghaiChotatsu() As Integer
            Dim stored As Integer? = SessionContext.Current.ShanghaiChotatsu
            Return If(stored.HasValue, stored.Value, 0)
        End Function
#End Region

#Region "内部処理（書式）"
        ''' <summary>
        ''' 日付を入力欄の書式にします。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>文字列。未設定のときは空文字。</returns>
        Private Function ToDateText(value As Date?) As String
            ' 値が無いとき
            If Not value.HasValue Then
                Return String.Empty
            End If
            Return value.Value.ToString(DateFormat, CultureInfo.InvariantCulture)
        End Function

        ''' <summary>
        ''' 整数をそのままの桁で文字列にします。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>文字列。未設定のときは空文字。</returns>
        Private Function ToIntegerText(value As Integer?) As String
            ' 値が無いとき
            If Not value.HasValue Then
                Return String.Empty
            End If
            Return value.Value.ToString(CultureInfo.InvariantCulture)
        End Function

        ''' <summary>
        ''' 整数を 3 桁区切りの文字列にします。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>文字列。未設定のときは空文字。</returns>
        Private Function ToNumberText(value As Integer?) As String
            ' 値が無いとき
            If Not value.HasValue Then
                Return String.Empty
            End If
            Return value.Value.ToString(NumberFormat, CultureInfo.CurrentCulture)
        End Function

        ''' <summary>
        ''' 実数を 3 桁区切りの文字列にします。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>文字列。未設定のときは空文字。</returns>
        Private Function ToNumberText(value As Double?) As String
            ' 値が無いとき
            If Not value.HasValue Then
                Return String.Empty
            End If
            Return value.Value.ToString(NumberFormat, CultureInfo.CurrentCulture)
        End Function
#End Region

    End Class
End Namespace
