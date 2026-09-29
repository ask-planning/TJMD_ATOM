Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports System.IO
Imports Atom.Core.Entities
Imports Atom.Core.Exceptions
Imports Atom.Core.Logging
Imports Atom.Data.Services
Imports Atom.Web.Common

Namespace Pages.Entry

    ''' <summary>
    ''' INVOICE / PACKING 取込画面（menuNo=101）です。
    ''' 旧 Access 版の F_1_上海輸入_INVOICE取込 に相当します。
    ''' </summary>
    ''' <remarks>
    ''' 処理の途中で利用者への確認が 4 回入るため、確認のたびにポストバックし、
    ''' どこから再開するかを ViewState に持たせて処理を進めます。
    ''' 業務処理は Atom.Data.Services.ShanghaiImportService が担います。
    ''' </remarks>
    Public Class InvoiceImport
        Inherits BasePage

#Region "定数"
        ''' <summary>ログの発生元名。</summary>
        Private Const LogSourceName As String = "InvoiceImport"

        ''' <summary>この画面のメニュー番号。</summary>
        Private Const ThisMenuNo As Integer = 101

        ''' <summary>INVOICE 内容確認画面のメニュー番号。</summary>
        Private Const InvoiceConfirmMenuNo As Integer = 102

        ''' <summary>INVOICE ヘッダー入力画面のメニュー番号。</summary>
        Private Const InvoiceHeaderMenuNo As Integer = 103

        ''' <summary>INVOICE ファイル名の末尾。</summary>
        Private Const InvoiceSuffix As String = "_INV"

        ''' <summary>PACKING ファイル名の末尾。</summary>
        Private Const PackingSuffix As String = "_PAC"

        ''' <summary>通貨の整合を判定するファイル名の先頭文字数。</summary>
        Private Const CurrencyKeyLength As Integer = 5

        ''' <summary>取込対象として認める拡張子。</summary>
        Private Shared ReadOnly AllowedExtensions As String() = {".xls", ".xlsx"}

        ''' <summary>ViewState のキー。</summary>
        Private Const RefNoKey As String = "RefNo"
        Private Const InvoiceNoMainKey As String = "InvoiceNoMain"
        Private Const FutureCountKey As String = "FutureCount"
        Private Const RegisteredKey As String = "Registered"
        Private Const IsReplaceKey As String = "IsReplace"
        Private Const ConfirmKindKey As String = "ConfirmKind"

        ''' <summary>メッセージ。</summary>
        Private Const NoChumonsakiMessage As String = "発注先を選択してください。"
        Private Const NoTsukaMessage As String = "通貨を選択してください。"
        Private Const NoInvoiceFileMessage As String = "INVOICEファイルを選択してください。"
        Private Const NoPackingFileMessage As String = "PACKINGファイルを選択してください。"
        Private Const InvalidExtensionMessage As String = "Excelファイル（.xls / .xlsx）を選択してください。"
        Private Const FileNameMismatchMessage As String = "INVOICEとPACKINGのファイル名が一致しません。"
        Private Const CurrencyMismatchMessage As String = "INVOICEと通貨が一致しません。"
        Private Const NoTempDataMessage As String = "取込データがありません。先に取り込みを行ってください。"
        Private Const PoDuplicatedMessage As String = "複数INVOICEに分かれているPOがあります。POを変更し、再度実行してください。"
        Private Const HasErrorMessage As String = "エラーデータがあります。EXCELを確認してください。"
        Private Const RegisterFailedMessage As String = "INVOICE、PACKINGの取込に失敗しました。EXCELデータをご確認ください。"
        Private Const CompletedMessage As String = "処理終了しました。"
        Private Const DatabaseErrorMessage As String = "処理に失敗しました。時間をおいて再度お試しください。"
        Private Const NotImplementedMessage As String = "現在未実装です。"

        ''' <summary>確認ダイアログの文言。</summary>
        Private Const FutureConfirmMessage As String = "翌月以降のBL DATEのINVOICEです。取り込みを中断しますか？"
        Private Const RegisteredConfirmMessage As String = "登録済みのINVOICEです。どのように処理しますか？"
        Private Const WarningConfirmMessage As String = "WARNINGデータがあります。このまま取り込みますか？"
        Private Const NormalConfirmMessage As String = "内容の確認を行いますか？"
#End Region

#Region "列挙型"
        ''' <summary>処理の再開位置です。</summary>
        Private Enum FlowStep
            Analyze = 0
            FutureBlDate = 1
            Registered = 2
            Checks = 3
            Register = 4
        End Enum

        ''' <summary>表示中の確認ダイアログの種類です。</summary>
        Private Enum ConfirmKind
            None = 0
            FutureBlDate = 1
            Registered = 2
            Warning = 3
            Normal = 4
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
            Me.HeaderActions.Add(New HeaderAction("INVヘッダ", InvoiceHeaderMenuNo))
            Me.HeaderActions.Add(New HeaderAction("INV確認", InvoiceConfirmMenuNo))
            Me.HeaderActions.Add(New HeaderAction("MENUへ", HeaderAction.MenuScreenNo))
            MyBase.OnInit(e)
        End Sub

        ''' <summary>
        ''' 読み込み時の処理です。初回表示のみ選択肢を用意します。
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
                Me.BindChumonsaki()
                Me.SetDefaultShanghaiChotatsu()
                Me.txtRefNo.Text = String.Empty
            End If
        End Sub

        ''' <summary>
        ''' 取り込みボタン押下時の処理です。Excel を一時テーブルへ取り込み、続けて再実行処理を行います。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
            Me.HideConfirm()

            ' 入力に誤りがあるとき
            If Not Me.ValidateForImport() Then
                Return
            End If

            Try
                Me.LoadUploadedFiles()
            Catch ex As ImportFormatException
                Me.ShowMessage(ex.Message, True)
                Return
            Catch ex As SqlException
                ' 詳細は Service 側でログへ記録済みです。
                Me.ShowMessage(DatabaseErrorMessage, True)
                Return
            End Try

            Me.RunFlow(FlowStep.Analyze)
        End Sub

        ''' <summary>
        ''' 再実行ボタン押下時の処理です。一時データに対してチェックと本登録を行います。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnRerun_Click(sender As Object, e As EventArgs) Handles btnRerun.Click
            Me.HideConfirm()

            ' 入力に誤りがあるとき
            If Not Me.ValidateForRerun() Then
                Return
            End If

            Me.RunFlow(FlowStep.Analyze)
        End Sub

        ''' <summary>
        ''' 取込内容確認ボタン押下時の処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnKakunin_Click(sender As Object, e As EventArgs) Handles btnKakunin.Click
            Me.HideConfirm()
            Me.NavigateToKakunin()
        End Sub

        ''' <summary>
        ''' 確認ダイアログの 1 つ目のボタン押下時の処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnConfirm1_Click(sender As Object, e As EventArgs) Handles btnConfirm1.Click
            Dim kind As ConfirmKind = Me.CurrentConfirmKind
            Me.HideConfirm()

            Select Case kind
                ' 翌月以降の確認で「中断する」を選んだとき
                Case ConfirmKind.FutureBlDate
                    Return
                ' 登録済みの確認で「置換」を選んだとき
                Case ConfirmKind.Registered
                    Me.IsReplace = True
                    Me.RunFlow(FlowStep.Checks)
                ' 警告の確認で「このまま取り込む」を選んだとき
                Case ConfirmKind.Warning
                    Me.RunFlow(FlowStep.Register)
                ' 正常時の確認で「取込内容を確認」を選んだとき
                Case ConfirmKind.Normal
                    Me.NavigateToKakunin()
            End Select
        End Sub

        ''' <summary>
        ''' 確認ダイアログの 2 つ目のボタン押下時の処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnConfirm2_Click(sender As Object, e As EventArgs) Handles btnConfirm2.Click
            Dim kind As ConfirmKind = Me.CurrentConfirmKind
            Me.HideConfirm()

            Select Case kind
                ' 翌月以降の確認で「続行する」を選んだとき
                Case ConfirmKind.FutureBlDate
                    Me.RunFlow(FlowStep.Registered)
                ' 登録済みの確認で「追加」を選んだとき
                Case ConfirmKind.Registered
                    Me.IsReplace = False
                    Me.RunFlow(FlowStep.Checks)
                ' 警告の確認で「キャンセル」を選んだとき
                Case ConfirmKind.Warning
                    Me.NavigateToKakunin()
                ' 正常時の確認で「このまま取り込む」を選んだとき
                Case ConfirmKind.Normal
                    Me.RunFlow(FlowStep.Register)
            End Select
        End Sub

        ''' <summary>
        ''' 確認ダイアログの 3 つ目のボタン押下時の処理です。登録済みの確認のキャンセルに使います。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnConfirm3_Click(sender As Object, e As EventArgs) Handles btnConfirm3.Click
            Me.HideConfirm()
        End Sub
#End Region

#Region "内部処理（画面の状態）"
        ''' <summary>表示中の確認ダイアログの種類。</summary>
        Private ReadOnly Property CurrentConfirmKind As ConfirmKind
            Get
                Dim stored As Object = Me.ViewState(ConfirmKindKey)
                ' 保持していないとき
                If stored Is Nothing Then
                    Return ConfirmKind.None
                End If
                Return CType(stored, ConfirmKind)
            End Get
        End Property

        ''' <summary>登録済みデータを置き換えるかどうか。</summary>
        Private Property IsReplace As Boolean
            Get
                Dim stored As Object = Me.ViewState(IsReplaceKey)
                ' 保持していないとき
                If stored Is Nothing Then
                    Return False
                End If
                Return CBool(stored)
            End Get
            Set(value As Boolean)
                Me.ViewState(IsReplaceKey) = value
            End Set
        End Property

        ''' <summary>INVOICE 番号（メイン）。</summary>
        Private ReadOnly Property InvoiceNoMain As String
            Get
                Dim stored As Object = Me.ViewState(InvoiceNoMainKey)
                ' 保持していないとき
                If stored Is Nothing Then
                    Return String.Empty
                End If
                Return CStr(stored)
            End Get
        End Property
#End Region

#Region "内部処理（処理の流れ）"
        ''' <summary>
        ''' 取込の一連の処理を、指定した位置から実行します。
        ''' </summary>
        ''' <param name="startStep">再開する位置。</param>
        Private Sub RunFlow(startStep As FlowStep)
            Try
                ' 一時データの確認から始めるとき
                If startStep <= FlowStep.Analyze Then
                    ' 一時データが無いとき
                    If Not Me.AnalyzeTempData() Then
                        Return
                    End If
                End If

                ' 翌月以降の BL DATE の確認
                If startStep <= FlowStep.FutureBlDate Then
                    If CInt(Me.ViewState(FutureCountKey)) > 0 Then
                        Me.ShowConfirm(ConfirmKind.FutureBlDate, FutureConfirmMessage,
                                       "中断する", "続行する", String.Empty)
                        Return
                    End If
                End If

                ' 登録済みかどうかの確認
                If startStep <= FlowStep.Registered Then
                    If CBool(Me.ViewState(RegisteredKey)) Then
                        Me.ShowConfirm(ConfirmKind.Registered, RegisteredConfirmMessage,
                                       "置換する", "追加する", "キャンセル")
                        Return
                    End If
                End If

                ' 業務チェック
                If startStep <= FlowStep.Checks Then
                    ' チェックの結果、先へ進めないとき
                    If Not Me.RunChecks() Then
                        Return
                    End If
                End If

                ' 本登録
                Me.RunRegister()
            Catch ex As SqlException
                ' 詳細は Service 側でログへ記録済みです。
                Me.ShowMessage(DatabaseErrorMessage, True)
            End Try
        End Sub

        ''' <summary>
        ''' 一時データを読み取り、REF_NO と確認に使う情報を求めます。
        ''' </summary>
        ''' <returns>一時データがあり、処理を続けられるとき True。</returns>
        Private Function AnalyzeTempData() As Boolean
            Dim analysis As ShanghaiImportAnalysis =
                Me._service.Analyze(Me.CurrentLoginCode, Me.SelectedShanghaiChotatsu)

            ' 一時データが無いとき
            If Not analysis.HasTempData Then
                Me.ShowMessage(NoTempDataMessage, True)
                Return False
            End If

            Me.ViewState(RefNoKey) = analysis.RefNo
            Me.ViewState(InvoiceNoMainKey) = analysis.InvoiceNoMain
            Me.ViewState(FutureCountKey) = analysis.FutureBlDateCount
            Me.ViewState(RegisteredKey) = analysis.Registered.HasAny
            Me.txtRefNo.Text = analysis.RefNo
            Return True
        End Function

        ''' <summary>
        ''' 業務チェックを行い、結果に応じて確認ダイアログや画面遷移を行います。
        ''' </summary>
        ''' <returns>そのまま本登録へ進めるとき True。</returns>
        Private Function RunChecks() As Boolean
            Dim result As ImportCheckResult =
                Me._service.RunChecks(Me.CurrentLoginCode, Me.SelectedShanghaiChotatsu)

            ' 1 つの PO が複数 INVOICE に分かれているとき
            If result.IsPoDuplicated Then
                Me.ShowMessage(PoDuplicatedMessage, True)
                Me.NavigateToKakunin()
                Return False
            End If

            Select Case result.Level
                ' エラーがあるとき
                Case ImportCheckLevel.Error
                    Me.ShowMessage(HasErrorMessage, True)
                    Me.NavigateToKakunin()
                    Return False
                ' 警告だけのとき
                Case ImportCheckLevel.Warning
                    Me.ShowConfirm(ConfirmKind.Warning, WarningConfirmMessage,
                                   "このまま取り込む", "キャンセル", String.Empty)
                    Return False
                ' どちらも無いとき
                Case Else
                    Me.ShowConfirm(ConfirmKind.Normal, NormalConfirmMessage,
                                   "取込内容を確認", "このまま取り込む", String.Empty)
                    Return False
            End Select
        End Function

        ''' <summary>
        ''' 本登録を行います。
        ''' </summary>
        Private Sub RunRegister()
            Dim counts As ShanghaiImportCounts = Me._service.Register(
                Me.CurrentLoginCode, Me.SelectedShanghaiChotatsu, Me.SelectedChumonsaki,
                Me.SelectedTsuka, Me.InvoiceNoMain, Me.IsReplace)

            ' 明細が登録されなかったとき
            If counts.InvoiceCount = 0 OrElse counts.PackingCount = 0 Then
                Me.ShowMessage(RegisterFailedMessage, True)
                Return
            End If

            Me.ShowMessage(CompletedMessage, False)
        End Sub

        ''' <summary>
        ''' 取込内容確認画面へ遷移します。
        ''' </summary>
        ''' <remarks>
        ''' 取込内容確認画面はメニューに載らない画面のため、画面定義にメニュー番号がありません。
        ''' 実装されるまでは未実装の案内だけを表示します。
        ''' </remarks>
        Private Sub NavigateToKakunin()
            Me.ShowMessage(NotImplementedMessage, False)
        End Sub
#End Region

#Region "内部処理（入力チェック）"
        ''' <summary>
        ''' 取り込みに必要な入力を検証します。
        ''' </summary>
        ''' <returns>すべて満たしているとき True。</returns>
        Private Function ValidateForImport() As Boolean
            ' 発注先と通貨のチェック
            If Not Me.ValidateForRerun() Then
                Return False
            End If

            ' INVOICE ファイルが未選択のとき
            If Not Me.fuInvoice.HasFile Then
                Me.ShowMessage(NoInvoiceFileMessage, True)
                Return False
            End If
            ' PACKING ファイルが未選択のとき
            If Not Me.fuPacking.HasFile Then
                Me.ShowMessage(NoPackingFileMessage, True)
                Return False
            End If

            Dim invoiceFileName As String = Me.fuInvoice.FileName
            Dim packingFileName As String = Me.fuPacking.FileName

            ' 拡張子が対象外のとき
            If Not Me.IsAllowedExtension(invoiceFileName) OrElse Not Me.IsAllowedExtension(packingFileName) Then
                Me.ShowMessage(InvalidExtensionMessage, True)
                Return False
            End If

            ' ベース名が一致しないとき
            Dim invoiceBase As String = Me.GetBaseName(invoiceFileName, InvoiceSuffix)
            Dim packingBase As String = Me.GetBaseName(packingFileName, PackingSuffix)
            If Not String.Equals(invoiceBase, packingBase, StringComparison.OrdinalIgnoreCase) Then
                Me.ShowMessage(FileNameMismatchMessage, True)
                Return False
            End If

            ' ファイル名から判定できる通貨と選択した通貨が違うとき
            If Not Me.IsCurrencyMatched(invoiceBase) Then
                Me.ShowMessage(CurrencyMismatchMessage, True)
                Return False
            End If

            Return True
        End Function

        ''' <summary>
        ''' 再実行に必要な入力を検証します。
        ''' </summary>
        ''' <returns>すべて満たしているとき True。</returns>
        Private Function ValidateForRerun() As Boolean
            ' 発注先が未選択のとき
            If Me.SelectedChumonsaki <= 0 Then
                Me.ShowMessage(NoChumonsakiMessage, True)
                Return False
            End If
            ' 通貨が未選択のとき
            If Not Tsuka.IsValid(Me.SelectedTsuka) Then
                Me.ShowMessage(NoTsukaMessage, True)
                Return False
            End If
            Return True
        End Function

        ''' <summary>
        ''' 取込対象として認める拡張子かどうかを返します。
        ''' </summary>
        ''' <param name="fileName">判定するファイル名。</param>
        ''' <returns>認める拡張子のとき True。</returns>
        Private Function IsAllowedExtension(fileName As String) As Boolean
            Dim extension As String = Path.GetExtension(fileName)
            For Each allowed As String In AllowedExtensions
                ' 一致したとき
                If String.Equals(extension, allowed, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' ファイル名から拡張子と末尾の識別子を除いた名称を取り出します。
        ''' </summary>
        ''' <param name="fileName">対象のファイル名。</param>
        ''' <param name="suffix">除く末尾の識別子。_INV または _PAC。</param>
        ''' <returns>ベースとなる名称。</returns>
        Private Function GetBaseName(fileName As String, suffix As String) As String
            Dim name As String = Path.GetFileNameWithoutExtension(fileName)
            ' 末尾に識別子が付いているとき
            If name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) Then
                Return name.Substring(0, name.Length - suffix.Length)
            End If
            Return name
        End Function

        ''' <summary>
        ''' ファイル名の先頭から判定できる通貨と、選択した通貨が一致するかを返します。
        ''' </summary>
        ''' <param name="baseName">INVOICE ファイルのベース名称。</param>
        ''' <returns>一致するとき、または判定対象外のとき True。</returns>
        ''' <remarks>
        ''' 現行仕様どおり SJ-TJ / SJ-TR / YJ-TR だけを判定します。それ以外は判定しません。
        ''' </remarks>
        Private Function IsCurrencyMatched(baseName As String) As Boolean
            ' 判定に必要な桁数が無いとき
            If baseName.Length < CurrencyKeyLength Then
                Return True
            End If

            Dim key As String = baseName.Substring(0, CurrencyKeyLength).ToUpperInvariant()
            Select Case key
                ' 円建てのとき
                Case "SJ-TJ"
                    Return String.Equals(Me.SelectedTsuka, Tsuka.Jpy, StringComparison.Ordinal)
                ' 元建てのとき
                Case "SJ-TR", "YJ-TR"
                    Return String.Equals(Me.SelectedTsuka, Tsuka.Rmb, StringComparison.Ordinal)
                ' 判定対象外のとき
                Case Else
                    Return True
            End Select
        End Function
#End Region

#Region "内部処理（入力値の取得）"
        ''' <summary>選択されている発注先区分。未選択のときは 0。</summary>
        Private ReadOnly Property SelectedChumonsaki As Integer
            Get
                Dim value As Integer = 0
                ' 選択されていないとき
                If Me.rblChumonsaki.SelectedIndex < 0 Then
                    Return 0
                End If
                ' 数値として読めないとき
                If Not Integer.TryParse(Me.rblChumonsaki.SelectedValue, value) Then
                    Return 0
                End If
                Return value
            End Get
        End Property

        ''' <summary>選択されている通貨。未選択のときは空文字。</summary>
        Private ReadOnly Property SelectedTsuka As String
            Get
                ' 選択されていないとき
                If Me.rblTsuka.SelectedIndex < 0 Then
                    Return String.Empty
                End If
                Return Me.rblTsuka.SelectedValue
            End Get
        End Property

        ''' <summary>選択されている担当区分。1=上海 / 2=調達 / 3=貿易。</summary>
        Private ReadOnly Property SelectedShanghaiChotatsu As Integer
            Get
                Dim value As Integer = 1
                ' 数値として読めないとき
                If Not Integer.TryParse(Me.ddlShanghaiChotatsu.SelectedValue, value) Then
                    Return 1
                End If
                Return value
            End Get
        End Property

        ''' <summary>操作している利用者のログインコード。</summary>
        Private ReadOnly Property CurrentLoginCode As String
            Get
                Return SessionContext.Current.LoginCode
            End Get
        End Property
#End Region

#Region "内部処理（画面の初期化と表示）"
        ''' <summary>
        ''' 発注先の選択肢を割り当てます。
        ''' </summary>
        Private Sub BindChumonsaki()
            Dim items As IList(Of Koumoku)
            Try
                items = Me._service.SelectChumonsakiList()
            Catch ex As SqlException
                ' 詳細は Service 側でログへ記録済みです。
                Me.ShowMessage(DatabaseErrorMessage, True)
                Return
            End Try

            Dim targets As New List(Of Koumoku)()
            For Each item As Koumoku In items
                ' 「全て」は取込画面では使わない
                If item.SeqNo > 0 Then
                    targets.Add(item)
                End If
            Next

            Me.rblChumonsaki.DataSource = targets
            Me.rblChumonsaki.DataTextField = "Contents2"
            Me.rblChumonsaki.DataValueField = "SeqNo"
            Me.rblChumonsaki.DataBind()
        End Sub

        ''' <summary>
        ''' 担当区分の初期値を、ログイン情報から設定します。
        ''' </summary>
        Private Sub SetDefaultShanghaiChotatsu()
            Dim stored As Integer? = SessionContext.Current.ShanghaiChotatsu
            ' ログイン情報に値が無いとき
            If Not stored.HasValue Then
                Return
            End If

            Dim item As ListItem = Me.ddlShanghaiChotatsu.Items.FindByValue(stored.Value.ToString())
            ' 選択肢に無い値のとき
            If item Is Nothing Then
                Return
            End If
            Me.ddlShanghaiChotatsu.SelectedValue = item.Value
        End Sub

        ''' <summary>
        ''' 確認ダイアログを表示します。
        ''' </summary>
        ''' <param name="kind">確認の種類。</param>
        ''' <param name="message">表示する文言。</param>
        ''' <param name="text1">1 つ目のボタンの文言。</param>
        ''' <param name="text2">2 つ目のボタンの文言。</param>
        ''' <param name="text3">3 つ目のボタンの文言。使わないときは空文字。</param>
        Private Sub ShowConfirm(kind As ConfirmKind, message As String,
                                text1 As String, text2 As String, text3 As String)
            Me.ViewState(ConfirmKindKey) = kind
            Me.litConfirmMessage.Text = message
            Me.btnConfirm1.Text = text1
            Me.btnConfirm2.Text = text2
            Me.btnConfirm3.Text = text3
            Me.btnConfirm3.Visible = text3.Length > 0
            Me.pnlConfirm.Visible = True
        End Sub

        ''' <summary>
        ''' 確認ダイアログを閉じます。
        ''' </summary>
        Private Sub HideConfirm()
            Me.ViewState(ConfirmKindKey) = ConfirmKind.None
            Me.pnlConfirm.Visible = False
        End Sub
#End Region

#Region "内部処理（ファイルの取り扱い）"
        ''' <summary>
        ''' アップロードされた 2 つの Excel を一時フォルダへ保存し、解析して一時テーブルへ登録します。
        ''' </summary>
        ''' <remarks>
        ''' 同名ファイルが同時に取り込まれても衝突しないよう、保存名に利用者と時刻を含めます。
        ''' 解析後は保存したファイルを削除します。
        ''' </remarks>
        Private Sub LoadUploadedFiles()
            Dim folderPath As String = AppSettings.UploadFolderPath
            ' 保存先が無いとき（初回利用時に作成する）
            If Not Directory.Exists(folderPath) Then
                Directory.CreateDirectory(folderPath)
            End If

            Dim invoicePath As String = Me.SaveUploadedFile(Me.fuInvoice, folderPath, InvoiceSuffix)
            Dim packingPath As String = String.Empty

            Try
                packingPath = Me.SaveUploadedFile(Me.fuPacking, folderPath, PackingSuffix)

                Using invoiceStream As FileStream = File.OpenRead(invoicePath)
                    Using packingStream As FileStream = File.OpenRead(packingPath)
                        Me._service.LoadToTemp(Me.CurrentLoginCode, Me.SelectedTsuka,
                                               invoiceStream, packingStream)
                    End Using
                End Using
            Finally
                Me.DeleteTemporaryFile(invoicePath)
                Me.DeleteTemporaryFile(packingPath)
            End Try
        End Sub

        ''' <summary>
        ''' アップロードされたファイルを一時フォルダへ保存します。
        ''' </summary>
        ''' <param name="upload">アップロードのコントロール。</param>
        ''' <param name="folderPath">保存先フォルダ。</param>
        ''' <param name="suffix">保存名に付ける識別子。</param>
        ''' <returns>保存したファイルのパス。</returns>
        Private Function SaveUploadedFile(upload As FileUpload, folderPath As String, suffix As String) As String
            Dim extension As String = Path.GetExtension(upload.FileName)
            Dim fileName As String = String.Concat(Me.CurrentLoginCode, "_",
                                                   Date.Now.ToString("yyyyMMddHHmmssfff"), "_",
                                                   Guid.NewGuid().ToString("N"), suffix, extension)
            Dim filePath As String = Path.Combine(folderPath, fileName)
            upload.SaveAs(filePath)
            Return filePath
        End Function

        ''' <summary>
        ''' 一時ファイルを削除します。削除できなくても業務処理は止めません。
        ''' </summary>
        ''' <param name="filePath">削除するファイルのパス。</param>
        Private Sub DeleteTemporaryFile(filePath As String)
            ' パスが無いとき
            If String.IsNullOrEmpty(filePath) Then
                Return
            End If

            Try
                ' ファイルがあるとき
                If File.Exists(filePath) Then
                    File.Delete(filePath)
                End If
            Catch ex As IOException
                Logger.WriteWarning(LogSourceName, "一時ファイルの削除に失敗しました。", Me.CurrentLoginCode)
            Catch ex As UnauthorizedAccessException
                Logger.WriteWarning(LogSourceName, "一時ファイルの削除が許可されていません。", Me.CurrentLoginCode)
            End Try
        End Sub
#End Region

    End Class
End Namespace
