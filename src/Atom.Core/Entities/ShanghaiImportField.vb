Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic

Namespace Entities

    ''' <summary>
    ''' 取込結果確認画面で編集できる項目の値の種類です。
    ''' </summary>
    Public Enum ShanghaiImportValueKind

        ''' <summary>文字列。</summary>
        Text = 0

        ''' <summary>整数。小数は保持できません。</summary>
        WholeNumber = 1

        ''' <summary>小数を含む数値。</summary>
        DecimalNumber = 2

        ''' <summary>日付。</summary>
        DateValue = 3
    End Enum

    ''' <summary>
    ''' 取込結果確認画面で編集できる項目 1 つ分の定義です。
    ''' </summary>
    ''' <remarks>
    ''' 更新する列名は必ずこの定義から取り出します。画面から受け取った文字列を
    ''' そのまま SQL へ埋め込まないためのホワイトリストとして機能します。
    ''' </remarks>
    Public NotInheritable Class ShanghaiImportField

#Region "コンストラクタ"
        ''' <summary>
        ''' 項目の定義を初期化します。
        ''' </summary>
        ''' <param name="key">画面上の入力欄の ID。</param>
        ''' <param name="columnName">更新するデータベースの列名。</param>
        ''' <param name="displayName">エラーメッセージに使う項目名。</param>
        ''' <param name="kind">値の種類。</param>
        ''' <param name="maxLength">文字列のときの桁数。文字列以外では 0。</param>
        Public Sub New(key As String, columnName As String, displayName As String,
                       kind As ShanghaiImportValueKind, maxLength As Integer)
            Me.Key = key
            Me.ColumnName = columnName
            Me.DisplayName = displayName
            Me.Kind = kind
            Me.MaxLength = maxLength
        End Sub
#End Region

#Region "プロパティ"
        ''' <summary>画面上の入力欄の ID。</summary>
        Public ReadOnly Property Key As String

        ''' <summary>更新するデータベースの列名。</summary>
        Public ReadOnly Property ColumnName As String

        ''' <summary>エラーメッセージに使う項目名。</summary>
        Public ReadOnly Property DisplayName As String

        ''' <summary>値の種類。</summary>
        Public ReadOnly Property Kind As ShanghaiImportValueKind

        ''' <summary>文字列のときの桁数。</summary>
        Public ReadOnly Property MaxLength As Integer
#End Region

    End Class

    ''' <summary>
    ''' 取込結果確認画面で編集できる項目の一覧です。
    ''' </summary>
    ''' <remarks>
    ''' ERR_CONTENTS はチェック処理が書き込む欄のため、編集対象に含めません。
    ''' shanghai_code（品目コード）は INVOICE と PACKING の突合キーのため、両方に登録しています。
    ''' 片方だけを直すと View_kaigai_shiire_jisseki_kbn1_invoice_gokei の内部結合が外れます。
    ''' また shanghai_code を更新すると R3_hinmoku_code も併せて更新する必要があるため、
    ''' 更新処理は ShanghaiImportRepository の専用メソッドが担当します。
    ''' </remarks>
    Public NotInheritable Class ShanghaiImportFieldCatalog

#Region "フィールド"
        ''' <summary>INVOICE の編集項目。</summary>
        Private Shared ReadOnly _invoiceFields As Dictionary(Of String, ShanghaiImportField) = CreateInvoiceFields()

        ''' <summary>PACKING の編集項目。</summary>
        Private Shared ReadOnly _packingFields As Dictionary(Of String, ShanghaiImportField) = CreatePackingFields()
#End Region

#Region "コンストラクタ"
        ''' <summary>インスタンス化を禁止します。</summary>
        Private Sub New()
        End Sub
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' INVOICE の編集項目を取得します。
        ''' </summary>
        ''' <param name="key">入力欄の ID。</param>
        ''' <returns>項目の定義。定義に無いときは Nothing。</returns>
        Public Shared Function FindInvoiceField(key As String) As ShanghaiImportField
            Return Find(_invoiceFields, key)
        End Function

        ''' <summary>
        ''' PACKING の編集項目を取得します。
        ''' </summary>
        ''' <param name="key">入力欄の ID。</param>
        ''' <returns>項目の定義。定義に無いときは Nothing。</returns>
        Public Shared Function FindPackingField(key As String) As ShanghaiImportField
            Return Find(_packingFields, key)
        End Function
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 一覧から項目の定義を探します。
        ''' </summary>
        ''' <param name="fields">探す対象の一覧。</param>
        ''' <param name="key">入力欄の ID。</param>
        ''' <returns>項目の定義。定義に無いときは Nothing。</returns>
        Private Shared Function Find(fields As Dictionary(Of String, ShanghaiImportField),
                                     key As String) As ShanghaiImportField
            ' ID が無いとき
            If String.IsNullOrEmpty(key) Then
                Return Nothing
            End If

            Dim found As ShanghaiImportField = Nothing
            ' 定義があるとき
            If fields.TryGetValue(key, found) Then
                Return found
            End If
            ' 定義に無いとき
            Return Nothing
        End Function

        ''' <summary>INVOICE の編集項目を作成します。</summary>
        ''' <returns>入力欄の ID をキーにした一覧。</returns>
        Private Shared Function CreateInvoiceFields() As Dictionary(Of String, ShanghaiImportField)
            Dim fields As New List(Of ShanghaiImportField)() From {
                New ShanghaiImportField("txtSuryo", "suryo", "数量", ShanghaiImportValueKind.WholeNumber, 0),
                New ShanghaiImportField("txtSuryoTani", "suryo_tani", "単位", ShanghaiImportValueKind.Text, 4),
                New ShanghaiImportField("txtTsuka", "tsuka", "通貨", ShanghaiImportValueKind.Text, 3),
                New ShanghaiImportField("txtUnitPrice", "unit_price", "単価", ShanghaiImportValueKind.DecimalNumber, 0),
                New ShanghaiImportField("txtAmount", "amount", "金額", ShanghaiImportValueKind.DecimalNumber, 0),
                New ShanghaiImportField("txtKoubaiDenpyoNo", "R3_koubai_denpyo_no", "購買伝票#", ShanghaiImportValueKind.Text, 50),
                New ShanghaiImportField("txtInvoiceNo", "invoice_no", "INVOICE#", ShanghaiImportValueKind.Text, 50),
                New ShanghaiImportField("txtBlDate", "bl_date", "B/L DATE", ShanghaiImportValueKind.DateValue, 0),
                New ShanghaiImportField("txtSyukkaHoho", "syukka_hoho", "出荷方法", ShanghaiImportValueKind.Text, 50),
                New ShanghaiImportField("txtShanghaiCode", "shanghai_code", "品目コード", ShanghaiImportValueKind.Text, 50),
                New ShanghaiImportField("txtShanghaiCodeText", "shanghai_code_text", "品名", ShanghaiImportValueKind.Text, 128)
            }
            Return ToDictionary(fields)
        End Function

        ''' <summary>PACKING の編集項目を作成します。</summary>
        ''' <returns>入力欄の ID をキーにした一覧。</returns>
        Private Shared Function CreatePackingFields() As Dictionary(Of String, ShanghaiImportField)
            Dim fields As New List(Of ShanghaiImportField)() From {
                New ShanghaiImportField("txtInvoiceNo", "invoice_no", "INVOICE_NO", ShanghaiImportValueKind.Text, 50),
                New ShanghaiImportField("txtTajimaPoNo", "tajima_po_no", "PO", ShanghaiImportValueKind.Text, 50),
                New ShanghaiImportField("txtTrueOrFalse", "true_or_false", "T_F", ShanghaiImportValueKind.Text, 10),
                New ShanghaiImportField("txtShanghaiCode", "shanghai_code", "品目コード", ShanghaiImportValueKind.Text, 50),
                New ShanghaiImportField("txtSyukkaHoho", "syukka_hoho", "出荷方法", ShanghaiImportValueKind.Text, 50),
                New ShanghaiImportField("txtBlDate", "bl_date", "BL_DATE", ShanghaiImportValueKind.DateValue, 0),
                New ShanghaiImportField("txtSuryo", "suryo", "数量", ShanghaiImportValueKind.WholeNumber, 0),
                New ShanghaiImportField("txtNetWeight", "net_weight", "NET", ShanghaiImportValueKind.DecimalNumber, 0),
                New ShanghaiImportField("txtGrossWeight", "gross_weight", "GRS", ShanghaiImportValueKind.DecimalNumber, 0),
                New ShanghaiImportField("txtM3", "m3", "M3", ShanghaiImportValueKind.DecimalNumber, 0),
                New ShanghaiImportField("txtCartonNoFrom", "carton_no_from", "C/NO 開始", ShanghaiImportValueKind.Text, 50),
                New ShanghaiImportField("txtCartonNoTo", "carton_no_to", "C/NO 終了", ShanghaiImportValueKind.Text, 50),
                New ShanghaiImportField("txtUnitPrice", "unit_price", "単価", ShanghaiImportValueKind.DecimalNumber, 0),
                New ShanghaiImportField("txtAmount", "amount", "合計", ShanghaiImportValueKind.DecimalNumber, 0)
            }
            Return ToDictionary(fields)
        End Function

        ''' <summary>一覧を辞書へ変換します。</summary>
        ''' <param name="fields">変換する一覧。</param>
        ''' <returns>入力欄の ID をキーにした辞書。</returns>
        Private Shared Function ToDictionary(fields As List(Of ShanghaiImportField)) As Dictionary(Of String, ShanghaiImportField)
            Dim results As New Dictionary(Of String, ShanghaiImportField)(StringComparer.Ordinal)
            For Each field As ShanghaiImportField In fields
                results.Add(field.Key, field)
            Next
            Return results
        End Function
#End Region

    End Class
End Namespace
