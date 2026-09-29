Option Strict On
Option Explicit On
Option Infer Off

Namespace Entities

    ''' <summary>
    ''' INVOICE の Excel 1 行分です。tbl_t_shanghai_invoice_import_check_temp へ登録します。
    ''' </summary>
    ''' <remarks>
    ''' Excel の A〜R 列に対応します。列の対応は docs/specs/101_INVOICE・PACKING取込.md を参照してください。
    ''' </remarks>
    Public NotInheritable Class ShanghaiInvoiceImportRow

#Region "プロパティ"
        ''' <summary>一時テーブルの連番。行を特定するために使います。</summary>
        Public Property SeqNo As Integer

        ''' <summary>A 列 项号。行番号。</summary>
        Public Property LineSeq As Double?

        ''' <summary>B 列 订单单号。受注番号。</summary>
        Public Property SoNo As String = String.Empty

        ''' <summary>C 列 订单行号。受注明細番号。</summary>
        Public Property SoNoSeq As Integer?

        ''' <summary>D 列 产品编号。上海コード。数値のときは 5 桁ゼロ埋め。</summary>
        Public Property ShanghaiCode As String = String.Empty

        ''' <summary>E 列 品名规格。品名。</summary>
        Public Property ShanghaiCodeText As String = String.Empty

        ''' <summary>F 列。規格。</summary>
        Public Property Kikaku As String = String.Empty

        ''' <summary>G 列 其他。商品群。2 文字。</summary>
        Public Property PartSort As String = String.Empty

        ''' <summary>H 列 销售单位。数量単位。</summary>
        Public Property SuryoTani As String = String.Empty

        ''' <summary>I 列 实际出货数量。数量。</summary>
        Public Property Suryo As Decimal?

        ''' <summary>J 列 原币单价。単価。</summary>
        Public Property UnitPrice As Double?

        ''' <summary>K 列 原币税前金额。金額。</summary>
        Public Property Amount As Double?

        ''' <summary>L 列 客户产品编号。</summary>
        Public Property DummyNo1 As String = String.Empty

        ''' <summary>M 列。</summary>
        Public Property DummyNo2 As String = String.Empty

        ''' <summary>N 列 客户订单单号。田島 PO 番号。</summary>
        Public Property TajimaPoNo As String = String.Empty

        ''' <summary>O 列 客户采购单号。R3 購買伝票番号。空のときは 9999999999。</summary>
        Public Property R3KoubaiDenpyoNo As String = String.Empty

        ''' <summary>P 列 发票号码。INVOICE 番号。</summary>
        Public Property InvoiceNo As String = String.Empty

        ''' <summary>Q 列 出货日期。BL DATE。</summary>
        Public Property BlDate As Date?

        ''' <summary>R 列 交运方式。出荷方法。SEA 等の文字列。</summary>
        Public Property SyukkaHoho As String = String.Empty

        ''' <summary>R3 品目コード。上海コードが数値のときは 18 桁ゼロ埋め。</summary>
        Public Property R3HinmokuCode As String = String.Empty

        ''' <summary>元の田島 PO 番号。N 列と同じ値。</summary>
        Public Property TajimaPoNoOriginal As String = String.Empty

        ''' <summary>金額の文字列表現。小数桁のチェックに使います。</summary>
        Public Property AmountText As String = String.Empty

        ''' <summary>通貨。画面で選択した値。</summary>
        Public Property Tsuka As String = String.Empty

        ''' <summary>チェック結果の内容。取込結果確認画面で表示します。</summary>
        Public Property ErrContents As String = String.Empty

        ''' <summary>チェック結果の区分。0=正常 / 1=エラー / 2=警告。</summary>
        Public Property ErrStatus As Integer
#End Region

    End Class

    ''' <summary>
    ''' PACKING LIST の Excel 1 行分です。tbl_t_shanghai_packing_import_check_temp へ登録します。
    ''' </summary>
    ''' <remarks>
    ''' Excel の A〜N 列に対応します。列の対応は docs/specs/101_INVOICE・PACKING取込.md を参照してください。
    ''' </remarks>
    Public NotInheritable Class ShanghaiPackingImportRow

#Region "プロパティ"
        ''' <summary>一時テーブルの連番。行を特定するために使います。</summary>
        Public Property SeqNo As Integer

        ''' <summary>A 列 T/F。</summary>
        Public Property TrueOrFalse As String = String.Empty

        ''' <summary>B 列 发票号码。INVOICE 番号。</summary>
        Public Property InvoiceNo As String = String.Empty

        ''' <summary>C 列 客户订单单号。田島 PO 番号。</summary>
        Public Property TajimaPoNo As String = String.Empty

        ''' <summary>D 列 产品编号。上海コード。</summary>
        Public Property ShanghaiCode As String = String.Empty

        ''' <summary>E 列 产品数量。数量。</summary>
        Public Property Suryo As Decimal?

        ''' <summary>F 列 原币单价。単価。</summary>
        Public Property UnitPrice As Double?

        ''' <summary>G 列 原币税前金额。金額。</summary>
        Public Property Amount As Double?

        ''' <summary>H 列 起始包装箱号。開始カートン番号。空のときは直前の値を引き継ぎます。</summary>
        Public Property CartonNoFrom As String = String.Empty

        ''' <summary>I 列 截止包装箱号。終了カートン番号。直前と同じ値のときは設定しません。</summary>
        Public Property CartonNoTo As String = String.Empty

        ''' <summary>J 列 总净重(Kg)。正味重量。</summary>
        Public Property NetWeight As Double?

        ''' <summary>K 列 总毛重(Kg)。総重量。</summary>
        Public Property GrossWeight As Double?

        ''' <summary>L 列 总材积(Cuft)。容積。見出しの単位は Cuft ですが列名は m3 です。</summary>
        Public Property M3 As Double?

        ''' <summary>M 列 出货日期。BL DATE。</summary>
        Public Property BlDate As Date?

        ''' <summary>N 列 交运方式。出荷方法。</summary>
        Public Property SyukkaHoho As String = String.Empty

        ''' <summary>R3 品目コード。上海コードが数値のときは 18 桁ゼロ埋め。</summary>
        Public Property R3HinmokuCode As String = String.Empty

        ''' <summary>元の田島 PO 番号。C 列と同じ値。</summary>
        Public Property TajimaPoNoOriginal As String = String.Empty

        ''' <summary>チェック結果の内容。取込結果確認画面で表示します。</summary>
        Public Property ErrContents As String = String.Empty

        ''' <summary>チェック結果の区分。0=正常 / 1=エラー / 2=警告。</summary>
        Public Property ErrStatus As Integer

        ''' <summary>
        ''' カートン番号の範囲です。取込結果確認画面の C/NO 列に表示します。
        ''' </summary>
        ''' <remarks>
        ''' 旧 Access 版は開始と終了を別のコントロールで並べていましたが、
        ''' 1 つの見出し（C/NO）にまとめられているため 1 列で表示します。
        ''' </remarks>
        Public ReadOnly Property CartonNoRange As String
            Get
                ' 開始も終了も無いとき
                If Me.CartonNoFrom.Length = 0 AndAlso Me.CartonNoTo.Length = 0 Then
                    Return String.Empty
                End If
                ' 終了が無いとき（1 箱）
                If Me.CartonNoTo.Length = 0 Then
                    Return Me.CartonNoFrom
                End If
                ' 開始が無いとき
                If Me.CartonNoFrom.Length = 0 Then
                    Return Me.CartonNoTo
                End If
                ' 開始と終了が同じとき（1 箱）
                If String.Equals(Me.CartonNoFrom, Me.CartonNoTo, StringComparison.Ordinal) Then
                    Return Me.CartonNoFrom
                End If
                Return Me.CartonNoFrom & "～" & Me.CartonNoTo
            End Get
        End Property
#End Region

    End Class
End Namespace
