Option Strict On
Option Explicit On
Option Infer Off

Namespace Sql

    ''' <summary>
    ''' 上海輸入 INVOICE / PACKING 取込（menuNo=101）で使う SQL 文を組み立てます。
    ''' </summary>
    ''' <remarks>
    ''' 業務チェックと本登録の本体はストアドプロシージャ側にあります。
    ''' ここに置くのは一時テーブルへの登録、置換時の削除、ヘッダ登録などの直接操作です。
    ''' </remarks>
    Public NotInheritable Class ShanghaiImportSql

#Region "コンストラクタ"
        ''' <summary>インスタンス化を禁止します。</summary>
        Private Sub New()
        End Sub
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' INVOICE のチェック用一時データを利用者単位で削除する SQL を作成します。
        ''' </summary>
        ''' <returns>DELETE 文。</returns>
        Public Shared Function CreateDeleteInvoiceCheckTempSql() As String
            Return "DELETE FROM dbo.tbl_t_shanghai_invoice_import_check_temp " &
                   "WHERE pk_login_code = @LoginCode"
        End Function

        ''' <summary>
        ''' PACKING のチェック用一時データを利用者単位で削除する SQL を作成します。
        ''' </summary>
        ''' <returns>DELETE 文。</returns>
        Public Shared Function CreateDeletePackingCheckTempSql() As String
            Return "DELETE FROM dbo.tbl_t_shanghai_packing_import_check_temp " &
                   "WHERE pk_login_code = @LoginCode"
        End Function

        ''' <summary>
        ''' INVOICE のチェック用一時データを 1 件登録する SQL を作成します。
        ''' </summary>
        ''' <returns>INSERT 文。</returns>
        Public Shared Function CreateInsertInvoiceCheckTempSql() As String
            Return "INSERT INTO dbo.tbl_t_shanghai_invoice_import_check_temp " &
                   "(pk_login_code, line_seq, so_no, so_no_seq, shanghai_code, shanghai_code_text, " &
                   "kikaku, part_sort, suryo_tani, suryo, unit_price, amount, dummy_no1, dummy_no2, " &
                   "tajima_po_no, R3_hinmoku_code, tsuka, tajima_po_no_original, invoice_no, amount_text, " &
                   "R3_koubai_denpyo_no, bl_date, syukka_hoho, err_status) " &
                   "VALUES (@LoginCode, @LineSeq, @SoNo, @SoNoSeq, @ShanghaiCode, @ShanghaiCodeText, " &
                   "@Kikaku, @PartSort, @SuryoTani, @Suryo, @UnitPrice, @Amount, @DummyNo1, @DummyNo2, " &
                   "@TajimaPoNo, @R3HinmokuCode, @Tsuka, @TajimaPoNoOriginal, @InvoiceNo, @AmountText, " &
                   "@R3KoubaiDenpyoNo, @BlDate, @SyukkaHoho, 0)"
        End Function

        ''' <summary>
        ''' PACKING のチェック用一時データを 1 件登録する SQL を作成します。
        ''' </summary>
        ''' <returns>INSERT 文。</returns>
        Public Shared Function CreateInsertPackingCheckTempSql() As String
            Return "INSERT INTO dbo.tbl_t_shanghai_packing_import_check_temp " &
                   "(pk_login_code, true_or_false, invoice_no, tajima_po_no, shanghai_code, suryo, " &
                   "unit_price, amount, carton_no_from, carton_no_to, net_weight, gross_weight, m3, " &
                   "bl_date, syukka_hoho, R3_hinmoku_code, tajima_po_no_original, err_status) " &
                   "VALUES (@LoginCode, @TrueOrFalse, @InvoiceNo, @TajimaPoNo, @ShanghaiCode, @Suryo, " &
                   "@UnitPrice, @Amount, @CartonNoFrom, @CartonNoTo, @NetWeight, @GrossWeight, @M3, " &
                   "@BlDate, @SyukkaHoho, @R3HinmokuCode, @TajimaPoNoOriginal, 0)"
        End Function

        ''' <summary>
        ''' INVOICE のチェック用一時データを 1 項目だけ更新する SQL を作成します。
        ''' </summary>
        ''' <param name="columnName">更新する列名。ShanghaiImportFieldCatalog から取得した値のみ渡します。</param>
        ''' <returns>UPDATE 文。</returns>
        ''' <remarks>
        ''' 列名だけは構文の一部なので連結しますが、値はパラメータで渡します。
        ''' 列名は呼び出し側でホワイトリスト照合済みのものに限ります。
        ''' </remarks>
        Public Shared Function CreateUpdateInvoiceCheckTempValueSql(columnName As String) As String
            Return "UPDATE dbo.tbl_t_shanghai_invoice_import_check_temp " &
                   "SET " & columnName & " = @Value " &
                   "WHERE pk_login_code = @LoginCode AND pk_seq_no = @SeqNo"
        End Function

        ''' <summary>
        ''' PACKING のチェック用一時データを 1 項目だけ更新する SQL を作成します。
        ''' </summary>
        ''' <param name="columnName">更新する列名。ShanghaiImportFieldCatalog から取得した値のみ渡します。</param>
        ''' <returns>UPDATE 文。</returns>
        Public Shared Function CreateUpdatePackingCheckTempValueSql(columnName As String) As String
            Return "UPDATE dbo.tbl_t_shanghai_packing_import_check_temp " &
                   "SET " & columnName & " = @Value " &
                   "WHERE pk_login_code = @LoginCode AND pk_seq_no = @SeqNo"
        End Function

        ''' <summary>
        ''' INVOICE のチェック用一時データの品目コードを更新する SQL を作成します。
        ''' </summary>
        ''' <returns>UPDATE 文。</returns>
        ''' <remarks>
        ''' 上海コードから R3 品目コードを作る決まりのため、2 列を同時に更新します。
        ''' 片方だけ更新すると本登録後の R3 連携がずれます。
        ''' </remarks>
        Public Shared Function CreateUpdateInvoiceShanghaiCodeSql() As String
            Return "UPDATE dbo.tbl_t_shanghai_invoice_import_check_temp " &
                   "SET shanghai_code = @ShanghaiCode, R3_hinmoku_code = @R3HinmokuCode " &
                   "WHERE pk_login_code = @LoginCode AND pk_seq_no = @SeqNo"
        End Function

        ''' <summary>
        ''' PACKING のチェック用一時データの品目コードを更新する SQL を作成します。
        ''' </summary>
        ''' <returns>UPDATE 文。</returns>
        ''' <remarks>
        ''' 上海コードから R3 品目コードを作る決まりのため、2 列を同時に更新します。
        ''' </remarks>
        Public Shared Function CreateUpdatePackingShanghaiCodeSql() As String
            Return "UPDATE dbo.tbl_t_shanghai_packing_import_check_temp " &
                   "SET shanghai_code = @ShanghaiCode, R3_hinmoku_code = @R3HinmokuCode " &
                   "WHERE pk_login_code = @LoginCode AND pk_seq_no = @SeqNo"
        End Function

        ''' <summary>
        ''' 一時データの件数を取得する SQL を作成します。再実行時に取込済みデータの有無を判定します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        Public Shared Function CreateSelectCheckTempCountSql() As String
            Return "SELECT COUNT(*) FROM dbo.tbl_t_shanghai_invoice_import_check_temp " &
                   "WHERE pk_login_code = @LoginCode"
        End Function

        ''' <summary>
        ''' 一時データに含まれる INVOICE 番号の一覧を取得する SQL を作成します。REF_NO の編集に使います。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        Public Shared Function CreateSelectInvoiceNoListSql() As String
            Return "SELECT invoice_no FROM dbo.tbl_t_shanghai_invoice_import_check_temp " &
                   "WHERE pk_login_code = @LoginCode AND invoice_no IS NOT NULL " &
                   "GROUP BY invoice_no ORDER BY invoice_no"
        End Function

        ''' <summary>
        ''' 取込結果確認画面（menuNo=108）で表示する INVOICE の一覧を取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        ''' <remarks>
        ''' INVOICE 番号の指定が無いとき（@InvoiceNo が NULL）は、その利用者の全件を返します。
        ''' 条件の有無で SQL を組み替えず、値はすべてパラメータで渡します。
        ''' </remarks>
        Public Shared Function CreateSelectInvoiceResultSql() As String
            Return "SELECT pk_seq_no, err_contents, err_status, suryo, suryo_tani, tsuka, unit_price, amount, " &
                   "R3_koubai_denpyo_no, invoice_no, bl_date, syukka_hoho, shanghai_code_text, " &
                   "shanghai_code, tajima_po_no " &
                   "FROM dbo.tbl_t_shanghai_invoice_import_check_temp " &
                   "WHERE pk_login_code = @LoginCode " &
                   "AND (@InvoiceNo IS NULL OR invoice_no = @InvoiceNo) " &
                   "ORDER BY invoice_no, pk_seq_no"
        End Function

        ''' <summary>
        ''' 取込結果確認画面（menuNo=108）で表示する PACKING の一覧を取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        Public Shared Function CreateSelectPackingResultSql() As String
            Return "SELECT pk_seq_no, err_contents, err_status, invoice_no, tajima_po_no, true_or_false, " &
                   "shanghai_code, syukka_hoho, bl_date, suryo, net_weight, gross_weight, m3, " &
                   "carton_no_from, carton_no_to, unit_price, amount " &
                   "FROM dbo.tbl_t_shanghai_packing_import_check_temp " &
                   "WHERE pk_login_code = @LoginCode " &
                   "AND (@InvoiceNo IS NULL OR invoice_no = @InvoiceNo) " &
                   "ORDER BY invoice_no, pk_seq_no"
        End Function

        ''' <summary>
        ''' 今回の取込対象のうち、INVOICE 明細として既に登録済みの INVOICE 番号を取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        Public Shared Function CreateSelectRegisteredInvoiceNoSql() As String
            Return "SELECT M.pk_invoice_no FROM dbo.View_shanghai_invoice_no_invoice AS M " &
                   "INNER JOIN dbo.View_shanghai_import_invoice_no AS T " &
                   "ON M.pk_invoice_no = T.invoice_no " &
                   "WHERE T.pk_login_code = @LoginCode " &
                   "GROUP BY M.pk_invoice_no"
        End Function

        ''' <summary>
        ''' 今回の取込対象のうち、PACKING 明細として既に登録済みの INVOICE 番号を取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        Public Shared Function CreateSelectRegisteredPackingInvoiceNoSql() As String
            ' 同じ INVOICE 番号を何度も削除しないよう GROUP BY で重複を除きます。
            Return "SELECT M.pk_invoice_no FROM dbo.View_shanghai_invoice_no_packing AS M " &
                   "INNER JOIN dbo.View_shanghai_import_invoice_no_packing AS T " &
                   "ON M.pk_invoice_no = T.invoice_no " &
                   "WHERE T.pk_login_code = @LoginCode " &
                   "GROUP BY M.pk_invoice_no"
        End Function

        ''' <summary>
        ''' INVOICE ヘッダを削除する SQL を作成します。
        ''' </summary>
        ''' <returns>DELETE 文。</returns>
        Public Shared Function CreateDeleteInvoiceHeaderSql() As String
            Return "DELETE FROM dbo.tbl_t_shanghai_invoice_header " &
                   "WHERE pk_invoice_no_main = @InvoiceNoMain"
        End Function

        ''' <summary>
        ''' INVOICE ヘッダサブを削除する SQL を作成します。
        ''' </summary>
        ''' <returns>DELETE 文。</returns>
        Public Shared Function CreateDeleteInvoiceHeaderSubSql() As String
            Return "DELETE FROM dbo.tbl_t_shanghai_invoice_header_sub " &
                   "WHERE pk_invoice_no = @InvoiceNo"
        End Function

        ''' <summary>
        ''' INVOICE 明細を削除する SQL を作成します。
        ''' </summary>
        ''' <returns>DELETE 文。</returns>
        Public Shared Function CreateDeleteInvoiceMeisaiSql() As String
            Return "DELETE FROM dbo.tbl_t_shanghai_invoice_meisai " &
                   "WHERE pk_invoice_no = @InvoiceNo"
        End Function

        ''' <summary>
        ''' PACKING 明細を削除する SQL を作成します。
        ''' </summary>
        ''' <returns>DELETE 文。</returns>
        Public Shared Function CreateDeletePackingMeisaiSql() As String
            Return "DELETE FROM dbo.tbl_t_shanghai_packing_meisai " &
                   "WHERE pk_invoice_no = @InvoiceNo"
        End Function

        ''' <summary>
        ''' INVOICE ヘッダを登録する SQL を作成します。
        ''' </summary>
        ''' <returns>INSERT 文。</returns>
        Public Shared Function CreateInsertInvoiceHeaderSql() As String
            ' ref_no と vessel は現行仕様どおり NULL で登録します。
            Return "INSERT INTO dbo.tbl_t_shanghai_invoice_header " &
                   "(pk_invoice_no_main, ref_no, bl_date, vessel, syukka_hoho, carton_qty, " &
                   "return_pallet, otsunaka_no, shiiresaki_code, update_ymd, update_login) " &
                   "VALUES (@InvoiceNoMain, NULL, @BlDate, NULL, @SyukkaHoho, @CartonQty, " &
                   "0, 0, @ShiiresakiCode, GETDATE(), @UpdateLogin)"
        End Function

        ''' <summary>
        ''' 項目マスタの一覧を取得する SQL を作成します。発注先の選択肢に使います。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        Public Shared Function CreateSelectKoumokuListSql() As String
            Return "SELECT pk_seq_no, contents1, contents2, contents3, contents4, contents_num " &
                   "FROM dbo.tbl_m_koumoku " &
                   "WHERE pk_koumoku_code = @KoumokuCode " &
                   "ORDER BY pk_seq_no"
        End Function
#End Region

    End Class
End Namespace
