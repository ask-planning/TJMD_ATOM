Option Strict On
Option Explicit On
Option Infer Off

Namespace Sql

    ''' <summary>
    ''' 配送手配書（menuNo=103 §10）で使う SQL 文を組み立てます。
    ''' </summary>
    ''' <remarks>
    ''' 値はすべて SqlParameter で渡します。文字列連結で値を埋め込みません。
    ''' 印刷用一時テーブルは利用者ごとに分かれているため、必ずログインコードで絞り込みます。
    ''' </remarks>
    Public NotInheritable Class ShanghaiHaisoTehaiSql

#Region "コンストラクタ"
        ''' <summary>インスタンス化を禁止します。</summary>
        Private Sub New()
        End Sub
#End Region

#Region "印刷対象の取得"
        ''' <summary>
        ''' 印刷対象を取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        ''' <remarks>
        ''' 並び順は設計書 §10.3 のとおり
        ''' 配送予定日 → 配送タイプ → コンテナ区分 → コンテナNO → INVOICE_NO です。
        ''' 配送予定日が未入力の行を後ろへ送らないよう、NULL も含めて昇順で並べます。
        ''' </remarks>
        Public Shared Function CreateSelectPickupSql() As String
            Return "SELECT pk_invoice_no_main, pk_invoice_no, bl_date, vessel, otsunaka_no, " &
                   "haiso_yotei_ymd, haiso_type, haiso_type_text, " &
                   "container_type, container_type_text, container_no, seal_no, " &
                   "carton_qty, total_suryo, haiso_saki1, haiso_saki2, haiso_saki3 " &
                   "FROM dbo.View_shanghai_haiso_tehai_pickup " &
                   "WHERE pk_invoice_no_main = @InvoiceNoMain " &
                   "ORDER BY haiso_yotei_ymd, haiso_type, container_type, container_no, pk_invoice_no"
        End Function
#End Region

#Region "印刷用の一時データ"
        ''' <summary>
        ''' 印刷用の一時データを削除する SQL を作成します。
        ''' </summary>
        ''' <returns>DELETE 文。</returns>
        Public Shared Function CreateDeletePrintTempSql() As String
            Return "DELETE FROM dbo.tbl_t_shanghai_haiso_tehai_print_temp " &
                   "WHERE pk_login_code = @LoginCode"
        End Function

        ''' <summary>
        ''' 印刷用の一時データを 1 行登録する SQL を作成します。
        ''' </summary>
        ''' <returns>INSERT 文。</returns>
        ''' <remarks>
        ''' `pk_seq_no` は採番列のため指定しません。
        ''' `ref_no` / `nisshin_plant` / `logistic_plant` は元データが特定できていないため
        ''' 登録しません（設計書 §18 C-37）。
        ''' `week_day` は Web 版では使わないため登録しません（同 C-38）。
        ''' </remarks>
        Public Shared Function CreateInsertPrintTempSql() As String
            Return "INSERT INTO dbo.tbl_t_shanghai_haiso_tehai_print_temp " &
                   "(pk_login_code, invoice_no_main, invoice_no, vessel, bl_date, haiso_type, " &
                   "container_type, container_no, seal_no, total_carton, total_quantity, " &
                   "haisyo_yotei_ymd, haiso_saki1, haiso_saki2, haiso_saki3, otsunaka_no) " &
                   "VALUES " &
                   "(@LoginCode, @InvoiceNoMain, @InvoiceNo, @Vessel, @BlDate, @HaisoType, " &
                   "@ContainerType, @ContainerNo, @SealNo, @TotalCarton, @TotalQuantity, " &
                   "@HaisoYoteiYmd, @HaisoSaki1, @HaisoSaki2, @HaisoSaki3, @OtsunakaNo)"
        End Function
#End Region

#Region "帳票の内容"
        ''' <summary>
        ''' 帳票に印字する内容を取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        ''' <remarks>
        ''' **ビューが採番列（`pk_seq_no`）を返さないため、項目で並べ替えます。**
        ''' §10.3 の並び順に寄せていますが、**コンテナ区分は外しています。**
        ''' 印刷用一時データに入っているのは名称で、拾うときに使ったコードとは順序が一致しないためです
        ''' （例: コード順と `20Ft` / `40Ft` / `40HQ` の名称順は別物）。
        ''' 登録順へ完全に揃えるには、ビューが `pk_seq_no` を返す必要があります（設計書 §18 C-40）。
        ''' </remarks>
        Public Shared Function CreateSelectReportRowsSql() As String
            Return "SELECT ref_no, invoice_no, vessel, bl_date, haiso_type, " &
                   "container_type, container_no, seal_no, total_carton, haisyo_yotei_ymd, " &
                   "otsunaka_no_text, otsunaka_tanto1, otsunaka_tanto2, " &
                   "haiso_saki1, haiso_saki2, haiso_saki3, " &
                   "haiso_saki1_name, haiso_saki1_address, haiso_saki1_phone_no, haiso_saki1_tanto1, " &
                   "haiso_saki2_name, haiso_saki2_address, haiso_saki2_phone_no, haiso_saki2_tanto1, " &
                   "haiso_saki3_name, haiso_saki3_address, haiso_saki3_phone_no, haiso_saki3_tanto1, " &
                   "display_name, bumon_code_text, phone_no, fax_no " &
                   "FROM dbo.View_Rpt_shanghai_haiso_tehai " &
                   "WHERE pk_login_code = @LoginCode " &
                   "ORDER BY haisyo_yotei_ymd, haiso_type, container_no, invoice_no"
        End Function
#End Region

    End Class
End Namespace
