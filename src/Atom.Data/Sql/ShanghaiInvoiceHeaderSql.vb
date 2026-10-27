Option Strict On
Option Explicit On
Option Infer Off

Namespace Sql

    ''' <summary>
    ''' INVOICEヘッダー入力（menuNo=103）で使う SQL 文を組み立てます。
    ''' </summary>
    ''' <remarks>
    ''' 登録状況の判定と一時データの作成はストアドプロシージャ側にあります。
    ''' ここに置くのは選択肢の取得、ヘッダの読み込み、一時データの読み書きです。
    ''' </remarks>
    Public NotInheritable Class ShanghaiInvoiceHeaderSql

#Region "コンストラクタ"
        ''' <summary>インスタンス化を禁止します。</summary>
        Private Sub New()
        End Sub
#End Region

#Region "選択肢"
        ''' <summary>
        ''' 元INV_NO の選択肢を取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        ''' <remarks>
        ''' ビュー側で B/L DATE の降順に並んでいますが、ビューの ORDER BY は保証されないため
        ''' ここでも明示します。B/L DATE は並び順にしか使わないため取得しません。
        ''' </remarks>
        Public Shared Function CreateSelectInvoiceNoMainOptionsSql() As String
            Return "SELECT pk_invoice_no_main " &
                   "FROM dbo.View__invoice_no_main " &
                   "ORDER BY bl_date DESC, pk_invoice_no_main DESC"
        End Function

        ''' <summary>
        ''' 出荷方法の選択肢を取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        Public Shared Function CreateSelectSyukkaHohoOptionsSql() As String
            Return "SELECT syukka_hoho, syukka_hoho_text " &
                   "FROM dbo.View__syukka_hoho " &
                   "ORDER BY sort_seq, syukka_hoho"
        End Function

        ''' <summary>
        ''' 乙仲の選択肢を取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        Public Shared Function CreateSelectOtsunakaOptionsSql() As String
            Return "SELECT otsunaka_no, otsunaka_no_text " &
                   "FROM dbo.View__otsunaka " &
                   "ORDER BY sort_seq, otsunaka_no"
        End Function

        ''' <summary>
        ''' 配送タイプの選択肢を取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        Public Shared Function CreateSelectHaisoTypeOptionsSql() As String
            Return "SELECT haiso_type, haiso_type_text " &
                   "FROM dbo.View__haiso_type " &
                   "ORDER BY haiso_type"
        End Function

        ''' <summary>
        ''' コンテナタイプの選択肢を取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        ''' <remarks>
        ''' 並び順は区分の値の順です。ビューに並び順の列があるかは未確認のため指定していません。
        ''' </remarks>
        Public Shared Function CreateSelectContainerTypeOptionsSql() As String
            Return "SELECT container_type, container_type_text " &
                   "FROM dbo.View__container_type " &
                   "ORDER BY container_type"
        End Function

        ''' <summary>
        ''' 船名の選択肢を項目マスタから取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        ''' <remarks>項目コード 'SHP' の行が船名の一覧です。</remarks>
        Public Shared Function CreateSelectVesselOptionsSql() As String
            Return "SELECT contents1 " &
                   "FROM dbo.tbl_m_koumoku " &
                   "WHERE pk_koumoku_code = @KoumokuCode " &
                   "ORDER BY sort_seq, pk_seq_no"
        End Function
#End Region

#Region "ヘッダ"
        ''' <summary>
        ''' INVOICE のヘッダを 1 件取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        Public Shared Function CreateSelectHeaderSql() As String
            Return "SELECT pk_invoice_no_main, bl_date, syukka_hoho, carton_qty, vessel, " &
                   "return_pallet, ref_no, otsunaka_no, shiiresaki_code, update_ymd, update_login " &
                   "FROM dbo.tbl_t_shanghai_invoice_header " &
                   "WHERE pk_invoice_no_main = @InvoiceNoMain"
        End Function
#End Region

#Region "ヘッダの登録"
        ''' <summary>
        ''' INVOICE のヘッダを登録する SQL を作成します。
        ''' </summary>
        ''' <returns>INSERT 文。</returns>
        Public Shared Function CreateInsertHeaderSql() As String
            Return "INSERT INTO dbo.tbl_t_shanghai_invoice_header " &
                   "(pk_invoice_no_main, ref_no, bl_date, vessel, syukka_hoho, carton_qty, " &
                   "return_pallet, otsunaka_no, update_ymd, update_login) " &
                   "VALUES (@InvoiceNoMain, @RefNo, @BlDate, @Vessel, @SyukkaHoho, @CartonQty, " &
                   "@ReturnPallet, @OtsunakaNo, GETDATE(), @LoginCode)"
        End Function

        ''' <summary>
        ''' INVOICE のヘッダを更新する SQL を作成します。
        ''' </summary>
        ''' <returns>UPDATE 文。</returns>
        Public Shared Function CreateUpdateHeaderSql() As String
            Return "UPDATE dbo.tbl_t_shanghai_invoice_header " &
                   "SET ref_no = @RefNo, " &
                   "bl_date = @BlDate, " &
                   "vessel = @Vessel, " &
                   "syukka_hoho = @SyukkaHoho, " &
                   "carton_qty = @CartonQty, " &
                   "return_pallet = @ReturnPallet, " &
                   "otsunaka_no = @OtsunakaNo, " &
                   "update_ymd = GETDATE(), " &
                   "update_login = @LoginCode " &
                   "WHERE pk_invoice_no_main = @InvoiceNoMain"
        End Function
#End Region

#Region "印刷用 INVOICE_NO"
        ''' <summary>
        ''' 印刷用 INVOICE_NO を組み立てる対象の明細を取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        ''' <remarks>
        ''' 配送タイプ 3（日新入荷）は対象外です。
        ''' コンテナ区分・コンテナNO・INVOICE_NO の順に並べ、コンテナNO ごとに区切ります。
        ''' </remarks>
        Public Shared Function CreateSelectPrintTargetsSql() As String
            Return "SELECT pk_invoice_no, container_type, container_no, haiso_yotei_ymd " &
                   "FROM dbo.tbl_t_shanghai_invoice_header_sub " &
                   "WHERE pk_invoice_no_main = @InvoiceNoMain " &
                   "AND haiso_type <> @ExcludedHaisoType " &
                   "ORDER BY container_type, container_no, pk_invoice_no"
        End Function

        ''' <summary>
        ''' 印刷用 INVOICE_NO を更新する SQL を作成します。
        ''' </summary>
        ''' <returns>UPDATE 文。</returns>
        ''' <remarks>
        ''' 旧 Access はコンテナNO だけで更新しており、同じコンテナNO を持つ別の INVOICE まで
        ''' 書き換わる可能性がありました。ここでは INVOICE 番号（メイン）でも絞り込みます（C-27）。
        ''' </remarks>
        Public Shared Function CreateUpdateInvoiceNoPrintSql() As String
            Return "UPDATE dbo.tbl_t_shanghai_invoice_header_sub " &
                   "SET invoice_no_print = @InvoiceNoPrint " &
                   "WHERE pk_invoice_no_main = @InvoiceNoMain " &
                   "AND container_no = @ContainerNo"
        End Function
#End Region

#Region "明細（作業用の一時データ）"
        ''' <summary>
        ''' 作業用の一時データから明細を取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        ''' <remarks>
        ''' 一時データは利用者ごとに分かれています。必ずログインコードで絞り込みます。
        ''' </remarks>
        Public Shared Function CreateSelectDetailTempSql() As String
            Return "SELECT pk_invoice_no_main, pk_invoice_no, haiso_type, container_type, " &
                   "haiso_yotei_ymd, container_no, seal_no, carton_qty, invoice_no_print, del_flg " &
                   "FROM dbo.tbl_t_shanghai_invoice_header_sub_temp " &
                   "WHERE pk_login_code = @LoginCode " &
                   "ORDER BY pk_invoice_no"
        End Function

        ''' <summary>
        ''' 作業用の一時データの 1 行を更新する SQL を作成します。
        ''' </summary>
        ''' <returns>UPDATE 文。</returns>
        Public Shared Function CreateUpdateDetailTempSql() As String
            Return "UPDATE dbo.tbl_t_shanghai_invoice_header_sub_temp " &
                   "SET haiso_type = @HaisoType, " &
                   "container_type = @ContainerType, " &
                   "haiso_yotei_ymd = @HaisoYoteiYmd, " &
                   "container_no = @ContainerNo, " &
                   "seal_no = @SealNo, " &
                   "carton_qty = @CartonQty " &
                   "WHERE pk_login_code = @LoginCode " &
                   "AND pk_invoice_no_main = @InvoiceNoMain " &
                   "AND pk_invoice_no = @InvoiceNo"
        End Function
#End Region

    End Class
End Namespace
