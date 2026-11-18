Option Strict On
Option Explicit On
Option Infer Off

Namespace Sql

    ''' <summary>
    ''' 入荷明細書（menuNo=103 §11）で使う SQL 文を組み立てます。
    ''' </summary>
    ''' <remarks>
    ''' 値はすべて SqlParameter で渡します。文字列連結で値を埋め込みません。
    ''' 印刷用一時テーブルは利用者ごとに分かれているため、必ずログインコードで絞り込みます。
    ''' </remarks>
    Public NotInheritable Class ShanghaiNyukaMeisaiSql

#Region "コンストラクタ"
        ''' <summary>インスタンス化を禁止します。</summary>
        Private Sub New()
        End Sub
#End Region

#Region "帳票の内容"
        ''' <summary>
        ''' 帳票に印字する内容を取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        ''' <remarks>
        ''' 並び順は旧レポート `R_上海タジマ入荷明細` のグループ構成に合わせています。
        ''' コンテナ → 枝番 → 納入先 → 在庫転送NO → 出荷伝票 → CARTON_NO → 品目コードの順です。
        '''
        ''' **カートン数はビューに無いため、枝番のヘッダーから持ってきます。**
        ''' 同じ枝番が複数行あることがある（実データで 24 件）ため、結合が増えないよう
        ''' あらかじめ 1 行にまとめてから結合します。
        ''' 削除済みを外す条件に `ISNULL` を使っているのは、`del_flg` が NULL の行を
        ''' 落とさないためです（`del_flg &lt;&gt; 1` では NULL が除かれます）。
        ''' </remarks>
        Public Shared Function CreateSelectReportRowsSql() As String
            Return "SELECT v.sort_seq, v.pk_invoice_no_main, v.pk_invoice_no, " &
                   "v.bl_date, v.container_type_text, v.container_no, v.seal_no, " &
                   "v.haiso_yotei_ymd, v.invoice_no_print, v.remarks, " &
                   "v.nounyu_plant, v.hokan_basyo_text, v.zaiko_tenso_no, v.syukka_denpyo_no, " &
                   "v.nukitori_kensa_flg, v.R3_hinmoku_code, " &
                   "v.carton_no_from, v.carton_no_to, v.carton_qty, v.nounyu_suryo, " &
                   "h.carton_qty AS total_carton " &
                   "FROM dbo.View_Rpt_shanghai_nyuka_meisai AS v " &
                   "LEFT OUTER JOIN (SELECT pk_invoice_no, MAX(carton_qty) AS carton_qty " &
                   "FROM dbo.tbl_t_shanghai_invoice_header_sub " &
                   "WHERE ISNULL(del_flg, 0) <> 1 " &
                   "GROUP BY pk_invoice_no) AS h " &
                   "ON h.pk_invoice_no = v.pk_invoice_no " &
                   "WHERE v.pk_login_code = @LoginCode " &
                   "ORDER BY v.sort_seq, v.container_no, v.pk_invoice_no, v.nounyu_plant, " &
                   "v.zaiko_tenso_no, v.syukka_denpyo_no, v.carton_no_from, v.R3_hinmoku_code"
        End Function
#End Region

    End Class
End Namespace
