-- 既存DBの参照用スナップショット（読み取り専用）
-- 取得日: 2026-10-30 / DB: ATOM
-- 入荷明細書（menuNo=103 §11 / §14）の元データ。
-- 印刷用一時テーブルに品目マスタと保管場所を結合したもの。
CREATE VIEW dbo.View_Rpt_shanghai_nyuka_meisai
AS
SELECT  TOP (1000000)
        t.pk_login_code, t.sort_seq,
        t.pk_invoice_no_main, t.pk_invoice_no,
        t.bl_date, t.container_type, t.container_type_text,
        t.container_no, t.seal_no, t.haiso_yotei_ymd,
        t.nounyu_plant, t.koubai_hacchu_no,
        t.zaiko_tenso_no, t.syukka_denpyo_no,
        t.nounyu_kijitsu_ymd, t.R3_hinmoku_code,
        t.hinmoku_text, t.hacchu_tani,
        t.nounyu_suryo, t.suryo_tani,
        t.syanai_sansyo_no, t.syanai_sansyo_no_tenso,
        t.carton_no_from, t.carton_no_to, t.carton_qty,
        t.hosozai_flg, t.nukitori_kensa_flg, t.invoice_no_print,
        TJM_master.dbo.tbl_m_hinmokumaster.hinmei,
        dbo.View__plant_hokan_basyo.hokan_basyo_text,
        t.remarks
FROM    dbo.tbl_t_shanghai_nyuka_meisai_temp AS t
        LEFT OUTER JOIN TJM_master.dbo.tbl_m_hinmokumaster
            ON t.R3_hinmoku_code = TJM_master.dbo.tbl_m_hinmokumaster.hinmokucode_0
        LEFT OUTER JOIN dbo.View__plant_hokan_basyo
            ON t.nounyu_plant = dbo.View__plant_hokan_basyo.hokan_basyo
ORDER BY t.pk_login_code, t.sort_seq, t.container_no, t.pk_invoice_no,
         t.nounyu_plant, t.zaiko_tenso_no, t.syukka_denpyo_no,
         t.carton_no_from, t.R3_hinmoku_code

-- 注意
--  * 元の定義は列を dbo.tbl_t_shanghai_nyuka_meisai_temp.… と完全修飾で並べている。
--    ここでは読みやすさのため別名 t を付けた。**内容は同じ**。
--  * ビューの ORDER BY は保証されない。呼び出し側で必ず並べ替えること。
--  * TJM_master を参照する（品目マスタ）。
