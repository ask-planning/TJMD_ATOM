-- 既存DBの参照用スナップショット（読み取り専用）
-- 取得日: 2026-10-29 / DB: ATOM
-- 配送手配書（menuNo=103 §10 / §14）の元データ。
-- 印刷用一時テーブルに各マスタを結合したもの。
CREATE VIEW dbo.View_Rpt_shanghai_haiso_tehai
AS
SELECT  dbo.tbl_t_shanghai_haiso_tehai_print_temp.pk_login_code,
        dbo.tbl_t_shanghai_haiso_tehai_print_temp.ref_no,
        dbo.tbl_t_shanghai_haiso_tehai_print_temp.invoice_no_main,
        dbo.tbl_t_shanghai_haiso_tehai_print_temp.invoice_no,
        dbo.tbl_t_shanghai_haiso_tehai_print_temp.vessel,
        dbo.tbl_t_shanghai_haiso_tehai_print_temp.bl_date,
        dbo.tbl_t_shanghai_haiso_tehai_print_temp.haiso_type,
        dbo.tbl_t_shanghai_haiso_tehai_print_temp.container_type,
        dbo.tbl_t_shanghai_haiso_tehai_print_temp.container_no,
        dbo.tbl_t_shanghai_haiso_tehai_print_temp.seal_no,
        dbo.tbl_t_shanghai_haiso_tehai_print_temp.total_carton,
        dbo.tbl_t_shanghai_haiso_tehai_print_temp.total_quantity,
        dbo.tbl_t_shanghai_haiso_tehai_print_temp.haisyo_yotei_ymd,
        dbo.tbl_t_shanghai_haiso_tehai_print_temp.week_day,
        dbo.View__week_text.weekday_no_text,
        dbo.tbl_t_shanghai_haiso_tehai_print_temp.haiso_saki1,
        dbo.View__haiso_saki.haiso_saki_name       AS haiso_saki1_name,
        dbo.View__haiso_saki.haiso_saki_address    AS haiso_saki1_address,
        dbo.View__haiso_saki.haiso_saki_phone_no   AS haiso_saki1_phone_no,
        dbo.View__haiso_saki.haiso_saki_tanto1     AS haiso_saki1_tanto1,
        dbo.View__otsunaka.otsunaka_no_text,
        dbo.View__otsunaka.otsunaka_tanto1,
        dbo.View__otsunaka.otsunaka_tanto2,
        View__otsunaka_1.otsunaka_no_text          AS atesaki_nisshin,
        View__otsunaka_1.otsunaka_tanto1           AS atesaki_nisshin_tanto1,
        View__otsunaka_1.otsunaka_tanto2           AS atesaki_nisshin_tanto2,
        dbo.tbl_t_shanghai_haiso_tehai_print_temp.haiso_saki2,
        View__haiso_saki_1.haiso_saki_name         AS haiso_saki2_name,
        View__haiso_saki_1.haiso_saki_address      AS haiso_saki2_address,
        View__haiso_saki_1.haiso_saki_phone_no     AS haiso_saki2_phone_no,
        View__haiso_saki_1.haiso_saki_tanto1       AS haiso_saki2_tanto1,
        dbo.tbl_t_shanghai_haiso_tehai_print_temp.haiso_saki3,
        View__haiso_saki_2.haiso_saki_name         AS haiso_saki3_name,
        View__haiso_saki_2.haiso_saki_address      AS haiso_saki3_address,
        View__haiso_saki_2.haiso_saki_phone_no     AS haiso_saki3_phone_no,
        View__haiso_saki_2.haiso_saki_tanto1       AS haiso_saki3_tanto1,
        TJM_master.dbo.tbl_m_login_sys.display_name,
        dbo.View__bumon_code.bumon_code_text,
        TJM_master.dbo.tbl_m_login_sys.phone_no,
        TJM_master.dbo.tbl_m_login_sys.fax_no
FROM    dbo.tbl_t_shanghai_haiso_tehai_print_temp
        INNER JOIN TJM_master.dbo.tbl_m_login_sys
            ON dbo.tbl_t_shanghai_haiso_tehai_print_temp.pk_login_code = TJM_master.dbo.tbl_m_login_sys.pk_shain_code
        INNER JOIN dbo.View__bumon_code
            ON TJM_master.dbo.tbl_m_login_sys.bumon_code = dbo.View__bumon_code.pk_bumon_code
        LEFT OUTER JOIN dbo.View__haiso_saki AS View__haiso_saki_2
            ON dbo.tbl_t_shanghai_haiso_tehai_print_temp.haiso_saki3 = View__haiso_saki_2.haiso_saki_no
        LEFT OUTER JOIN dbo.View__haiso_saki AS View__haiso_saki_1
            ON dbo.tbl_t_shanghai_haiso_tehai_print_temp.haiso_saki2 = View__haiso_saki_1.haiso_saki_no
        LEFT OUTER JOIN dbo.View__otsunaka AS View__otsunaka_1
            ON dbo.tbl_t_shanghai_haiso_tehai_print_temp.nisshin_plant = View__otsunaka_1.otsunaka_no
        LEFT OUTER JOIN dbo.View__otsunaka
            ON dbo.tbl_t_shanghai_haiso_tehai_print_temp.otsunaka_no = dbo.View__otsunaka.otsunaka_no
        LEFT OUTER JOIN dbo.View__haiso_saki
            ON dbo.tbl_t_shanghai_haiso_tehai_print_temp.haiso_saki1 = dbo.View__haiso_saki.haiso_saki_no
        LEFT OUTER JOIN dbo.View__week_text
            ON dbo.tbl_t_shanghai_haiso_tehai_print_temp.week_day = dbo.View__week_text.weekday_no
