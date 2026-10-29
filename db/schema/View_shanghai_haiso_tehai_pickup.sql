-- 既存DBの参照用スナップショット（読み取り専用）
-- 取得日: 2026-10-29 / DB: ATOM
-- 配送手配書の印刷対象を拾うビュー（menuNo=103 §10.3）。
-- 配送先 1〜3 は View__haiso_type に紐づく固定値（§10.7）。
CREATE VIEW dbo.View_shanghai_haiso_tehai_pickup
AS
SELECT  dbo.tbl_t_shanghai_invoice_header.pk_invoice_no_main,
        dbo.tbl_t_shanghai_invoice_header.bl_date,
        dbo.tbl_t_shanghai_invoice_header.vessel,
        dbo.tbl_t_shanghai_invoice_header.otsunaka_no,
        dbo.tbl_t_shanghai_invoice_header_sub.haiso_yotei_ymd,
        dbo.tbl_t_shanghai_invoice_header_sub.haiso_type,
        dbo.View__haiso_type.haiso_type_text,
        dbo.tbl_t_shanghai_invoice_header_sub.container_type,
        dbo.tbl_t_shanghai_invoice_header_sub.container_no,
        dbo.tbl_t_shanghai_invoice_header_sub.seal_no,
        dbo.tbl_t_shanghai_invoice_header_sub.pk_invoice_no,
        dbo.tbl_t_shanghai_invoice_header_sub.carton_qty,
        dbo.View_shanghai_packing_total_suryo.total_suryo,
        dbo.View__container_type.container_type_text,
        dbo.View__haiso_type.haiso_saki1,
        dbo.View__haiso_type.haiso_saki2,
        dbo.View__haiso_type.haiso_saki3
FROM    dbo.tbl_t_shanghai_invoice_header
        INNER JOIN dbo.tbl_t_shanghai_invoice_header_sub
            ON dbo.tbl_t_shanghai_invoice_header.pk_invoice_no_main = dbo.tbl_t_shanghai_invoice_header_sub.pk_invoice_no_main
        INNER JOIN dbo.View_shanghai_packing_total_suryo
            ON dbo.tbl_t_shanghai_invoice_header_sub.pk_invoice_no = dbo.View_shanghai_packing_total_suryo.pk_invoice_no
           AND dbo.tbl_t_shanghai_invoice_header_sub.pk_invoice_no_main = dbo.View_shanghai_packing_total_suryo.pk_invoice_no_main
        LEFT OUTER JOIN dbo.View__haiso_type
            ON dbo.tbl_t_shanghai_invoice_header_sub.haiso_type = dbo.View__haiso_type.haiso_type
        LEFT OUTER JOIN dbo.View__container_type
            ON dbo.tbl_t_shanghai_invoice_header_sub.container_type = dbo.View__container_type.container_type
