/*
    既存DBの定義を調べるための参照クエリ集
    ------------------------------------------------------------------
    ・参照のみです。DDL（CREATE / ALTER / DROP / TRUNCATE）は含みません。
    ・名称は各クエリの WHERE 句で指定します。調べたい対象を足し引きしてください。
    ・SSMS では「結果をテキストで表示」（Ctrl+T）にしてから実行してください。
      グリッド表示だと長い定義が途中で切れます。
    ・結果はそのまま貼り付けてもらえれば読めます。ファイル保存は任意です。
*/

USE ATOM;
GO

-- ==================================================================
-- 1. ストアドプロシージャ・ビュー・関数の定義
--    調べたい名称を WHERE の IN に並べてください。
-- ==================================================================
SELECT
    o.name                              AS [オブジェクト名],
    o.type_desc                         AS [種別],
    m.definition                        AS [定義]
FROM sys.sql_modules AS m
INNER JOIN sys.objects AS o ON m.object_id = o.object_id
WHERE o.name IN (
        -- INVOICE/PACKING 取込（menuNo=101）
        'usp_shanghai_import_count_get_next_month',
        'usp_shanghai_import_tourokuzumi_count',
        'usp_shanghai_import_check_err_clear',
        'usp_shanghai_import_jyufuku_check',
        'usp_shanghai_import_err_check_invoice',
        'usp_shanghai_import_err_check_packing',
        'usp_shanghai_import_check_amount',
        'usp_shanghai_import_invoice_update',
        'usp_shanghai_invoice_gokei_get',
        'usp_POEM_tokuisaki_master_get',
        'usp_koumoku_master_pickup',
        'usp_koumoku_contents_get',
        -- 進捗更新（schedule_update）
        'usp_calender_mode_get',
        'usp_shanghai_invoice_header_pickup',
        'usp_shanghai_invoice_header_get',
        'usp_shinchoku_jyokyo_sonzai_check',
        'usp_TJM_master_calender_AM_PM_sonzai_check',
        -- ビュー
        'View_shanghai_invoice_no_invoice',
        'View_shanghai_invoice_no_packing',
        'View_shanghai_import_invoice_no',
        'View_shanghai_import_invoice_no_packing',
        'View__haiso_type'
      )
ORDER BY o.type_desc, o.name;
GO

-- ==================================================================
-- 2. ストアドプロシージャの引数定義（序数・型・長さ・入出力）
--    現行コードが序数（Parameters(n)）で受け渡ししているため必須です。
-- ==================================================================
SELECT
    o.name                              AS [ストアド名],
    p.parameter_id                      AS [序数],
    p.name                              AS [引数名],
    TYPE_NAME(p.user_type_id)           AS [型],
    p.max_length                        AS [最大長],
    p.precision                         AS [精度],
    p.scale                             AS [位取り],
    p.is_output                         AS [OUTPUTか]
FROM sys.parameters AS p
INNER JOIN sys.objects AS o ON p.object_id = o.object_id
WHERE o.name IN (
        'usp_shanghai_import_count_get_next_month',
        'usp_shanghai_import_tourokuzumi_count',
        'usp_shanghai_import_check_err_clear',
        'usp_shanghai_import_jyufuku_check',
        'usp_shanghai_import_err_check_invoice',
        'usp_shanghai_import_err_check_packing',
        'usp_shanghai_import_check_amount',
        'usp_shanghai_import_invoice_update',
        'usp_shanghai_invoice_gokei_get',
        'usp_POEM_tokuisaki_master_get',
        'usp_koumoku_master_pickup',
        'usp_koumoku_contents_get',
        'usp_calender_mode_get',
        'usp_shanghai_invoice_header_pickup',
        'usp_shanghai_invoice_header_get',
        'usp_shinchoku_jyokyo_sonzai_check',
        'usp_TJM_master_calender_AM_PM_sonzai_check'
      )
ORDER BY o.name, p.parameter_id;
GO

-- ==================================================================
-- 3. テーブル・ビューの列定義
-- ==================================================================
SELECT
    t.name                              AS [テーブル名],
    c.column_id                         AS [列順],
    c.name                              AS [列名],
    TYPE_NAME(c.user_type_id)           AS [型],
    c.max_length                        AS [最大長],
    c.precision                         AS [精度],
    c.scale                             AS [位取り],
    c.is_nullable                       AS [NULL許容],
    c.is_identity                       AS [ID列か]
FROM sys.columns AS c
INNER JOIN sys.objects AS t ON c.object_id = t.object_id
WHERE t.name IN (
        'tbl_t_shanghai_invoice_import_check_temp',
        'tbl_t_shanghai_packing_import_check_temp',
        'tbl_t_shanghai_invoice_header',
        'tbl_t_shanghai_invoice_header_sub',
        'tbl_t_shanghai_invoice_meisai',
        'tbl_t_shanghai_packing_meisai',
        'tbl_t_import_progress',
        'tbl_t_chotatsu_invoice_header',
        'tbl_m_invoice_default',
        'tbl_m_koumoku'
      )
ORDER BY t.name, c.column_id;
GO

-- ==================================================================
-- 4. 主キー・一意制約（置換処理の削除条件を確定させるため）
-- ==================================================================
SELECT
    t.name                              AS [テーブル名],
    i.name                              AS [インデックス名],
    CASE WHEN i.is_primary_key = 1 THEN N'PK'
         WHEN i.is_unique_constraint = 1 THEN N'UQ'
         ELSE N'UNIQUE INDEX' END       AS [種別],
    ic.key_ordinal                      AS [キー順],
    c.name                              AS [列名]
FROM sys.indexes AS i
INNER JOIN sys.objects AS t ON i.object_id = t.object_id
INNER JOIN sys.index_columns AS ic
    ON i.object_id = ic.object_id AND i.index_id = ic.index_id
INNER JOIN sys.columns AS c
    ON ic.object_id = c.object_id AND ic.column_id = c.column_id
WHERE i.is_unique = 1
  AND t.name IN (
        'tbl_t_shanghai_invoice_import_check_temp',
        'tbl_t_shanghai_packing_import_check_temp',
        'tbl_t_shanghai_invoice_header',
        'tbl_t_shanghai_invoice_header_sub',
        'tbl_t_shanghai_invoice_meisai',
        'tbl_t_shanghai_packing_meisai',
        'tbl_t_import_progress'
      )
ORDER BY t.name, i.name, ic.key_ordinal;
GO

-- ==================================================================
-- 5. 項目マスタの区分 'COS'（発注先の選択肢）
--    B-3「customer_kbn の値と表示名の対応」を確定させるため。
-- ==================================================================
EXEC usp_koumoku_master_pickup 'COS';
GO

-- ==================================================================
-- 6. 通貨の内部値の裏取り（B-4）
--    INVOICE_NO の先頭と通貨の対応を「件数だけ」確認します。
--    明細データそのものは出力しません。
-- ==================================================================
SELECT
    LEFT(pk_invoice_no, 5)              AS [INVOICE_NO先頭5文字],
    tsuka                               AS [通貨],
    COUNT(*)                            AS [件数]
FROM tbl_t_shanghai_invoice_meisai
GROUP BY LEFT(pk_invoice_no, 5), tsuka
ORDER BY [INVOICE_NO先頭5文字], [通貨];
GO

-- ==================================================================
-- 7. 別データベース（TJM_master）側の列定義
-- ==================================================================
USE TJM_master;
GO

SELECT
    t.name                              AS [テーブル名],
    c.column_id                         AS [列順],
    c.name                              AS [列名],
    TYPE_NAME(c.user_type_id)           AS [型],
    c.max_length                        AS [最大長],
    c.is_nullable                       AS [NULL許容]
FROM sys.columns AS c
INNER JOIN sys.objects AS t ON c.object_id = t.object_id
WHERE t.name IN ('tbl_m_login_sys', 'tbl_m_calender_jigyosyo')
ORDER BY t.name, c.column_id;
GO

-- ==================================================================
-- 補助. 名前が分からないときの探し方
--    LIKE で当たりを付けてから、上のクエリの IN へ名称を足してください。
-- ==================================================================
/*
USE ATOM;

-- 名前に「shanghai」を含むストアド・ビューを探す
SELECT name, type_desc
FROM sys.objects
WHERE type IN ('P', 'V', 'FN', 'IF', 'TF')
  AND name LIKE '%shanghai%'
ORDER BY type_desc, name;

-- 名前に「invoice」を含むテーブルを探す
SELECT name
FROM sys.tables
WHERE name LIKE '%invoice%'
ORDER BY name;

-- 定義の中に特定の語（例: tbl_t_shanghai_invoice_meisai）が出てくるオブジェクトを探す
SELECT o.name, o.type_desc
FROM sys.sql_modules AS m
INNER JOIN sys.objects AS o ON m.object_id = o.object_id
WHERE m.definition LIKE '%tbl_t_shanghai_invoice_meisai%'
ORDER BY o.type_desc, o.name;
*/
