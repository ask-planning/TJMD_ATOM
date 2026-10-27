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
        'View__haiso_type',
        -- INVOICEヘッダー入力（menuNo=103）
        'usp_shanghai_invoice_sonzai_check',
        'usp_shanghai_invoice_header_sub_temp_create',
        'usp_shanghai_invoice_header_sub_update',
        'usp_shanghai_invoice_container_count',
        'usp_shanghai_ref_no_get',
        'usp_shanghai_ref_no_update',
        -- 登録の後続処理（§8.6）と明細更新（§9）。引数が未確認
        'usp_shanghai_invoice_shikyu_hosozai_nouki_update',
        'usp_poem_makishin_nyu_syukko_update',
        'usp_shanghai_container_type_check',
        'usp_shanghai_container_no_update',
        'View__container_type',
        'View__invoice_no_main',
        'View__syukka_hoho',
        'View__otsunaka',
        'View_shanghai_haiso_tehai_pickup',
        -- 101 の中で使われているが定義が未確認のビュー
        'View_shanghai_special_hinmoku',
        'View_shanghai_import_suryo_check_invoice',
        'View_shanghai_import_suryo_check_packing',
        'View_shanghai_import_invoice_check_tajima_po',
        'View_shanghai_import_packing_check_tajima_po',
        'View_shanghai_import_packing_check_invoice_no',
        'View_shanghai_packing_carton_no',
        'View_shanghai_invoice_total_amount',
        'View_shanghai_invoice_container_no',
        'View_shanghai_invoice_bl_date_check',
        'View_shanghai_packing_total_suryo'
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
        'usp_TJM_master_calender_AM_PM_sonzai_check',
        -- INVOICEヘッダー入力（menuNo=103）
        'usp_shanghai_invoice_sonzai_check',
        'usp_shanghai_invoice_header_sub_temp_create',
        'usp_shanghai_invoice_header_sub_update',
        'usp_shanghai_invoice_container_count',
        'usp_shanghai_ref_no_get',
        'usp_shanghai_ref_no_update',
        'usp_shanghai_invoice_shikyu_hosozai_nouki_update',
        'usp_poem_makishin_nyu_syukko_update',
        'usp_shanghai_container_type_check',
        'usp_shanghai_container_no_update'
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

-- ==================================================================
-- 補助2. 103 の明細が入る一時テーブルを探す
--    usp_shanghai_invoice_header_sub_temp_create が何を作るのかが
--    旧 Access の VBA からは読み取れないため、DB 側から特定します。
--    結果に出たテーブル名を、上の「3. 列定義」の IN へ足してください。
-- ==================================================================
/*
USE ATOM;

-- 名前に「header_sub」を含むテーブル・ビュー
SELECT name, type_desc
FROM sys.objects
WHERE type IN ('U', 'V')
  AND name LIKE '%header_sub%'
ORDER BY type_desc, name;

-- ストアドの中で INSERT されている先を直接見る（定義を読むのが確実）
SELECT m.definition
FROM sys.sql_modules AS m
INNER JOIN sys.objects AS o ON m.object_id = o.object_id
WHERE o.name = 'usp_shanghai_invoice_header_sub_temp_create';
*/

-- ==================================================================
-- 補助3. 接続先に追従しないストアド・ビューを洗い出す
--    POEM.dbo. のようにデータベース名を直接書いているものは、
--    検証環境から実行しても本番を参照・更新してしまいます。
--    特に UPDATE / DELETE / INSERT を含むものは要注意です。
-- ==================================================================
/*
USE ATOM;

SELECT
    o.name                                      AS [オブジェクト名],
    o.type_desc                                 AS [種別],
    CASE
        WHEN m.definition LIKE '%DELETE%' OR m.definition LIKE '%UPDATE%' OR m.definition LIKE '%INSERT%'
        THEN '更新あり'
        ELSE '参照のみ'
    END                                         AS [更新の有無]
FROM sys.sql_modules AS m
INNER JOIN sys.objects AS o ON m.object_id = o.object_id
WHERE m.definition LIKE '%POEM.dbo.%'
   OR m.definition LIKE '%TJM_master.dbo.%'
ORDER BY [更新の有無] DESC, o.type_desc, o.name;
*/

-- ==================================================================
-- 補助4.【精度が低い。補助6 を使うこと】
--    別データベースを「更新している」ものだけに絞る（文字列一致）
--    補助3 は件数が多く（2026-10-23 時点で 91 件）、目視が大変です。
--    別DB参照があり、かつ更新文を含むものだけに絞ります。
--    ビューは更新文を持てないため、実質ストアドだけになります。
--
--    ※【2026-10-23 判明】このクエリは誤検出が多いです。
--      update_ymd / delete_ymd という列名が 'UPDATE' / 'DELETE' に当たるため、
--      ほぼすべてのオブジェクトで U と D が立ちます。
--      （ビューに D が付くのがその証拠。ビューは DELETE できません）
--      実際の判定には補助6 を使ってください。
-- ==================================================================
/*
USE ATOM;

SELECT
    o.name                                  AS [オブジェクト名],
    o.type_desc                             AS [種別],
    CASE WHEN m.definition LIKE '%POEM.dbo.%'       THEN '●' ELSE '' END AS [POEM],
    CASE WHEN m.definition LIKE '%TJM_master.dbo.%' THEN '●' ELSE '' END AS [TJM],
    CASE WHEN m.definition LIKE '%INSERT%' THEN 'I' ELSE '' END
  + CASE WHEN m.definition LIKE '%UPDATE%' THEN 'U' ELSE '' END
  + CASE WHEN m.definition LIKE '%DELETE%' THEN 'D' ELSE '' END
  + CASE WHEN m.definition LIKE '%MERGE%'  THEN 'M' ELSE '' END
                                            AS [更新文の有無]
FROM sys.sql_modules AS m
INNER JOIN sys.objects AS o ON m.object_id = o.object_id
WHERE (m.definition LIKE '%POEM.dbo.%' OR m.definition LIKE '%TJM_master.dbo.%')
  AND (m.definition LIKE '%INSERT%'
    OR m.definition LIKE '%UPDATE%'
    OR m.definition LIKE '%DELETE%'
    OR m.definition LIKE '%MERGE%')
ORDER BY o.type_desc, o.name;
*/

-- ==================================================================
-- 補助5. 使われていなさそうなオブジェクトを洗い出す
--    del_ / del_new_ / new_ / _test / z_ などの名前が残っています。
--    消すかどうかは DB 側の判断です（本ファイルは参照のみ）。
-- ==================================================================
/*
USE ATOM;

SELECT name, type_desc, create_date, modify_date
FROM sys.objects
WHERE type IN ('P', 'V', 'FN', 'IF', 'TF', 'U')
  AND (name LIKE 'del[_]%' OR name LIKE 'new[_]%' OR name LIKE 'z[_]%' OR name LIKE '%[_]test')
ORDER BY type_desc, name;
*/

-- ==================================================================
-- 補助6. 別データベースを更新しているものを確実に洗い出す
--    SQL Server のパーサが解析した依存関係を使います。
--    is_updated が「本当に更新しているか」を表すため、
--    UPDATE t SET … FROM POEM.dbo.xxx AS t のように
--    別名を挟む書き方も正しく判定できます。
--
--    カーソルを使っているのは、解析できないオブジェクトで
--    処理が止まらないようにするためです。参照のみで、何も変更しません。
-- ==================================================================
/*
USE ATOM;

SET NOCOUNT ON;

DECLARE @results TABLE (
    [オブジェクト名] sysname,
    [種別]           nvarchar(60),
    [参照先DB]       sysname,
    [参照先]         nvarchar(300),
    [更新するか]     bit
);

DECLARE @id int, @name sysname;

DECLARE cur CURSOR LOCAL FAST_FORWARD FOR
    SELECT o.object_id, o.name
    FROM sys.objects AS o
    INNER JOIN sys.sql_modules AS m ON o.object_id = m.object_id
    WHERE o.type IN ('P', 'V', 'FN', 'IF', 'TF', 'TR');

OPEN cur;
FETCH NEXT FROM cur INTO @id, @name;

WHILE @@FETCH_STATUS = 0
BEGIN
    BEGIN TRY
        INSERT INTO @results
        SELECT
            @name,
            (SELECT type_desc FROM sys.objects WHERE object_id = @id),
            re.referenced_database_name,
            ISNULL(re.referenced_schema_name, '') + '.' + re.referenced_entity_name,
            re.is_updated
        FROM sys.dm_sql_referenced_entities(
                 OBJECT_SCHEMA_NAME(@id) + '.' + @name, 'OBJECT') AS re
        WHERE re.referenced_database_name IS NOT NULL
          AND re.referenced_database_name <> DB_NAME();
    END TRY
    BEGIN CATCH
        -- 解析できないものは読み飛ばす
    END CATCH

    FETCH NEXT FROM cur INTO @id, @name;
END

CLOSE cur;
DEALLOCATE cur;

SELECT DISTINCT
    [オブジェクト名], [種別], [参照先DB], [参照先],
    CASE WHEN [更新するか] = 1 THEN '★更新する' ELSE '参照のみ' END AS [操作]
FROM @results
ORDER BY [更新するか] DESC, [参照先DB], [オブジェクト名];
*/
