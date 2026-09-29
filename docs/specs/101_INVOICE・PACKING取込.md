# ATOM_INVOICE・PACKING取込画面 設計書

## 1. 画面概要

### 1.1 画面名
**INVOICE／PACKING取込画面**

### 1.2 目的
INVOICEおよびPACKING LISTのExcelファイルを取り込み、内容チェックを実施したうえで、INVOICE・PACKING関連テーブルへ登録する。

現行Accessではフォルダパスおよびテキスト名を入力し、対象となる `_INV.xls` / `_PAC.xls` を検索しているが、Web版では**利用者がINVOICEファイルとPACKINGファイルを直接選択する方式**へ変更する。

### 1.3 Web化に伴う主な変更

| 項目 | Access版 | Web版 |
|---|---|---|
| PASS | フォルダパスを入力 | **廃止** |
| テキスト名 | ファイル検索用名称を入力 | **廃止** |
| INVOICEファイル | PASS＋テキスト名から自動検索 | **ファイル選択** |
| PACKINGファイル | PASS＋テキスト名から自動検索 | **ファイル選択** |
| REF_NO | 自動設定 | 自動設定 |
| 取込処理 | Access TransferSpreadsheet | サーバー側でExcelを解析 |
| メッセージ | Access MsgBox | Webダイアログ／メッセージ表示 |

---

# 2. 画面レイアウト

## 2.1 レイアウトイメージ

```text
┌───────────────────────────────────────────────────────────────────────┐
│ 使用者：XXXXXXXX       INVOICE/PACKING取込み     [INVヘッダ][INV確認][MENUへ] │
├───────────────────────────────────────────────────────────────────────┤
│                                                                       │
│ 発注先  ○ 上海ダンマ  ○ 庸助貿易  ○ 上海ダンマ(円)                   │
│                                                                       │
│ 通貨    ○ JPY        ○ RMB                                           │
│                                                                       │
│ INVOICEファイル  [ ファイルを選択 ]  AAAA_INV.xls                    │
│                                                                       │
│ PACKINGファイル  [ ファイルを選択 ]  AAAA_PAC.xls                    │
│                                                                       │
│ REF_NO            [ 自動表示                                    ]    │
│                                                                       │
│ ┌──────────────┐                                                      │
│ │   取り込み   │                                                      │
│ ├──────────────┤                                                      │
│ │   再実行     │                                                      │
│ ├──────────────┤                                                      │
│ │ 取込内容確認 │                                                      │
│ └──────────────┘                                                      │
│                                                                       │
└───────────────────────────────────────────────────────────────────────┘
```

---

# 3. 画面デザイン

現行画面の配色を基本的に踏襲する。

## 3.1 ヘッダー

| 項目 | 設定 |
|---|---|
| 背景色 | 濃紫 |
| 参考色 | `#400080` 相当 |
| タイトル文字色 | 白 |
| タイトル | `INVOICE/PACKING取込み` |
| タイトル配置 | 中央 |
| タイトル | 太字 |
| 操作ボタン背景 | ライトグレー |
| 操作ボタン文字 | 濃紫 |

※正確なRGB値についてはWeb画面実装時に現行画面との見た目を確認し微調整する。

## 3.2 入力エリア

背景色は白とする。

入力項目は横方向を揃え、INVOICEファイルとPACKINGファイルはファイル選択ボタンおよび選択済みファイル名を表示する。

## 3.3 REF_NO

利用者による直接入力は不可。

背景色をグレー系として、参照専用項目であることを視覚的に判別できるようにする。

---

# 4. 項目定義

| No | 項目名 | UI | 入力 | 必須 | 内容 |
|---:|---|---|---|---|---|
| 1 | 使用者 | 表示 | × | - | ログインユーザーを表示 |
| 2 | 発注先 | ラジオボタン | ○ | ○ | 発注先区分を選択 |
| 3 | 通貨 | ラジオボタン | ○ | ○ | JPY / RMB |
| 4 | INVOICEファイル | ファイル選択 | ○ | ○ | `_INV.xls` ファイル |
| 5 | PACKINGファイル | ファイル選択 | ○ | ○ | `_PAC.xls` ファイル |
| 6 | REF_NO | テキスト表示 | × | - | 取込データから自動生成 |
| 7 | 取り込み | ボタン | - | - | Excel取込およびチェック開始 |
| 8 | 再実行 | ボタン | - | - | 取り込んだデータを再チェック |
| 9 | 取込内容確認 | ボタン | - | - | 取込内容確認画面へ遷移 |
| 10 | INVヘッダ | ボタン | - | - | INVOICEヘッダ入力画面へ遷移 |
| 11 | INV確認 | ボタン | - | - | INVOICE内容確認画面へ遷移 |
| 12 | MENUへ | ボタン | - | - | メインメニューへ遷移 |

---

# 5. 発注先

発注先は単一選択とする。

画面表示上、以下を選択可能とする。

- 上海ダンマ
- 庸助貿易
- 上海ダンマ(円)

発注先に対応する内部値については、既存の `customer_kbn` / `shanghai_chotatsu` の設定を踏襲する。

発注先未選択で取込処理を実行した場合、以下を表示する。

> 発注先を選択してください。

---

# 6. 通貨

以下から単一選択する。

- JPY
- RMB

通貨未選択の場合、以下を表示する。

> 通貨を選択してください。

また、ファイル名と選択通貨の組み合わせをチェックする。

| ファイル名先頭 | 必要となる通貨 |
|---|---|
| `SJ-TJ` | JPY |
| `SJ-TR` | RMB |
| `YJ-TR` | RMB |

一致しない場合は以下を表示し、処理を中止する。

> INVOICEと通貨が一致しません。

---

# 7. ファイル選択仕様

## 7.1 INVOICEファイル

利用者が端末上のExcelファイルを選択する。

想定ファイル名：

```text
AAAA_INV.xls
SJ-9999_INV.xls
```

## 7.2 PACKINGファイル

利用者が端末上のExcelファイルを選択する。

想定ファイル名：

```text
AAAA_PAC.xls
SJ-9999_PAC.xls
```

## 7.3 ファイル形式

現行仕様を踏襲し、対象形式は以下とする。

```text
.xls
```

Excelファイルはヘッダー行ありとする。

現行処理ではExcelのA～R列を対象として取り込んでいるため、Web版についても現行Excelフォーマットを維持する。

## 7.4 INV/PAC組み合わせチェック

INVOICEファイルとPACKINGファイルから、それぞれ末尾の `_INV.xls` / `_PAC.xls` を除いた名称を取得する。

例：

```text
AAAA_INV.xls → AAAA
AAAA_PAC.xls → AAAA
```

両者が一致すること。

### 正常例

```text
AAAA_INV.xls
AAAA_PAC.xls
```

### エラー例

```text
AAAA_INV.xls
BBBB_PAC.xls
```

一致しない場合は以下を表示する。

> INVOICEとPACKINGのファイル名が一致しません。

※本チェックはWeb化に伴い追加するチェックとする。

---

# 8. 「取り込み」ボタン

## 8.1 処理概要

以下の順番で処理する。

1. 入力チェック
2. ファイルチェック
3. Excel読込
4. 一時データ作成
5. INVOICE/PACKINGチェック用データ作成
6. 再実行処理を実施

---

## 8.2 入力チェック

以下を確認する。

### 発注先

未選択の場合：

> 発注先を選択してください。

### 通貨

未選択の場合：

> 通貨を選択してください。

### INVOICEファイル

未選択の場合：

> INVOICEファイルを選択してください。

### PACKINGファイル

未選択の場合：

> PACKINGファイルを選択してください。

### 拡張子

`.xls` 以外の場合：

> Excelファイル（.xls）を選択してください。

### INV/PACファイル名

ベース名称が一致しない場合：

> INVOICEとPACKINGのファイル名が一致しません。

### 通貨整合性

ファイル名から判定される通貨と選択通貨が一致しない場合：

> INVOICEと通貨が一致しません。

---

# 9. Excel取込処理

## 9.1 処理順序

INVOICE → PACKINGの順で処理する。

### INVOICE

取込内容を以下のチェック用一時テーブルへ展開する。

```text
tbl_t_shanghai_invoice_import_check_temp
```

### PACKING

取込内容を以下のチェック用一時テーブルへ展開する。

```text
tbl_t_shanghai_packing_import_check_temp
```

ログインユーザー単位で処理対象を管理する。

---

# 10. INVOICE取込項目

現行処理ではINVOICE取込データから以下の項目をチェック用データへ設定する。

| 項目 |
|---|
| line_seq |
| so_no |
| so_no_seq |
| shanghai_code |
| shanghai_code_text |
| kikaku |
| part_sort |
| suryo_tani |
| suryo |
| unit_price |
| amount |
| dummy_no1 |
| dummy_no2 |
| tajima_po_no |
| R3_koubai_denpyo_no |
| invoice_no |
| bl_date |
| syukka_hoho |
| R3_hinmoku_code |
| tsuka |
| tajima_po_no_original |
| amount_text |
| err_status |

### 上海コード

数値の場合は5桁になるようゼロ埋めする。

### R3品目コード

数値の場合は18桁になるようゼロ埋めする。

### 通貨

選択された通貨から設定する。

---

# 11. PACKING取込項目

PACKINGについて以下の項目をチェック用データへ設定する。

| 項目 |
|---|
| true_or_false |
| invoice_no |
| tajima_po_no |
| shanghai_code |
| suryo |
| unit_price |
| amount |
| carton_no_from |
| carton_no_to |
| net_weight |
| gross_weight |
| m3 |
| bl_date |
| syukka_hoho |
| R3_hinmoku_code |
| tajima_po_no_original |
| err_status |

## 11.1 カートン番号補完

PACKING LISTのカートン番号について、開始番号が空白の場合は直前の開始番号を引き継ぐ。

終了番号が直前の終了番号と同一の場合は、重複値として扱わずNULLを設定する。

---

# 12. 再実行処理

## 12.1 処理概要

「再実行」はExcelファイルを再アップロードする処理ではなく、**一時領域へ取り込まれたINVOICE/PACKINGデータに対して業務チェックおよび本登録を再実行する処理**とする。

取り込みボタンによるExcel読込完了後にも、自動的に本処理を実行する。

---

# 13. 将来日付チェック

取り込んだINVOICEに将来月のBL DATEが存在するかチェックする。

使用処理：

```text
usp_shanghai_import_count_get_next_month
```

対象が存在する場合、確認ダイアログを表示する。

> 今月以降のBL DATEのINVOICEです。  
> 取り込みを中断しますか？

### 中断

処理終了。

### 続行

後続処理へ進む。

---

# 14. 登録済みチェック

以下により、対象INVOICE/PACKINGが登録済みか確認する。

```text
usp_shanghai_import_tourokuzumi_count
```

登録済みの場合、以下を確認する。

> 登録済みのINVOICEです。  
> 置換／追加／キャンセル

### 置換

既存データを削除して今回のデータへ置き換える。

### 追加

既存データを残した状態で処理を続行する。

### キャンセル

処理を中断する。

---

# 15. 置換処理

「置換」が選択された場合、対象INVOICEの既存データを削除する。

対象テーブル：

```text
tbl_t_shanghai_invoice_header
tbl_t_shanghai_invoice_header_sub
tbl_t_shanghai_invoice_meisai
tbl_t_shanghai_packing_meisai
```

削除完了後、今回の取込データを登録する。

**Web版では削除～再登録までを同一DBトランザクション内で実行し、登録途中でエラーが発生した場合はロールバックすること。**

---

# 16. REF_NO

取込対象のINVOICE_NOから表示用文字列を生成し、REF_NO欄へ設定する。

REF_NOは利用者による編集不可とする。

複数の商品群・INVOICEが存在する場合は、現行ロジックに従い `/` 等を使用して表示用文字列を編集する。

---

# 17. 重複チェック

貿易用INVOICE以外について、1つのPOが複数INVOICEに分割されていないかチェックする。

使用処理：

```text
usp_shanghai_import_jyufuku_check
```

該当する場合：

> 複数INVOICEに分かれているPOがあります。  
> POを変更し、再度実行してください。

処理を中断し、「取込内容確認画面」へ遷移する。

---

# 18. データチェック

## 18.1 エラー情報初期化

```text
usp_shanghai_import_check_err_clear
```

## 18.2 INVOICEチェック

```text
usp_shanghai_import_err_check_invoice
```

## 18.3 PACKINGチェック

```text
usp_shanghai_import_err_check_packing
```

## 18.4 金額チェック

INVOICE/PACKINGの個別チェックでERRORが存在しない場合、以下を実行する。

```text
usp_shanghai_import_check_amount
```

チェック結果は以下の3種類とする。

| 結果 | 処理 |
|---|---|
| ERROR | 取込不可 |
| WARNING | 利用者確認後、取込可能 |
| 正常 | 取込可能 |

---

# 19. ERROR時

ERRORデータが存在する場合：

> エラーデータがあります。  
> EXCELを確認してください。

本登録は実施せず、取込内容確認画面へ遷移する。

---

# 20. WARNING時

WARNINGのみ存在する場合：

> WARNINGデータがあります、このまま取り込みますか？

### 取り込み

本登録処理へ進む。

### キャンセル

本登録せず、取込内容確認画面へ遷移する。

---

# 21. 正常時

ERROR・WARNINGともに存在しない場合：

> 内容の確認を行いますか？

### 確認する

取込内容確認画面へ遷移する。

### 取り込む

本登録処理へ進む。

Web版ではダイアログの選択肢を以下のように明示する。

```text
[取込内容を確認] [このまま取り込む]
```

---

# 22. 本登録処理

## 22.1 INVOICE/PACKING登録

以下を実行する。

```text
usp_shanghai_import_invoice_update
```

ログインユーザーおよび担当区分を引数として、チェック済み一時データを本テーブルへ登録する。

処理結果が正常でない場合：

> INVOICE、PACKINGの取込に失敗しました。  
> EXCELデータをご確認ください。

---

# 23. INVOICEヘッダ作成

明細登録後、以下により合計情報を取得する。

```text
usp_shanghai_invoice_gokei_get
```

主に以下を取得する。

- BL DATE
- 出荷方法
- カートン数

取得結果を使用して以下へINVOICEヘッダを登録する。

```text
tbl_t_shanghai_invoice_header
```

主な登録項目：

| 項目 | 内容 |
|---|---|
| pk_invoice_no_main | メインINVOICE NO |
| REF_NO | NULL（現行仕様） |
| bl_date | 取得したBL DATE |
| vessel | NULL |
| syukka_hoho | 出荷方法 |
| carton_qty | カートン数 |
| return_pallet | 0 |
| otsunaka_no | 0 |
| shiiresaki_code | 発注先から取得 |
| update_ymd | 更新日時 |
| update_login | ログインユーザー |

---

# 24. 進捗更新

本登録完了後、対象INVOICEについて進捗状況を更新する。

現行処理：

```text
schedule_update(5, 処理日, INVOICE_NO, 1)
```

---

# 25. 処理完了

すべて正常終了した場合：

> 処理終了しました。

表示後、画面はINVOICE/PACKING取込画面に留まる。

---

# 26. 取込内容確認ボタン

押下時、「取込内容確認画面」へ遷移する。

現行遷移先：

```text
F_1_取込内容確認_MAIN
```

INVOICE内容確認用として以下を使用する。

```text
F_1_取込内容確認_SUB_INVOICE
```

遷移時に以下の情報を引き継ぐ。

- REF_NO
- 選択ファイル名／取込識別情報
- 通貨
- ログインユーザー情報

※現行の `text_name` はWeb版では廃止するため、必要な場合は選択したINVOICEファイルのベース名を引き継ぐ。

---

# 27. ヘッダボタン

## 27.1 INVヘッダ

INVOICEヘッダ入力画面へ遷移する。

現行：

```text
F_1_上海輸入_INVOICE_HEADER
```

## 27.2 INV確認

INVOICE内容確認画面へ遷移する。

現行：

```text
F_1_上海輸入_INVOICE_MAIN
```

## 27.3 MENUへ

メインメニューへ遷移する。

現行：

```text
F_0_START_MENU
```

Web版では各画面に対応するURLへ遷移する。

遷移先画面が未実装の場合は、システム共通処理として以下を表示する。

> 現在未実装です。

---

# 28. 初期表示処理

画面表示時に以下を実施する。

1. ログインセッション確認
2. ログインユーザー情報表示
3. 発注先選択肢取得
4. 通貨選択肢表示
5. INVOICEファイル未選択状態
6. PACKINGファイル未選択状態
7. REF_NO空白

現行Accessで実施しているPASS・テキスト名の初期値取得処理は、ファイル選択方式への変更に伴い**廃止**する。

---

# 29. セッション管理

Access版で使用していた以下のグローバル変数相当の情報は、Web版ではセッション等で管理する。

例：

```text
login_code
intShanghaiChotatsu
intMenuNo
intMenuNoSeq
strInvoiceNo
```

一時取込データについてもログインユーザー単位で識別する。

---

# 30. セッション切れ

処理実行時にセッション切れを検出した場合、ログイン画面を表示する。

ファイルアップロード後・取込確認中などにセッションが切れた場合を考慮し、可能な範囲で処理状態を一時保存する。

ただし、ブラウザのセキュリティ仕様上、ユーザーが選択したローカルファイルをセッション復帰後に自動的に再選択することはできない。

そのため、**サーバーへのアップロード完了前にセッションが切れた場合はファイルの再選択が必要**となる。

---

# 31. 二重実行防止

「取り込み」「再実行」の処理開始後は、処理完了まで対象ボタンを非活性とする。

画面上に処理中であることを表示する。

例：

```text
INVOICE/PACKINGファイルを取り込んでいます。
しばらくお待ちください。
```

同一ユーザーによる連打・二重登録を防止する。

---

# 32. トランザクション

本登録処理について、以下の処理は原則として一連のトランザクションとして扱う。

1. 既存データ削除（置換時）
2. INVOICE明細登録
3. PACKING明細登録
4. ヘッダ／サブヘッダ登録
5. INVOICEヘッダ作成
6. 進捗更新

途中でDBエラーが発生した場合はロールバックし、不完全な状態でデータを残さないこと。

---

# 33. ファイル取扱い

アップロードされたExcelファイルは、ブラウザからサーバーへ送信して処理する。

サーバー上で処理する際は、利用者が指定したローカルファイルパスには依存しない。

また、同名ファイルを複数ユーザーが同時に取り込んでも干渉しないよう、アップロードファイルはユーザー／処理単位で一意に管理する。

取込完了後、不要となった一時ファイルは削除する。

---

# 34. エラーメッセージ一覧

| No | 条件 | メッセージ |
|---:|---|---|
| 1 | 発注先未選択 | 発注先を選択してください。 |
| 2 | 通貨未選択 | 通貨を選択してください。 |
| 3 | INV未選択 | INVOICEファイルを選択してください。 |
| 4 | PAC未選択 | PACKINGファイルを選択してください。 |
| 5 | 拡張子不正 | Excelファイル（.xls）を選択してください。 |
| 6 | INV/PAC名称不一致 | INVOICEとPACKINGのファイル名が一致しません。 |
| 7 | 通貨不一致 | INVOICEと通貨が一致しません。 |
| 8 | INVフォーマット不正 | INVOICEのフォーマットが違います。 |
| 9 | PACフォーマット不正 | PACKINGのフォーマットが違います。 |
| 10 | INV読込エラー | INVOICEのEXCELの取り込み処理のエラーです。EXCELの内容等を確認してください。 |
| 11 | PAC読込エラー | PACKINGのEXCELの取り込み処理のエラーです。EXCELの内容等を確認してください。 |
| 12 | PO重複 | 複数INVOICEに分かれているPOがあります。POを変更し、再度実行してください。 |
| 13 | ERRORあり | エラーデータがあります。EXCELを確認してください。 |
| 14 | WARNINGあり | WARNINGデータがあります、このまま取り込みますか？ |
| 15 | 本登録失敗 | INVOICE、PACKINGの取込に失敗しました。EXCELデータをご確認ください。 |
| 16 | 正常終了 | 処理終了しました。 |
| 17 | 未実装画面 | 現在未実装です。 |

---

# 35. Web化時の廃止仕様

以下のAccess固有仕様はWeb版では使用しない。

| Access仕様 | Web版 |
|---|---|
| PASS入力 | 廃止 |
| テキスト名入力 | 廃止 |
| ローカルフォルダ直接参照 | 廃止 |
| `DoCmd.TransferSpreadsheet` | サーバー側Excel解析へ置換 |
| `DoCmd.Hourglass` | ローディング表示へ置換 |
| `MsgBox` | Webダイアログへ置換 |
| Accessフォームグローバル変数 | セッション等へ置換 |
| `form_open_close` | Web画面遷移へ置換 |

---

# 36. 処理フロー

```text
画面表示
   ↓
発注先選択
   ↓
通貨選択
   ↓
INVOICEファイル選択
   ↓
PACKINGファイル選択
   ↓
[取り込み]
   ↓
必須チェック
   ↓
拡張子チェック
   ↓
INV/PACファイル名チェック
   ↓
ファイル名－通貨整合チェック
   ↓
Excel解析
   ↓
一時テーブル作成
   ↓
INVOICE/PACKINGデータチェック
   ↓
登録済み？
 ┌─Yes─────────────┐
 │ 置換 / 追加 / キャンセル │
 └─────────────────┘
   ↓
業務チェック
   ↓
 ┌ ERROR ─────→ 取込内容確認
 │
 ├ WARNING ──→ 続行確認
 │
 └ 正常 ─────→ 内容確認する？
                    │
              ┌─────┴─────┐
             Yes           No
              ↓             ↓
        取込内容確認       本登録
                            ↓
                    INVOICE/PACKING登録
                            ↓
                      ヘッダ作成
                            ↓
                       進捗更新
                            ↓
                     「処理終了しました。」
```

---

# 37. 備考・要確認事項

### 37.1 Excel列定義

現行VBAから取込先項目および基本的な列位置は判断可能だが、Excelそのものの正式な列名・データ型・入力規則までは確定できない。

**実際のINVOICE/PACKING Excelサンプル入手後、ファイルレイアウト定義を別途追加すること。**

### 37.2 ストアドプロシージャ

本設計では既存ストアドプロシージャの処理内容は現行踏襲を前提とする。

以下などの内部仕様については別途DB設計／処理設計の対象とする。

```text
usp_shanghai_import_count_get_next_month
usp_shanghai_import_tourokuzumi_count
usp_shanghai_import_check_err_clear
usp_shanghai_import_jyufuku_check
usp_shanghai_import_err_check_invoice
usp_shanghai_import_err_check_packing
usp_shanghai_import_check_amount
usp_shanghai_import_invoice_update
usp_shanghai_invoice_gokei_get
```

### 37.3 ファイル形式

初期リプレイスでは現行互換性を優先し `.xls` を対象とする。

`.xlsx` 対応が必要となった場合は別途対応する。

---

## 確認事項・回答履歴

### 2026-08-13

| # | 確認内容 | 回答 | 反映箇所 |
|---|---|---|---|
| A-1 | Excel のファイルレイアウト（列位置・型） | サンプル `SJ-TJ7694_INV.xls` / `SJ-TJ7694_PAC.xls` を提示。**ただし開発環境では .xls を直接開けないため、A~R 列のヘッダー行と data 数行をテキスト（CSV 等）で別途受領する** | §37.1 |
| A-2 | Excel 解析ライブラリ（NuGet）の追加 | **追加可**。将来 `.xlsx` も発生し得るため、**.xls と .xlsx の両対応が望ましい**（NPOI を採用予定） | §7.3, §37.3 |
| A-3 | ストアドプロシージャの定義 | 後日提示 | §37.2 |
| A-4 | `schedule_update` の実装 | 現行 VBA を提示。**多数の画面から呼ばれるため共通処理として実装する** | §24 |
| A-5 | 一時テーブルの列定義 | 後日提示 | §9 |
| B-1 | 発注先の名称 | **「上海タジマ」が正**（設計書の「上海ダンマ」は誤記） | §5 |
| B-2 | 発注先の選択肢の取得方法 | **マスタから取得する**（現行 `usp_koumoku_master_pickup('COS')` を踏襲） | §5, §28 |
| B-3 | `customer_kbn` の値と意味、区分1＋通貨1 のとき区分3として仕入先を引く理由 | 未確定。DB 情報の提示待ち | §5 |
| B-4 | 通貨の内部値（2=JPY / 1=RMB） | 未確定（その対応で合っている見込み）。DB 情報または Access フォームの定義で裏取りする | §6 |
| B-5 | 調達／貿易区分（`shanghai_chotatsu`） | **画面で選択させる**。設計書§4 の項目定義に追加が必要 | §4, §16, §17, §22 |
| B-6 | 「今月以降」か「翌月以降」か | 未確定。`usp_shanghai_import_count_get_next_month` の定義（A-3）で確定させる | §13 |
| B-7 | 本登録失敗時の挙動 | **中断してロールバックする**（現行は続行していたが改める） | §22, §32 |
| B-8 | 「追加」時にヘッダを作成しない分岐の意味 | 未確定。`usp_shanghai_invoice_gokei_get` の定義（A-3）とヘッダテーブルのキー構成（A-5）で確定させる | §23 |
| B-9 | REF_NO の桁数ルール | **現行どおり**（貿易は `/` の位置まで、それ以外は5文字目が数値なら8文字・非数値なら9文字） | §16 |
| B-10 | 中間テーブル `tbl_t_shanghai_invoice_import_temp_<社員コード>` | **不要**。サーバーで Excel を解析し `check_temp` へ直接登録する | §9 |
| B-11 | PACKING が INVOICE 用の中間テーブルを使い回している件 | **不要**（B-10 により解消） | §9 |
| B-12 | 「再実行」単独実行 | 現行どおり維持。一時データが無いときは「取込データがありません。」を表示する | §12 |
| B-13 | INVOICE / PACKING の片方のみ取込 | **同時取込のみ**とする | §9 |
| B-14 | 通貨整合チェックの対象外ファイル名 | **現行どおり**（先頭が `SJ-TJ` / `SJ-TR` / `YJ-TR` 以外はチェックしない） | §6 |
| C-1 | 共通ヘッダーへの画面固有ボタン領域の追加 | **追加する**（共通ヘッダー部品設計書も改訂する） | §3.1, §27 |
| C-2 | 利用者名の表示位置 | **共通ヘッダーの仕様（右端）に合わせる** | §2.1 |
| C-3 | 画面の配置先 | `src/Atom.Web/Pages/Entry/` とする | — |
| C-4 | 画像にある黄色の説明パネル | **不要**（PASS・テキスト名の廃止に伴い） | §35 |
| C-5 | 確認ダイアログの実装方式 | **(a) 画面内モーダル＋ポストバックで処理を再開する方式** | §13, §14, §20, §21 |
| C-6 | アップロード先とサイズ上限 | 上限 10MB。保存先はテスト用に `C:\Users\itvendor06\Desktop\Input\ATOM\INVOICE取込`（`Web.config` の設定キーで管理し、本番は環境ごとに変更する） | §33 |
| C-7 | 同一ユーザーの同時実行 | 制限せず、二重送信防止のみ行う | §31 |
| D | 進め方 | **必要な情報が全部揃ってから着手する** | — |

### 2026-08-13（追記：ストアドプロシージャ定義の受領）

ストアド 17 本とビュー 5 本の定義を受領し、以下が確定した。

| # | 確認内容 | 判明した内容 | 反映箇所 |
|---|---|---|---|
| B-4 | 通貨の内部値 | **画面の 1 / 2 という内部値は不要。**`check_temp.tsuka` には `'JPY'` / `'RMB'` の文字列が入り、各チェックストアドもこの文字列で分岐している（`usp_shanghai_import_check_amount` 等）。Web版では通貨を **`'JPY'` / `'RMB'` の文字列で保持する** | §6, §10 |
| B-6 | 「今月以降」か「翌月以降」か | **翌月以降が正。**`usp_shanghai_import_count_get_next_month` は `WHERE BL_DATE >= @bl_date` で、Access は `@bl_date` に翌月1日を渡している。**メッセージの「今月以降」は誤記**のため、Web版は「翌月以降のBL DATEのINVOICEです。」とする | §13, §34 |
| B-8 | 「追加」時にヘッダを作成しない分岐の意味 | **二重登録の防止だった。**`usp_shanghai_invoice_gokei_get` の第8引数は `@header_count`（`tbl_t_shanghai_invoice_header` に対象 INVOICE のヘッダが既に存在する件数）。よってヘッダ作成条件は「**追加を選んでいない、かつ既存ヘッダが無い**」とき | §23 |
| B-3 | 発注先区分の意味 | **経路は判明。**画面の `customer_kbn` は `POEM.dbo.tbl_m_tokuisaki.customer_kbn_num` に対応し、`usp_POEM_tokuisaki_master_get`（第1引数=2）で `shiiresaki_code` を引いている。**値と発注先名の対応はマスタのデータ待ち** | §5, §23 |
| E-1 | `pk_invoice_no_main` の切り出し | **REF_NO 表示用とは別ロジック。**本登録（`usp_shanghai_import_invoice_update`）では `INVOICE_NO` の9文字目が数値なら先頭9文字、そうでなければ先頭8文字を `pk_invoice_no_main` とする。REF_NO 表示用（§16）の桁数判定（5文字目で判定）とは別物なので、**両方をそのまま実装する** | §16, §22 |
| E-2 | 重複品目チェックの副作用 | `usp_shanghai_import_err_check_invoice` / `_packing` の重複品目集計（`#jyufuku_hinmoku` への INSERT）に **`pk_login_code` の絞り込みが無く、全利用者の一時データを対象にしている**。同時刻に別の利用者が取込中だと誤検知し得る（既存の不具合。DB は変更しないため現状のまま） | §18 |
| E-3 | 補完処理の担当 | `usp_shanghai_import_check_err_clear` が R3品目コードのゼロ埋め・PO 未設定時の `CCBF+年月日` 補完・購買伝票番号の `9999999999` 補完まで行う。**アプリ側で重複して実装しない**（Access は購買伝票番号を二重に補完していた） | §10, §18.1 |
| E-4 | ビュー名の不一致 | `View_shanghai_import_invoice_no_packing` の定義文中の名前が `View_shanghai_invoice_no_packing_import` になっている（過去のリネームの痕跡）。動作影響は無いため記録のみ | — |
| E-5 | 一時テーブルの利用 | 複数のチェックストアドが内部で `CREATE TABLE #...` を行う。**アプリ側のトランザクションと同一接続で実行すれば問題ない** | §32 |

### 2026-08-13（追記：ストアドプロシージャの引数定義）

引数定義を受領した。`SqlParameter` の型と長さはこの表に従う（`sys.parameters.max_length` はバイト数のため、`nvarchar` の桁数はその半分）。

#### この画面（menuNo=101）で使うもの

| ストアド | 序数 | 引数 | 型 | 方向 |
|---|---:|---|---|---|
| `usp_shanghai_import_count_get_next_month` | 1 | `@login_code` | `nvarchar(6)` | IN |
| | 2 | `@bl_date` | `datetime` | IN |
| | 3 | `@next_count` | `int` | OUT |
| `usp_shanghai_import_tourokuzumi_count` | 1 | `@login_code` | `nvarchar(6)` | IN |
| | 2 | `@invoice_count` | `int` | OUT |
| | 3 | `@packing_count` | `int` | OUT |
| `usp_shanghai_import_check_err_clear` | 1 | `@login_code` | `nvarchar(6)` | IN |
| `usp_shanghai_import_jyufuku_check` | 1 | `@login_code` | `nvarchar(6)` | IN |
| | 2 | `@err_count` | `int` | OUT |
| `usp_shanghai_import_err_check_invoice` | 1 | `@login_code` | `nvarchar(6)` | IN |
| | 2 | `@err_count1`（ERROR 件数） | `int` | OUT |
| | 3 | `@err_count2`（WARNING 件数） | `int` | OUT |
| `usp_shanghai_import_err_check_packing` | 1〜3 | 上記と同じ | | |
| `usp_shanghai_import_check_amount` | 1〜3 | 上記と同じ | | |
| `usp_shanghai_import_invoice_update` | 1 | `@login_code` | `nvarchar(6)` | IN |
| | 2 | `@shanghai_chotatsu` | `int` | IN |
| | 3 | `@invoice_count` | `int` | OUT |
| | 4 | `@packing_count` | `int` | OUT |
| `usp_shanghai_invoice_gokei_get` | 1 | `@invoice_no` | `nvarchar(20)` | IN |
| | 2 | `@bl_date` | `datetime` | OUT |
| | 3 | `@syukka_hoho` | `int` | OUT |
| | 4 | `@total_carton` | `int` | OUT |
| | 5 | `@total_amount` | `float` | OUT |
| | 6 | `@total_suryo` | `float` | OUT |
| | 7 | `@container_text` | `nvarchar(50)` | OUT |
| | 8 | `@header_count` | `int` | OUT |
| `usp_POEM_tokuisaki_master_get` | 1 | `@syori_flg`（1=得意先コードから / 2=区分から） | `int` | IN |
| | 2 | `@tokui_code` | `nvarchar(6)` | IN |
| | 3 | `@customer_kbn` | `int` | IN |
| | 4〜13 | 得意先名・通貨・国等 | | OUT |
| | 14 | `@return_shiiresaki_code` | `nvarchar(10)` | OUT |
| `usp_koumoku_master_pickup` | 1 | `@koumoku_code` | `nvarchar(3)` | IN |
| | — | 戻りは**結果セット**（`pk_seq_no` / `contents1`〜`contents4` / `contents_num`） | | |

#### 進捗更新（`schedule_update` 相当）で使うもの

| ストアド | 序数 | 引数 | 型 | 方向 |
|---|---:|---|---|---|
| `usp_shinchoku_jyokyo_sonzai_check` | 1 | `@invoice_no` | `nvarchar(20)` | IN |
| | 2 | `@touroku_count` | `int` | OUT |
| `usp_shanghai_invoice_header_get` | 1 | `@shanghai_chotatsu` | `int` | IN |
| | 2 | `@invoice_no` | `nvarchar(20)` | IN |
| | 3 | `@bl_date` | `datetime` | OUT |
| | 4 | `@syukka_hoho` | `int` | OUT |
| `usp_shanghai_invoice_header_pickup` | 1 | `@invoice_no` | `nvarchar(20)` | IN |
| | — | 戻りは**結果セット** | | |
| `usp_calender_mode_get` | 1 | `@update_mode` | `int` | OUT |
| `usp_TJM_master_calender_AM_PM_sonzai_check` | 1 | `@check_ymd` | `datetime` | IN |
| | 2〜5 | `@container_AM` / `AM2` / `PM` / `PM2` | `nvarchar(50)` | OUT |
| `usp_koumoku_contents_get` | 1 | `@koumoku_code` | `nvarchar(3)` | IN |
| | 2 | `@seq_no` | `int` | IN |
| | 3〜4 | `@contents1` / `@contents2` | `nvarchar(128)` | OUT |
| | 5〜6 | `@contents3` / `@contents4` | `nvarchar(50)` | OUT |
| | 7 | `@contents_num` | `float` | OUT |

### 2026-08-13（追記：Excel レイアウトとテーブル定義）

#### INVOICE Excel の列マッピング（確定）

`SJ-TJ7694_INV.xls` の 1 行目はヘッダー（中国語）。2 行目以降がデータ。A〜R の 18 列が
`tbl_t_shanghai_invoice_import_check_temp` の INSERT 列順とそのまま一致する（F 列・M 列は非表示）。

| 列 | Excel の見出し | 登録先の列 | 型 | 備考 |
|---|---|---|---|---|
| A | 项号 | `line_seq` | `float` | 行番号 |
| B | 订单单号 | `so_no` | `nvarchar(50)` | 例 `SO1-2607000071` |
| C | 订单行号 | `so_no_seq` | `int` | |
| D | 产品编号 | `shanghai_code` | `nvarchar(50)` | **数値のときは5桁ゼロ埋め**。同じ値を18桁ゼロ埋めして `R3_hinmoku_code` にも設定 |
| E | 品名规格 | `shanghai_code_text` | `nvarchar(128)` | |
| F | （非表示） | `kikaku` | `nvarchar(128)` | |
| G | 其他 | `part_sort` | `nvarchar(2)` | 例 `A2` `A3` `P5`。**2文字しか入らない** |
| H | 销售单位 | `suryo_tani` | `nvarchar(4)` | 例 `PCS` |
| I | 实际出货数量 | `suryo` | `decimal(18,0)` | **小数を保持できない**（要確認 F-1） |
| J | 原币单价 | `unit_price` | `float` | |
| K | 原币税前金额 | `amount` | `float` | 同じ値を文字列で `amount_text` にも設定（小数桁チェック用） |
| L | 客户产品编号 | `dummy_no1` | `nvarchar(50)` | |
| M | （非表示） | `dummy_no2` | `nvarchar(50)` | |
| N | 客户订单单号 | `tajima_po_no` | `nvarchar(50)` | 同じ値を `tajima_po_no_original` にも設定 |
| O | 客户采购单号 | `R3_koubai_denpyo_no` | `nvarchar(50)` | **空のときは `'9999999999'`** |
| P | 发票号码 | `invoice_no` | `nvarchar(50)` | 例 `SJ-TJ7694A` |
| Q | 出货日期 | `bl_date` | `datetime` | 例 `26/08/14` |
| R | 交运方式 | `syukka_hoho` | `nvarchar(50)` | 例 `SEA`。**文字列で登録**し、後段で `View__syukka_hoho` により数値化される |

`tsuka` には画面で選択した通貨（`'JPY'` / `'RMB'`）、`err_status` には `0` を設定する。

#### PACKING Excel の列マッピング（確定）

`SJ-TJ7694_PAC.xls` の 1 行目はヘッダー。A〜N の 14 列。

| 列 | Excel の見出し | 登録先の列 | 型 | 備考 |
|---|---|---|---|---|
| A | T/F | `true_or_false` | `nvarchar(10)` | 値は `F` |
| B | 发票号码 | `invoice_no` | `nvarchar(50)` | 例 `SJ-TJ7694A` |
| C | 客户订单单号 | `tajima_po_no` | `nvarchar(50)` | 同じ値を `tajima_po_no_original` にも設定 |
| D | 产品编号 | `shanghai_code` | `nvarchar(50)` | 18桁ゼロ埋めして `R3_hinmoku_code` にも設定 |
| E | 产品数量 | `suryo` | `decimal(18,0)` | |
| F | 原币单价 | `unit_price` | `float` | |
| G | 原币税前金额 | `amount` | `float` | |
| H | 起始包装箱号 | `carton_no_from` | `nvarchar(50)` | 例 `0001`（4桁ゼロ埋めの文字列）。§11.1 の補完対象 |
| I | 截止包装箱号 | `carton_no_to` | `nvarchar(50)` | 例 `0002`。§11.1 の補完対象 |
| J | 总净重(Kg) | `net_weight` | `float` | |
| K | 总毛重(Kg) | `gross_weight` | `float` | |
| L | 总材积(Cuft) | `m3` | `float` | **見出しの単位は Cuft（立方フィート）だが列名は `m3`。**値はそのまま登録する |
| M | 出货日期 | `bl_date` | `datetime` | |
| N | 交运方式 | `syukka_hoho` | `nvarchar(50)` | |

`err_status` には `0` を設定する。PACKING には `tsuka` の列が無い。

カートン番号の補完（§11.1）は H 列・I 列に適用する。`R3_hinmoku_code` は D 列から、
`tajima_po_no_original` は C 列から設定し、`err_status` は `0` とする。

#### 発注先区分（B-3 の一部が確定）

`usp_koumoku_master_pickup 'COS'` の結果。

| `pk_seq_no` | `contents1` | `contents2`（画面表示） |
|---:|---|---|
| 0 | `ZZ` | 全て |
| 1 | `SH` | 上海タジマ |
| 2 | `YS` | 庸助貿易 |
| 9 | `KN` | 上海タジマ（円） |

画面のラジオは `1` / `2` / `9` を使う（`0`「全て」は取込画面では使わない）。

`POEM.dbo.tbl_m_tokuisaki` 側の区分は「**得意先 × 通貨**」の組み合わせになっている。

| `customer_kbn_num` | 略号 | 得意先 | 通貨 | 得意先コード | `shiiresaki_code` |
|---:|---|---|---|---|---|
| 1 | SH | 上海タジマ | JPY | 897404 | 133128 |
| 3 | SH | 上海タジマ | RMB | 897405 | 133128 |
| 2 | YS | 庸助貿易 | JPY | 897861 | 143534 |
| 4 | YS | 庸助貿易 | RMB | 897865 | 143534 |
| 9 | KR | 上海タジマ | JPY | 999999 | 144142 |

これで「区分1 かつ RMB のとき 3 に読み替える」理由が判明した。ただし
**`shiiresaki_code` は同一得意先なら通貨によらず同じ値**（1 と 3 はともに 133128、2 と 4 はともに 143534）。

現行コードは区分1のときだけ読み替えており、庸助貿易（区分2）で RMB を選んでも区分4 へ読み替えない。
上記のとおり `shiiresaki_code` は同じなので**この画面での実害は無い**。現行ロジックをそのまま踏襲する。

#### 通貨（B-4 確定）

実データの `LEFT(pk_invoice_no, 5)` と `tsuka` の対応。

| 先頭5文字 | 通貨 | 件数 | 現行の整合チェック対象 |
|---|---|---:|---|
| `SJ-TJ` | JPY | 49,243 | ○ |
| `SJ-TR` | RMB | 193,529 | ○ |
| `YJ-TR` | RMB | 26,677 | ○ |
| `SJ-T3` | JPY | 7,670 | ×（チェックされない） |
| `SJ-T4` | JPY | 47,540 | ×（同上） |
| `YJ-T0` | JPY | 2,675 | ×（同上） |
| `YJ-TJ` | JPY | 997 | ×（同上） |

#### テーブル定義から判明した実装上の注意

| # | 内容 |
|---|---|
| 1 | 両 `check_temp` の `pk_seq_no` は **IDENTITY**。INSERT の列に含めない |
| 2 | `tbl_t_shanghai_invoice_header` の**主キーは `pk_invoice_no_main`**（`pk_key` は IDENTITY だが PK ではない）。1 INVOICE につきヘッダは 1 件。B-8 の二重登録防止と整合する |
| 3 | `check_temp.syukka_hoho` は `nvarchar(50)`（`SEA` 等の文字列）。数値化は `usp_shanghai_invoice_gokei_get` が `View__syukka_hoho` 経由で行う |
| 4 | `check_temp.carton_no_from` / `_to` は `nvarchar(50)`。本テーブル `tbl_t_shanghai_packing_meisai` 側は `int`。文字種チェックは `usp_shanghai_import_err_check_packing` が行う |
| 5 | `check_temp.part_sort` は **`nvarchar(2)`**。Excel の G 列（`A2` 等）は2文字だが、3文字以上が来ると切り詰めまたはエラーになる |
| 6 | `tbl_t_import_progress` に**主キー・一意制約が無い**。`schedule_update` は存在チェックしてから INSERT している |
| 7 | `tbl_m_login_sys.shain_name` は `nvarchar(50)`。既存の `LoginRepository` が `nvarchar(20)` としていたため **50 に修正済み**（ストアド `usp_syain_jyoho_get` の `@login_name` が 20 だったための誤り） |

#### 新たな要確認事項

| # | 箇所 | 内容 | こちらの想定（推測） |
|---|---|---|---|
| # | 箇所 | 内容 | 回答 |
|---|---|---|---|
| F-1 | 数量の小数 | `check_temp.suryo` が `decimal(18,0)`、`tbl_t_shanghai_invoice_meisai.suryo` が `int` のため小数を保持できない | **現行どおり**（DB 側の丸めに任せる） |
| F-2 | PACKING レイアウト | 実ファイルでの確認 | **確定**（A〜N の14列。上表のとおり） |
| F-3 | 通貨チェックの対象 | `SJ-T3` / `SJ-T4` / `YJ-T0` / `YJ-TJ` は現行の整合チェック対象外 | **現行どおり追加しない**（B-14 の回答どおり） |
| F-4 | 発注先の区分値 | `customer_kbn_num` と `shiiresaki_code` の対応 | **確定**（上表のとおり） |

#### 未受領（実装前に必要）

| # | 内容 | 状況 |
|---|---|---|
| G-1 | `syori_log_create` のコード | **受領。**登録先は `tbl_t_syori_log`（`pk_menu_no` / `pk_menu_no_seq` / `invoice_no` / `syori_naiyo` / `syori_ymd` / `syori_login`） |
| G-2 | `schedule_update` の VBA | **受領。**`docs/archive/ATOM_Access_共通処理_schedule_update.txt` に原文保管 |

### 2026-08-14（実装完了）

#### 実装したファイル

| 層 | ファイル | 役割 |
|---|---|---|
| 画面 | `src/Atom.Web/Pages/Entry/InvoiceImport.aspx`（`.vb` / `.designer.vb`） | 画面と入力チェック、確認ダイアログの制御 |
| 共通 | `src/Atom.Web/Common/HeaderAction.vb` | ヘッダーの画面固有ボタン |
| 業務 | `src/Atom.Core/Services/ShanghaiImportExcelParser.vb` | Excel の解析（NPOI 2.5.6） |
| 業務 | `src/Atom.Core/Services/ShanghaiInvoiceNoFormatter.vb` | REF_NO と pk_invoice_no_main の編集 |
| 業務 | `src/Atom.Core/Entities/ShanghaiImportRow.vb` ほか | 取込行・チェック結果・通貨 |
| 手順 | `src/Atom.Data/Services/ShanghaiImportService.vb` | 一連の手順とトランザクション |
| データ | `src/Atom.Data/Sql/ShanghaiImportSql.vb`／`Repositories/ShanghaiImportRepository.vb` | 一時テーブル・置換削除・ヘッダ登録・ストアド呼び出し |
| データ | `src/Atom.Data/Sql/ImportProgressSql.vb`／`Repositories/ImportProgressRepository.vb` | 進捗更新（処理NO 5） |

#### 設計書からの逸脱

| # | 箇所 | 設計書の記載 | 実装 | 理由 |
|---|---|---|---|---|
| 1 | §7.3, §8.2 | 対象形式は `.xls` のみ。拡張子が違うとき「Excelファイル（.xls）を選択してください。」 | `.xls` と `.xlsx` の両方を受け付け、文言は「Excelファイル（.xls / .xlsx）を選択してください。」 | 回答 A-2「今後 xlsx も発生する可能性があるので両方対応可能だと望ましい」による |
| 2 | §13, §34 | 「今月以降のBL DATEのINVOICEです。」 | 「翌月以降のBL DATEのINVOICEです。」 | 回答 B-6。ストアドの判定は翌月1日以降のため |
| 3 | §22 | 本登録に失敗してもメッセージ表示のみで処理を続行 | ロールバックして中断 | 回答 B-7 |

#### 未実装として残した部分

| # | 内容 | 影響 |
|---|---|---|
| 1 | **取込内容確認画面への遷移**。この画面はメニューに載らないため menuNo が未採番 | エラー時・確認時に「現在未実装です。」を表示するだけ。エラー内容を画面で確認できない |
| 2 | **進捗更新のカレンダー自動更新**（処理NO 10 の経路）と `syori_log_create` | この画面からは呼ばれないため影響なし。配送手配書作成の画面を作るときに追加する |
| 3 | **権限による分岐** | 全メニューが押下可能（共通の課題） |

#### 要確認事項（残り）

| # | 内容 | 回答 |
|---|---|---|
| I-1 | 取込内容確認画面に menuNo を採番するか | **採番する。menuNo = 108。**メニューには表示しないため、画面定義に `"Hidden": true` を追加して登録した。画面を実装したら `Url` を設定するだけで遷移するようになる |

#### テスト仕様書

`docs/UT/ATOM_テスト仕様書.xlsx` の「INVOICE取込画面」シートに 57 件を起票済み。
未実装の 2 件（取込内容確認画面への遷移）は状態を「未実装」としている。

### 2026-08-21（実データによる動作確認）

`SJ-TJ7699_INV.xls` / `SJ-TJ7699_PAC.xls` で取り込みから本登録までを通し確認した。

#### 照合結果

| 項目 | ファイルの値 | 登録された値 | 判定 |
|---|---|---|---|
| 明細件数 | 51 | 一時 51 / 本明細 51（INVOICE・PACKING とも） | 一致 |
| INVOICE 番号 | `SJ-TJ7699A` / `SJ-TJ7699E` | サブヘッダ 2 件 | 一致 |
| `pk_invoice_no_main` | — | `SJ-TJ7699`（9 文字目が数値のため先頭 9 文字） | 正しい |
| REF_NO | — | `SJ-TJ7699A/E` | 正しい |
| カートン数 | 161（A=148 / E=13） | `carton_qty` = 161 | 一致 |
| 出货日期 | `26/08/21`（文字列） | `bl_date` = 2026-08-21 | 一致 |
| 交运方式 | `SEA` | `syukka_hoho` = 0 | 一致（`View__syukka_hoho` で SEA = 0） |
| 発注先 | 上海タジマ / JPY | `shiiresaki_code` = 133128 | 一致 |
| 進捗 | — | `shinchoku_jyokyo` = 3、日付 3 項目に当日、`syurui` = 1（出荷方法 + 1） | 正しい |

#### Excel の列定義（実ファイルで確定）

非表示だった 2 列の内容も確定した。推定どおりだった。

| 列 | 見出し | 登録先 |
|---|---|---|
| F | 规格 | `kikaku` |
| M | CKD工单编号 | `dummy_no2` |
| G | 其他分群码 二 | `part_sort`（`A2` / `E1` などの 2 文字） |

出货日期は**日付書式ではなく `26/08/21` の文字列**で入っている。解析側で `yy/MM/dd` として読めている。

なおカートン番号は、このファイルでは開始の空欄も終了の重複も 0 件だったため、
**補完ロジック（§11.1）は未検証**。該当データを含むファイルで別途確認が必要。

#### 出荷方法の区分値

| 値 | 内容 |
|---:|---|
| 0 | SEA |
| 1 | AIR |
| 2 | HDS |
| 3 | TRUCKAGE |
| 4 | DHL |
| 5 | HAND CARRY |
| 9 | その他 |

### 2026-08-21（判明した不具合と要確認事項）

いずれも**旧 Access 版から存在するもの**で、データベース側の変更が必要なため当方では実施していない。

| # | 内容 | 影響 | 対応案 |
|---|---|---|---|
| J-1 | `usp_shanghai_import_err_check_invoice` の「購買伝票番号」の UPDATE（定義 156 行目付近）に **`pk_login_code` の絞り込みが無い**。他の利用者の行にも ` W:購買伝票番号無し` を追記し続け、`err_contents`（`nvarchar(100)`）を超えると **エラー 8152 で取込が失敗する** | 複数人で使うと必ず発生する。実際に本開発中に発生した | 当該 UPDATE の WHERE に `pk_login_code = @login_code` を追加する。**要判断** |
| J-2 | 同ストアドの重複品目チェックが、`#jyufuku_hinmoku` へ**全利用者分を集計**している（`INSERT` に絞り込みが無い） | 他の利用者のデータで重複品目と誤検知し得る | 同上。**要判断** |
| J-3 | 出荷方法の **0 は SEA を意味する**ため、集計で取得できなかったときの既定値 0（現行仕様）と区別できない | 今回は結果的に正しい値だが、取得失敗が SEA として登録される余地がある | 現行踏襲とするか、取得失敗を区別するか。**要判断** |

#### 暫定の回避策（実施済み）

J-1 の発生時は、一時テーブルの古い行を削除すれば取り込めるようになる。

```sql
DELETE FROM tbl_t_shanghai_invoice_import_check_temp WHERE pk_login_code <> '自分の社員コード';
DELETE FROM tbl_t_shanghai_packing_import_check_temp WHERE pk_login_code <> '自分の社員コード';
```

#### 進捗更新（`schedule_update`）— 今回実装する経路

この画面からの呼び出しは `schedule_update(5, 処理日, pk_invoice_no_main, 1)` の一通りのみ。
`syori_no = 5`（上海連絡・ヘッダ）／`shanghai_chotatsu = 1`（上海）の経路だけを実装し、
他の経路（`syori_no` = 3 / 4 / 6〜17）は該当画面の実装時に追加する。

処理順序は以下のとおり。

1. `pk_invoice_no_main` が空または `"NULL"` のときは何もしない
2. INVOICE_NO が `SJ-` または `YJ-` 始まりのとき（上海経路）
   - `usp_shanghai_invoice_header_pickup` で `syukka_hoho` と `bl_date` を取得し、**種類 = `syukka_hoho` + 1**
   - 取得できないときは INVOICE_NO から種類を決める
     - `SJ-` かつ 5 文字目が `2` → `99`
     - 4 文字目が `R` → `2`（RMB 建て）
     - 上記以外 → `1`
3. `usp_shanghai_invoice_header_get(1, INVOICE_NO)` で `bl_date` を取り直す
4. `usp_shinchoku_jyokyo_sonzai_check` で進捗データの有無を確認し、無ければ
   `tbl_t_import_progress` へ INSERT（`pk_invoice_no_main` / `syurui` / `syukko_ymd` = BL DATE / `update_ymd` / `update_login`）
5. `tbl_t_import_progress` を UPDATE
   - `syukko_ymd` = BL DATE、`shanghai_syorui_ymd` = 当日、`shanghai_header_ymd` = 当日、`shinchoku_jyokyo` = 3
   - 条件は `pk_invoice_no_main` 一致 **かつ `shanghai_header_ymd IS NULL`**（二重更新の防止）

**カレンダー自動更新は `syori_no = 10`（配送連絡→ロジ）のときだけ動くため、この画面では動作しない。**
そのため `usp_calender_mode_get` / `usp_TJM_master_calender_AM_PM_sonzai_check` /
`usp_koumoku_contents_get('HKB')` / `tbl_m_calender_jigyosyo` の更新は今回の実装対象外とする。

`syori_log_create` も上記カレンダー更新の中でのみ呼ばれるため、今回は実装対象外。

**移植時の注意（現行 VBA の問題点）**

- 値を文字列連結で SQL に埋め込んでいる。移植先ではすべて `SqlParameter` にする
- カレンダー更新では列名の一部（`AM` / `AM2` / `PM` / `PM2`）を連結している。実装するときはホワイトリスト照合が必須
- `On Error Resume Next` でエラーを無視している箇所がある。移植先ではログ記録のうえ、業務を止めるか続けるかを明示する
- 進捗更新で使う BL DATE は Access のグローバル変数 `datBLDate` 経由で、取得できないと**前の処理の値が残る**。移植先では明示的に受け渡す

#### 実装前の確認事項

| # | 箇所 | 内容 | こちらの想定（推測） |
|---|---|---|---|
| H-1 | フォルダ構成 | 一連の取込手順（Excel 解析 → 一時登録 → チェック → 本登録 → 進捗更新）をまとめる置き場が基本方針§3 に無い。`.aspx.vb` に書くと §5.2 違反、`Atom.Core/Services` は `Atom.Data` を参照できない。**`Atom.Data/Services/` の新設**を許可してほしい | 新設を許可 |
| H-2 | NuGet | NPOI を `Atom.Core` へ追加する（Excel 解析）。`Atom.Core` は DB に依存しないという方針には反しない | `Atom.Core` に追加 |

---

## 2026-09-25（本登録ストアドの解析）

`usp_shanghai_import_invoice_update`（ATOM_dev）のソースを確認し、これまで「ストアドの中身は未確認」としていた点を確定させた。

### 担当（`@shanghai_chotatsu`）の用途

**このストアドでは `boueki_flg` を決めるためだけに使われる。**

```sql
boueki_flg =
  CASE WHEN @shanghai_chotatsu = 3 THEN 1             -- 貿易なら必ず 1
  ELSE
    CASE WHEN LEFT(tajima_po_no, 3) IN ('NSN','HON')  -- それ以外は PO の頭3文字で判定
         THEN 1 ELSE 0 END
  END
```

PACKING 側の登録では一切参照していない。

したがって **1=上海 と 2=調達 はこのストアドでは同じ動き**で、違うのは 3=貿易 だけ。
画面側の分岐（REF_NO の桁数判定、PO 重複チェックの有無）とあわせても、
担当の実質的な意味は「**貿易かどうか**」である。

`boueki_flg` は輸入仕入実績（menuNo=107）で `= 1` / `<> 1` により
処理を完全に別ルートへ振り分ける列のため、ここを誤ると後工程が変わる。

### 登録先と流れ

| 登録先 | 内容 |
|---|---|
| `tbl_t_shanghai_invoice_meisai` | INVOICE 明細 |
| `tbl_t_shanghai_packing_meisai` | PACKING 明細 |
| `tbl_t_shanghai_invoice_header_sub` | ヘッダサブ（箱数のみ。既にあれば作らない） |

`pk_invoice_no_main` の切り出し規則（9文字目が数値なら9文字、違えば8文字）は
画面側の `ShanghaiInvoiceNoFormatter.GetInvoiceNoMain` と同じ。

### 登録対象の条件

```sql
WHERE pk_login_code = @login_code
  AND tajima_po_no is not null
  AND suryo is not null
  AND ( shanghai_code is not null OR ... )
  AND View_shanghai_invoice_no_invoice.pk_invoice_no is null   -- 既登録の番号を除外
```

- **`err_status` を見ていない。**PO・数量・品目コードが揃っていればエラー行も登録される。
  現在は画面側でエラー時に本登録へ進ませないため実害は無いが、ストアド単体には防御が無い
- 最後の条件が**二重登録の防止**にあたる。「追加」を選んでも同じ INVOICE 番号は二重に入らない
- WHERE 句の `shanghai_code = 'MISC'` は `is not null` に必ず含まれるため**意味を持たない**
  （旧来からの記述。害は無いが意図と実際の動きがずれている可能性。保守マップ §8-42）

### 項目の入り方で注意が要るもの

| 項目 | 入り方 |
|---|---|
| **PO** | `tajima_po_no` → `syanai_sansyo_no`、`tajima_po_no_original` → `syanai_sansyo_no_original` の**2列に入る** |
| **箱数** | `SUM(carton_no_to - carton_no_from + 1)`。**引き算**のため C/NO は数値である前提 |
| **MISC / FREIGHT** | 品目コードが空＋品名が `FREIGHT` → 品目コードを `MISC` に。品目コードが `MISC` ＋品名が空 → 品名を `FREIGHT` に。運賃行を通すための読み替え |
| 固定値 | `del_flg`=0 / `nukitori_flg`=0 / `haiso_type`=0 / `container_type`=0 / `container_no`・`seal_no`=NULL |

### 取込結果確認画面（108）の編集との関係

この解析により、108 画面の編集で壊れうる箇所が 2 つ判明した。詳細は
`docs/specs/108_取込結果確認.md` の 2026-09-25 の節に記載。
