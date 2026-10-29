# tbl_t_shanghai_haiso_tehai_print_temp

既存DBの参照用スナップショット（読み取り専用）。取得日 2026-10-29 / DB: `ATOM`。

配送手配書の印刷用一時テーブル。**利用者ごと**に作り直して使う（menuNo=103 §10.3）。

| 順 | 列名 | 型 | 長さ(byte) | NULL |
|---:|---|---|---:|:---:|
| 1 | `pk_seq_no` | int | 4 | 不可 |
| 2 | `pk_login_code` | nvarchar | 12 | 可 |
| 3 | `ref_no` | nvarchar | 100 | 可 |
| 4 | `invoice_no_main` | nvarchar | 40 | 可 |
| 5 | `invoice_no` | nvarchar | 100 | 可 |
| 6 | `vessel` | nvarchar | 100 | 可 |
| 7 | `bl_date` | datetime | 8 | 可 |
| 8 | `haiso_type` | int | 4 | 可 |
| 9 | `container_type` | nvarchar | 60 | 可 |
| 10 | `container_no` | nvarchar | 40 | 可 |
| 11 | `seal_no` | nvarchar | 40 | 可 |
| 12 | `total_carton` | int | 4 | 可 |
| 13 | `total_quantity` | int | 4 | 可 |
| 14 | `haisyo_yotei_ymd` | datetime | 8 | 可 |
| 15 | `week_day` | int | 4 | 可 |
| 16 | `haiso_saki1` | int | 4 | 可 |
| 17 | `haiso_saki2` | int | 4 | 可 |
| 18 | `haiso_saki3` | int | 4 | 可 |
| 19 | `otsunaka_no` | int | 4 | 可 |
| 20 | `nisshin_plant` | int | 4 | 可 |
| 21 | `logistic_plant` | int | 4 | 可 |

**長さは byte 数**（`sys.columns.max_length`）。`nvarchar` なので文字数はその半分。
たとえば `invoice_no` は **50 文字**まで。

## 補足

- `pk_seq_no` は採番列と見られる（**要確認**）。INSERT では指定しない。
- `ref_no` / `nisshin_plant` / `logistic_plant` は `View_shanghai_haiso_tehai_pickup` に
  対応する列が無く、**どこから埋めていたかが不明**（§18 C-37）。
- `week_day` は `View__week_text.weekday_no` と結合するためのもの。
  Web 版は曜日を ATOM 側で組み立てるため使わない（§18 C-38）。
