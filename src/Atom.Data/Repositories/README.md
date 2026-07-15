# Repositories
この層に SQL を一本化する（UIにSQLを置かない）。実行SQL・バインド用SELECTとも、
機能・テーブル別のRepositoryクラス内にインラインで記述する。

規約:
- SQLはクラス先頭の Private Shared ReadOnly 定数にまとめる（長文は VB の XMLリテラルで複数行可）。
- 値は必ず SqlParameter 化（"...=" & v は禁止 → "=@v"）。
- 動的テーブル名はパラメータ化不可 → 許可名のホワイトリスト検証を通す。
- クロスDB(TJM_master)参照は接続/権限方針に留意。
例: ExpenseCodeRepository, ShanghaiInvoiceRepository, ProgressRepository, CalendarRepository。
