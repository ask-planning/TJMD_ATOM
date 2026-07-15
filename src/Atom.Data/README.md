# Atom.Data
SQL Server との境界（UIにSQLを置かない）。SQLはこの層に集約する。
- 接続基盤（現行 BaseConnectionString 相当）
- StoredProcedures/ : ストアド呼出ラッパ（usp_* 1本＝1メソッド。現状維持の契約）
- Repositories/     : べた書きSQL（実行SQL・バインド用SELECT）を集約。クラス内にインライン、値は SqlParameter 化
- TableGateways/    : 直接バインド系マスタ（例 tbl_m_import_keihi_code）の入出力
※ SQLの外部 .sql ファイル分離は行わない（Repositoriesにインライン統一）。
