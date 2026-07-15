# ExpenseCodeMaintenanceForm（予定）
現行: F_M_経費コード登録_MAIN + _SUB を統合。
- DataGridView を tbl_m_import_keihi_code にバインド（BindingSource + SqlDataAdapter）
- 経費区分列 = DataGridViewComboBoxColumn（View__keihi_kbn）
- 更新時に update_ymd/update_login を自動セット（現行AfterUpdate相当）
- 権限: menuNo=907 の管理者フラグ / 戻り先: return_form_name（既定=諸経費入力）
設計書: docs/04_画面設計/経費/D1_画面設計_経費コード登録_v0.1.md
