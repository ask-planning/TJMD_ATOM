# ATOM リプレイス プロジェクト

Git管理。Access 2003 (.adp) 「ATOM」の一部画面を **VB.NET / WinForms（.NET Framework 4.8.1）＋ SVF** へ移行する。
SQL Server のストアドプロシージャ（ロジック）は現状維持し、VB.NETから同名で呼び出す。

## リポジトリ構成
- `docs/` … 設計・管理文書（Markdown正、Word/Excel併用）
- `src/`  … VB.NETソリューション（Atom.sln 以下）
- `db/`   … SQL Server資産の写し（参照用。本番非接続）
- `tools/`… 解析・変換ツール

## 技術前提（確定）
| 区分 | 内容 |
|---|---|
| UI | WinForms |
| 実行基盤 | .NET Framework 4.8.1 |
| 帳票 | SVF（設計=SVFX-Designer / 実行時=SVF Client） |
| DB | SQL Server（ストアド現状維持） |
| メール | .NET で再実装（現行BASP21を置換） |
| 文書形式 | Markdown / Excel（Word不要） |

## 移行スコープ（赤枠4ブロック=17メニュー）
対象フォーム40本 / 帳票18本 / ストアド93本。詳細は `docs/02_現行分析/ATOM_移行スコープ定義書.xlsx`。

## 主要ドキュメント
- ドキュメント体系・フォルダ構成: `docs/03_基本設計/`
- 画面設計テンプレート: `docs/_templates/`
- 画面デモ: `docs/_prototypes/`
