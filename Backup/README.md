# ATOM（ASP.NET Web Forms 版）

旧 Microsoft Access 2003（.adp）アプリケーション「ATOM」の移行先です。

## 技術スタック

| 項目 | 内容 |
|---|---|
| フレームワーク | ASP.NET Web Forms / .NET Framework 4.8.1 |
| 言語 | VB.NET（`Option Strict On` / `Explicit On` / `Infer Off`） |
| データベース | SQL Server |
| 帳票 | SVF |

## プロジェクト構成

| プロジェクト | 役割 | 参照 |
|---|---|---|
| `Atom.Web` | 画面層（ASP.NET Web Forms） | Core / Data / Mail / Report |
| `Atom.Core` | エンティティ・業務ロジック・ログ抽象 | — |
| `Atom.Data` | データアクセス | Core |
| `Atom.Mail` | メール送信 | Core |
| `Atom.Report` | SVF 帳票 | Core |
| `Atom.Tests` | 単体テスト | Core / Web |

`Atom.Core` はデータアクセスに依存しません。ログの書き出し先は `ILogWriter` で抽象化し、
実装（`SqlLogWriter`）を `Atom.Data` に置いて `Global.asax` の起動時に登録しています。

## フォルダ構成

```
ATOM/
├ ATOM.sln
├ src/          ソースコード
├ db/schema/    DDL
├ docs/         コーディング規約・プロンプト集
└ tools/        補助ツール
```

## 起動手順

1. Visual Studio で `ATOM.sln` を開く
2. `Atom.Web` を「スタートアップ プロジェクト」に設定する
3. `src/Atom.Web/Web.config` の `connectionStrings` に接続先を設定する
   （マスクされた状態でコミットされています。実値はコミットしないでください）
4. `db/schema/*.sql` を対象データベースへ適用する
5. F5 で起動する

## 実装済みの画面

| menuNo | 画面名 | 状態 |
|---|---|---|
| — | メインメニュー（`Default.aspx`） | 実装済み |
| 907 | 輸入経費コード登録 | URL 登録のみ（画面未作成） |

その他のメニューは `src/Atom.Web/Common/PageUrlMap.vb` に URL を追記すると有効になります。

## 未実装・要確認事項

- ログイン画面と権限取得処理（現在 `Default.aspx.vb` の `SetTemporaryLogin` で仮設定）
- SVF のサーバー実行ライセンス（`Atom.Report/SvfReportGenerator.vb` 参照）
- `db/schema/m_import_expense_code.sql` は画面項目からの推測。既存テーブル定義との突き合わせが必要
- 単体テストの実行には NuGet で `MSTest.TestFramework` / `MSTest.TestAdapter` の追加が必要
