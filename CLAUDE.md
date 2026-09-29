# CLAUDE.md — ATOM（TJM_ATOM）

このリポジトリで作業する AI エージェント向けの指示書。

## 必読

**作業前に必ず [`docs/00_開発基本方針.md`](docs/00_開発基本方針.md) を読むこと。**
規約・フォルダ構成・命名規則・禁止事項の正本はそちら。本ファイルはその要約で、矛盾する場合は基本方針を正とする。

**どのファイルを触ればよいかは [`docs/02_保守マップ.md`](docs/02_保守マップ.md) を見ること。**
機能から担当ファイルを逆引きできる。**実装を変更したら、同じ作業の中で保守マップも更新する**（更新ルールは同ファイル §9）。

## 技術スタック

ASP.NET Web Forms / VB.NET / .NET Framework 4.8.1 / SQL Server（**既存DB**）/ 帳票 SVF / Windows 認証

| プロジェクト | 役割 |
|---|---|
| `Atom.Web` | 画面層 |
| `Atom.Core` | エンティティ・業務ロジック・ログ抽象 |
| `Atom.Data` | データアクセス（`Sql/` と `Repositories/`） |
| `Atom.Mail` | メール送信 |
| `Atom.Report` | SVF 帳票 |
| `Atom.Tests` | 単体テスト |

## 判断の優先順位

**① 画面設計書（`docs/specs/`） → ② 本方針 → ③ 既存コード**

設計書と既存コードが食い違うときは設計書に従い、食い違いを報告する。既存コードが規約に反していても、触った箇所だけ寄せる。`Backup/` は旧版で参照可・編集禁止。

## 絶対ルール

1. **DBを勝手に変更しない。** DDL（CREATE / ALTER / DROP / TRUNCATE）の生成・実行、テーブルや列の追加を前提にした実装をしない。不足があれば実装を止めて「要確認」で報告する。
2. **本番環境（DB / IIS / ファイルサーバ）への直接操作を提案も実行もしない。**
3. **接続文字列・パスワード・個人情報を出力に含めない。** 例示は `Password=***` にマスクする。
4. **SQL の値は必ず `SqlParameter`。** 文字列連結・`Replace` による埋め込みは全面禁止。
5. **推測で実装しない。** 読み取れない箇所は「要確認」として質問する。推測を述べるときは「推測」と明示する。
6. **依頼されていないファイルを変更しない。** 差分は最小に。
7. **ファイル作成前に基本方針のフォルダ構成を参照し、定められた場所に置く。** 該当が無ければ勝手に新設せず確認する。
8. 破壊的操作を含むコードを生成したら必ず明記する。
9. 1回の変更で扱う目的は1つだけ。

## 画面作成の手順（省略しない）

画面の作成・改修は**必ず画面設計書（md）を受け取ってから**着手する。設計書が無ければ提示を求める。

1. 受け取った設計書を `docs/specs/<menuNo>_<画面名>.md` に**そのまま保存する**（原文を改変しない）
2. 本方針・既存コードと突き合わせ、矛盾・不足・曖昧な点を洗い出す
3. 「要確認事項」として番号付きで質問する（推測で埋めない）
4. 回答を得たら設計書末尾の「確認事項・回答履歴」に日付付きで**追記して保存し直す**
5. 確定した設計書に基づいて実装する

## フォルダ構成

```
TJM_ATOM/
├ CLAUDE.md
├ ATOM.sln
├ docs/            00_開発基本方針.md / specs/（設計書）/ archive/
├ src/
│  ├ Atom.Web/     Pages/{Master,Entry,Search,Report} Controls/ Common/
│  │               Content/ Scripts/ Images/ Site.Master Global.asax Web.config
│  ├ Atom.Core/    Entities/ Services/ Logging/ Exceptions/
│  ├ Atom.Data/    Database.vb Sql/ Repositories/ Logging/
│  ├ Atom.Mail/    Atom.Report/    Atom.Tests/
├ db/schema/       既存DBの参照用スナップショット（読み取り専用）
├ tools/
└ Backup/          旧版一式（参照可・編集禁止）
```

新規画面は `src/Atom.Web/Pages/<領域>/` の下に置く。`.aspx` + `.aspx.vb` + `.aspx.designer.vb` の3点セット。

## 画面デザインはマークアップで定義する

**Visual Studio のデザイナで開いて編集できる状態を保つこと。**

- コントロールの配置・種類・見た目は `.aspx` / `.ascx` に**静的に**書く
- コードビハインドでの動的生成を禁止（`New TextBox()` / `Controls.Add` / `Columns.Add`）
- 出し分けはマークアップに全パターンを書き、`Visible` / `Enabled` を切り替える
- **例外**（理由をコメントに書けば可）: 構造上マークアップで表現できない場合／ポストバック・イベント発生後の状態変更（`Click`・`RowDataBound` 等での `Visible`・`Text`・`DataSource` 変更）／`FindControl` での値詰め込み
- `.aspx.designer.vb` は手編集しない

## VB.NET 要点

- `Option Strict On` / `Explicit On` / `Infer Off` を全ファイルで宣言。UTF-8（BOM付き）/ CRLF / スペース4
- 命名: クラス・メソッド・定数 = PascalCase / 引数・ローカル = camelCase / フィールド = `_camelCase`
- VB識別子にスネークケース禁止（DBのテーブル名・列名はスネークのまま）。ローマ字はヘボン式に統一
- `Me.` を常に付ける。文字列結合は `&`（`+` 禁止）。`Call` 禁止
- リソースは `Using`。再送出は `Throw`（**`Throw ex` 禁止**）。空 `Catch` 禁止
- Public 要素に日本語 XML コメント。`#Region` でグループ化
- 分岐コメントは「Aのとき / Bのとき」形式。旧コードのコメントアウト・履歴コメントを残さない

## コントロールID接頭辞

```
lbl txt ddl chk cbl rbl btn lnk img
gv(GridView) lv rpt dl  hdn pnl ph upn lit fu
rfv rev cv vsm  cnt(Content)
```
`接頭辞 + PascalCase`（例: `btnSave`, `gvExpenseCode`）。`Label4` のようなデザイナ既定名を残さない。

## 画面実装の必須事項

- `Page_Load` の初期化は `If Not IsPostBack Then` で囲む
- 業務ロジックを `.aspx.vb` に書かない（`Atom.Core` / `Atom.Data` へ）
- `Shared` 変数に状態を持たせない
- グリッドの列は名前で識別（`Cells(7)` 禁止）
- `Literal` は `Mode="Encode"`。`Page_Init` で `ViewStateUserKey` を設定
- サーバー側で入力を再検証。二重送信対策。削除はサーバー側でも再確認
- 権限チェック（menuNo単位）は `BasePage` で一元化し、処理の入口でも判定
- 遷移は `Response.Redirect(url, False)` + `Context.ApplicationInstance.CompleteRequest()`
- 例外の詳細を画面に出さない

## データアクセス

- 値はすべて `SqlParameter`。`Parameters.Add(名前, 型, 長さ)` で型と長さを明示（`AddWithValue` 不可）
- `SELECT *` 禁止。NULL は `DBNull.Value`
- 動的条件でも値はパラメータ。ソート列・テーブル名の可変はホワイトリスト照合
- 複数更新はトランザクション。接続はフィールドやセッションに保持しない
- 命名: `Sql/` は `Create○○Sql`、`Repositories/` は `Select`/`Insert`/`Update`/`Delete`/`Merge` 始まり、論理削除は `SoftDelete`
- 排他制御は更新日時による楽観ロック。競合時は `ConcurrencyException`

## 設定は Web.config から取得する

| 用途 | キー |
|---|---|
| ログ出力先 | `Log.Writer`（Sql / File / Trace）, `Log.ConnectionName`, `Log.TableName`, `Log.FolderPath`, `Log.MinimumLevel` |
| メール | `Mail.Enabled`（false で送信抑止・ログのみ）, `Mail.FromAddress`, `Mail.RedirectTo`（検証用の宛先差し替え） |
| SVF | `Svf.FormFolderPath`, `Svf.WorkFolderPath` |
| DB接続 | `connectionStrings` の `AtomDb` |

パス・出力先・送信可否をコードに直書きしない。ログは `Logger` 経由のみ（画面から直接書かない）。接続文字列の実値はコミットしない。

## 提出前チェック

- [ ] 設計書を `docs/specs/` に保存し、要確認事項の回答を追記した
- [ ] フォルダ構成に従って配置した／依頼範囲外を触っていない
- [ ] `Option Strict On` / `Infer Off` / 型明記
- [ ] SQL は全てパラメータ化（連結ゼロ）／`Using` で解放／複数更新はトランザクション
- [ ] **DDL を含まない。テーブル・列の追加変更を前提にしていない**
- [ ] コントロールを動的生成していない／designer を手編集していない
- [ ] XSS・CSRF・権限チェック・サーバー側再検証あり
- [ ] `Throw ex` なし／空 `Catch` なし
- [ ] 機密情報がコード・ログ・コミットに含まれていない
