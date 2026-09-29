# ATOM リプレイス（ASP.NET Web Forms / VB.NET）プロンプト集

作成日: 2026-08-10
対象: ATOM（旧 Microsoft Access 2003 .adp）の Web 化
技術スタック: ASP.NET Web Forms / VB.NET / .NET Framework 4.8.1 / SQL Server / SVF

---

## 0. このドキュメントの使い方

再現性を高めるため、AI への依頼は必ず以下の順で組み立てます。

| 手順 | 使うプロンプト | 目的 |
|---|---|---|
| 1 | **P1 基盤プロンプト** | 会話の冒頭に必ず貼る。規約を毎回明示する |
| 2 | **P2 旧画面の仕様抽出** | Access 側の画面から仕様書を起こす |
| 3 | **P3 画面生成** | .aspx / .aspx.vb を生成する |
| 4 | **P4 データアクセス層生成** | Atom.Data のクエリメソッドを生成する |
| 5 | **P5 帳票（SVF）** | 帳票出力処理を生成する |
| 6 | **P6 レビュー** | 規約適合と脆弱性をチェックする |

**原則**: P1 は毎回貼り直す。会話が長くなったら新しいスレッドを立て、P1 から再開する。

---

## P1. 基盤プロンプト（毎回冒頭に貼る）

```text
あなたは ATOM システムの ASP.NET Web Forms 移行を担当する開発者です。
以下の前提と規約を厳守してください。

# システム概要
- 旧システム: Microsoft Access 2003 (.adp) + SQL Server
- 業務ドメイン: 輸入・購買・INVOICE・物流
- 新システム: ASP.NET Web Forms / VB.NET / .NET Framework 4.8.1
- 帳票: SVF
- リポジトリ: ask-planning/TJMD_ATOM

# ソリューション構成
- Atom.Web     : ASP.NET Web Forms（新規・画面層）
- Atom.Core    : 業務ロジック・エンティティ（既存資産を再利用）
- Atom.Data    : データアクセス（既存資産を再利用）
- Atom.Mail    : メール送信（既存資産を再利用）
- Atom.Report  : SVF 帳票
- Atom.Tests   : 単体テスト

# コーディング規約
- Option Strict On / Option Explicit On / Option Infer Off を全ファイルで宣言
- プライベートフィールド: _camelCase
- 定数: PascalCase
- コントロールフィールド: プレフィックスのみ（アンダースコアなし）
- Public 要素には日本語の XML ドキュメントコメントを必ず付ける
- #Region でメンバーをグループ化する
- インスタンスメンバーには常に Me. を付ける
- 文字列結合は & を使う（+ は使わない）
- クエリメソッドは Select / Insert / Update / Delete のいずれかで始める
- 例外は SQL Server のログテーブルへ記録する
- 難解な英単語は使わない
- 分岐コメントは「Aのとき / Bのとき」の形式で書く

# Web Forms コントロールのプレフィックス
lbl / txt / ddl / chk / chkList / rbl / btn / lnk / img
gv (GridView) / lv (ListView) / rpt (Repeater) / dl (DataList)
hdn (HiddenField) / pnl / ph (PlaceHolder) / upn (UpdatePanel)
rfv (RequiredFieldValidator) / rev / cv (CustomValidator) / vsm (ValidationSummary)

# 設計方針
- 業務ロジックは .aspx.vb に書かず、Atom.Core 側のクラスへ寄せる
  （将来 Blazor 等へ移行できる余地を残すため）
- .aspx.vb はイベント受付・画面とエンティティの詰め替えのみを担当する
- SQL は必ずパラメータ化する。文字列結合による SQL 組み立ては禁止
- 画面共通レイアウトは Site.Master、共通ヘッダーは HeaderControl.ascx
- ViewState は必要な項目のみ有効にし、GridView の肥大化を避ける
- 一覧のページングとソートはサーバー側（SQL）で行う

# 回答ルール
- 回答は日本語で行う
- 推測を含む場合は「推測」と明示する
- 接続文字列・パスワード等はマスクして提示する
- 本番環境への直接操作は提案しない
- 短いスニペットで足りる場合はコードのみ提示する
- 大規模変更の場合はファイル単位で全文を出力する
```

---

## P2. 旧画面の仕様抽出プロンプト

Access 側の画面情報（フォーム定義、SQL、VBA）を渡して仕様書を起こすときに使います。

```text
以下は旧 ATOM（Access .adp）の画面情報です。
ASP.NET Web Forms へ移植するための画面仕様書を Markdown で作成してください。

# 入力情報
- menuNo:
- 画面名（日本語）:
- 対象テーブル / ビュー:
- フォーム定義（貼り付け）:
- 埋め込み SQL（貼り付け）:
- VBA コード（貼り付け）:

# 出力してほしい章立て
1. 画面概要（目的・利用者・利用頻度）
2. 画面項目一覧
   | 項目名 | 物理名 | 型 | 桁 | 必須 | 初期値 | 入力制御 | 備考 |
3. 一覧グリッドの仕様（表示列・ソート既定・ページング単位）
4. ボタンと動作（新規 / 検索 / 登録 / 更新 / 削除 / 閉じる など）
5. バリデーション仕様（単項目・相関チェック）
6. データベース操作（実行される SQL の意図を日本語で説明）
7. 画面遷移（呼び出し元・呼び出し先）
8. 移植時の論点（Access 固有機能で Web に置き換えが必要な箇所）

不明な点は推測で埋めず「要確認」と明記してください。
```

---

## P3. 画面生成プロンプト

```text
P2 で作成した画面仕様書に基づき、ASP.NET Web Forms の画面を実装してください。

# 対象画面
- 画面名（日本語）:
- ファイル名: Pages/<領域>/<PascalCase>.aspx
- menuNo:
- ヘッダー背景色:
- 対象テーブル:

# 出力してほしいファイル
1. <画面名>.aspx          … Site.Master を使用したマークアップ
2. <画面名>.aspx.vb       … コードビハインド
3. <画面名>.aspx.designer.vb … コントロール宣言

# 実装要件
- Site.Master の ContentPlaceHolder に配置する
- 画面タイトル・ヘッダー色は Page_Load でヘッダーコントロールへ設定する
- 一覧は GridView、明細編集は TemplateField を使う
- 検索条件は入力後の再検索でも保持する
- 登録／更新／削除は必ず確認を挟む（削除は JavaScript の confirm ＋ サーバー側再確認）
- 例外は握りつぶさず、ログ記録のうえ画面へ日本語メッセージを表示する
- 二重送信を防止する（ボタンの UseSubmitBehavior と クライアント側の無効化）
- CSRF 対策として Page_Init で ViewStateUserKey にセッション ID を設定する

# 出力形式
ファイルごとに全文を出力してください。省略記号（...）は使わないでください。
```

---

## P4. データアクセス層プロンプト

```text
Atom.Data に以下のデータアクセスクラスを追加してください。

# 対象
- テーブル名:
- クラス名: <エンティティ名>Repository
- エンティティ: Atom.Core.<エンティティ名>

# 実装要件
- Option Strict On / Explicit On / Infer Off
- メソッド名は Select / Insert / Update / Delete で始める
- SQL は必ず SqlParameter でパラメータ化する
- 接続文字列は Web.config の connectionStrings から取得する（値はマスクして記載）
- 排他制御は更新日時（または RowVersion）による楽観ロックとする
  競合時は専用の例外クラスを送出する
- 複数テーブルの更新はトランザクションで囲む
- 例外は SQL Server のログテーブルへ記録したうえで再送出する
- Public メソッドには日本語の XML ドキュメントコメントを付ける
- #Region でグループ化する（フィールド / コンストラクタ / 参照系 / 更新系 / 内部処理）

# 実装してほしいメソッド
- SelectAll
- SelectByKey
- SelectByCondition（検索条件クラスを引数に取る）
- Insert / Update / Delete
```

---

## P5. 帳票（SVF）プロンプト

```text
ASP.NET Web Forms から SVF 帳票を出力する処理を実装してください。

# 前提
- 帳票定義ファイル:
- 出力形式: PDF
- Web サーバー上で生成し、ブラウザへダウンロードさせる

# 実装要件
- 帳票生成処理は Atom.Report 側のクラスに実装し、画面からは呼び出すだけにする
- 一時ファイルは必ず Finally で削除する
- 同時実行を考慮し、一時ファイル名は GUID を含める
- Response への書き出しは Content-Disposition を設定し、
  日本語ファイル名は URL エンコードする
- 出力後は Response.End ではなく
  HttpContext.Current.ApplicationInstance.CompleteRequest を使う
- 例外時は帳票を出さず、画面へ日本語のエラーメッセージを表示する

なお、SVF の Web サーバー実行に別ライセンスが必要かどうかは
実装前に確認が必要な事項として明記してください。
```

---

## P6. レビュープロンプト

```text
以下のコードを ATOM の規約に照らしてレビューしてください。

# チェック観点
## 規約
- Option Strict On / Explicit On / Infer Off が宣言されているか
- 命名規則（_camelCase / PascalCase / コントロールプレフィックス）
- Me. の付与
- 文字列結合が & になっているか
- Public 要素に日本語 XML コメントがあるか
- #Region によるグループ化
- クエリメソッドの命名（Select / Insert / Update / Delete）
- 難解な英単語が使われていないか

## 設計
- 業務ロジックが .aspx.vb に漏れていないか
- Atom.Core / Atom.Data の責務分離ができているか

## セキュリティ
- SQL がパラメータ化されているか
- 画面出力時に HTML エスケープされているか（XSS）
- CSRF 対策（ViewStateUserKey）
- 認可チェック（menuNo に対する権限判定）が入っているか
- 接続文字列やパスワードがソースに直書きされていないか

## 品質
- 例外処理とログ記録
- リソースの解放（Using）
- 二重送信対策
- ViewState の肥大化

# 出力形式
| 重要度 | 箇所 | 指摘 | 修正案 |
重要度は 高 / 中 / 低 の3段階。修正案はコード断片で示してください。
```

---

## 付録 A. WinForms → Web Forms 対応表

| WinForms（現行） | Web Forms（移行後） |
|---|---|
| `BaseForm` | `Site.Master` ＋ 基底クラス `BasePage` |
| `HeaderControl`（UserControl） | `HeaderControl.ascx` |
| `HeaderConfig` / `HeaderButton` | `BasePage` のプロパティ ＋ `HeaderControl.Apply` |
| `DataGridView` | `GridView` |
| `DataGridViewComboBoxColumn` | `TemplateField` ＋ `DropDownList` |
| `CellFormatting` | `RowDataBound` |
| `CellPainting`（見た目） | CSS クラス（`RowDataBound` で付与） |
| `ErrorProvider` | `Validator` コントロール ＋ `ValidationSummary` |
| `MessageBox.Show` | クライアント側 `confirm` / 画面上のメッセージ領域 |
| `Form.ShowDialog` | 別ページ遷移 または モーダル |
| フォーム間の値渡し | クエリ文字列（改ざん前提で必ず再検証）／セッション |
| `Application.Exit` | ログアウト（セッション破棄） |

---

## 付録 B. Web 化で新たに必要になる検討事項

移行時に必ず設計判断が必要な項目です。着手前に方針を決めておきます。

| 項目 | 論点 |
|---|---|
| 認証方式 | Windows 認証（社内 AD）か Forms 認証か |
| 認可 | menuNo 単位の権限をどこで判定するか（BasePage で一元化を推奨） |
| セッション管理 | タイムアウト時間、保持方式（InProc / StateServer / SQL Server） |
| 排他制御 | 悲観ロックが使えないため楽観ロックへ変更が必要 |
| ファイル出力 | Excel / CSV 出力のサーバー側実装方式 |
| ファイル取込 | アップロードのサイズ上限とウイルス対策 |
| 印刷 | クライアント直接印刷ができないため PDF ダウンロードへ変更 |
| 同時実行 | 静的変数の共有事故に注意（WinForms 感覚で書くと事故る） |
| ブラウザ | 対応ブラウザとバージョンの決定 |
| ログ | 操作ログ・エラーログのテーブル設計 |

---

## 付録 C. 記入例（輸入経費コード登録 / menuNo=907）

P3 の記入例です。

```text
# 対象画面
- 画面名（日本語）: 輸入経費コード登録
- ファイル名: Pages/Master/ExpenseCode.aspx
- menuNo: 907
- ヘッダー背景色: #FF8040
- 対象テーブル: 輸入経費コードマスタ

# この画面固有の要件
- ヘッダーに「閉じる」ボタンを配置し、メニュー画面へ戻る
- 一覧に伝票区分の選択列（旧 DataGridViewComboBoxColumn）がある
  - 表示は4桁ゼロパディングのコード（例: 0001）
  - 選択肢は「コード　名称」形式（区切りは全角スペース）
  - Web Forms では TemplateField 内の DropDownList で実装し、
    DataTextField に「コード　名称」、DataValueField にコードを設定する
- 選択列のセル背景色は黄色（CSS クラスで指定）
```

---

## 更新履歴

| 日付 | 内容 |
|---|---|
| 2026-08-10 | 初版作成（ASP.NET Web Forms / VB.NET 方式決定に伴い） |
