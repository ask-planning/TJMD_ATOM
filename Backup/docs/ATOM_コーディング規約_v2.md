# ATOM コーディング規約 v2.0

制定日: 2026-08-10
対象: ATOM リプレイス（ASP.NET Web Forms / VB.NET / .NET Framework 4.8.1 / SQL Server / SVF）
位置づけ: 既存規約 v1 と、既存 FPI 実装（FS036_02 一式）の実態分析を統合した版

---

## 0. 本規約の運用方針

| 区分 | 意味 | 扱い |
|---|---|---|
| **必須** | 違反はレビューで差し戻し | ビルド設定・アナライザで機械的に検出する |
| **推奨** | 原則従う。逸脱時はコメントで理由を書く | レビューで指摘する |
| **参考** | 判断材料 | 指摘しない |

既存コードの一括修正は求めない。**触った箇所から本規約に寄せる**（ボーイスカウトルール）。

---

## 1. コンパイラ設定

**必須**

- 全 `.vb` ファイルの先頭に以下を宣言する。プロジェクト設定にも同じ値を設定する。

```vb
Option Strict On
Option Explicit On
Option Infer Off
```

- 文字コードは **UTF-8（BOM 付き）**、改行は **CRLF**
  - BOM の付与は Python で行う。bash の `printf '\xEF\xBB\xBF'` はリテラル文字列を書き込むため使用禁止

### 1.1 Option Strict On で必要になる対応

既存コードには以下のパターンがあり、Option Strict On では必ずコンパイルエラーになる。移植時は機械的に潰す。

| 既存パターン | 対応 |
|---|---|
| `Public Sub Foo(dgv, textA, textB)` | 全引数に型を明記する |
| `Public Function Bar(...) `（戻り値型なし） | `As DataSet` などを明記する |
| `Dim x = ...` | `Dim x As String = ...` と型を明記する |
| `obj.DataSource.Rows.Count`（遅延バインディング） | `CType(obj.DataSource, DataTable).Rows.Count` |
| `Dim s As String = row.ItemArray(0)` | `CStr(row("列名"))` / `Convert.ToString(...)` |

---

## 2. 命名規則

### 2.1 基本

**必須**

| 対象 | 記法 | 例 |
|---|---|---|
| クラス / モジュール / 構造体 | PascalCase | `ExpenseCodeRepository` |
| メソッド（Sub / Function） | PascalCase | `SelectByKey` |
| プロパティ | PascalCase | `IsDecided` |
| 引数 | camelCase | `tenpoCode` |
| ローカル変数 | camelCase | `iraiNo` |
| プライベートフィールド | `_camelCase` | `_connectionString` |
| 定数（Const） | PascalCase | `TransactionTimeoutSeconds` |
| 列挙型 / メンバー | PascalCase | `ShukkaMode.Selected` |

**禁止事項（必須）**

- スネークケースの VB 変数名（`tenpo_code_sub` → `tenpoCodeSub`）
  - ただし **DB の列名・テーブル名はスネークケースのまま**扱う。文字列リテラルとして現れる分には変更しない
- 定数への `C_` プレフィックス、大文字スネーク（`C_TRAN_TIMEOUT` → `TransactionTimeoutSeconds`）
- 型指定のない `Dim`（アクセス修飾子も必ず書く。`Dim Lg As ...` → `Private ReadOnly _logic As ...`）

### 2.2 ローマ字の扱い

**必須**

- 業務用語はローマ字を許容する（`Haisouirai`, `Bukken` 等）が、**ヘボン式に統一**する。
  - 既存の不統一例: `Jyucyuzan` / `Jyuchuzan`、`taisyo` / `tausyo`、`Shubetu` / `Shubetsu`
- 新規コードでは以下に従う。

| 誤りやすい語 | 採用表記 |
|---|---|
| 受注残 | `Juchuzan` |
| 対象 | `Taisho` |
| 種別 | `Shubetsu` |
| 出荷 | `Shukka` |
| 配送依頼 | `HaisoIrai` |
| 物件 | `Bukken` |

- 綴り誤りを見つけたら修正する（`Lavel` → `Label`、`Dicide` → `Decide`、`Trauncate` → `Truncate`）

### 2.3 コントロール名（ASP.NET Web Forms）

**必須。プレフィックス方式に統一する。**

既存 FPI はサフィックス方式（`bukkenAddBtn`）だが、Web Forms では以下を採用する。理由は、Web Forms の標準的な慣習であること、`gv` / `ddl` のように種別が一目で判別できること。

```
lbl  Label            txt  TextBox          ddl  DropDownList
chk  CheckBox         cbl  CheckBoxList     rbl  RadioButtonList
btn  Button           lnk  LinkButton       img  Image
gv   GridView         lv   ListView         rpt  Repeater
hdn  HiddenField      pnl  Panel            ph   PlaceHolder
upn  UpdatePanel      lit  Literal          fu   FileUpload
rfv  RequiredFieldValidator    rev  RegularExpressionValidator
cv   CustomValidator           vsm  ValidationSummary
```

- 記法は `プレフィックス + PascalCase`（アンダースコアなし）。例: `btnSave`, `gvExpenseCode`, `ddlTruckShubetsu`
- **デザイナ既定名の残存を禁止**する。`Label4`, `MetroLabel1`, `DataGridViewTextBoxColumn27` のような名前はコミット前に必ず改名する
  - 既存 FPI では全103コントロール中の大半が既定名のままであり、可読性を著しく損なっている

---

## 3. 記述スタイル

**必須**

- インスタンスメンバーの参照には常に `Me.` を付ける
- 文字列結合は `&` を使う。`+` は使用禁止（`String + Object` の暗黙変換事故を防ぐため）
- 3つ以上の連結、またはループ内の連結は `String.Concat` / `StringBuilder` / `String.Format` を使う
- `Call` キーワードは使わない（`Call dispCalendar()` → `Me.DispCalendar()`）

**推奨**

- 1メソッドは50行以内を目安とする
- ネストは3段までを目安とする

---

## 4. コメント

### 4.1 XML ドキュメントコメント

**必須**

- Public / Protected 要素には日本語の XML コメントを付ける
- `<param>` と `<returns>` は**本文を必ず書く**。空タグは禁止
- **実在しない引数の `<param>` を書かない**
  - 既存例: `BackgroundColorChange(dgv)` に `<param name="shukkaDateText">` と `<param name="d">` が記載されており、実装と乖離している
- `<returns>` に型名だけを書かない（`<returns>Text.StringBuilder</returns>` は不可。何を返すのかを書く）

```vb
''' <summary>
''' 出荷先と出荷日を条件に、選択中の物件一覧を取得します。
''' </summary>
''' <param name="shukkasakiName">出荷先名</param>
''' <param name="shukkaDate">出荷日</param>
''' <returns>物件一覧。該当がない場合は空のリストを返します。</returns>
Public Function SelectBukkenIchiran(shukkasakiName As String, shukkaDate As Date) As IReadOnlyList(Of BukkenIchiran)
```

### 4.2 コード内コメント

**必須**

- 分岐コメントは「Aのとき / Bのとき」の形式で書く

```vb
'生産指示済みのとき
If ... Then
    ...
'生産指示未済のとき
Else
    ...
End If
```

- **コメントアウトした旧コードを残さない**。履歴は Git で追う
- **変更履歴コメントをコード内に書かない**（`'2023/10/25 田尻 追加` 等）。コミットメッセージと Issue 番号で管理する
  - 例外: 業務上の判断理由が必要な場合は、履歴ではなく「なぜそうするか」を書く

### 4.3 `#Region`

**推奨**

- 以下の順でグループ化する。

```
#Region "定数"
#Region "フィールド"
#Region "プロパティ"
#Region "コンストラクタ"
#Region "イベントハンドラ"
#Region "参照系"
#Region "更新系"
#Region "内部処理"
```

---

## 5. データアクセス（最重要）

### 5.1 SQL の組み立て

**必須。既存 FPI 実装で最も深刻な問題であり、Web 化にあたって例外なく是正する。**

- **SQL のパラメータは必ず `SqlParameter` を使う。文字列連結・`String.Replace` による値の埋め込みを全面的に禁止する。**

既存の以下のパターンはすべて禁止。

```vb
' 禁止例 1: 直接連結
.AppendLine("WHERE truck_type = '" + truckShubetsu + "'")

' 禁止例 2: 置換マーカー（@ 始まりだがパラメータではない）
.AppendLine("SET @ShukkaDate")
query.Replace("@ShukkaDate", "shukka_date = '" & shukkaDate & "'")

' 禁止例 3: Boolean や日付の文字列化
.AppendLine("    '" & deleteFlag & "',")
```

正しい書き方。

```vb
Const SelectCabiSinkSql As String =
    "SELECT cabi_count, sink_count " &
    "FROM dbo.m_truck_shubetsu " &
    "WHERE truck_type = @TruckType"

Using command As New SqlCommand(SelectCabiSinkSql, connection, transaction)
    command.Parameters.Add("@TruckType", SqlDbType.NVarChar, 20).Value = truckType
    ...
End Using
```

- 型は `Parameters.Add(名前, 型, 長さ)` で明示する。`AddWithValue` は型推論による性能劣化を招くため**非推奨**
- NULL を渡す場合は `DBNull.Value` を使う

### 5.2 動的 SQL が避けられない場合

**必須**

- 条件の有無で `WHERE` 句を組み替える場合も、**値は必ずパラメータ**にする。組み立てて良いのは SQL の構文部分のみ

```vb
Dim conditions As New List(Of String)()
If shukkaDate.HasValue Then
    conditions.Add("HI.shukka_date = @ShukkaDate")
    command.Parameters.Add("@ShukkaDate", SqlDbType.Date).Value = shukkaDate.Value
End If
```

- ソート列やテーブル名を可変にする場合は、**ホワイトリスト照合**を必須とする（外部入力をそのまま埋め込まない）

### 5.3 型の扱い

**必須**

- 日付は `Date` / `Date?`、真偽値は `Boolean` / `Boolean?` で受け渡す。`String` で持ち回らない
- DB 側の NULL は `Nullable(Of T)` で表現する。`IsDBNull` の判定は取得箇所で完結させ、上位に `Object` のまま渡さない
- ハードコードされた番兵値（`'1900/01/01 0:00:00'` 等）を新規に増やさない。既存データの都合で必要な場合は定数化する

```vb
''' <summary>未設定を表す日時。既存データとの互換のために使用します。</summary>
Public Const UnsetDateTime As String = "1900/01/01 00:00:00"
```

### 5.4 接続・トランザクション

**必須**

- `SqlConnection` / `SqlCommand` / `SqlDataAdapter` は `Using` で囲む。`Finally` での手動 `Dispose` は禁止
- **接続の Open / Close をデータアクセスメソッドごとに行わない**。トランザクション単位を呼び出し側で制御できるようにする
  - 既存 FPI はメソッドごとに `connection.Open()` / `connection.Close()` を行いながら、別途 `Public tran As SqlTransaction` を保持しており、整合が崩れやすい構造になっている
- 複数テーブルを更新する処理は必ずトランザクションで囲む
- ASP.NET では接続オブジェクトを**フィールドやセッションに保持しない**。リクエスト内で完結させる

### 5.5 メソッド命名

**必須。役割によって命名体系を分ける。**

| レイヤ | 役割 | 命名 | 例 |
|---|---|---|---|
| Sql クラス | SQL 文字列の生成のみ | `Create○○Sql` | `CreateSelectBukkenIchiranSql` |
| Data 層（Repository） | SQL の実行 | `Select` / `Insert` / `Update` / `Delete` / `Merge` で始める | `SelectBukkenIchiran`, `UpdateShukkaDate` |
| Core 層（業務ロジック） | 業務処理 | 動詞 + 目的語の PascalCase | `DecideShukkaDate` |

- 既存の `GetIchiran`, `truckTypeMerge`, `HaisouiraiSubMatomeSoftDelete` のような混在は解消する
- 論理削除は `Delete` ではなく **`SoftDelete`** を接頭辞に使い、物理削除と区別する

---

## 6. 例外処理とログ

**必須**

- 再送出は **`Throw`** を使う。**`Throw ex` は禁止**（スタックトレースが失われる。既存 FPI に18箇所存在）
- 例外を握りつぶさない。`Catch` でメッセージ表示のみ行い処理を続行するのは禁止。継続できない場合は処理を中断する
- `Catch ex As Exception` で握る場合は、必ずログ記録と再送出、または明示的な業務的リカバリのいずれかを行う
- ログには実行 SQL とパラメータを記録する。**パラメータ値に接続文字列・パスワード等が含まれる場合はマスクする**
- ログの出力先（SQL Server のログテーブル）は共通ライブラリ経由で行い、各画面で直接書き込まない

```vb
Catch ex As SqlException
    Me.WriteErrorLog("物件一覧の取得に失敗しました。", ex)
    Throw
End Try
```

---

## 7. 画面層（ASP.NET Web Forms）

**必須**

- **業務ロジックを `.aspx.vb` に書かない**。`Atom.Core` / `Atom.Data` に配置する
  - `.aspx.vb` の責務はイベント受付、入力値の詰め替え、結果の表示のみ
- **静的（`Shared`）変数に画面やユーザーの状態を保持しない**。全ユーザーで共有され、データ混在事故を起こす
- グリッドの列は**列名で識別**する。`Cells(7)` のようなインデックス直指定を禁止する
  - 既存 FPI には `Cells(7)`, `Cells(15)`, `Cells(18)`, `ItemArray(0)` 等が多数あり、列順変更で破綻する
  - Web Forms では `e.Row.FindControl("lblShukkaDate")` または `DataBinder.Eval(e.Row.DataItem, "shukka_date")` を使う
- 画面出力時は HTML エスケープを行う（XSS 対策）。`Literal` の `Mode` は `Encode` を既定とする
- `Page_Init` で `ViewStateUserKey` にセッション ID を設定する（CSRF 対策）
- 更新系ボタンは二重送信を防止する
- 削除はクライアント側の確認に加え、**サーバー側でも再確認**する

**禁止（WinForms からの持ち込み）**

| WinForms のコード | Web での扱い |
|---|---|
| `MessageBox.Show(...)` | 画面上のメッセージ領域 または クライアントスクリプト |
| `Cursor.Current = Cursors.WaitCursor` | クライアント側のローディング表示 |
| `Me.Size = New Size(Width + 1, Height + 1)` 的な再描画ハック | 不要。CSS で解決する |
| `Form.Show()` による子画面表示 | ページ遷移 または モーダル |
| フォームの Public フィールド経由の値渡し | クエリ文字列（**改ざん前提で必ずサーバー側再検証**）／セッション |

---

## 8. 重複コードの排除

**必須**

- 同一ロジックの3回以上の重複を禁止する
  - 既存 FPI の決定ボタン処理では、選択中一覧と追加候補一覧に対する約80行のブロックがほぼ丸ごと重複している
- ループ対象や区分のみが異なる場合は、引数で切り替える共通メソッドに切り出す

---

## 9. レビューチェックリスト

コミット前・レビュー時に確認する。

### コンパイル設定
- [ ] `Option Strict On` / `Explicit On` / `Infer Off` が宣言されている
- [ ] 引数と戻り値に型が明記されている
- [ ] 遅延バインディングがない

### 命名
- [ ] フィールドが `_camelCase`、定数が PascalCase
- [ ] VB の変数にスネークケースがない
- [ ] コントロールがプレフィックス方式で命名されている
- [ ] デザイナ既定名（`Label4` 等）が残っていない
- [ ] ローマ字がヘボン式に統一されている / 綴り誤りがない

### 記述
- [ ] `Me.` が付いている
- [ ] 文字列結合が `&` になっている
- [ ] `Call` を使っていない

### コメント
- [ ] Public 要素に日本語 XML コメントがある
- [ ] `<param>` が実在の引数と一致し、本文が書かれている
- [ ] 分岐コメントが「Aのとき / Bのとき」形式
- [ ] コメントアウトした旧コード・履歴コメントがない

### データアクセス
- [ ] **すべての値が `SqlParameter` 経由である（文字列連結・`Replace` による埋め込みがない）**
- [ ] `Using` でリソースを解放している
- [ ] 日付・真偽値を `String` で持ち回っていない
- [ ] 複数更新がトランザクションで囲まれている
- [ ] メソッド名が `Create○○Sql` / `Select○○` 等の体系に沿っている

### 例外・ログ
- [ ] `Throw ex` がない
- [ ] 例外を握りつぶしていない
- [ ] ログに機密情報が平文で出ていない

### Web 固有
- [ ] 業務ロジックが `.aspx.vb` に漏れていない
- [ ] `Shared` 変数に状態を持たせていない
- [ ] 列をインデックスではなく名前で識別している
- [ ] XSS / CSRF 対策が入っている
- [ ] 権限チェック（menuNo 単位）が入っている

---

## 10. 未確定事項（要確認）

| # | 項目 | 確認先 |
|---|---|---|
| 1 | `LogCommon.LogError` の実際の出力先。規約は「SQL Server ロギング」だが、FPI 実装がファイル出力の場合は共通ライブラリの見直しが必要 | 開発チーム |
| 2 | コントロール命名をプレフィックス方式に統一することの合意。既存 FPI 資産はサフィックス方式のため、両システムで異なる方式になる | 開発チーム |
| 3 | 既存 FPI コードを ATOM 側で流用する範囲。流用する場合、Option Strict On 化の工数見積もりが必要 | 開発チーム |
| 4 | `MetroFramework` への依存。Web Forms では使用できないため代替の UI 方針が必要 | 開発チーム |

---

## 更新履歴

| 版 | 日付 | 内容 |
|---|---|---|
| 1.0 | — | WinForms 移行時の初版 |
| 2.0 | 2026-08-10 | ASP.NET Web Forms 対応。既存 FPI 実装（FS036_02）の分析結果を反映し、SQL パラメータ化・例外処理・命名体系を追加 |
