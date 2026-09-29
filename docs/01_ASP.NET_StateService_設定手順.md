# ASP.NET State Service 設定手順（セッション保存先の変更）

- **作成日**: 2026-08-13
- **対象**: ATOM を配置する IIS サーバー
- **目的**: セッションの保存先を InProc から StateServer へ変更し、IIS のアプリケーションプール再起動でログインが切れないようにする

> **【2026-08-13 現在：この手順は未適用です】**
> 運用が煩雑なため、`Web.config` は `InProc` を採用しています。
> 本書は、将来「セッションが切れると困る画面」が出てきて StateServer へ切り替える際の手順として保管します。
> 切り替える場合は、本書の手順に加えて `Web.config` の `sessionState` を `StateServer` へ戻してください。

> **この手順書は人間が実行してください。** AI エージェントは本番環境（DB / IIS / ファイルサーバ）へ直接操作を行いません。
> 適用は検証環境で確認してから本番へ行ってください。

---

## 1. 前提と影響

### 前提

| 項目 | 内容 |
|---|---|
| OS | Windows Server（IIS が動作しているサーバー） |
| 必要な権限 | 対象サーバーの管理者権限 |
| サービス名 | ASP.NET State Service（サービス名 `aspnet_state`） |
| 既定ポート | TCP 42424 |
| アプリ側の設定 | `src/Atom.Web/Web.config` の `sessionState` は設定済み（本手順 §5 で確認） |

### 影響

- **作業中、既存のログインセッションはすべて失われます。**利用者は再ログインが必要です。業務時間外に実施してください。
- StateServer はセッションの内容をプロセス外に持つため、**セッションへ入れるオブジェクトは `Serializable` である必要があります。**
  - ATOM で保存しているのは `Atom.Web.Common.SessionContext` のみで、`<Serializable()>` 属性を付与済みです。
  - 今後セッションへ独自クラスを入れる場合は、必ず `<Serializable()>` を付けてください。付け忘れると実行時に例外になります。
- `aspnet_state` サービス自体を再起動すると、その時点の全セッションが失われます。

---

## 2. サービスの有効化（同一サーバーで動かす場合）

ATOM の Web サーバーとセッション保存先を同じサーバーにする、最も単純な構成です。**この構成なら §3 のレジストリ変更は不要です。**

### 2.1 GUI で行う場合

1. `services.msc` を開く
2. 一覧から **ASP.NET State Service** を選択して右クリック →「プロパティ」
3. 「スタートアップの種類」を **自動** に変更
4. 「開始」を押してサービスを起動
5. 「OK」で閉じる

### 2.2 PowerShell で行う場合（管理者として実行）

```powershell
# 状態の確認
Get-Service aspnet_state

# 自動起動に設定して開始
Set-Service -Name aspnet_state -StartupType Automatic
Start-Service -Name aspnet_state

# 起動できたか確認（Status が Running であること）
Get-Service aspnet_state | Format-List Name, Status, StartType
```

### 2.3 サービスが一覧に無い場合

`aspnet_state` は .NET Framework に含まれます。サービスが存在しないときは、サーバーマネージャーの「役割と機能の追加」で以下を確認してください。

- 機能 → **.NET Framework 4.8 Features**
- 役割 → Web サーバー (IIS) → アプリケーション開発 → **ASP.NET 4.8**

---

## 3. 別サーバーで動かす場合（複数台構成のみ）

Web サーバーとセッション保存先を分ける場合のみ実施してください。**同一サーバー構成では不要です。**

### 3.1 レジストリの変更（セッション保存先サーバー側）

対象キー:

```text
HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\aspnet_state\Parameters
```

| 値の名前 | 種類 | 設定値 | 意味 |
|---|---|---|---|
| `AllowRemoteConnection` | DWORD | `1` | 他サーバーからの接続を許可する |
| `Port` | DWORD | `42424`（10進） | 待ち受けポート |

PowerShell で設定する場合（管理者として実行）:

```powershell
$key = 'HKLM:\SYSTEM\CurrentControlSet\Services\aspnet_state\Parameters'
Set-ItemProperty -Path $key -Name 'AllowRemoteConnection' -Value 1 -Type DWord
Set-ItemProperty -Path $key -Name 'Port' -Value 42424 -Type DWord

# 設定を反映するためサービスを再起動（この時点で全セッションが失われます）
Restart-Service -Name aspnet_state
```

### 3.2 ファイアウォール

**TCP 42424 を全体に開放しないでください。**Web サーバーの IP からの接続のみ許可します。

```powershell
New-NetFirewallRule -DisplayName 'ASP.NET State Service (ATOM)' `
    -Direction Inbound -Protocol TCP -LocalPort 42424 `
    -RemoteAddress '<WebサーバーのIPアドレス>' -Action Allow
```

### 3.3 Web.config の接続先

`src/Atom.Web/Web.config` の `stateConnectionString` を、セッション保存先サーバーの IP に変更します。

```xml
<sessionState mode="StateServer"
              stateConnectionString="tcpip=<セッションサーバーのIP>:42424"
              timeout="60"
              cookieless="false" />
```

---

## 4. machineKey の明示（複数台構成、または再起動時の ViewState 対策）

`machineKey` が自動生成のままだと、IIS の再起動や複数サーバー構成で ViewState の検証に失敗します（「ビューステート MAC の検証に失敗しました」）。

1. IIS マネージャーでサイトを選択 → **コンピューターキー**（Machine Key）を開く
2. 「検証キーを自動的に生成する」「暗号化キーを自動的に生成する」の**チェックを外す**
3. 「キーの生成」を押す
4. 「適用」

生成された値は `Web.config` に書き込まれます。**複数サーバー構成では、全サーバーで同じ値にしてください。**

> 生成されたキーは機密情報です。リポジトリへコミットしないでください（`Web.<環境>.config` 変換または `configSource` による外部ファイル化で管理します）。

---

## 5. アプリ側の設定の確認

`src/Atom.Web/Web.config` に以下が入っていることを確認します（コミット済みの状態です）。

```xml
<sessionState mode="StateServer"
              stateConnectionString="tcpip=127.0.0.1:42424"
              timeout="60"
              cookieless="false" />
```

| 項目 | 値 | 備考 |
|---|---|---|
| `mode` | `StateServer` | — |
| `stateConnectionString` | `tcpip=127.0.0.1:42424` | 同一サーバー構成の場合。別サーバーなら §3.3 で変更 |
| `timeout` | `60` | セッションの有効時間（分） |
| `cookieless` | `false` | URL にセッション ID を含めない |

---

## 6. 動作確認

1. ATOM のログイン画面を開き、社員コードでログインする
2. メニュー画面が表示され、ヘッダー右端に氏名が出ることを確認する
3. **IIS のアプリケーションプールをリサイクルする**

   ```powershell
   Restart-WebAppPool -Name '<ATOMのアプリケーションプール名>'
   ```

4. ブラウザでメニュー画面を再読み込みする
5. **ログイン画面へ戻されず、メニュー画面がそのまま表示されれば成功**
   （InProc のままだとここでログイン画面へ戻されます）

### うまくいかないとき

| 症状 | 確認すること |
|---|---|
| 「状態サーバーへの接続を確立できません」 | `aspnet_state` が起動しているか（§2）、ポートとファイアウォール（§3.2）、`stateConnectionString` の値 |
| 「型 …… はシリアル化可能としてマークされていません」 | セッションへ入れているクラスに `<Serializable()>` が付いているか |
| 「ビューステート MAC の検証に失敗しました」 | `machineKey` を明示しているか（§4） |
| リサイクル後にログインが切れる | `mode` が `StateServer` になっているか、アプリケーションプールが正しい `Web.config` を読んでいるか |

---

## 7. 切り戻し手順

問題が起きた場合は、`Web.config` を InProc に戻せば元の動作に戻ります（サービスは起動したままで構いません）。

```xml
<sessionState mode="InProc" timeout="60" cookieless="false" />
```

変更後、アプリケーションプールをリサイクルしてください。**この時点で全セッションが失われます。**

---

## 8. 実施記録（記入欄）

| 日付 | 環境 | 実施者 | 結果 | 備考 |
|---|---|---|---|---|
|  | 検証 |  |  |  |
|  | 本番 |  |  |  |
