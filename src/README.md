# src（VB.NET ソリューション）

WinForms / .NET Framework 4.8.1（クラシック形式・非SDK）。`ATOM.sln` を Visual Studio 2022 で開く。

## プロジェクト構成（6）と参照関係
- **Atom.App**（WinExe/WinForms）… 画面・起動。→ Core, Data, Mail, Report を参照
- **Atom.Core**（Library）… 定数・enum(ShanghaiChotatsu)・共通関数(NumberHelper)。無依存
- **Atom.Data**（Library）… 接続(Database)・Repositories(SQL集約)・StoredProcedures・TableGateways。→ Core
- **Atom.Mail**（Library）… メール送信(MailSender=BASP21置換)。→ Core
- **Atom.Report**（Library）… SVF Client連携(ReportService)。→ Core
- **Atom.Tests**（Library/MSTest）… 単体テスト。→ Core, Data

## ビルド前提
- Visual Studio 2022（VB.NET）
- .NET Framework 4.8.1 Targeting Pack（未導入ならVS Installerで追加）
- 起動プロジェクト = Atom.App
- 接続文字列は `Atom.App/app.config` の `AtomDb` を環境に合わせて設定（資格情報は埋め込まずWindows認証推奨）
- Atom.Report は SVF Client 導入後にアセンブリ参照を追加（現状はIF雛形）

## 現状
- .sln / 各 .vbproj / My Project / 起動フォーム(MainForm) / 主要クラスの**雛形まで作成済み**。
- 画面・帳票・メールの本実装は設計書(D-1等)に沿って順次追加する。
