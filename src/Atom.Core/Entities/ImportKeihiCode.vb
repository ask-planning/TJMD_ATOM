Option Strict On
Option Explicit On
Option Infer Off

Namespace Entities

    ''' <summary>
    ''' 輸入経費コード（menuNo=907）の 1 件です。
    ''' </summary>
    ''' <remarks>
    ''' `dbo.tbl_m_import_keihi_code` の 1 行に対応します。
    ''' 桁数と NULL 可否は実テーブルの定義から確定しています（設計書 §13.1 ①）。
    ''' </remarks>
    Public Class ImportKeihiCode

        ''' <summary>伝票コード。主キー。4 文字・必須。</summary>
        Public Property KeihiCode As String = String.Empty

        ''' <summary>経費名称。50 文字・任意。</summary>
        Public Property KeihiCodeText As String = String.Empty

        ''' <summary>REF。英語の正式名称。50 文字・任意。</summary>
        Public Property KeihiCodeRef As String = String.Empty

        ''' <summary>区分。`View__keihi_kbn` の `contents1`。3 文字・任意。</summary>
        Public Property KeihiKbn As String = String.Empty

        ''' <summary>
        ''' 保険対象。
        ''' </summary>
        ''' <remarks>
        ''' 旧はテキストボックスで、0 / 1 以外も入れられました（設計書 C-10）。
        ''' 列は NULL 可ですが実データに NULL はありません。読み込み時の NULL は 0 として扱います。
        ''' </remarks>
        Public Property HokenTaisyoFlg As Integer = 0

        ''' <summary>
        ''' 更新日。
        ''' </summary>
        ''' <remarks>
        ''' **楽観ロックの比較に使います。**既存行の 39 件が NULL のため、
        ''' NULL 同士も一致とみなします（設計書 §13.1 ③）。
        ''' </remarks>
        Public Property UpdateYmd As Date? = Nothing

        ''' <summary>更新者の社員コード。6 文字。</summary>
        Public Property UpdateLogin As String = String.Empty

    End Class

    ''' <summary>
    ''' 経費区分の選択肢です。
    ''' </summary>
    ''' <remarks>
    ''' `dbo.View__keihi_kbn`（項目マスタ `tbl_m_koumoku` の `pk_koumoku_code = 'DKB'`）の 1 行です。
    ''' **保存するのは `Code`、画面に出すのは `Name`** です（設計書 §13.1 ②）。
    ''' </remarks>
    Public Class ImportKeihiKbn

        ''' <summary>区分の値（`contents1`）。経費コードの `keihi_kbn` に保存します。</summary>
        Public Property Code As String = String.Empty

        ''' <summary>区分の表示名（`contents2`）。</summary>
        Public Property Name As String = String.Empty

    End Class

End Namespace
