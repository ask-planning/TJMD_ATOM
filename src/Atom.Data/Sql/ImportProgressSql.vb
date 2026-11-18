Option Strict On
Option Explicit On
Option Infer Off

Namespace Sql

    ''' <summary>
    ''' 輸入進捗（dbo.tbl_t_import_progress）の SQL 文を組み立てます。
    ''' 旧 Access 版の共通処理 schedule_update に相当します。
    ''' </summary>
    Public NotInheritable Class ImportProgressSql

#Region "コンストラクタ"
        ''' <summary>インスタンス化を禁止します。</summary>
        Private Sub New()
        End Sub
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' 進捗データを 1 件登録する SQL を作成します。
        ''' </summary>
        ''' <returns>INSERT 文。</returns>
        Public Shared Function CreateInsertProgressSql() As String
            Return "INSERT INTO dbo.tbl_t_import_progress " &
                   "(pk_invoice_no_main, syurui, syukko_ymd, update_ymd, update_login) " &
                   "VALUES (@InvoiceNoMain, @Syurui, @SyukkoYmd, GETDATE(), @UpdateLogin)"
        End Function

        ''' <summary>
        ''' 上海連絡（ヘッダ）の進捗を更新する SQL を作成します。処理NO 5 に対応します。
        ''' </summary>
        ''' <returns>UPDATE 文。</returns>
        ''' <remarks>
        ''' 二重更新を避けるため、上海ヘッダ日付が未設定の行だけを対象にします。
        ''' </remarks>
        Public Shared Function CreateUpdateShanghaiHeaderProgressSql() As String
            Return "UPDATE dbo.tbl_t_import_progress " &
                   "SET syukko_ymd = @SyukkoYmd, shanghai_syorui_ymd = @SyoriYmd, " &
                   "shanghai_header_ymd = @SyoriYmd, shinchoku_jyokyo = 3, " &
                   "update_ymd = GETDATE(), update_login = @UpdateLogin " &
                   "WHERE pk_invoice_no_main = @InvoiceNoMain AND shanghai_header_ymd IS NULL"
        End Function

        ''' <summary>
        ''' 通関依頼の進捗を更新する SQL を作成します。処理NO 7 に対応します。
        ''' </summary>
        ''' <returns>UPDATE 文。</returns>
        ''' <remarks>
        ''' 配送手配書を発行したときに進めます（menuNo=103 §16）。
        ''' 旧 VBA では `syori_no - 2` が進捗状況の値になるため、7 のときは 5 です。
        ''' </remarks>
        Public Shared Function CreateUpdateTsukanIraiProgressSql() As String
            Return "UPDATE dbo.tbl_t_import_progress " &
                   "SET tsukan_irai_ymd = @SyoriYmd, shinchoku_jyokyo = 5, " &
                   "update_ymd = GETDATE(), update_login = @UpdateLogin " &
                   "WHERE pk_invoice_no_main = @InvoiceNoMain AND tsukan_irai_ymd IS NULL"
        End Function

        ''' <summary>
        ''' 配送連絡（ロジ）の進捗を更新する SQL を作成します。処理NO 10 に対応します。
        ''' </summary>
        ''' <returns>UPDATE 文。</returns>
        ''' <remarks>
        ''' 入荷明細書を発行したときに進めます（menuNo=103 §16）。進捗状況の値は 8 です。
        ''' </remarks>
        Public Shared Function CreateUpdateHaisoLogiProgressSql() As String
            Return "UPDATE dbo.tbl_t_import_progress " &
                   "SET haiso_logi_ymd = @SyoriYmd, shinchoku_jyokyo = 8, " &
                   "update_ymd = GETDATE(), update_login = @UpdateLogin " &
                   "WHERE pk_invoice_no_main = @InvoiceNoMain AND haiso_logi_ymd IS NULL"
        End Function

        ''' <summary>
        ''' 配送連絡（貿易・日新）の進捗を更新する SQL を作成します。処理NO 12 に対応します。
        ''' </summary>
        ''' <returns>UPDATE 文。</returns>
        ''' <remarks>
        ''' 入荷スケジュールを出力したときに進めます（menuNo=103 §16）。
        ''' **この処理だけ 2 つの日付を同時に埋めます。**旧 VBA が 12 と 13 をまとめて
        ''' 扱っており、日新と貿易の両方を同じ日付にしているためです。進捗状況の値は 11 です。
        ''' </remarks>
        Public Shared Function CreateUpdateHaisoNisshinProgressSql() As String
            Return "UPDATE dbo.tbl_t_import_progress " &
                   "SET haiso_nisshin_ymd = @SyoriYmd, haiso_boueki_ymd = @SyoriYmd, " &
                   "shinchoku_jyokyo = 11, " &
                   "update_ymd = GETDATE(), update_login = @UpdateLogin " &
                   "WHERE pk_invoice_no_main = @InvoiceNoMain AND haiso_nisshin_ymd IS NULL"
        End Function
#End Region

    End Class
End Namespace
