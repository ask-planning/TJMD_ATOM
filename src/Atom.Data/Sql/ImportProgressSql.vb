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
#End Region

    End Class
End Namespace
