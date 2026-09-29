Option Strict On
Option Explicit On
Option Infer Off

Namespace Sql

    ''' <summary>
    ''' ログインに関する SQL 文を組み立てます。
    ''' 旧 Access 版の usp_syain_jyoho_get / usp_atom_login_log に相当します。
    ''' </summary>
    ''' <remarks>
    ''' ストアドプロシージャは保守性のため使用せず、同等の SQL をアプリ側で保持します。
    ''' 社員マスタは別データベースにあるため、3 部構成の名前で参照します。
    ''' </remarks>
    Public NotInheritable Class LoginSql

#Region "コンストラクタ"
        ''' <summary>インスタンス化を禁止します。</summary>
        Private Sub New()
        End Sub
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' 社員コードから社員情報を取得する SQL を作成します。
        ''' </summary>
        ''' <returns>社員情報を 1 件取得する SELECT 文。</returns>
        Public Shared Function CreateSelectLoginUserSql() As String
            Return "SELECT shain_name, login_mode_atom, bumon_code, shanghai_chotatsu, mail_address " &
                   "FROM TJM_master.dbo.tbl_m_login_sys " &
                   "WHERE pk_shain_code = @LoginCode"
        End Function

        ''' <summary>
        ''' ログイン履歴を登録する SQL を作成します。
        ''' </summary>
        ''' <returns>ログイン履歴を 1 件登録する INSERT 文。</returns>
        Public Shared Function CreateInsertLoginLogSql() As String
            Return "INSERT INTO dbo.tbl_t_atom_login_log (pk_login_code, menu_no, login_ymd, version_no) " &
                   "VALUES (@LoginCode, @MenuNo, GETDATE(), @VersionNo)"
        End Function
#End Region

    End Class
End Namespace
