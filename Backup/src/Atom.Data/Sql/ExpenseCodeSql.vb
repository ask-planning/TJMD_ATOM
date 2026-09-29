Option Strict On
Option Explicit On
Option Infer Off

Namespace Sql

    ''' <summary>
    ''' 経費コード（menuNo=907）の SQL を保持します。
    ''' 値は必ずパラメータで渡し、文字列連結で埋め込みません。
    ''' </summary>
    Public NotInheritable Class ExpenseCodeSql

#Region "定数"
        ''' <summary>
        ''' 対象テーブル。
        ''' ※ テーブル名は命名規則（tbl_m_○○）からの推測です。実名に合わせて修正してください。
        ''' 修正はこの 1 行のみで済みます。
        ''' </summary>
        Private Const TableName As String = "dbo.tbl_m_keihi_code"

        ''' <summary>取得する列。</summary>
        Private Const SelectColumns As String =
            "SELECT pk_keihi_code" &
            "     , keihi_code_text" &
            "     , keihi_kbn" &
            "     , update_ymd" &
            "     , update_login" &
            "  FROM " & TableName

        ''' <summary>全件を取得する SQL。</summary>
        Public Const SelectAll As String =
            SelectColumns &
            " ORDER BY pk_keihi_code"

        ''' <summary>経費コードを条件に 1 件取得する SQL。</summary>
        Public Const SelectByKey As String =
            SelectColumns &
            " WHERE pk_keihi_code = @KeihiCode"

        ''' <summary>1 件登録する SQL。</summary>
        Public Const Insert As String =
            "INSERT INTO " & TableName & " (" &
            "    pk_keihi_code, keihi_code_text, keihi_kbn, update_ymd, update_login" &
            ") VALUES (" &
            "    @KeihiCode, @KeihiCodeText, @KeihiKbn, GETDATE(), @UpdateLogin" &
            ")"

        ''' <summary>
        ''' 1 件更新する SQL。
        ''' 経費コード自体の変更にも対応するため、更新前のコードを条件に指定します。
        ''' 更新日時が一致しないときは 0 件になります（NULL 同士も一致として扱います）。
        ''' </summary>
        Public Const Update As String =
            "UPDATE " & TableName &
            "   SET pk_keihi_code   = @KeihiCode" &
            "     , keihi_code_text = @KeihiCodeText" &
            "     , keihi_kbn       = @KeihiKbn" &
            "     , update_ymd      = GETDATE()" &
            "     , update_login    = @UpdateLogin" &
            " WHERE pk_keihi_code = @OriginalKeihiCode" &
            "   AND (update_ymd = @UpdateYmd OR (update_ymd IS NULL AND @UpdateYmd IS NULL))"

        ''' <summary>1 件削除する SQL。更新日時が一致しないときは 0 件になります。</summary>
        Public Const Delete As String =
            "DELETE FROM " & TableName &
            " WHERE pk_keihi_code = @KeihiCode" &
            "   AND (update_ymd = @UpdateYmd OR (update_ymd IS NULL AND @UpdateYmd IS NULL))"
#End Region

#Region "コンストラクタ"
        ''' <summary>インスタンス化を禁止します。</summary>
        Private Sub New()
        End Sub
#End Region

    End Class
End Namespace
