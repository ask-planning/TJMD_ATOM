Option Strict On
Option Explicit On
Option Infer Off

Namespace Sql

    ''' <summary>
    ''' 輸入経費コード登録（menuNo=907）で使う SQL 文を組み立てます。
    ''' </summary>
    ''' <remarks>
    ''' 旧 Access はフォームをテーブルに直結し、Access の標準動作で更新していました。
    ''' ストアドは使っていなかったため、Web 版でも**テーブルを直接更新します**（設計書 §6.3）。
    ''' 値はすべて SqlParameter で渡します。文字列連結で値を埋め込みません。
    ''' </remarks>
    Public NotInheritable Class ImportKeihiCodeSql

#Region "定数"
        ''' <summary>
        ''' 楽観ロックの条件。
        ''' </summary>
        ''' <remarks>
        ''' **NULL 同士も一致とみなします。**既存行の 39 件で `update_ymd` が NULL のため、
        ''' 単純な `=` では 1 件も更新できなくなります（設計書 §13.1 ③）。
        ''' </remarks>
        Private Const ConcurrencyCondition As String =
            "pk_keihi_code = @OriginalKeihiCode " &
            "AND (update_ymd = @OriginalUpdateYmd " &
            "OR (update_ymd IS NULL AND @OriginalUpdateYmd IS NULL))"
#End Region

#Region "コンストラクタ"
        ''' <summary>インスタンス化を禁止します。</summary>
        Private Sub New()
        End Sub
#End Region

#Region "参照系"
        ''' <summary>
        ''' 経費コードを全件取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        ''' <remarks>並び順は伝票コードの昇順です（設計書 §4.2）。件数が少ないため絞り込みはしません。</remarks>
        Public Shared Function CreateSelectAllSql() As String
            Return "SELECT pk_keihi_code, keihi_code_text, keihi_code_ref, keihi_kbn, " &
                   "hoken_taisyo_flg, update_ymd, update_login " &
                   "FROM dbo.tbl_m_import_keihi_code " &
                   "ORDER BY pk_keihi_code"
        End Function

        ''' <summary>
        ''' 経費コードを 1 件取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        ''' <remarks>保存前に、更新日と「どの項目が変わったか」を調べるために使います（設計書 C-11）。</remarks>
        Public Shared Function CreateSelectByCodeSql() As String
            Return "SELECT pk_keihi_code, keihi_code_text, keihi_code_ref, keihi_kbn, " &
                   "hoken_taisyo_flg, update_ymd, update_login " &
                   "FROM dbo.tbl_m_import_keihi_code " &
                   "WHERE pk_keihi_code = @KeihiCode"
        End Function

        ''' <summary>
        ''' 経費区分の選択肢を取得する SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        ''' <remarks>
        ''' **ビュー側の `ORDER BY` は当てにしません。**`TOP` と組み合わせたときしか効かないためです。
        ''' 並びはビューの定義と同じ `sort_seq, pk_seq_no` を明示します（設計書 §13.1 ②）。
        ''' </remarks>
        Public Shared Function CreateSelectKbnSql() As String
            Return "SELECT contents1, contents2 " &
                   "FROM dbo.View__keihi_kbn " &
                   "ORDER BY sort_seq, pk_seq_no"
        End Function

        ''' <summary>
        ''' 伝票コードの重複を調べる SQL を作成します。
        ''' </summary>
        ''' <returns>SELECT 文。</returns>
        ''' <remarks>
        ''' 主キー違反を待たずに、保存前に確かめます。
        ''' コードを変更するときは、変更前のコード自身を数えないよう除きます。
        ''' </remarks>
        Public Shared Function CreateSelectDuplicateCountSql() As String
            Return "SELECT COUNT(*) " &
                   "FROM dbo.tbl_m_import_keihi_code " &
                   "WHERE pk_keihi_code = @KeihiCode AND pk_keihi_code <> @OriginalKeihiCode"
        End Function
#End Region

#Region "更新系"
        ''' <summary>
        ''' 経費コードを 1 件登録する SQL を作成します。
        ''' </summary>
        ''' <returns>INSERT 文。</returns>
        Public Shared Function CreateInsertSql() As String
            Return "INSERT INTO dbo.tbl_m_import_keihi_code " &
                   "(pk_keihi_code, keihi_code_text, keihi_code_ref, keihi_kbn, " &
                   "hoken_taisyo_flg, update_ymd, update_login) " &
                   "VALUES (@KeihiCode, @KeihiCodeText, @KeihiCodeRef, @KeihiKbn, " &
                   "@HokenTaisyoFlg, GETDATE(), @UpdateLogin)"
        End Function

        ''' <summary>
        ''' 経費コードを 1 件更新する SQL を作成します。更新日と更新者も書き換えます。
        ''' </summary>
        ''' <returns>UPDATE 文。</returns>
        ''' <remarks>
        ''' 伝票コード・経費名称・区分のいずれかが変わったときに使います（設計書 C-11）。
        ''' </remarks>
        Public Shared Function CreateUpdateWithStampSql() As String
            Return "UPDATE dbo.tbl_m_import_keihi_code " &
                   "SET pk_keihi_code = @KeihiCode, keihi_code_text = @KeihiCodeText, " &
                   "keihi_code_ref = @KeihiCodeRef, keihi_kbn = @KeihiKbn, " &
                   "hoken_taisyo_flg = @HokenTaisyoFlg, " &
                   "update_ymd = GETDATE(), update_login = @UpdateLogin " &
                   "WHERE " & ConcurrencyCondition
        End Function

        ''' <summary>
        ''' 経費コードを 1 件更新する SQL を作成します。更新日と更新者は変えません。
        ''' </summary>
        ''' <returns>UPDATE 文。</returns>
        ''' <remarks>
        ''' **REF と保険対象だけを変えたときに使います。**旧 Access はこの 2 項目に
        ''' `AfterUpdate` を書いておらず、更新日・更新者が変わりませんでした。
        ''' 旧と同じ挙動にするため、ここでも変えません（設計書 C-11 / 既知の不具合 K-1）。
        '''
        ''' **この経路では楽観ロックが働きません。**更新日が進まないため、
        ''' 同じ 2 項目を 2 人が同時に直すと、後から保存した方が黙って上書きします。
        ''' 旧と同じ挙動として残しています。
        ''' </remarks>
        Public Shared Function CreateUpdateWithoutStampSql() As String
            Return "UPDATE dbo.tbl_m_import_keihi_code " &
                   "SET pk_keihi_code = @KeihiCode, keihi_code_text = @KeihiCodeText, " &
                   "keihi_code_ref = @KeihiCodeRef, keihi_kbn = @KeihiKbn, " &
                   "hoken_taisyo_flg = @HokenTaisyoFlg " &
                   "WHERE " & ConcurrencyCondition
        End Function

        ''' <summary>
        ''' 経費コードを 1 件削除する SQL を作成します。
        ''' </summary>
        ''' <returns>DELETE 文。</returns>
        ''' <remarks>
        ''' **物理削除です。**削除フラグに当たる列がありません（設計書 §7）。
        ''' 明細（`tbl_t_import_keihi_meisai`）で使用中のコードも削除できます。
        ''' 旧にも止める仕組みと外部キーが無かったためです（C-7 / 既知の不具合 K-4）。
        ''' </remarks>
        Public Shared Function CreateDeleteSql() As String
            Return "DELETE FROM dbo.tbl_m_import_keihi_code " &
                   "WHERE " & ConcurrencyCondition
        End Function
#End Region

    End Class
End Namespace
