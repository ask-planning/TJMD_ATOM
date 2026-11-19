Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports Atom.Core.Entities
Imports Atom.Data.Sql

Namespace Repositories

    ''' <summary>
    ''' 輸入経費コード登録（menuNo=907）のデータアクセスを担います。
    ''' </summary>
    ''' <remarks>
    ''' 値はすべて SqlParameter で渡し、型と長さを明示します。
    ''' 長さは実テーブルの定義に合わせています（設計書 §13.1 ①）。
    ''' 接続とトランザクションは呼び出し側が持ちます。
    ''' </remarks>
    Public Class ImportKeihiCodeRepository

#Region "定数"
        ''' <summary>伝票コードの長さ。</summary>
        Private Const KeihiCodeLength As Integer = 4

        ''' <summary>経費名称の長さ。</summary>
        Private Const KeihiCodeTextLength As Integer = 50

        ''' <summary>REF の長さ。</summary>
        Private Const KeihiCodeRefLength As Integer = 50

        ''' <summary>区分の長さ。</summary>
        Private Const KeihiKbnLength As Integer = 3

        ''' <summary>更新者の長さ。</summary>
        Private Const UpdateLoginLength As Integer = 6
#End Region

#Region "参照系"
        ''' <summary>
        ''' 経費コードを全件取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <returns>伝票コードの昇順の一覧。</returns>
        Public Function SelectAll(connection As SqlConnection,
                                  transaction As SqlTransaction) As IList(Of ImportKeihiCode)

            Dim results As New List(Of ImportKeihiCode)()

            Using command As New SqlCommand(ImportKeihiCodeSql.CreateSelectAllSql(), connection, transaction)
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        results.Add(ReadRow(reader))
                    End While
                End Using
            End Using

            Return results
        End Function

        ''' <summary>
        ''' 経費コードを 1 件取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="keihiCode">伝票コード。</param>
        ''' <returns>見つかった 1 件。無いときは Nothing。</returns>
        Public Function SelectByCode(connection As SqlConnection, transaction As SqlTransaction,
                                     keihiCode As String) As ImportKeihiCode

            Using command As New SqlCommand(ImportKeihiCodeSql.CreateSelectByCodeSql(), connection, transaction)
                command.Parameters.Add("@KeihiCode", SqlDbType.NVarChar, KeihiCodeLength).Value = keihiCode

                Using reader As SqlDataReader = command.ExecuteReader()
                    ' 見つからないとき
                    If Not reader.Read() Then
                        Return Nothing
                    End If
                    Return ReadRow(reader)
                End Using
            End Using
        End Function

        ''' <summary>
        ''' 経費区分の選択肢を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <returns>ビューの並び順のままの一覧。</returns>
        Public Function SelectKbnList(connection As SqlConnection,
                                      transaction As SqlTransaction) As IList(Of ImportKeihiKbn)

            Dim results As New List(Of ImportKeihiKbn)()

            Using command As New SqlCommand(ImportKeihiCodeSql.CreateSelectKbnSql(), connection, transaction)
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        Dim kbn As New ImportKeihiKbn()
                        kbn.Code = GetString(reader, "contents1")
                        kbn.Name = GetString(reader, "contents2")
                        results.Add(kbn)
                    End While
                End Using
            End Using

            Return results
        End Function

        ''' <summary>
        ''' 伝票コードが既に使われているかを調べます。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="keihiCode">これから使う伝票コード。</param>
        ''' <param name="originalKeihiCode">変更前の伝票コード。新規のときは空文字。</param>
        ''' <returns>既に使われているとき True。</returns>
        Public Function ExistsDuplicate(connection As SqlConnection, transaction As SqlTransaction,
                                        keihiCode As String, originalKeihiCode As String) As Boolean

            Using command As New SqlCommand(ImportKeihiCodeSql.CreateSelectDuplicateCountSql(),
                                            connection, transaction)
                command.Parameters.Add("@KeihiCode", SqlDbType.NVarChar, KeihiCodeLength).Value = keihiCode
                command.Parameters.Add("@OriginalKeihiCode", SqlDbType.NVarChar, KeihiCodeLength).Value =
                    If(originalKeihiCode, String.Empty)

                Dim count As Integer = Convert.ToInt32(command.ExecuteScalar(),
                                                       Globalization.CultureInfo.InvariantCulture)
                Return count > 0
            End Using
        End Function
#End Region

#Region "更新系"
        ''' <summary>
        ''' 経費コードを 1 件登録します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="entity">登録する内容。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        Public Sub Insert(connection As SqlConnection, transaction As SqlTransaction,
                          entity As ImportKeihiCode, loginCode As String)

            Using command As New SqlCommand(ImportKeihiCodeSql.CreateInsertSql(), connection, transaction)
                AddValueParameters(command, entity)
                command.Parameters.Add("@UpdateLogin", SqlDbType.NVarChar, UpdateLoginLength).Value = loginCode
                command.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>
        ''' 経費コードを 1 件更新します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="entity">更新後の内容。</param>
        ''' <param name="original">更新前の内容。楽観ロックの比較に使います。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="updateStamp">更新日と更新者を書き換えるときは True。</param>
        ''' <returns>更新した件数。0 のときは他の利用者が先に更新しています。</returns>
        ''' <remarks>
        ''' `updateStamp` が False のときは更新日・更新者を変えません。
        ''' 旧 Access が REF・保険対象の変更では更新日を書き換えなかったためです
        ''' （設計書 C-11 / 既知の不具合 K-1）。
        ''' </remarks>
        Public Function Update(connection As SqlConnection, transaction As SqlTransaction,
                               entity As ImportKeihiCode, original As ImportKeihiCode,
                               loginCode As String, updateStamp As Boolean) As Integer

            Dim sql As String = If(updateStamp,
                                   ImportKeihiCodeSql.CreateUpdateWithStampSql(),
                                   ImportKeihiCodeSql.CreateUpdateWithoutStampSql())

            Using command As New SqlCommand(sql, connection, transaction)
                AddValueParameters(command, entity)
                AddConcurrencyParameters(command, original)

                ' 更新日を書き換えるとき
                If updateStamp Then
                    command.Parameters.Add("@UpdateLogin", SqlDbType.NVarChar, UpdateLoginLength).Value = loginCode
                End If

                Return command.ExecuteNonQuery()
            End Using
        End Function

        ''' <summary>
        ''' 経費コードを 1 件削除します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="original">削除する行。楽観ロックの比較に使います。</param>
        ''' <returns>削除した件数。0 のときは他の利用者が先に更新または削除しています。</returns>
        Public Function Delete(connection As SqlConnection, transaction As SqlTransaction,
                               original As ImportKeihiCode) As Integer

            Using command As New SqlCommand(ImportKeihiCodeSql.CreateDeleteSql(), connection, transaction)
                AddConcurrencyParameters(command, original)
                Return command.ExecuteNonQuery()
            End Using
        End Function
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 値のパラメータを追加します。
        ''' </summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="entity">内容。</param>
        Private Shared Sub AddValueParameters(command As SqlCommand, entity As ImportKeihiCode)
            command.Parameters.Add("@KeihiCode", SqlDbType.NVarChar, KeihiCodeLength).Value = entity.KeihiCode
            AddNullableString(command, "@KeihiCodeText", KeihiCodeTextLength, entity.KeihiCodeText)
            AddNullableString(command, "@KeihiCodeRef", KeihiCodeRefLength, entity.KeihiCodeRef)
            AddNullableString(command, "@KeihiKbn", KeihiKbnLength, entity.KeihiKbn)
            command.Parameters.Add("@HokenTaisyoFlg", SqlDbType.Int).Value = entity.HokenTaisyoFlg
        End Sub

        ''' <summary>
        ''' 楽観ロックのパラメータを追加します。
        ''' </summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="original">更新前の内容。</param>
        Private Shared Sub AddConcurrencyParameters(command As SqlCommand, original As ImportKeihiCode)
            command.Parameters.Add("@OriginalKeihiCode", SqlDbType.NVarChar, KeihiCodeLength).Value =
                original.KeihiCode

            Dim parameter As SqlParameter = command.Parameters.Add("@OriginalUpdateYmd", SqlDbType.DateTime)
            ' 更新日が無いとき（既存行に多い）
            If Not original.UpdateYmd.HasValue Then
                parameter.Value = DBNull.Value
                Return
            End If
            parameter.Value = original.UpdateYmd.Value
        End Sub

        ''' <summary>
        ''' 空文字を NULL として扱う文字列のパラメータを追加します。
        ''' </summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="name">パラメータ名。</param>
        ''' <param name="length">長さ。</param>
        ''' <param name="value">値。</param>
        ''' <remarks>
        ''' 旧 Access は未入力の欄を NULL で保存していました。空文字と NULL が混ざると
        ''' 検索や比較で扱いが分かれるため、空文字は NULL に寄せます。
        ''' </remarks>
        Private Shared Sub AddNullableString(command As SqlCommand, name As String,
                                             length As Integer, value As String)
            Dim parameter As SqlParameter = command.Parameters.Add(name, SqlDbType.NVarChar, length)
            ' 未入力のとき
            If String.IsNullOrEmpty(value) Then
                parameter.Value = DBNull.Value
                Return
            End If
            parameter.Value = value
        End Sub

        ''' <summary>
        ''' 1 行を読み取ります。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <returns>読み取った内容。</returns>
        Private Shared Function ReadRow(reader As SqlDataReader) As ImportKeihiCode
            Dim entity As New ImportKeihiCode()
            entity.KeihiCode = GetString(reader, "pk_keihi_code")
            entity.KeihiCodeText = GetString(reader, "keihi_code_text")
            entity.KeihiCodeRef = GetString(reader, "keihi_code_ref")
            entity.KeihiKbn = GetString(reader, "keihi_kbn")
            entity.HokenTaisyoFlg = GetInteger(reader, "hoken_taisyo_flg")
            entity.UpdateYmd = GetNullableDate(reader, "update_ymd")
            entity.UpdateLogin = GetString(reader, "update_login")
            Return entity
        End Function

        ''' <summary>
        ''' 文字列の列を読み取ります。NULL は空文字にします。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>読み取った値。</returns>
        Private Shared Function GetString(reader As SqlDataReader, columnName As String) As String
            Dim index As Integer = reader.GetOrdinal(columnName)
            ' NULL のとき
            If reader.IsDBNull(index) Then
                Return String.Empty
            End If
            Return Convert.ToString(reader.GetValue(index), Globalization.CultureInfo.InvariantCulture)
        End Function

        ''' <summary>
        ''' 整数の列を読み取ります。NULL は 0 にします。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>読み取った値。</returns>
        ''' <remarks>
        ''' 保険対象は列が NULL 可ですが、実データに NULL はありません。
        ''' 万一 NULL があっても「対象外」として扱います（設計書 C-32）。
        ''' </remarks>
        Private Shared Function GetInteger(reader As SqlDataReader, columnName As String) As Integer
            Dim index As Integer = reader.GetOrdinal(columnName)
            ' NULL のとき
            If reader.IsDBNull(index) Then
                Return 0
            End If
            Return Convert.ToInt32(reader.GetValue(index), Globalization.CultureInfo.InvariantCulture)
        End Function

        ''' <summary>
        ''' 日付の列を読み取ります。NULL は Nothing にします。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>読み取った値。</returns>
        Private Shared Function GetNullableDate(reader As SqlDataReader, columnName As String) As Date?
            Dim index As Integer = reader.GetOrdinal(columnName)
            ' NULL のとき
            If reader.IsDBNull(index) Then
                Return Nothing
            End If
            Return Convert.ToDateTime(reader.GetValue(index), Globalization.CultureInfo.InvariantCulture)
        End Function
#End Region

    End Class
End Namespace
