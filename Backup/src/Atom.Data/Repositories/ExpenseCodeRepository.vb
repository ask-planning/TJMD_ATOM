Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports Atom.Core.Entities
Imports Atom.Core.Exceptions
Imports Atom.Core.Logging
Imports Atom.Data.Sql

Namespace Repositories

    ''' <summary>
    ''' 経費コード（menuNo=907）のデータアクセスを担います。
    ''' </summary>
    Public Class ExpenseCodeRepository

#Region "定数"
        ''' <summary>経費コードの桁数上限。※ 実テーブル定義に合わせて調整してください。</summary>
        Private Const KeihiCodeMaxLength As Integer = 10

        ''' <summary>経費名称の桁数上限。※ 実テーブル定義に合わせて調整してください。</summary>
        Private Const KeihiCodeTextMaxLength As Integer = 100

        ''' <summary>経費区分の桁数上限。※ 実テーブル定義に合わせて調整してください。</summary>
        Private Const KeihiKbnMaxLength As Integer = 10

        ''' <summary>ログインコードの桁数上限。</summary>
        Private Const LoginCodeMaxLength As Integer = 20

        ''' <summary>ログの発生元名。</summary>
        Private Const LogSource As String = "ExpenseCodeRepository"
#End Region

#Region "フィールド"
        ''' <summary>操作している利用者のログインコード。</summary>
        Private ReadOnly _loginCode As String
#End Region

#Region "コンストラクタ"
        ''' <summary>
        ''' 利用者を指定して初期化します。
        ''' </summary>
        ''' <param name="loginCode">操作している利用者のログインコード。</param>
        Public Sub New(loginCode As String)
            Me._loginCode = loginCode
        End Sub
#End Region

#Region "参照系"
        ''' <summary>
        ''' 経費コードを全件取得します。
        ''' </summary>
        ''' <returns>経費コード順に並んだ一覧。該当が無いときは空の一覧。</returns>
        Public Function SelectAll() As IList(Of ExpenseCode)
            Dim results As New List(Of ExpenseCode)()

            Try
                Using connection As SqlConnection = Database.CreateOpenConnection()
                    Using command As New SqlCommand(ExpenseCodeSql.SelectAll, connection)
                        Using reader As SqlDataReader = command.ExecuteReader()
                            While reader.Read()
                                results.Add(Me.CreateEntity(reader))
                            End While
                        End Using
                    End Using
                End Using
            Catch ex As SqlException
                Logger.WriteError(LogSource, "経費コードの一覧取得に失敗しました。", ex, Me._loginCode)
                Throw
            End Try

            Return results
        End Function

        ''' <summary>
        ''' 経費コードを条件に 1 件取得します。
        ''' </summary>
        ''' <param name="keihiCode">経費コード。</param>
        ''' <returns>該当する 1 件。存在しないときは Nothing。</returns>
        Public Function SelectByKey(keihiCode As String) As ExpenseCode
            Try
                Using connection As SqlConnection = Database.CreateOpenConnection()
                    Using command As New SqlCommand(ExpenseCodeSql.SelectByKey, connection)
                        Me.AddKeihiCodeParameter(command, "@KeihiCode", keihiCode)
                        Using reader As SqlDataReader = command.ExecuteReader()
                            ' 該当があるとき
                            If reader.Read() Then
                                Return Me.CreateEntity(reader)
                            End If
                        End Using
                    End Using
                End Using
            Catch ex As SqlException
                Logger.WriteError(LogSource, "経費コードの取得に失敗しました。", ex, Me._loginCode)
                Throw
            End Try

            ' 該当が無いとき
            Return Nothing
        End Function
#End Region

#Region "更新系"
        ''' <summary>
        ''' 経費コードを 1 件登録します。
        ''' </summary>
        ''' <param name="entity">登録する内容。</param>
        Public Sub Insert(entity As ExpenseCode)
            Try
                Using connection As SqlConnection = Database.CreateOpenConnection()
                    Using command As New SqlCommand(ExpenseCodeSql.Insert, connection)
                        Me.AddKeihiCodeParameter(command, "@KeihiCode", entity.KeihiCode)
                        Me.AddValueParameters(command, entity)
                        command.ExecuteNonQuery()
                    End Using
                End Using
            Catch ex As SqlException
                Logger.WriteError(LogSource, "経費コードの登録に失敗しました。", ex, Me._loginCode)
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' 経費コードを 1 件更新します。経費コード自体の変更にも対応します。
        ''' 更新日時が一致しないときは競合として扱います。
        ''' </summary>
        ''' <param name="entity">更新後の内容。</param>
        ''' <param name="originalKeihiCode">更新前の経費コード。</param>
        Public Sub Update(entity As ExpenseCode, originalKeihiCode As String)
            Try
                Using connection As SqlConnection = Database.CreateOpenConnection()
                    Using command As New SqlCommand(ExpenseCodeSql.Update, connection)
                        Me.AddKeihiCodeParameter(command, "@KeihiCode", entity.KeihiCode)
                        Me.AddKeihiCodeParameter(command, "@OriginalKeihiCode", originalKeihiCode)
                        Me.AddValueParameters(command, entity)
                        Me.AddUpdateYmdParameter(command, entity.UpdateYmd)

                        Dim affected As Integer = command.ExecuteNonQuery()
                        ' 1 件も更新されなかったとき（他の利用者が先に更新している）
                        If affected = 0 Then
                            Throw New ConcurrencyException()
                        End If
                    End Using
                End Using
            Catch ex As SqlException
                Logger.WriteError(LogSource, "経費コードの更新に失敗しました。", ex, Me._loginCode)
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' 経費コードを 1 件削除します。更新日時が一致しないときは競合として扱います。
        ''' </summary>
        ''' <param name="keihiCode">経費コード。</param>
        ''' <param name="updateYmd">画面に表示していた時点の更新日時。</param>
        Public Sub Delete(keihiCode As String, updateYmd As Date?)
            Try
                Using connection As SqlConnection = Database.CreateOpenConnection()
                    Using command As New SqlCommand(ExpenseCodeSql.Delete, connection)
                        Me.AddKeihiCodeParameter(command, "@KeihiCode", keihiCode)
                        Me.AddUpdateYmdParameter(command, updateYmd)

                        Dim affected As Integer = command.ExecuteNonQuery()
                        ' 1 件も削除されなかったとき
                        If affected = 0 Then
                            Throw New ConcurrencyException()
                        End If
                    End Using
                End Using
            Catch ex As SqlException
                Logger.WriteError(LogSource, "経費コードの削除に失敗しました。", ex, Me._loginCode)
                Throw
            End Try
        End Sub
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 読み取り結果からエンティティを作成します。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <returns>作成したエンティティ。</returns>
        Private Function CreateEntity(reader As SqlDataReader) As ExpenseCode
            Dim entity As New ExpenseCode()
            entity.KeihiCode = Me.GetString(reader, "pk_keihi_code")
            entity.KeihiCodeText = Me.GetString(reader, "keihi_code_text")
            entity.KeihiKbn = Me.GetString(reader, "keihi_kbn")
            entity.UpdateYmd = Me.GetNullableDate(reader, "update_ymd")
            entity.UpdateLogin = Me.GetString(reader, "update_login")
            Return entity
        End Function

        ''' <summary>経費コードのパラメータを追加します。</summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="parameterName">パラメータ名。</param>
        ''' <param name="keihiCode">経費コード。</param>
        Private Sub AddKeihiCodeParameter(command As SqlCommand, parameterName As String, keihiCode As String)
            command.Parameters.Add(parameterName, SqlDbType.NVarChar, KeihiCodeMaxLength).Value = keihiCode
        End Sub

        ''' <summary>更新対象の値のパラメータを追加します。</summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="entity">値を持つエンティティ。</param>
        Private Sub AddValueParameters(command As SqlCommand, entity As ExpenseCode)
            command.Parameters.Add("@KeihiCodeText", SqlDbType.NVarChar, KeihiCodeTextMaxLength).Value = entity.KeihiCodeText
            command.Parameters.Add("@KeihiKbn", SqlDbType.NVarChar, KeihiKbnMaxLength).Value = entity.KeihiKbn
            command.Parameters.Add("@UpdateLogin", SqlDbType.NVarChar, LoginCodeMaxLength).Value = Me._loginCode
        End Sub

        ''' <summary>楽観ロック用の更新日時パラメータを追加します。</summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="updateYmd">更新日時。</param>
        Private Sub AddUpdateYmdParameter(command As SqlCommand, updateYmd As Date?)
            Dim parameter As SqlParameter = command.Parameters.Add("@UpdateYmd", SqlDbType.DateTime)
            ' 更新日時が記録されているとき
            If updateYmd.HasValue Then
                parameter.Value = updateYmd.Value
            Else
                parameter.Value = DBNull.Value
            End If
        End Sub

        ''' <summary>文字列の列を取得します。NULL のときは空文字を返します。</summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>列の値。</returns>
        Private Function GetString(reader As SqlDataReader, columnName As String) As String
            Dim index As Integer = reader.GetOrdinal(columnName)
            ' NULL のとき
            If reader.IsDBNull(index) Then
                Return String.Empty
            End If
            Return reader.GetValue(index).ToString()
        End Function

        ''' <summary>日時の列を取得します。NULL のときは Nothing を返します。</summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>列の値。</returns>
        Private Function GetNullableDate(reader As SqlDataReader, columnName As String) As Date?
            Dim index As Integer = reader.GetOrdinal(columnName)
            ' NULL のとき
            If reader.IsDBNull(index) Then
                Return Nothing
            End If
            Return reader.GetDateTime(index)
        End Function
#End Region

    End Class
End Namespace
