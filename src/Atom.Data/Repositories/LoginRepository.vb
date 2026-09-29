Option Strict On
Option Explicit On
Option Infer Off

Imports System.Data
Imports System.Data.SqlClient
Imports Atom.Core.Entities
Imports Atom.Core.Logging
Imports Atom.Data.Sql

Namespace Repositories

    ''' <summary>
    ''' ログインに関するデータアクセスを担います。
    ''' </summary>
    Public Class LoginRepository

#Region "定数"
        ''' <summary>社員コードの桁数。</summary>
        Private Const LoginCodeLength As Integer = 6

        ''' <summary>社員名の桁数上限。</summary>
        Private Const LoginNameMaxLength As Integer = 50

        ''' <summary>部門コードの桁数上限。</summary>
        Private Const BumonCodeMaxLength As Integer = 10

        ''' <summary>メールアドレスの桁数上限。</summary>
        Private Const MailAddressMaxLength As Integer = 50

        ''' <summary>バージョン番号の桁数上限。</summary>
        Private Const VersionNoMaxLength As Integer = 20

        ''' <summary>ログの発生元名。</summary>
        Private Const LogSource As String = "LoginRepository"
#End Region

#Region "参照系"
        ''' <summary>
        ''' 社員コードから社員情報を 1 件取得します。
        ''' </summary>
        ''' <param name="loginCode">社員コード（6桁）。</param>
        ''' <returns>該当する社員情報。存在しないときは Nothing。</returns>
        Public Function SelectLoginUser(loginCode As String) As LoginUser
            Try
                Using connection As SqlConnection = Database.CreateOpenConnection()
                    Using command As New SqlCommand(LoginSql.CreateSelectLoginUserSql(), connection)
                        command.Parameters.Add("@LoginCode", SqlDbType.NVarChar, LoginCodeLength).Value = loginCode

                        Using reader As SqlDataReader = command.ExecuteReader()
                            ' 該当があるとき
                            If reader.Read() Then
                                Return Me.CreateEntity(reader, loginCode)
                            End If
                        End Using
                    End Using
                End Using
            Catch ex As SqlException
                Logger.WriteError(LogSource, "社員情報の取得に失敗しました。", ex, loginCode)
                Throw
            End Try

            ' 該当が無いとき
            Return Nothing
        End Function
#End Region

#Region "更新系"
        ''' <summary>
        ''' ログイン履歴を 1 件登録します。
        ''' </summary>
        ''' <param name="loginCode">社員コード。</param>
        ''' <param name="versionNo">システムバージョン。</param>
        ''' <param name="menuNo">操作したメニュー番号。ログイン時は 0 を渡します。</param>
        Public Sub InsertLoginLog(loginCode As String, versionNo As String, menuNo As Integer)
            Try
                Using connection As SqlConnection = Database.CreateOpenConnection()
                    Using command As New SqlCommand(LoginSql.CreateInsertLoginLogSql(), connection)
                        command.Parameters.Add("@LoginCode", SqlDbType.NVarChar, LoginCodeLength).Value = loginCode
                        command.Parameters.Add("@MenuNo", SqlDbType.Int).Value = menuNo
                        command.Parameters.Add("@VersionNo", SqlDbType.NVarChar, VersionNoMaxLength).Value = versionNo
                        command.ExecuteNonQuery()
                    End Using
                End Using
            Catch ex As SqlException
                ' 履歴の記録に失敗してもログインは継続させるため、記録のみ行い再送出しません。
                Logger.WriteError(LogSource, "ログイン履歴の登録に失敗しました。", ex, loginCode)
            End Try
        End Sub
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 読み取り結果から社員情報を作成します。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="loginCode">検索に使った社員コード。</param>
        ''' <returns>作成した社員情報。</returns>
        Private Function CreateEntity(reader As SqlDataReader, loginCode As String) As LoginUser
            Dim entity As New LoginUser()
            entity.LoginCode = loginCode
            entity.LoginName = Me.GetString(reader, "shain_name", LoginNameMaxLength)
            entity.ModeAtom = Me.GetNullableInteger(reader, "login_mode_atom")
            entity.BumonCode = Me.GetString(reader, "bumon_code", BumonCodeMaxLength)
            entity.ShanghaiChotatsu = Me.GetNullableInteger(reader, "shanghai_chotatsu")
            entity.MailAddress = Me.GetString(reader, "mail_address", MailAddressMaxLength)
            Return entity
        End Function

        ''' <summary>
        ''' 文字列の列を取得します。NULL のときは空文字を返します。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <param name="maxLength">列の桁数上限。読み取り結果の検証には使いません。</param>
        ''' <returns>列の値。</returns>
        Private Function GetString(reader As SqlDataReader, columnName As String, maxLength As Integer) As String
            Dim index As Integer = reader.GetOrdinal(columnName)
            ' NULL のとき
            If reader.IsDBNull(index) Then
                Return String.Empty
            End If

            Dim value As String = reader.GetString(index)
            ' 想定外に長い値が入っていたとき（画面表示の崩れを防ぐため切り詰める）
            If value.Length > maxLength Then
                Return value.Substring(0, maxLength)
            End If
            Return value
        End Function

        ''' <summary>
        ''' 数値の列を取得します。NULL のときは Nothing を返します。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>列の値。</returns>
        Private Function GetNullableInteger(reader As SqlDataReader, columnName As String) As Integer?
            Dim index As Integer = reader.GetOrdinal(columnName)
            ' NULL のとき
            If reader.IsDBNull(index) Then
                Return Nothing
            End If
            Return reader.GetInt32(index)
        End Function
#End Region

    End Class
End Namespace
