Option Strict On
Option Explicit On
Option Infer Off

Imports System.Data
Imports System.Data.SqlClient
Imports Atom.Core.Logging

Namespace Logging

    ''' <summary>
    ''' SQL Server のログテーブルへログを書き出します。
    ''' </summary>
    Public Class SqlLogWriter
        Implements ILogWriter

#Region "定数"
        ''' <summary>ログを 1 件登録する SQL。</summary>
        Private Const InsertLogSql As String =
            "INSERT INTO dbo.t_app_log (" &
            "    log_level, log_source, log_message, log_detail, login_code, creation_time" &
            ") VALUES (" &
            "    @LogLevel, @LogSource, @LogMessage, @LogDetail, @LoginCode, GETDATE()" &
            ")"

        ''' <summary>発生元の桁数上限。</summary>
        Private Const SourceMaxLength As Integer = 200

        ''' <summary>文言の桁数上限。</summary>
        Private Const MessageMaxLength As Integer = 1000
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' ログを 1 件書き出します。
        ''' </summary>
        ''' <param name="level">重要度。</param>
        ''' <param name="source">発生元。</param>
        ''' <param name="message">記録する文言。</param>
        ''' <param name="detail">詳細。</param>
        ''' <param name="loginCode">ログインコード。</param>
        Public Sub Write(level As LogLevel,
                         source As String,
                         message As String,
                         detail As String,
                         loginCode As String) Implements ILogWriter.Write

            Using connection As SqlConnection = Database.CreateOpenConnection()
                Using command As New SqlCommand(InsertLogSql, connection)
                    command.Parameters.Add("@LogLevel", SqlDbType.Int).Value = CInt(level)
                    command.Parameters.Add("@LogSource", SqlDbType.NVarChar, SourceMaxLength).Value = Truncate(source, SourceMaxLength)
                    command.Parameters.Add("@LogMessage", SqlDbType.NVarChar, MessageMaxLength).Value = Truncate(message, MessageMaxLength)
                    command.Parameters.Add("@LogDetail", SqlDbType.NVarChar, -1).Value = ToDbValue(detail)
                    command.Parameters.Add("@LoginCode", SqlDbType.NVarChar, 20).Value = ToDbValue(loginCode)
                    command.ExecuteNonQuery()
                End Using
            End Using
        End Sub
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 指定桁数に収まるよう文言を切り詰めます。
        ''' </summary>
        ''' <param name="value">対象の文言。</param>
        ''' <param name="maxLength">桁数上限。</param>
        ''' <returns>切り詰めた文言。</returns>
        Private Shared Function Truncate(value As String, maxLength As Integer) As String
            ' 未設定のとき
            If String.IsNullOrEmpty(value) Then
                Return String.Empty
            End If
            ' 上限を超えるとき
            If value.Length > maxLength Then
                Return value.Substring(0, maxLength)
            End If
            Return value
        End Function

        ''' <summary>
        ''' 空文字を DBNull に変換します。
        ''' </summary>
        ''' <param name="value">対象の文言。</param>
        ''' <returns>データベースへ渡す値。</returns>
        Private Shared Function ToDbValue(value As String) As Object
            ' 未設定のとき
            If String.IsNullOrEmpty(value) Then
                Return DBNull.Value
            End If
            Return value
        End Function
#End Region

    End Class
End Namespace
