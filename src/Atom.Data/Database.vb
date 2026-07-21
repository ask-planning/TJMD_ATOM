Imports System.Configuration
Imports System.Data.SqlClient

''' <summary>SQL Server 接続基盤（現行 CurrentProject.BaseConnectionString 相当）。</summary>
Public Module Database
    ''' <summary>app.config の接続文字列 "AtomDb" から接続を生成する。</summary>
    ''' <returns>未オープンの SqlConnection。</returns>
    Public Function CreateConnection() As SqlConnection
        ' 接続文字列を app.config から取得する
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("AtomDb")
        ' 未定義のとき：エラーにする
        If setting Is Nothing OrElse String.IsNullOrEmpty(setting.ConnectionString) Then
            Throw New InvalidOperationException("接続文字列 'AtomDb' が app.config に定義されていません。")
        End If
        Return New SqlConnection(setting.ConnectionString)
    End Function
End Module
