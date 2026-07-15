\xEF\xBB\xBFImports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient

Namespace Repositories
    ''' <summary>
    ''' ログイン者の権限メニューを取得（現行 usp_koumoku_contents_get('MEN', ...) 相当）。
    ''' ※ 引数名・戻り列は現行DB定義で要確認（ここでは仮）。
    ''' </summary>
    Public Class MenuPermissionRepository

#Region "公開メソッド"
        ''' <summary>ログイン者が使える許可メニュー番号の一覧を取得する。</summary>
        ''' <param name="loginCode">ログイン社員コード。</param>
        ''' <returns>許可メニュー番号の一覧。</returns>
        Public Function SelectAllowedMenus(loginCode As String) As List(Of Integer)
            Dim menus As New List(Of Integer)()
            ' 接続を開いてストアドを呼ぶ
            Using conn As SqlConnection = Database.CreateConnection()
                Using command As New SqlCommand("usp_koumoku_contents_get", conn)
                    command.CommandType = CommandType.StoredProcedure
                    command.Parameters.Add("@kbn", SqlDbType.NVarChar).Value = "MEN"          ' 仮
                    command.Parameters.Add("@login_code", SqlDbType.NVarChar).Value = loginCode ' 仮
                    conn.Open()
                    ' 取得結果を1件ずつ一覧へ詰める
                    Using reader As SqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            menus.Add(Convert.ToInt32(reader("menu_no")))                     ' 仮の列名
                        End While
                    End Using
                End Using
            End Using
            Return menus
        End Function
#End Region

    End Class
End Namespace
