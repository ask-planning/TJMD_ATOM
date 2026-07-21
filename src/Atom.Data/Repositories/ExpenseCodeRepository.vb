Imports System.Data
Imports System.Data.SqlClient

Namespace Repositories
    ''' <summary>経費コードマスタ（tbl_m_import_keihi_code）のデータアクセス。SQLはこのクラスに集約（インライン統一）。</summary>
    Public Class ExpenseCodeRepository

#Region "定数"
        ' 一覧取得SQL（長文は XMLリテラルで複数行）
        Private Shared ReadOnly SelectAllSql As String =
            <sql>
                SELECT pk_keihi_code, keihi_code_text, keihi_code_ref, keihi_kbn,
                       hoken_taisyo_flg, update_ymd, update_login
                FROM   tbl_m_import_keihi_code
                ORDER  BY pk_keihi_code
            </sql>.Value

        ' 1件削除SQL（値はパラメータ化）
        Private Shared ReadOnly DeleteSql As String =
            "DELETE FROM tbl_m_import_keihi_code WHERE pk_keihi_code = @pk"
#End Region

#Region "公開メソッド"
        ''' <summary>経費コードマスタを全件取得する。</summary>
        ''' <returns>画面表示用の一覧（DataTable）。</returns>
        Public Function SelectAll() As DataTable
            Dim table As New DataTable()
            ' 接続を開いて一覧を取得する
            Using conn As SqlConnection = Database.CreateConnection()
                Using adapter As New SqlDataAdapter(SelectAllSql, conn)
                    adapter.Fill(table)
                End Using
            End Using
            Return table
        End Function

        ''' <summary>経費コードを1件削除する。</summary>
        ''' <param name="pkKeihiCode">削除する経費コード。</param>
        Public Sub DeleteExpenseCode(pkKeihiCode As String)
            ' 接続を開いてパラメータ付きで削除する
            Using conn As SqlConnection = Database.CreateConnection()
                Using command As New SqlCommand(DeleteSql, conn)
                    command.Parameters.Add("@pk", SqlDbType.NVarChar).Value = pkKeihiCode
                    conn.Open()
                    command.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        ''' <summary>伝票コードコンボ用：コードと名称の一覧を取得する（マスタ自身から）。</summary>
        ''' <returns>pk_keihi_code / keihi_code_text の一覧。</returns>
        Public Function SelectCodeList() As DataTable
            Dim table As New DataTable()
            ' コードと名称を取得する
            Using conn As SqlConnection = Database.CreateConnection()
                Using adapter As New SqlDataAdapter("SELECT pk_keihi_code, keihi_code_text FROM tbl_m_import_keihi_code ORDER BY pk_keihi_code", conn)
                    adapter.Fill(table)
                End Using
            End Using
            Return table
        End Function

        ''' <summary>区分コンボ用：区分の（コード, 名称）一覧を取得する（View__keihi_kbn の先頭2列）。※列順は現行DBで要確認。</summary>
        ''' <returns>code / name の2列を持つ一覧。</returns>
        Public Function SelectKbnList() As DataTable
            Dim result As New DataTable()
            result.Columns.Add("code", GetType(String))
            result.Columns.Add("name", GetType(String))
            ' 区分ビューから候補を取得する（0列目=コード, 1列目=名称 を想定）
            Using conn As SqlConnection = Database.CreateConnection()
                Using command As New SqlCommand("SELECT * FROM View__keihi_kbn", conn)
                    conn.Open()
                    Using reader As SqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            Dim code As String = Convert.ToString(reader(0))
                            Dim displayName As String = code
                            If reader.FieldCount > 1 Then
                                displayName = Convert.ToString(reader(1))
                            End If
                            result.Rows.Add(code, displayName)
                        End While
                    End Using
                End Using
            End Using
            Return result
        End Function

        ' TODO: SaveExpenseCode(DataTable) は SqlDataAdapter + SqlCommandBuilder 等で実装。
        ' 更新時の update_ymd / update_login スタンプは画面側 or ここで付与（設計書D-1参照）。
#End Region

    End Class
End Namespace
