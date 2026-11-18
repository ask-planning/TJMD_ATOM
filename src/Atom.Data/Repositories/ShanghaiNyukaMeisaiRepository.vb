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
    ''' 入荷明細書（menuNo=103 §11）のデータアクセスを担います。
    ''' </summary>
    ''' <remarks>
    ''' 値はすべて SqlParameter で渡し、型と長さを明示します。
    ''' 接続とトランザクションは呼び出し側が持ちます。
    ''' </remarks>
    Public Class ShanghaiNyukaMeisaiRepository

#Region "定数"
        ''' <summary>ログインコードの長さ。</summary>
        Private Const LoginCodeLength As Integer = 6

        ''' <summary>作成済み件数を数えるストアド。</summary>
        Private Const KensuProcedure As String = "dbo.usp_shanghai_nyuka_meisai_kensu"

        ''' <summary>入荷明細を作り直すストアド。</summary>
        Private Const CreateProcedure As String = "dbo.usp_shanghai_nyuka_meisai_create"

        ''' <summary>印刷用一時テーブルへ写すストアド。</summary>
        Private Const TempCreateProcedure As String = "dbo.usp_shanghai_nyuka_meisai_temp_create"

        ''' <summary>件数を数えるストアドの INVOICE 番号の長さ。</summary>
        Private Const KensuInvoiceNoLength As Integer = 50

        ''' <summary>
        ''' 作成するストアドの INVOICE 番号の長さ。
        ''' </summary>
        ''' <remarks>**このストアドだけ 128 文字**です。ほかの 2 つは 50 文字で、長さが揃っていません。</remarks>
        Private Const CreateInvoiceNoLength As Integer = 128

        ''' <summary>社内参照番号の長さ。</summary>
        Private Const SyanaiSansyoNoLength As Integer = 20

        ''' <summary>納入先プラントの長さ。</summary>
        Private Const NounyuPlantLength As Integer = 10
#End Region

#Region "ストアドの呼び出し"
        ''' <summary>
        ''' 作成済みの入荷明細の件数と最終更新日を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="shanghaiChotatsu">担当区分。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <param name="syanaiSansyoNo">社内参照番号の絞り込み。</param>
        ''' <param name="nounyuPlant">納入先プラントの絞り込み。</param>
        ''' <returns>件数と最終更新日。</returns>
        ''' <remarks>
        ''' 設計書 §11.2。このストアドだけ `@login_code` を取りません。
        ''' </remarks>
        Public Function SelectKensu(connection As SqlConnection, transaction As SqlTransaction,
                                    shanghaiChotatsu As Integer, invoiceNoMain As String,
                                    syanaiSansyoNo As String, nounyuPlant As String) As ShanghaiNyukaMeisaiKensu

            Using command As New SqlCommand(KensuProcedure, connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@shanghai_chotatsu", SqlDbType.Int).Value = shanghaiChotatsu
                command.Parameters.Add("@invoice_no", SqlDbType.NVarChar, KensuInvoiceNoLength).Value = invoiceNoMain
                command.Parameters.Add("@syanai_sansyo_no", SqlDbType.NVarChar, SyanaiSansyoNoLength).Value = syanaiSansyoNo
                command.Parameters.Add("@nounyu_plant", SqlDbType.NVarChar, NounyuPlantLength).Value = nounyuPlant

                Dim koushinYmd As SqlParameter = command.Parameters.Add("@koushin_ymd", SqlDbType.DateTime)
                koushinYmd.Direction = ParameterDirection.Output
                Dim meisaiCount As SqlParameter = command.Parameters.Add("@meisai_count", SqlDbType.Int)
                meisaiCount.Direction = ParameterDirection.Output

                command.ExecuteNonQuery()

                Dim result As New ShanghaiNyukaMeisaiKensu()
                result.Count = ToInteger(meisaiCount.Value)
                result.UpdatedAt = ToNullableDate(koushinYmd.Value)
                Return result
            End Using
        End Function

        ''' <summary>
        ''' 入荷明細を作り直します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="shanghaiChotatsu">担当区分。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <param name="syanaiSansyoNo">社内参照番号の絞り込み。</param>
        ''' <param name="nounyuPlant">納入先プラントの絞り込み。</param>
        ''' <returns>作成した件数。</returns>
        ''' <remarks>
        ''' **このストアドは `tbl_t_shanghai_nyuka_meisai` を削除してから作り直します。**
        ''' 設計書 §11.3 の 1。選択が「1:作り直す」のときだけ呼びます。
        ''' </remarks>
        Public Function ExecuteCreate(connection As SqlConnection, transaction As SqlTransaction,
                                      loginCode As String, shanghaiChotatsu As Integer,
                                      invoiceNoMain As String,
                                      syanaiSansyoNo As String, nounyuPlant As String) As Integer

            Using command As New SqlCommand(CreateProcedure, connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@login_code", SqlDbType.NVarChar, LoginCodeLength).Value = loginCode
                command.Parameters.Add("@shanghai_chotatsu", SqlDbType.Int).Value = shanghaiChotatsu
                command.Parameters.Add("@invoice_no", SqlDbType.NVarChar, CreateInvoiceNoLength).Value = invoiceNoMain
                command.Parameters.Add("@syanai_sansyo_no", SqlDbType.NVarChar, SyanaiSansyoNoLength).Value = syanaiSansyoNo
                command.Parameters.Add("@nounyu_plant", SqlDbType.NVarChar, NounyuPlantLength).Value = nounyuPlant

                Dim meisaiCount As SqlParameter = command.Parameters.Add("@meisai_count", SqlDbType.Int)
                meisaiCount.Direction = ParameterDirection.Output

                command.ExecuteNonQuery()
                Return ToInteger(meisaiCount.Value)
            End Using
        End Function

        ''' <summary>
        ''' 印刷用の一時テーブルへ写します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="shanghaiChotatsu">担当区分。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <param name="syanaiSansyoNo">社内参照番号の絞り込み。</param>
        ''' <param name="nounyuPlant">納入先プラントの絞り込み。</param>
        ''' <returns>写した件数。</returns>
        ''' <remarks>
        ''' 設計書 §11.3 の 2。自分のログインコードの一時データを削除してから写します。
        ''' </remarks>
        Public Function ExecuteTempCreate(connection As SqlConnection, transaction As SqlTransaction,
                                          loginCode As String, shanghaiChotatsu As Integer,
                                          invoiceNoMain As String,
                                          syanaiSansyoNo As String, nounyuPlant As String) As Integer

            Using command As New SqlCommand(TempCreateProcedure, connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@login_code", SqlDbType.NVarChar, LoginCodeLength).Value = loginCode
                command.Parameters.Add("@shanghai_chotatsu", SqlDbType.Int).Value = shanghaiChotatsu
                command.Parameters.Add("@invoice_no", SqlDbType.NVarChar, KensuInvoiceNoLength).Value = invoiceNoMain
                command.Parameters.Add("@syanai_sansyo_no", SqlDbType.NVarChar, SyanaiSansyoNoLength).Value = syanaiSansyoNo
                command.Parameters.Add("@nounyu_plant", SqlDbType.NVarChar, NounyuPlantLength).Value = nounyuPlant

                Dim outputCount As SqlParameter = command.Parameters.Add("@output_count", SqlDbType.Int)
                outputCount.Direction = ParameterDirection.Output

                command.ExecuteNonQuery()
                Return ToInteger(outputCount.Value)
            End Using
        End Function
#End Region

#Region "帳票の内容"
        ''' <summary>
        ''' 帳票に印字する内容を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>印字する行の一覧。並び順は SQL 側で決めています。</returns>
        Public Function SelectReportRows(connection As SqlConnection, transaction As SqlTransaction,
                                         loginCode As String) As IList(Of ShanghaiNyukaMeisaiReportRow)

            Dim results As New List(Of ShanghaiNyukaMeisaiReportRow)()

            Using command As New SqlCommand(ShanghaiNyukaMeisaiSql.CreateSelectReportRowsSql(), connection, transaction)
                command.Parameters.Add("@LoginCode", SqlDbType.NVarChar, LoginCodeLength).Value = loginCode

                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        results.Add(Me.ReadReportRow(reader))
                    End While
                End Using
            End Using

            Return results
        End Function
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 1 行を読み取ります。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <returns>読み取った行。</returns>
        Private Function ReadReportRow(reader As SqlDataReader) As ShanghaiNyukaMeisaiReportRow
            Dim row As New ShanghaiNyukaMeisaiReportRow()

            row.SortSeq = Me.GetNullableInteger(reader, "sort_seq")
            row.InvoiceNoMain = Me.GetString(reader, "pk_invoice_no_main")
            row.InvoiceNo = Me.GetString(reader, "pk_invoice_no")
            row.BlDate = Me.GetNullableDate(reader, "bl_date")
            row.ContainerTypeText = Me.GetString(reader, "container_type_text")
            row.ContainerNo = Me.GetString(reader, "container_no")
            row.SealNo = Me.GetString(reader, "seal_no")
            row.HaisoYoteiYmd = Me.GetNullableDate(reader, "haiso_yotei_ymd")
            row.InvoiceNoPrint = Me.GetString(reader, "invoice_no_print")
            row.Remarks = Me.GetString(reader, "remarks")
            row.NounyuPlant = Me.GetString(reader, "nounyu_plant")
            row.HokanBasyoText = Me.GetString(reader, "hokan_basyo_text")
            row.ZaikoTensoNo = Me.GetString(reader, "zaiko_tenso_no")
            row.SyukkaDenpyoNo = Me.GetString(reader, "syukka_denpyo_no")
            row.NukitoriKensaFlg = Me.GetNullableInteger(reader, "nukitori_kensa_flg")
            row.R3HinmokuCode = Me.GetString(reader, "R3_hinmoku_code")
            row.CartonNoFrom = Me.GetNullableInteger(reader, "carton_no_from")
            row.CartonNoTo = Me.GetNullableInteger(reader, "carton_no_to")
            row.CartonQty = Me.GetNullableInteger(reader, "carton_qty")
            row.NounyuSuryo = Me.GetNullableDecimal(reader, "nounyu_suryo")
            row.TotalCarton = Me.GetNullableInteger(reader, "total_carton")

            Return row
        End Function

        ''' <summary>
        ''' 文字列の列を読み取ります。NULL は空文字にします。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>読み取った値。</returns>
        Private Function GetString(reader As SqlDataReader, columnName As String) As String
            Dim index As Integer = reader.GetOrdinal(columnName)
            ' NULL のとき
            If reader.IsDBNull(index) Then
                Return String.Empty
            End If
            Return Convert.ToString(reader.GetValue(index), Globalization.CultureInfo.InvariantCulture)
        End Function

        ''' <summary>
        ''' 数値の列を読み取ります。NULL は Nothing にします。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>読み取った値。</returns>
        Private Function GetNullableInteger(reader As SqlDataReader, columnName As String) As Integer?
            Dim index As Integer = reader.GetOrdinal(columnName)
            ' NULL のとき
            If reader.IsDBNull(index) Then
                Return Nothing
            End If
            Return Convert.ToInt32(reader.GetValue(index), Globalization.CultureInfo.InvariantCulture)
        End Function

        ''' <summary>
        ''' 小数を含む数値の列を読み取ります。NULL は Nothing にします。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>読み取った値。</returns>
        Private Function GetNullableDecimal(reader As SqlDataReader, columnName As String) As Decimal?
            Dim index As Integer = reader.GetOrdinal(columnName)
            ' NULL のとき
            If reader.IsDBNull(index) Then
                Return Nothing
            End If
            Return Convert.ToDecimal(reader.GetValue(index), Globalization.CultureInfo.InvariantCulture)
        End Function

        ''' <summary>
        ''' 日付の列を読み取ります。NULL は Nothing にします。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>読み取った値。</returns>
        Private Function GetNullableDate(reader As SqlDataReader, columnName As String) As Date?
            Dim index As Integer = reader.GetOrdinal(columnName)
            ' NULL のとき
            If reader.IsDBNull(index) Then
                Return Nothing
            End If
            Return Convert.ToDateTime(reader.GetValue(index), Globalization.CultureInfo.InvariantCulture)
        End Function

        ''' <summary>
        ''' 出力引数の値を整数にします。
        ''' </summary>
        ''' <param name="value">出力引数の値。</param>
        ''' <returns>整数。値が無いときは 0。</returns>
        ''' <remarks>
        ''' **件数のストアドは、該当が無いと出力引数へ何も入れません。**
        ''' その場合 `DBNull` ではなく `Nothing` が返るため、両方を 0 として扱います。
        ''' </remarks>
        Private Shared Function ToInteger(value As Object) As Integer
            ' 値が無いとき
            If value Is Nothing OrElse Convert.IsDBNull(value) Then
                Return 0
            End If
            Return Convert.ToInt32(value, Globalization.CultureInfo.InvariantCulture)
        End Function

        ''' <summary>
        ''' 出力引数の値を日付にします。
        ''' </summary>
        ''' <param name="value">出力引数の値。</param>
        ''' <returns>日付。値が無いときは Nothing。</returns>
        Private Shared Function ToNullableDate(value As Object) As Date?
            ' 値が無いとき
            If value Is Nothing OrElse Convert.IsDBNull(value) Then
                Return Nothing
            End If
            Return Convert.ToDateTime(value, Globalization.CultureInfo.InvariantCulture)
        End Function
#End Region

    End Class
End Namespace
