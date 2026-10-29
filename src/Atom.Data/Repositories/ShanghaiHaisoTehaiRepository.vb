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
    ''' 配送手配書（menuNo=103 §10）のデータアクセスを担います。
    ''' </summary>
    ''' <remarks>
    ''' 値はすべて SqlParameter で渡し、型と長さを明示します。
    ''' 接続とトランザクションは呼び出し側が持ちます。
    ''' </remarks>
    Public Class ShanghaiHaisoTehaiRepository

#Region "定数"
        ''' <summary>ログインコードの長さ。</summary>
        Private Const LoginCodeLength As Integer = 6

        ''' <summary>INVOICE 番号（メイン）の長さ。</summary>
        Private Const InvoiceNoMainLength As Integer = 20

        ''' <summary>印字する INVOICE 番号の長さ。</summary>
        Private Const InvoiceNoLength As Integer = 50

        ''' <summary>船名の長さ。</summary>
        Private Const VesselLength As Integer = 50

        ''' <summary>コンテナ区分名の長さ。</summary>
        Private Const ContainerTypeLength As Integer = 30

        ''' <summary>コンテナ番号の長さ。</summary>
        Private Const ContainerNoLength As Integer = 20

        ''' <summary>シール番号の長さ。</summary>
        Private Const SealNoLength As Integer = 20
#End Region

#Region "印刷対象の取得"
        ''' <summary>
        ''' 印刷対象を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <returns>拾った行の一覧。</returns>
        Public Function SelectPickups(connection As SqlConnection, transaction As SqlTransaction,
                                      invoiceNoMain As String) As IList(Of ShanghaiHaisoTehaiPickup)

            Dim results As New List(Of ShanghaiHaisoTehaiPickup)()

            Using command As New SqlCommand(ShanghaiHaisoTehaiSql.CreateSelectPickupSql(), connection, transaction)
                command.Parameters.Add("@InvoiceNoMain", SqlDbType.NVarChar, InvoiceNoMainLength).Value = invoiceNoMain

                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        results.Add(Me.ReadPickup(reader))
                    End While
                End Using
            End Using

            Return results
        End Function
#End Region

#Region "印刷用の一時データ"
        ''' <summary>
        ''' 印刷用の一時データを削除します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        Public Sub DeletePrintTemp(connection As SqlConnection, transaction As SqlTransaction, loginCode As String)
            Using command As New SqlCommand(ShanghaiHaisoTehaiSql.CreateDeletePrintTempSql(), connection, transaction)
                command.Parameters.Add("@LoginCode", SqlDbType.NVarChar, LoginCodeLength).Value = loginCode
                command.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>
        ''' 印刷用の一時データを 1 行登録します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="row">登録する行。</param>
        Public Sub InsertPrintTemp(connection As SqlConnection, transaction As SqlTransaction,
                                   loginCode As String, row As ShanghaiHaisoTehaiPrintRow)

            Using command As New SqlCommand(ShanghaiHaisoTehaiSql.CreateInsertPrintTempSql(), connection, transaction)
                command.Parameters.Add("@LoginCode", SqlDbType.NVarChar, LoginCodeLength).Value = loginCode
                command.Parameters.Add("@InvoiceNoMain", SqlDbType.NVarChar, InvoiceNoMainLength).Value =
                    ToDbValue(row.InvoiceNoMain)
                command.Parameters.Add("@InvoiceNo", SqlDbType.NVarChar, InvoiceNoLength).Value =
                    ToDbValue(row.InvoiceNo)
                command.Parameters.Add("@Vessel", SqlDbType.NVarChar, VesselLength).Value =
                    ToDbValue(row.Vessel)
                command.Parameters.Add("@BlDate", SqlDbType.DateTime).Value = ToDbValue(row.BlDate)
                command.Parameters.Add("@HaisoType", SqlDbType.Int).Value = ToDbValue(row.HaisoType)
                command.Parameters.Add("@ContainerType", SqlDbType.NVarChar, ContainerTypeLength).Value =
                    ToDbValue(row.ContainerType)
                command.Parameters.Add("@ContainerNo", SqlDbType.NVarChar, ContainerNoLength).Value =
                    ToDbValue(row.ContainerNo)
                command.Parameters.Add("@SealNo", SqlDbType.NVarChar, SealNoLength).Value =
                    ToDbValue(row.SealNo)
                command.Parameters.Add("@TotalCarton", SqlDbType.Int).Value = row.TotalCarton
                command.Parameters.Add("@TotalQuantity", SqlDbType.Int).Value = row.TotalQuantity
                command.Parameters.Add("@HaisoYoteiYmd", SqlDbType.DateTime).Value = ToDbValue(row.HaisoYoteiYmd)
                command.Parameters.Add("@HaisoSaki1", SqlDbType.Int).Value = ToDbValue(row.HaisoSaki1)
                command.Parameters.Add("@HaisoSaki2", SqlDbType.Int).Value = ToDbValue(row.HaisoSaki2)
                command.Parameters.Add("@HaisoSaki3", SqlDbType.Int).Value = ToDbValue(row.HaisoSaki3)
                command.Parameters.Add("@OtsunakaNo", SqlDbType.Int).Value = ToDbValue(row.OtsunakaNo)

                command.ExecuteNonQuery()
            End Using
        End Sub
#End Region

#Region "帳票の内容"
        ''' <summary>
        ''' 帳票に印字する内容を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>帳票の行の一覧。</returns>
        Public Function SelectReportRows(connection As SqlConnection, transaction As SqlTransaction,
                                         loginCode As String) As IList(Of ShanghaiHaisoTehaiReportRow)

            Dim results As New List(Of ShanghaiHaisoTehaiReportRow)()

            Using command As New SqlCommand(ShanghaiHaisoTehaiSql.CreateSelectReportRowsSql(), connection, transaction)
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
        ''' 印刷対象の 1 行を読み取ります。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <returns>読み取った行。</returns>
        Private Function ReadPickup(reader As SqlDataReader) As ShanghaiHaisoTehaiPickup
            Dim pickup As New ShanghaiHaisoTehaiPickup()
            pickup.InvoiceNoMain = ReadString(reader, "pk_invoice_no_main")
            pickup.InvoiceNo = ReadString(reader, "pk_invoice_no")
            pickup.BlDate = ReadNullableDate(reader, "bl_date")
            pickup.Vessel = ReadString(reader, "vessel")
            pickup.OtsunakaNo = ReadNullableInteger(reader, "otsunaka_no")
            pickup.HaisoYoteiYmd = ReadNullableDate(reader, "haiso_yotei_ymd")
            pickup.HaisoType = ReadNullableInteger(reader, "haiso_type")
            pickup.HaisoTypeText = ReadString(reader, "haiso_type_text")
            pickup.ContainerType = ReadNullableInteger(reader, "container_type")
            pickup.ContainerTypeText = ReadString(reader, "container_type_text")
            pickup.ContainerNo = ReadString(reader, "container_no")
            pickup.SealNo = ReadString(reader, "seal_no")
            pickup.CartonQty = ReadNullableInteger(reader, "carton_qty")
            pickup.TotalSuryo = ReadNullableInteger(reader, "total_suryo")
            pickup.HaisoSaki1 = ReadNullableInteger(reader, "haiso_saki1")
            pickup.HaisoSaki2 = ReadNullableInteger(reader, "haiso_saki2")
            pickup.HaisoSaki3 = ReadNullableInteger(reader, "haiso_saki3")
            Return pickup
        End Function

        ''' <summary>
        ''' 帳票の 1 行を読み取ります。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <returns>読み取った行。</returns>
        Private Function ReadReportRow(reader As SqlDataReader) As ShanghaiHaisoTehaiReportRow
            Dim row As New ShanghaiHaisoTehaiReportRow()
            row.RefNo = ReadString(reader, "ref_no")
            row.InvoiceNo = ReadString(reader, "invoice_no")
            row.Vessel = ReadString(reader, "vessel")
            row.BlDate = ReadNullableDate(reader, "bl_date")
            row.HaisoType = ReadNullableInteger(reader, "haiso_type")
            row.ContainerType = ReadString(reader, "container_type")
            row.ContainerNo = ReadString(reader, "container_no")
            row.SealNo = ReadString(reader, "seal_no")
            row.TotalCarton = ReadNullableInteger(reader, "total_carton")
            row.HaisoYoteiYmd = ReadNullableDate(reader, "haisyo_yotei_ymd")
            row.OtsunakaNoText = ReadString(reader, "otsunaka_no_text")
            row.OtsunakaTanto1 = ReadString(reader, "otsunaka_tanto1")
            row.OtsunakaTanto2 = ReadString(reader, "otsunaka_tanto2")
            row.HaisoSaki1No = ReadNullableInteger(reader, "haiso_saki1")
            row.HaisoSaki2No = ReadNullableInteger(reader, "haiso_saki2")
            row.HaisoSaki3No = ReadNullableInteger(reader, "haiso_saki3")
            row.HaisoSaki1 = Me.ReadHaisoSaki(reader, "haiso_saki1_")
            row.HaisoSaki2 = Me.ReadHaisoSaki(reader, "haiso_saki2_")
            row.HaisoSaki3 = Me.ReadHaisoSaki(reader, "haiso_saki3_")
            row.DisplayName = ReadString(reader, "display_name")
            row.BumonCodeText = ReadString(reader, "bumon_code_text")
            row.PhoneNo = ReadString(reader, "phone_no")
            row.FaxNo = ReadString(reader, "fax_no")
            Return row
        End Function

        ''' <summary>
        ''' 配送先 1 箇所分を読み取ります。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="prefix">列名の接頭辞。</param>
        ''' <returns>読み取った配送先。</returns>
        ''' <remarks>
        ''' ビューの列名は `haiso_saki1_tanto1` ですが、SVF へ渡す項目名は `haiso_saki1_tanto` です。
        ''' 配送先の担当は 1 人分しか無いため、末尾の 1 を落としています。
        ''' </remarks>
        Private Function ReadHaisoSaki(reader As SqlDataReader, prefix As String) As ShanghaiHaisoSaki
            Dim saki As New ShanghaiHaisoSaki()
            saki.Name = ReadString(reader, prefix & "name")
            saki.Address = ReadString(reader, prefix & "address")
            saki.PhoneNo = ReadString(reader, prefix & "phone_no")
            saki.Tanto = ReadString(reader, prefix & "tanto1")
            Return saki
        End Function

        ''' <summary>
        ''' 文字列の列を読み取ります。NULL のときは空文字を返します。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>読み取った値。</returns>
        Private Shared Function ReadString(reader As SqlDataReader, columnName As String) As String
            Dim value As Object = reader(columnName)
            ' NULL のとき
            If value Is Nothing OrElse Convert.IsDBNull(value) Then
                Return String.Empty
            End If
            Return Convert.ToString(value, Globalization.CultureInfo.InvariantCulture)
        End Function

        ''' <summary>
        ''' 数値の列を読み取ります。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>読み取った値。NULL のときは Nothing。</returns>
        Private Shared Function ReadNullableInteger(reader As SqlDataReader, columnName As String) As Integer?
            Dim value As Object = reader(columnName)
            ' NULL のとき
            If value Is Nothing OrElse Convert.IsDBNull(value) Then
                Return Nothing
            End If
            Return Convert.ToInt32(value, Globalization.CultureInfo.InvariantCulture)
        End Function

        ''' <summary>
        ''' 日付の列を読み取ります。
        ''' </summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>読み取った値。NULL のときは Nothing。</returns>
        Private Shared Function ReadNullableDate(reader As SqlDataReader, columnName As String) As Date?
            Dim value As Object = reader(columnName)
            ' NULL のとき
            If value Is Nothing OrElse Convert.IsDBNull(value) Then
                Return Nothing
            End If
            Return Convert.ToDateTime(value, Globalization.CultureInfo.InvariantCulture)
        End Function

        ''' <summary>
        ''' 文字列をデータベースへ渡せる形にします。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>空のときは DBNull、それ以外は元の値。</returns>
        Private Shared Function ToDbValue(value As String) As Object
            ' 値が無いとき
            If String.IsNullOrEmpty(value) Then
                Return DBNull.Value
            End If
            Return value
        End Function

        ''' <summary>
        ''' 数値をデータベースへ渡せる形にします。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>値が無いときは DBNull、それ以外は元の値。</returns>
        Private Shared Function ToDbValue(value As Integer?) As Object
            ' 値が無いとき
            If Not value.HasValue Then
                Return DBNull.Value
            End If
            Return value.Value
        End Function

        ''' <summary>
        ''' 日付をデータベースへ渡せる形にします。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>値が無いときは DBNull、それ以外は元の値。</returns>
        Private Shared Function ToDbValue(value As Date?) As Object
            ' 値が無いとき
            If Not value.HasValue Then
                Return DBNull.Value
            End If
            Return value.Value
        End Function
#End Region

    End Class
End Namespace
