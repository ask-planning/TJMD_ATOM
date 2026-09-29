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
    ''' 上海輸入 INVOICE / PACKING 取込（menuNo=101）のデータアクセスを担います。
    ''' </summary>
    ''' <remarks>
    ''' 接続とトランザクションは呼び出し側が用意します。
    ''' 一連の手順は Atom.Data.Services.ShanghaiImportService が制御します。
    ''' </remarks>
    Public Class ShanghaiImportRepository

#Region "定数"
        ''' <summary>ログインコードの桁数。</summary>
        Private Const LoginCodeLength As Integer = 6

        ''' <summary>一時テーブルの一般的な文字列項目の桁数。</summary>
        Private Const TempTextLength As Integer = 50

        ''' <summary>一時テーブルの名称項目の桁数。</summary>
        Private Const TempNameLength As Integer = 128

        ''' <summary>商品群の桁数。</summary>
        Private Const PartSortLength As Integer = 2

        ''' <summary>数量単位の桁数。</summary>
        Private Const SuryoTaniLength As Integer = 4

        ''' <summary>通貨の桁数。</summary>
        Private Const TsukaLength As Integer = 3

        ''' <summary>PACKING の真偽区分の桁数。</summary>
        Private Const TrueOrFalseLength As Integer = 10

        ''' <summary>INVOICE 番号（メイン）の桁数。</summary>
        Private Const InvoiceNoMainLength As Integer = 20

        ''' <summary>仕入先コードの桁数。</summary>
        Private Const ShiiresakiCodeLength As Integer = 10

        ''' <summary>項目コードの桁数。</summary>
        Private Const KoumokuCodeLength As Integer = 3

        ''' <summary>数量の精度。</summary>
        Private Const SuryoPrecision As Byte = 18

        ''' <summary>数量の位取り。整数のみを保持します。</summary>
        Private Const SuryoScale As Byte = 0
#End Region

#Region "参照系"
        ''' <summary>
        ''' 一時データの件数を取得します。再実行時に取込済みデータの有無を判定します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>一時データの件数。</returns>
        Public Function SelectCheckTempCount(connection As SqlConnection, transaction As SqlTransaction,
                                             loginCode As String) As Integer
            Using command As New SqlCommand(ShanghaiImportSql.CreateSelectCheckTempCountSql(), connection, transaction)
                Me.AddLoginCode(command, loginCode)
                Return Me.ToInteger(command.ExecuteScalar())
            End Using
        End Function

        ''' <summary>
        ''' 一時データに含まれる INVOICE 番号の一覧を取得します。REF_NO の編集に使います。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>INVOICE 番号の一覧。昇順。</returns>
        Public Function SelectInvoiceNoList(connection As SqlConnection, transaction As SqlTransaction,
                                            loginCode As String) As IList(Of String)
            Return Me.SelectStringList(connection, transaction,
                                       ShanghaiImportSql.CreateSelectInvoiceNoListSql(), loginCode)
        End Function

        ''' <summary>
        ''' 今回の取込対象のうち、INVOICE 明細として既に登録済みの INVOICE 番号を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>登録済みの INVOICE 番号の一覧。</returns>
        Public Function SelectRegisteredInvoiceNoList(connection As SqlConnection, transaction As SqlTransaction,
                                                      loginCode As String) As IList(Of String)
            Return Me.SelectStringList(connection, transaction,
                                       ShanghaiImportSql.CreateSelectRegisteredInvoiceNoSql(), loginCode)
        End Function

        ''' <summary>
        ''' 今回の取込対象のうち、PACKING 明細として既に登録済みの INVOICE 番号を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>登録済みの INVOICE 番号の一覧。</returns>
        Public Function SelectRegisteredPackingInvoiceNoList(connection As SqlConnection, transaction As SqlTransaction,
                                                             loginCode As String) As IList(Of String)
            Return Me.SelectStringList(connection, transaction,
                                       ShanghaiImportSql.CreateSelectRegisteredPackingInvoiceNoSql(), loginCode)
        End Function

        ''' <summary>
        ''' 取込結果確認画面で表示する INVOICE の一覧を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="invoiceNo">絞り込む INVOICE 番号。空のときは全件。</param>
        ''' <returns>取込結果の一覧。</returns>
        Public Function SelectInvoiceResults(connection As SqlConnection, transaction As SqlTransaction,
                                             loginCode As String, invoiceNo As String) As IList(Of ShanghaiInvoiceImportRow)
            Dim results As New List(Of ShanghaiInvoiceImportRow)()

            Using command As New SqlCommand(ShanghaiImportSql.CreateSelectInvoiceResultSql(), connection, transaction)
                Me.AddLoginCode(command, loginCode)
                Me.AddInvoiceNoFilter(command, invoiceNo)

                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        Dim entity As New ShanghaiInvoiceImportRow()
                        entity.SeqNo = Me.ReadInteger(reader, "pk_seq_no")
                        entity.ErrContents = Me.ReadText(reader, "err_contents")
                        entity.ErrStatus = Me.ReadInteger(reader, "err_status")
                        entity.Suryo = Me.ReadNullableDecimal(reader, "suryo")
                        entity.SuryoTani = Me.ReadText(reader, "suryo_tani")
                        entity.Tsuka = Me.ReadText(reader, "tsuka")
                        entity.UnitPrice = Me.ReadNullableDouble(reader, "unit_price")
                        entity.Amount = Me.ReadNullableDouble(reader, "amount")
                        entity.R3KoubaiDenpyoNo = Me.ReadText(reader, "R3_koubai_denpyo_no")
                        entity.InvoiceNo = Me.ReadText(reader, "invoice_no")
                        entity.BlDate = Me.ReadNullableDate(reader, "bl_date")
                        entity.SyukkaHoho = Me.ReadText(reader, "syukka_hoho")
                        entity.ShanghaiCodeText = Me.ReadText(reader, "shanghai_code_text")
                        entity.ShanghaiCode = Me.ReadText(reader, "shanghai_code")
                        entity.TajimaPoNo = Me.ReadText(reader, "tajima_po_no")
                        results.Add(entity)
                    End While
                End Using
            End Using

            Return results
        End Function

        ''' <summary>
        ''' 取込結果確認画面で表示する PACKING の一覧を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="invoiceNo">絞り込む INVOICE 番号。空のときは全件。</param>
        ''' <returns>取込結果の一覧。</returns>
        Public Function SelectPackingResults(connection As SqlConnection, transaction As SqlTransaction,
                                             loginCode As String, invoiceNo As String) As IList(Of ShanghaiPackingImportRow)
            Dim results As New List(Of ShanghaiPackingImportRow)()

            Using command As New SqlCommand(ShanghaiImportSql.CreateSelectPackingResultSql(), connection, transaction)
                Me.AddLoginCode(command, loginCode)
                Me.AddInvoiceNoFilter(command, invoiceNo)

                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        Dim entity As New ShanghaiPackingImportRow()
                        entity.SeqNo = Me.ReadInteger(reader, "pk_seq_no")
                        entity.ErrContents = Me.ReadText(reader, "err_contents")
                        entity.ErrStatus = Me.ReadInteger(reader, "err_status")
                        entity.InvoiceNo = Me.ReadText(reader, "invoice_no")
                        entity.TajimaPoNo = Me.ReadText(reader, "tajima_po_no")
                        entity.TrueOrFalse = Me.ReadText(reader, "true_or_false")
                        entity.ShanghaiCode = Me.ReadText(reader, "shanghai_code")
                        entity.SyukkaHoho = Me.ReadText(reader, "syukka_hoho")
                        entity.BlDate = Me.ReadNullableDate(reader, "bl_date")
                        entity.Suryo = Me.ReadNullableDecimal(reader, "suryo")
                        entity.NetWeight = Me.ReadNullableDouble(reader, "net_weight")
                        entity.GrossWeight = Me.ReadNullableDouble(reader, "gross_weight")
                        entity.M3 = Me.ReadNullableDouble(reader, "m3")
                        entity.CartonNoFrom = Me.ReadText(reader, "carton_no_from")
                        entity.CartonNoTo = Me.ReadText(reader, "carton_no_to")
                        entity.UnitPrice = Me.ReadNullableDouble(reader, "unit_price")
                        entity.Amount = Me.ReadNullableDouble(reader, "amount")
                        results.Add(entity)
                    End While
                End Using
            End Using

            Return results
        End Function

        ''' <summary>
        ''' 指定日以降の BL DATE を持つ一時データの件数を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="blDate">判定の基準日。この日以降を数えます。</param>
        ''' <returns>該当する件数。</returns>
        Public Function SelectFutureBlDateCount(connection As SqlConnection, transaction As SqlTransaction,
                                                loginCode As String, blDate As Date) As Integer
            Using command As New SqlCommand("dbo.usp_shanghai_import_count_get_next_month", connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                Me.AddLoginCode(command, loginCode)
                command.Parameters.Add("@bl_date", SqlDbType.DateTime).Value = blDate
                Dim count As SqlParameter = Me.AddOutputInteger(command, "@next_count")
                command.ExecuteNonQuery()
                Return Me.ToInteger(count.Value)
            End Using
        End Function

        ''' <summary>
        ''' 今回の取込対象が既に登録済みかどうかの件数を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>INVOICE と PACKING それぞれの登録済み件数。</returns>
        Public Function SelectRegisteredCounts(connection As SqlConnection, transaction As SqlTransaction,
                                               loginCode As String) As ShanghaiImportCounts
            Using command As New SqlCommand("dbo.usp_shanghai_import_tourokuzumi_count", connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                Me.AddLoginCode(command, loginCode)
                Dim invoiceCount As SqlParameter = Me.AddOutputInteger(command, "@invoice_count")
                Dim packingCount As SqlParameter = Me.AddOutputInteger(command, "@packing_count")
                command.ExecuteNonQuery()

                Dim result As New ShanghaiImportCounts()
                result.InvoiceCount = Me.ToInteger(invoiceCount.Value)
                result.PackingCount = Me.ToInteger(packingCount.Value)
                Return result
            End Using
        End Function

        ''' <summary>
        ''' INVOICE 明細から集計したヘッダ情報を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <returns>集計したヘッダ情報。</returns>
        Public Function SelectInvoiceTotal(connection As SqlConnection, transaction As SqlTransaction,
                                           invoiceNoMain As String) As ShanghaiInvoiceTotal
            Using command As New SqlCommand("dbo.usp_shanghai_invoice_gokei_get", connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@invoice_no", SqlDbType.NVarChar, InvoiceNoMainLength).Value = invoiceNoMain

                Dim blDate As SqlParameter = Me.AddOutputParameter(command, "@bl_date", SqlDbType.DateTime, 0)
                Dim syukkaHoho As SqlParameter = Me.AddOutputInteger(command, "@syukka_hoho")
                Dim totalCarton As SqlParameter = Me.AddOutputInteger(command, "@total_carton")
                Dim totalAmount As SqlParameter = Me.AddOutputParameter(command, "@total_amount", SqlDbType.Float, 0)
                Dim totalSuryo As SqlParameter = Me.AddOutputParameter(command, "@total_suryo", SqlDbType.Float, 0)
                Dim containerText As SqlParameter = Me.AddOutputParameter(command, "@container_text", SqlDbType.NVarChar, TempTextLength)
                Dim headerCount As SqlParameter = Me.AddOutputInteger(command, "@header_count")
                command.ExecuteNonQuery()

                Dim result As New ShanghaiInvoiceTotal()
                result.BlDate = Me.ToNullableDate(blDate.Value)
                result.SyukkaHoho = Me.ToNullableInteger(syukkaHoho.Value)
                result.TotalCarton = Me.ToNullableInteger(totalCarton.Value)
                result.TotalAmount = Me.ToNullableDouble(totalAmount.Value)
                result.TotalSuryo = Me.ToNullableDouble(totalSuryo.Value)
                result.ContainerText = Me.ToText(containerText.Value)
                result.HeaderCount = Me.ToInteger(headerCount.Value)
                Return result
            End Using
        End Function

        ''' <summary>
        ''' 項目マスタの一覧を取得します。発注先の選択肢に使います。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="koumokuCode">項目コード。発注先は 'COS'。</param>
        ''' <returns>連番順に並んだ項目マスタの一覧。</returns>
        Public Function SelectKoumokuList(connection As SqlConnection, transaction As SqlTransaction,
                                          koumokuCode As String) As IList(Of Koumoku)
            Dim results As New List(Of Koumoku)()

            Using command As New SqlCommand(ShanghaiImportSql.CreateSelectKoumokuListSql(), connection, transaction)
                command.Parameters.Add("@KoumokuCode", SqlDbType.NVarChar, KoumokuCodeLength).Value = koumokuCode
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        Dim entity As New Koumoku()
                        entity.KoumokuCode = koumokuCode
                        entity.SeqNo = reader.GetInt32(reader.GetOrdinal("pk_seq_no"))
                        entity.Contents1 = Me.ReadText(reader, "contents1")
                        entity.Contents2 = Me.ReadText(reader, "contents2")
                        entity.Contents3 = Me.ReadText(reader, "contents3")
                        entity.Contents4 = Me.ReadText(reader, "contents4")
                        results.Add(entity)
                    End While
                End Using
            End Using

            Return results
        End Function

        ''' <summary>
        ''' 発注先区分から仕入先コードを取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="customerKbn">得意先マスタの customer_kbn_num。</param>
        ''' <returns>仕入先コード。取得できないときは空文字。</returns>
        Public Function SelectShiiresakiCode(connection As SqlConnection, transaction As SqlTransaction,
                                             customerKbn As Integer) As String
            Using command As New SqlCommand("dbo.usp_POEM_tokuisaki_master_get", connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                ' 第 1 引数 2 は「customer_kbn から引く」を意味します。
                command.Parameters.Add("@syori_flg", SqlDbType.Int).Value = 2
                command.Parameters.Add("@tokui_code", SqlDbType.NVarChar, LoginCodeLength).Value = "NON"
                command.Parameters.Add("@customer_kbn", SqlDbType.Int).Value = customerKbn

                Me.AddOutputParameter(command, "@return_tokui_code", SqlDbType.NVarChar, 6)
                Me.AddOutputParameter(command, "@return_tokui_name", SqlDbType.NVarChar, TempNameLength)
                Me.AddOutputParameter(command, "@return_tokui_name_short", SqlDbType.NVarChar, 20)
                Me.AddOutputParameter(command, "@return_customer_kbn", SqlDbType.NVarChar, 2)
                Me.AddOutputInteger(command, "@return_customer_kbn_num")
                Me.AddOutputParameter(command, "@return_tsuka", SqlDbType.NVarChar, TsukaLength)
                Me.AddOutputParameter(command, "@return_destination", SqlDbType.NVarChar, 20)
                Me.AddOutputInteger(command, "@return_rate_check")
                Me.AddOutputParameter(command, "@return_country", SqlDbType.NVarChar, 20)
                Me.AddOutputParameter(command, "@return_invoice_kigo", SqlDbType.NVarChar, 5)
                Dim shiiresakiCode As SqlParameter = Me.AddOutputParameter(command, "@return_shiiresaki_code",
                                                                           SqlDbType.NVarChar, ShiiresakiCodeLength)
                command.ExecuteNonQuery()
                Return Me.ToText(shiiresakiCode.Value)
            End Using
        End Function
#End Region

#Region "更新系"
        ''' <summary>
        ''' チェック用一時データを利用者単位で削除します。INVOICE と PACKING の両方を消します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        Public Sub DeleteCheckTemp(connection As SqlConnection, transaction As SqlTransaction, loginCode As String)
            Using command As New SqlCommand(ShanghaiImportSql.CreateDeleteInvoiceCheckTempSql(), connection, transaction)
                Me.AddLoginCode(command, loginCode)
                command.ExecuteNonQuery()
            End Using

            Using command As New SqlCommand(ShanghaiImportSql.CreateDeletePackingCheckTempSql(), connection, transaction)
                Me.AddLoginCode(command, loginCode)
                command.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>
        ''' INVOICE のチェック用一時データを登録します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="rows">登録する取込行。</param>
        Public Sub InsertInvoiceCheckTemp(connection As SqlConnection, transaction As SqlTransaction,
                                          loginCode As String, rows As IList(Of ShanghaiInvoiceImportRow))
            ' 登録する行が無いとき
            If rows Is Nothing OrElse rows.Count = 0 Then
                Return
            End If

            Using command As New SqlCommand(ShanghaiImportSql.CreateInsertInvoiceCheckTempSql(), connection, transaction)
                Me.AddLoginCode(command, loginCode)
                command.Parameters.Add("@LineSeq", SqlDbType.Float)
                command.Parameters.Add("@SoNo", SqlDbType.NVarChar, TempTextLength)
                command.Parameters.Add("@SoNoSeq", SqlDbType.Int)
                command.Parameters.Add("@ShanghaiCode", SqlDbType.NVarChar, TempTextLength)
                command.Parameters.Add("@ShanghaiCodeText", SqlDbType.NVarChar, TempNameLength)
                command.Parameters.Add("@Kikaku", SqlDbType.NVarChar, TempNameLength)
                command.Parameters.Add("@PartSort", SqlDbType.NVarChar, PartSortLength)
                command.Parameters.Add("@SuryoTani", SqlDbType.NVarChar, SuryoTaniLength)
                Me.AddSuryoParameter(command, "@Suryo")
                command.Parameters.Add("@UnitPrice", SqlDbType.Float)
                command.Parameters.Add("@Amount", SqlDbType.Float)
                command.Parameters.Add("@DummyNo1", SqlDbType.NVarChar, TempTextLength)
                command.Parameters.Add("@DummyNo2", SqlDbType.NVarChar, TempTextLength)
                command.Parameters.Add("@TajimaPoNo", SqlDbType.NVarChar, TempTextLength)
                command.Parameters.Add("@R3HinmokuCode", SqlDbType.NVarChar, TempTextLength)
                command.Parameters.Add("@Tsuka", SqlDbType.NVarChar, TsukaLength)
                command.Parameters.Add("@TajimaPoNoOriginal", SqlDbType.NVarChar, TempTextLength)
                command.Parameters.Add("@InvoiceNo", SqlDbType.NVarChar, TempTextLength)
                command.Parameters.Add("@AmountText", SqlDbType.NVarChar, TempTextLength)
                command.Parameters.Add("@R3KoubaiDenpyoNo", SqlDbType.NVarChar, TempTextLength)
                command.Parameters.Add("@BlDate", SqlDbType.DateTime)
                command.Parameters.Add("@SyukkaHoho", SqlDbType.NVarChar, TempTextLength)

                For Each row As ShanghaiInvoiceImportRow In rows
                    Me.SetValue(command, "@LineSeq", row.LineSeq)
                    Me.SetText(command, "@SoNo", row.SoNo, TempTextLength)
                    Me.SetValue(command, "@SoNoSeq", row.SoNoSeq)
                    Me.SetText(command, "@ShanghaiCode", row.ShanghaiCode, TempTextLength)
                    Me.SetText(command, "@ShanghaiCodeText", row.ShanghaiCodeText, TempNameLength)
                    Me.SetText(command, "@Kikaku", row.Kikaku, TempNameLength)
                    Me.SetText(command, "@PartSort", row.PartSort, PartSortLength)
                    Me.SetText(command, "@SuryoTani", row.SuryoTani, SuryoTaniLength)
                    Me.SetValue(command, "@Suryo", row.Suryo)
                    Me.SetValue(command, "@UnitPrice", row.UnitPrice)
                    Me.SetValue(command, "@Amount", row.Amount)
                    Me.SetText(command, "@DummyNo1", row.DummyNo1, TempTextLength)
                    Me.SetText(command, "@DummyNo2", row.DummyNo2, TempTextLength)
                    Me.SetText(command, "@TajimaPoNo", row.TajimaPoNo, TempTextLength)
                    Me.SetText(command, "@R3HinmokuCode", row.R3HinmokuCode, TempTextLength)
                    Me.SetText(command, "@Tsuka", row.Tsuka, TsukaLength)
                    Me.SetText(command, "@TajimaPoNoOriginal", row.TajimaPoNoOriginal, TempTextLength)
                    Me.SetText(command, "@InvoiceNo", row.InvoiceNo, TempTextLength)
                    Me.SetText(command, "@AmountText", row.AmountText, TempTextLength)
                    Me.SetText(command, "@R3KoubaiDenpyoNo", row.R3KoubaiDenpyoNo, TempTextLength)
                    Me.SetValue(command, "@BlDate", row.BlDate)
                    Me.SetText(command, "@SyukkaHoho", row.SyukkaHoho, TempTextLength)
                    command.ExecuteNonQuery()
                Next
            End Using
        End Sub

        ''' <summary>
        ''' PACKING のチェック用一時データを登録します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="rows">登録する取込行。</param>
        Public Sub InsertPackingCheckTemp(connection As SqlConnection, transaction As SqlTransaction,
                                          loginCode As String, rows As IList(Of ShanghaiPackingImportRow))
            ' 登録する行が無いとき
            If rows Is Nothing OrElse rows.Count = 0 Then
                Return
            End If

            Using command As New SqlCommand(ShanghaiImportSql.CreateInsertPackingCheckTempSql(), connection, transaction)
                Me.AddLoginCode(command, loginCode)
                command.Parameters.Add("@TrueOrFalse", SqlDbType.NVarChar, TrueOrFalseLength)
                command.Parameters.Add("@InvoiceNo", SqlDbType.NVarChar, TempTextLength)
                command.Parameters.Add("@TajimaPoNo", SqlDbType.NVarChar, TempTextLength)
                command.Parameters.Add("@ShanghaiCode", SqlDbType.NVarChar, TempTextLength)
                Me.AddSuryoParameter(command, "@Suryo")
                command.Parameters.Add("@UnitPrice", SqlDbType.Float)
                command.Parameters.Add("@Amount", SqlDbType.Float)
                command.Parameters.Add("@CartonNoFrom", SqlDbType.NVarChar, TempTextLength)
                command.Parameters.Add("@CartonNoTo", SqlDbType.NVarChar, TempTextLength)
                command.Parameters.Add("@NetWeight", SqlDbType.Float)
                command.Parameters.Add("@GrossWeight", SqlDbType.Float)
                command.Parameters.Add("@M3", SqlDbType.Float)
                command.Parameters.Add("@BlDate", SqlDbType.DateTime)
                command.Parameters.Add("@SyukkaHoho", SqlDbType.NVarChar, TempTextLength)
                command.Parameters.Add("@R3HinmokuCode", SqlDbType.NVarChar, TempTextLength)
                command.Parameters.Add("@TajimaPoNoOriginal", SqlDbType.NVarChar, TempTextLength)

                For Each row As ShanghaiPackingImportRow In rows
                    Me.SetText(command, "@TrueOrFalse", row.TrueOrFalse, TrueOrFalseLength)
                    Me.SetText(command, "@InvoiceNo", row.InvoiceNo, TempTextLength)
                    Me.SetText(command, "@TajimaPoNo", row.TajimaPoNo, TempTextLength)
                    Me.SetText(command, "@ShanghaiCode", row.ShanghaiCode, TempTextLength)
                    Me.SetValue(command, "@Suryo", row.Suryo)
                    Me.SetValue(command, "@UnitPrice", row.UnitPrice)
                    Me.SetValue(command, "@Amount", row.Amount)
                    Me.SetText(command, "@CartonNoFrom", row.CartonNoFrom, TempTextLength)
                    Me.SetText(command, "@CartonNoTo", row.CartonNoTo, TempTextLength)
                    Me.SetValue(command, "@NetWeight", row.NetWeight)
                    Me.SetValue(command, "@GrossWeight", row.GrossWeight)
                    Me.SetValue(command, "@M3", row.M3)
                    Me.SetValue(command, "@BlDate", row.BlDate)
                    Me.SetText(command, "@SyukkaHoho", row.SyukkaHoho, TempTextLength)
                    Me.SetText(command, "@R3HinmokuCode", row.R3HinmokuCode, TempTextLength)
                    Me.SetText(command, "@TajimaPoNoOriginal", row.TajimaPoNoOriginal, TempTextLength)
                    command.ExecuteNonQuery()
                Next
            End Using
        End Sub

        ''' <summary>
        ''' INVOICE のチェック用一時データを 1 項目だけ更新します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="seqNo">更新する行の連番。</param>
        ''' <param name="field">更新する項目の定義。</param>
        ''' <param name="value">設定する値。Nothing のときは NULL を設定します。</param>
        Public Sub UpdateInvoiceCheckTempValue(connection As SqlConnection, transaction As SqlTransaction,
                                               loginCode As String, seqNo As Integer,
                                               field As ShanghaiImportField, value As Object)
            Dim commandText As String = ShanghaiImportSql.CreateUpdateInvoiceCheckTempValueSql(
                Me.GetVerifiedColumnName(field, True))
            Me.ExecuteValueUpdate(connection, transaction, commandText, loginCode, seqNo, field, value)
        End Sub

        ''' <summary>
        ''' PACKING のチェック用一時データを 1 項目だけ更新します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="seqNo">更新する行の連番。</param>
        ''' <param name="field">更新する項目の定義。</param>
        ''' <param name="value">設定する値。Nothing のときは NULL を設定します。</param>
        Public Sub UpdatePackingCheckTempValue(connection As SqlConnection, transaction As SqlTransaction,
                                               loginCode As String, seqNo As Integer,
                                               field As ShanghaiImportField, value As Object)
            Dim commandText As String = ShanghaiImportSql.CreateUpdatePackingCheckTempValueSql(
                Me.GetVerifiedColumnName(field, False))
            Me.ExecuteValueUpdate(connection, transaction, commandText, loginCode, seqNo, field, value)
        End Sub

        ''' <summary>
        ''' INVOICE のチェック用一時データの品目コードを更新します。R3 品目コードも併せて更新します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="seqNo">更新する行の連番。</param>
        ''' <param name="shanghaiCode">設定する品目コード。空のときは NULL を設定します。</param>
        Public Sub UpdateInvoiceShanghaiCode(connection As SqlConnection, transaction As SqlTransaction,
                                             loginCode As String, seqNo As Integer, shanghaiCode As String)
            Me.ExecuteShanghaiCodeUpdate(connection, transaction,
                                         ShanghaiImportSql.CreateUpdateInvoiceShanghaiCodeSql(),
                                         loginCode, seqNo, shanghaiCode)
        End Sub

        ''' <summary>
        ''' PACKING のチェック用一時データの品目コードを更新します。R3 品目コードも併せて更新します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="seqNo">更新する行の連番。</param>
        ''' <param name="shanghaiCode">設定する品目コード。空のときは NULL を設定します。</param>
        Public Sub UpdatePackingShanghaiCode(connection As SqlConnection, transaction As SqlTransaction,
                                             loginCode As String, seqNo As Integer, shanghaiCode As String)
            Me.ExecuteShanghaiCodeUpdate(connection, transaction,
                                         ShanghaiImportSql.CreateUpdatePackingShanghaiCodeSql(),
                                         loginCode, seqNo, shanghaiCode)
        End Sub

        ''' <summary>
        ''' 一時データのエラー情報を初期化します。R3 品目コードや PO 番号の補完もここで行われます。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        Public Sub UpdateErrorClear(connection As SqlConnection, transaction As SqlTransaction, loginCode As String)
            Me.ExecuteLoginCodeProcedure(connection, transaction, "dbo.usp_shanghai_import_check_err_clear", loginCode)
        End Sub

        ''' <summary>
        ''' 1 つの PO が複数 INVOICE に分かれていないかを判定します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>該当する件数。0 以外のとき取込を中止します。</returns>
        Public Function UpdateDuplicateCheck(connection As SqlConnection, transaction As SqlTransaction,
                                             loginCode As String) As Integer
            Using command As New SqlCommand("dbo.usp_shanghai_import_jyufuku_check", connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                Me.AddLoginCode(command, loginCode)
                Dim errCount As SqlParameter = Me.AddOutputInteger(command, "@err_count")
                command.ExecuteNonQuery()
                Return Me.ToInteger(errCount.Value)
            End Using
        End Function

        ''' <summary>
        ''' INVOICE のデータチェックを行います。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>エラーと警告の件数。</returns>
        Public Function UpdateInvoiceCheck(connection As SqlConnection, transaction As SqlTransaction,
                                           loginCode As String) As ImportCheckResult
            Return Me.ExecuteCheckProcedure(connection, transaction,
                                            "dbo.usp_shanghai_import_err_check_invoice", loginCode)
        End Function

        ''' <summary>
        ''' PACKING のデータチェックを行います。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>エラーと警告の件数。</returns>
        Public Function UpdatePackingCheck(connection As SqlConnection, transaction As SqlTransaction,
                                           loginCode As String) As ImportCheckResult
            Return Me.ExecuteCheckProcedure(connection, transaction,
                                            "dbo.usp_shanghai_import_err_check_packing", loginCode)
        End Function

        ''' <summary>
        ''' 金額のチェックを行います。INVOICE と PACKING の数量比較も含みます。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>エラーと警告の件数。</returns>
        Public Function UpdateAmountCheck(connection As SqlConnection, transaction As SqlTransaction,
                                          loginCode As String) As ImportCheckResult
            Return Me.ExecuteCheckProcedure(connection, transaction,
                                            "dbo.usp_shanghai_import_check_amount", loginCode)
        End Function

        ''' <summary>
        ''' チェック済みの一時データを INVOICE / PACKING の明細テーブルへ登録します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="shanghaiChotatsu">担当区分。1=上海 / 2=調達 / 3=貿易。</param>
        ''' <returns>登録された INVOICE と PACKING の件数。</returns>
        Public Function InsertMeisai(connection As SqlConnection, transaction As SqlTransaction,
                                     loginCode As String, shanghaiChotatsu As Integer) As ShanghaiImportCounts
            Using command As New SqlCommand("dbo.usp_shanghai_import_invoice_update", connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                Me.AddLoginCode(command, loginCode)
                command.Parameters.Add("@shanghai_chotatsu", SqlDbType.Int).Value = shanghaiChotatsu
                Dim invoiceCount As SqlParameter = Me.AddOutputInteger(command, "@invoice_count")
                Dim packingCount As SqlParameter = Me.AddOutputInteger(command, "@packing_count")
                command.ExecuteNonQuery()

                Dim result As New ShanghaiImportCounts()
                result.InvoiceCount = Me.ToInteger(invoiceCount.Value)
                result.PackingCount = Me.ToInteger(packingCount.Value)
                Return result
            End Using
        End Function

        ''' <summary>
        ''' INVOICE ヘッダを削除します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        Public Sub DeleteInvoiceHeader(connection As SqlConnection, transaction As SqlTransaction,
                                       invoiceNoMain As String)
            Using command As New SqlCommand(ShanghaiImportSql.CreateDeleteInvoiceHeaderSql(), connection, transaction)
                command.Parameters.Add("@InvoiceNoMain", SqlDbType.NVarChar, InvoiceNoMainLength).Value = invoiceNoMain
                command.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>
        ''' INVOICE ヘッダサブを削除します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="invoiceNo">INVOICE 番号。</param>
        Public Sub DeleteInvoiceHeaderSub(connection As SqlConnection, transaction As SqlTransaction,
                                          invoiceNo As String)
            Me.ExecuteInvoiceNoCommand(connection, transaction,
                                       ShanghaiImportSql.CreateDeleteInvoiceHeaderSubSql(), invoiceNo)
        End Sub

        ''' <summary>
        ''' INVOICE 明細を削除します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="invoiceNo">INVOICE 番号。</param>
        Public Sub DeleteInvoiceMeisai(connection As SqlConnection, transaction As SqlTransaction, invoiceNo As String)
            Me.ExecuteInvoiceNoCommand(connection, transaction,
                                       ShanghaiImportSql.CreateDeleteInvoiceMeisaiSql(), invoiceNo)
        End Sub

        ''' <summary>
        ''' PACKING 明細を削除します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="invoiceNo">INVOICE 番号。</param>
        Public Sub DeletePackingMeisai(connection As SqlConnection, transaction As SqlTransaction, invoiceNo As String)
            Me.ExecuteInvoiceNoCommand(connection, transaction,
                                       ShanghaiImportSql.CreateDeletePackingMeisaiSql(), invoiceNo)
        End Sub

        ''' <summary>
        ''' INVOICE ヘッダを登録します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <param name="total">明細から集計したヘッダ情報。</param>
        ''' <param name="shiiresakiCode">仕入先コード。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        Public Sub InsertInvoiceHeader(connection As SqlConnection, transaction As SqlTransaction,
                                       invoiceNoMain As String, total As ShanghaiInvoiceTotal,
                                       shiiresakiCode As String, loginCode As String)
            Using command As New SqlCommand(ShanghaiImportSql.CreateInsertInvoiceHeaderSql(), connection, transaction)
                command.Parameters.Add("@InvoiceNoMain", SqlDbType.NVarChar, InvoiceNoMainLength).Value = invoiceNoMain
                command.Parameters.Add("@BlDate", SqlDbType.DateTime)
                command.Parameters.Add("@SyukkaHoho", SqlDbType.Int)
                command.Parameters.Add("@CartonQty", SqlDbType.Int)
                command.Parameters.Add("@ShiiresakiCode", SqlDbType.NVarChar, ShiiresakiCodeLength)
                command.Parameters.Add("@UpdateLogin", SqlDbType.NVarChar, LoginCodeLength).Value = loginCode

                Me.SetValue(command, "@BlDate", total.BlDate)
                ' 出荷方法とカートン数は取得できないとき 0 とします（現行仕様）。
                Me.SetValue(command, "@SyukkaHoho", New Integer?(Me.ToZeroIfNothing(total.SyukkaHoho)))
                Me.SetValue(command, "@CartonQty", New Integer?(Me.ToZeroIfNothing(total.TotalCarton)))
                Me.SetText(command, "@ShiiresakiCode", shiiresakiCode, ShiiresakiCodeLength)
                command.ExecuteNonQuery()
            End Using
        End Sub
#End Region

#Region "内部処理"
        ''' <summary>ログインコードのパラメータを追加します。</summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="loginCode">ログインコード。</param>
        Private Sub AddLoginCode(command As SqlCommand, loginCode As String)
            Dim parameterName As String = "@LoginCode"
            ' ストアドプロシージャのときは定義済みの引数名に合わせる
            If command.CommandType = CommandType.StoredProcedure Then
                parameterName = "@login_code"
            End If
            command.Parameters.Add(parameterName, SqlDbType.NVarChar, LoginCodeLength).Value = loginCode
        End Sub

        ''' <summary>
        ''' 更新して差し支えない列名かを確かめ、その列名を返します。
        ''' </summary>
        ''' <param name="field">更新する項目の定義。</param>
        ''' <param name="isInvoice">INVOICE 側のとき True、PACKING 側のとき False。</param>
        ''' <returns>ホワイトリストに載っている列名。</returns>
        ''' <remarks>
        ''' 列名は SQL の構文として連結するため、定義一覧に載っているものだけを通します。
        ''' </remarks>
        Private Function GetVerifiedColumnName(field As ShanghaiImportField, isInvoice As Boolean) As String
            ' 定義が渡されていないとき
            If field Is Nothing Then
                Throw New ArgumentNullException(NameOf(field))
            End If

            Dim verified As ShanghaiImportField
            ' INVOICE のとき
            If isInvoice Then
                verified = ShanghaiImportFieldCatalog.FindInvoiceField(field.Key)
            ' PACKING のとき
            Else
                verified = ShanghaiImportFieldCatalog.FindPackingField(field.Key)
            End If

            ' 定義一覧に無い、または列名が一致しないとき
            If verified Is Nothing OrElse
               Not String.Equals(verified.ColumnName, field.ColumnName, StringComparison.Ordinal) Then
                Throw New ArgumentException("更新できない項目が指定されました。", NameOf(field))
            End If
            Return verified.ColumnName
        End Function

        ''' <summary>
        ''' 1 項目だけを更新するコマンドを実行します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="commandText">実行する SQL 文。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="seqNo">更新する行の連番。</param>
        ''' <param name="field">更新する項目の定義。</param>
        ''' <param name="value">設定する値。</param>
        Private Sub ExecuteValueUpdate(connection As SqlConnection, transaction As SqlTransaction,
                                       commandText As String, loginCode As String, seqNo As Integer,
                                       field As ShanghaiImportField, value As Object)
            Using command As New SqlCommand(commandText, connection, transaction)
                Me.AddLoginCode(command, loginCode)
                command.Parameters.Add("@SeqNo", SqlDbType.Int).Value = seqNo
                Me.AddValueParameter(command, field, value)
                command.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>
        ''' 品目コードと R3 品目コードを同時に更新するコマンドを実行します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="commandText">実行する SQL 文。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="seqNo">更新する行の連番。</param>
        ''' <param name="shanghaiCode">設定する品目コード。</param>
        ''' <remarks>
        ''' ゼロ埋めの規則は取込時と同じ（Atom.Core.Services.ShanghaiCodeFormatter）です。
        ''' 名前空間は Atom.Data.Services と紛れるため、ここでは完全な名前で書いています。
        ''' </remarks>
        Private Sub ExecuteShanghaiCodeUpdate(connection As SqlConnection, transaction As SqlTransaction,
                                              commandText As String, loginCode As String, seqNo As Integer,
                                              shanghaiCode As String)
            Dim trimmed As String = String.Empty
            ' 値があるとき
            If shanghaiCode IsNot Nothing Then
                trimmed = shanghaiCode.Trim()
            End If

            Dim formattedCode As String = Atom.Core.Services.ShanghaiCodeFormatter.FormatShanghaiCode(trimmed)
            Dim hinmokuCode As String = Atom.Core.Services.ShanghaiCodeFormatter.FormatR3HinmokuCode(trimmed)

            Using command As New SqlCommand(commandText, connection, transaction)
                Me.AddLoginCode(command, loginCode)
                command.Parameters.Add("@SeqNo", SqlDbType.Int).Value = seqNo
                command.Parameters.Add("@ShanghaiCode", SqlDbType.NVarChar, TempTextLength)
                command.Parameters.Add("@R3HinmokuCode", SqlDbType.NVarChar, TempTextLength)
                Me.SetText(command, "@ShanghaiCode", formattedCode, TempTextLength)
                Me.SetText(command, "@R3HinmokuCode", hinmokuCode, TempTextLength)
                command.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>
        ''' 更新する値のパラメータを、項目の種類に合わせて追加します。
        ''' </summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="field">更新する項目の定義。</param>
        ''' <param name="value">設定する値。</param>
        Private Sub AddValueParameter(command As SqlCommand, field As ShanghaiImportField, value As Object)
            Dim parameter As SqlParameter

            Select Case field.Kind
                ' 整数のとき
                Case ShanghaiImportValueKind.WholeNumber
                    parameter = command.Parameters.Add("@Value", SqlDbType.Decimal)
                    parameter.Precision = SuryoPrecision
                    parameter.Scale = SuryoScale
                ' 小数を含む数値のとき
                Case ShanghaiImportValueKind.DecimalNumber
                    parameter = command.Parameters.Add("@Value", SqlDbType.Float)
                ' 日付のとき
                Case ShanghaiImportValueKind.DateValue
                    parameter = command.Parameters.Add("@Value", SqlDbType.DateTime)
                ' 文字列のとき
                Case Else
                    parameter = command.Parameters.Add("@Value", SqlDbType.NVarChar, field.MaxLength)
            End Select

            ' 値が無いとき
            If value Is Nothing Then
                parameter.Value = DBNull.Value
                Return
            End If
            parameter.Value = value
        End Sub

        ''' <summary>
        ''' INVOICE 番号での絞り込みパラメータを追加します。空のときは NULL を渡して全件対象にします。
        ''' </summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="invoiceNo">絞り込む INVOICE 番号。</param>
        Private Sub AddInvoiceNoFilter(command As SqlCommand, invoiceNo As String)
            Dim parameter As SqlParameter = command.Parameters.Add("@InvoiceNo", SqlDbType.NVarChar, TempTextLength)
            ' 絞り込みを指定しないとき
            If String.IsNullOrEmpty(invoiceNo) Then
                parameter.Value = DBNull.Value
                Return
            End If
            parameter.Value = invoiceNo
        End Sub

        ''' <summary>読み取り結果の整数を取得します。NULL のときは 0 を返します。</summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>列の値。</returns>
        Private Function ReadInteger(reader As SqlDataReader, columnName As String) As Integer
            Dim index As Integer = reader.GetOrdinal(columnName)
            ' NULL のとき
            If reader.IsDBNull(index) Then
                Return 0
            End If
            Return Convert.ToInt32(reader.GetValue(index))
        End Function

        ''' <summary>読み取り結果の 10 進数を取得します。NULL のときは Nothing を返します。</summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>列の値。</returns>
        Private Function ReadNullableDecimal(reader As SqlDataReader, columnName As String) As Decimal?
            Dim index As Integer = reader.GetOrdinal(columnName)
            ' NULL のとき
            If reader.IsDBNull(index) Then
                Return Nothing
            End If
            Return Convert.ToDecimal(reader.GetValue(index))
        End Function

        ''' <summary>読み取り結果の数値を取得します。NULL のときは Nothing を返します。</summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>列の値。</returns>
        Private Function ReadNullableDouble(reader As SqlDataReader, columnName As String) As Double?
            Dim index As Integer = reader.GetOrdinal(columnName)
            ' NULL のとき
            If reader.IsDBNull(index) Then
                Return Nothing
            End If
            Return Convert.ToDouble(reader.GetValue(index))
        End Function

        ''' <summary>読み取り結果の日付を取得します。NULL のときは Nothing を返します。</summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>列の値。</returns>
        Private Function ReadNullableDate(reader As SqlDataReader, columnName As String) As Date?
            Dim index As Integer = reader.GetOrdinal(columnName)
            ' NULL のとき
            If reader.IsDBNull(index) Then
                Return Nothing
            End If
            Return reader.GetDateTime(index)
        End Function

        ''' <summary>数量のパラメータを追加します。精度と位取りを明示します。</summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="parameterName">パラメータ名。</param>
        Private Sub AddSuryoParameter(command As SqlCommand, parameterName As String)
            Dim parameter As SqlParameter = command.Parameters.Add(parameterName, SqlDbType.Decimal)
            parameter.Precision = SuryoPrecision
            parameter.Scale = SuryoScale
        End Sub

        ''' <summary>出力パラメータを追加します。</summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="parameterName">パラメータ名。</param>
        ''' <param name="type">パラメータの型。</param>
        ''' <param name="size">文字列の桁数。桁数が不要な型では 0 を指定します。</param>
        ''' <returns>追加したパラメータ。</returns>
        Private Function AddOutputParameter(command As SqlCommand, parameterName As String,
                                            type As SqlDbType, size As Integer) As SqlParameter
            Dim parameter As SqlParameter
            ' 桁数の指定があるとき
            If size > 0 Then
                parameter = command.Parameters.Add(parameterName, type, size)
            Else
                parameter = command.Parameters.Add(parameterName, type)
            End If
            parameter.Direction = ParameterDirection.Output
            Return parameter
        End Function

        ''' <summary>整数の出力パラメータを追加します。</summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="parameterName">パラメータ名。</param>
        ''' <returns>追加したパラメータ。</returns>
        Private Function AddOutputInteger(command As SqlCommand, parameterName As String) As SqlParameter
            Return Me.AddOutputParameter(command, parameterName, SqlDbType.Int, 0)
        End Function

        ''' <summary>ログインコードだけを渡すストアドプロシージャを実行します。</summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="procedureName">ストアドプロシージャ名。</param>
        ''' <param name="loginCode">ログインコード。</param>
        Private Sub ExecuteLoginCodeProcedure(connection As SqlConnection, transaction As SqlTransaction,
                                              procedureName As String, loginCode As String)
            Using command As New SqlCommand(procedureName, connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                Me.AddLoginCode(command, loginCode)
                command.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>エラー件数と警告件数を返すチェック用ストアドプロシージャを実行します。</summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="procedureName">ストアドプロシージャ名。</param>
        ''' <param name="loginCode">ログインコード。</param>
        ''' <returns>エラーと警告の件数。</returns>
        Private Function ExecuteCheckProcedure(connection As SqlConnection, transaction As SqlTransaction,
                                               procedureName As String, loginCode As String) As ImportCheckResult
            Using command As New SqlCommand(procedureName, connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                Me.AddLoginCode(command, loginCode)
                Dim errorCount As SqlParameter = Me.AddOutputInteger(command, "@err_count1")
                Dim warningCount As SqlParameter = Me.AddOutputInteger(command, "@err_count2")
                command.ExecuteNonQuery()

                Dim result As New ImportCheckResult()
                result.ErrorCount = Me.ToInteger(errorCount.Value)
                result.WarningCount = Me.ToInteger(warningCount.Value)
                Return result
            End Using
        End Function

        ''' <summary>INVOICE 番号だけを条件にするコマンドを実行します。</summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="commandText">実行する SQL 文。</param>
        ''' <param name="invoiceNo">INVOICE 番号。</param>
        Private Sub ExecuteInvoiceNoCommand(connection As SqlConnection, transaction As SqlTransaction,
                                            commandText As String, invoiceNo As String)
            Using command As New SqlCommand(commandText, connection, transaction)
                command.Parameters.Add("@InvoiceNo", SqlDbType.NVarChar, TempTextLength).Value = invoiceNo
                command.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>ログインコードを条件に文字列の一覧を取得します。</summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="commandText">実行する SQL 文。</param>
        ''' <param name="loginCode">ログインコード。</param>
        ''' <returns>取得した文字列の一覧。</returns>
        Private Function SelectStringList(connection As SqlConnection, transaction As SqlTransaction,
                                          commandText As String, loginCode As String) As IList(Of String)
            Dim results As New List(Of String)()

            Using command As New SqlCommand(commandText, connection, transaction)
                Me.AddLoginCode(command, loginCode)
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        ' NULL でないとき
                        If Not reader.IsDBNull(0) Then
                            results.Add(reader.GetString(0))
                        End If
                    End While
                End Using
            End Using

            Return results
        End Function

        ''' <summary>パラメータへ文字列を設定します。空のときは NULL にし、桁数を超える分は切り詰めます。</summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="parameterName">パラメータ名。</param>
        ''' <param name="value">設定する値。</param>
        ''' <param name="maxLength">列の桁数。</param>
        Private Sub SetText(command As SqlCommand, parameterName As String, value As String, maxLength As Integer)
            ' 値が無いとき
            If String.IsNullOrEmpty(value) Then
                command.Parameters(parameterName).Value = DBNull.Value
                Return
            End If
            ' 桁数を超えているとき（登録時のエラーを避けるため切り詰める）
            If value.Length > maxLength Then
                command.Parameters(parameterName).Value = value.Substring(0, maxLength)
                Return
            End If
            command.Parameters(parameterName).Value = value
        End Sub

        ''' <summary>パラメータへ数値を設定します。値が無いときは NULL にします。</summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="parameterName">パラメータ名。</param>
        ''' <param name="value">設定する値。</param>
        Private Sub SetValue(command As SqlCommand, parameterName As String, value As Double?)
            Me.SetObjectValue(command, parameterName, If(value.HasValue, CType(value.Value, Object), Nothing))
        End Sub

        ''' <summary>パラメータへ 10 進数を設定します。値が無いときは NULL にします。</summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="parameterName">パラメータ名。</param>
        ''' <param name="value">設定する値。</param>
        Private Sub SetValue(command As SqlCommand, parameterName As String, value As Decimal?)
            Me.SetObjectValue(command, parameterName, If(value.HasValue, CType(value.Value, Object), Nothing))
        End Sub

        ''' <summary>パラメータへ整数を設定します。値が無いときは NULL にします。</summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="parameterName">パラメータ名。</param>
        ''' <param name="value">設定する値。</param>
        Private Sub SetValue(command As SqlCommand, parameterName As String, value As Integer?)
            Me.SetObjectValue(command, parameterName, If(value.HasValue, CType(value.Value, Object), Nothing))
        End Sub

        ''' <summary>パラメータへ日付を設定します。値が無いときは NULL にします。</summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="parameterName">パラメータ名。</param>
        ''' <param name="value">設定する値。</param>
        Private Sub SetValue(command As SqlCommand, parameterName As String, value As Date?)
            Me.SetObjectValue(command, parameterName, If(value.HasValue, CType(value.Value, Object), Nothing))
        End Sub

        ''' <summary>パラメータへ値を設定します。Nothing のときは NULL にします。</summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="parameterName">パラメータ名。</param>
        ''' <param name="value">設定する値。</param>
        Private Sub SetObjectValue(command As SqlCommand, parameterName As String, value As Object)
            ' 値が無いとき
            If value Is Nothing Then
                command.Parameters(parameterName).Value = DBNull.Value
                Return
            End If
            command.Parameters(parameterName).Value = value
        End Sub

        ''' <summary>読み取り結果の文字列を取得します。NULL のときは空文字を返します。</summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>列の値。</returns>
        Private Function ReadText(reader As SqlDataReader, columnName As String) As String
            Dim index As Integer = reader.GetOrdinal(columnName)
            ' NULL のとき
            If reader.IsDBNull(index) Then
                Return String.Empty
            End If
            Return reader.GetString(index)
        End Function

        ''' <summary>取得した値を整数に変換します。NULL のときは 0 を返します。</summary>
        ''' <param name="value">変換する値。</param>
        ''' <returns>整数の値。</returns>
        Private Function ToInteger(value As Object) As Integer
            Dim result As Integer? = Me.ToNullableInteger(value)
            ' 値が無いとき
            If Not result.HasValue Then
                Return 0
            End If
            Return result.Value
        End Function

        ''' <summary>取得した値を整数に変換します。NULL のときは Nothing を返します。</summary>
        ''' <param name="value">変換する値。</param>
        ''' <returns>整数の値。</returns>
        Private Function ToNullableInteger(value As Object) As Integer?
            ' NULL のとき
            If value Is Nothing OrElse Convert.IsDBNull(value) Then
                Return Nothing
            End If
            Return Convert.ToInt32(value)
        End Function

        ''' <summary>取得した値を数値に変換します。NULL のときは Nothing を返します。</summary>
        ''' <param name="value">変換する値。</param>
        ''' <returns>数値の値。</returns>
        Private Function ToNullableDouble(value As Object) As Double?
            ' NULL のとき
            If value Is Nothing OrElse Convert.IsDBNull(value) Then
                Return Nothing
            End If
            Return Convert.ToDouble(value)
        End Function

        ''' <summary>取得した値を日付に変換します。NULL のときは Nothing を返します。</summary>
        ''' <param name="value">変換する値。</param>
        ''' <returns>日付の値。</returns>
        Private Function ToNullableDate(value As Object) As Date?
            ' NULL のとき
            If value Is Nothing OrElse Convert.IsDBNull(value) Then
                Return Nothing
            End If
            Return Convert.ToDateTime(value)
        End Function

        ''' <summary>取得した値を文字列に変換します。NULL のときは空文字を返します。</summary>
        ''' <param name="value">変換する値。</param>
        ''' <returns>文字列の値。</returns>
        Private Function ToText(value As Object) As String
            ' NULL のとき
            If value Is Nothing OrElse Convert.IsDBNull(value) Then
                Return String.Empty
            End If
            Return Convert.ToString(value)
        End Function

        ''' <summary>値が無いときに 0 を返します。</summary>
        ''' <param name="value">判定する値。</param>
        ''' <returns>値、または 0。</returns>
        Private Function ToZeroIfNothing(value As Integer?) As Integer
            ' 値が無いとき
            If Not value.HasValue Then
                Return 0
            End If
            Return value.Value
        End Function
#End Region

    End Class
End Namespace
