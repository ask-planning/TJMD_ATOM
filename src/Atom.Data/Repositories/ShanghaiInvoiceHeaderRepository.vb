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
    ''' INVOICEヘッダー入力（menuNo=103）のデータアクセスを担います。
    ''' </summary>
    ''' <remarks>
    ''' 接続とトランザクションは呼び出し側が用意します。
    ''' 一連の手順は Atom.Data.Services.ShanghaiInvoiceHeaderService が制御します。
    ''' </remarks>
    Public Class ShanghaiInvoiceHeaderRepository

#Region "定数"
        ''' <summary>ログインコードの桁数。</summary>
        Private Const LoginCodeLength As Integer = 6

        ''' <summary>INVOICE 番号（メイン）の桁数。</summary>
        Private Const InvoiceNoMainLength As Integer = 20

        ''' <summary>INVOICE 番号（枝番付き）の桁数。</summary>
        Private Const InvoiceNoLength As Integer = 50

        ''' <summary>コンテナ番号・シール番号の桁数。</summary>
        Private Const ContainerTextLength As Integer = 50

        ''' <summary>コンテナタイプの表示名の桁数。</summary>
        Private Const ContainerTypeTextLength As Integer = 50

        ''' <summary>項目コードの桁数。</summary>
        Private Const KoumokuCodeLength As Integer = 3

        ''' <summary>船名の一覧を保持する項目コード。</summary>
        Private Const VesselKoumokuCode As String = "SHP"

        ''' <summary>REF_NO の桁数。</summary>
        Private Const RefNoLength As Integer = 50

        ''' <summary>船名の桁数。</summary>
        Private Const VesselLength As Integer = 50

        ''' <summary>印刷用 INVOICE_NO の組み立て対象から外す配送タイプ（日新入荷）。</summary>
        Private Const NisshinHaisoType As Integer = 3

        ''' <summary>ヘッダサブの入力情報から明細を更新することを表す値。</summary>
        Private Const UpdateFromHeaderSub As Integer = 0

        ''' <summary>既存の INVOICE から一時データを作ることを表す値。</summary>
        Private Const CreateFromExisting As Integer = 0

        ''' <summary>PACKING から新規に一時データを作ることを表す値。</summary>
        Private Const CreateFromPacking As Integer = 1
#End Region

#Region "選択肢"
        ''' <summary>
        ''' 元INV_NO の選択肢を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <returns>INVOICE 番号（メイン）の一覧。B/L DATE の新しい順に並びます。</returns>
        Public Function SelectInvoiceNoMainOptions(connection As SqlConnection, transaction As SqlTransaction) As IList(Of String)
            Dim results As New List(Of String)()

            Using command As New SqlCommand(ShanghaiInvoiceHeaderSql.CreateSelectInvoiceNoMainOptionsSql(), connection, transaction)
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        Dim invoiceNoMain As String = Me.ToText(reader("pk_invoice_no_main"))
                        ' 値が入っている行だけを選択肢にする
                        If invoiceNoMain.Length > 0 Then
                            results.Add(invoiceNoMain)
                        End If
                    End While
                End Using
            End Using

            Return results
        End Function

        ''' <summary>
        ''' 出荷方法の選択肢を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <returns>選択肢の一覧。</returns>
        Public Function SelectSyukkaHohoOptions(connection As SqlConnection, transaction As SqlTransaction) As IList(Of CodeOption)
            Return Me.SelectCodeOptions(connection, transaction,
                                        ShanghaiInvoiceHeaderSql.CreateSelectSyukkaHohoOptionsSql(),
                                        "syukka_hoho", "syukka_hoho_text")
        End Function

        ''' <summary>
        ''' 乙仲の選択肢を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <returns>選択肢の一覧。</returns>
        Public Function SelectOtsunakaOptions(connection As SqlConnection, transaction As SqlTransaction) As IList(Of CodeOption)
            Return Me.SelectCodeOptions(connection, transaction,
                                        ShanghaiInvoiceHeaderSql.CreateSelectOtsunakaOptionsSql(),
                                        "otsunaka_no", "otsunaka_no_text")
        End Function

        ''' <summary>
        ''' 配送タイプの選択肢を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <returns>選択肢の一覧。</returns>
        Public Function SelectHaisoTypeOptions(connection As SqlConnection, transaction As SqlTransaction) As IList(Of CodeOption)
            Return Me.SelectCodeOptions(connection, transaction,
                                        ShanghaiInvoiceHeaderSql.CreateSelectHaisoTypeOptionsSql(),
                                        "haiso_type", "haiso_type_text")
        End Function

        ''' <summary>
        ''' コンテナタイプの選択肢を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <returns>選択肢の一覧。</returns>
        Public Function SelectContainerTypeOptions(connection As SqlConnection, transaction As SqlTransaction) As IList(Of CodeOption)
            Return Me.SelectCodeOptions(connection, transaction,
                                        ShanghaiInvoiceHeaderSql.CreateSelectContainerTypeOptionsSql(),
                                        "container_type", "container_type_text")
        End Function

        ''' <summary>
        ''' 船名の選択肢を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <returns>船名の一覧。</returns>
        Public Function SelectVesselOptions(connection As SqlConnection, transaction As SqlTransaction) As IList(Of String)
            Dim results As New List(Of String)()

            Using command As New SqlCommand(ShanghaiInvoiceHeaderSql.CreateSelectVesselOptionsSql(), connection, transaction)
                command.Parameters.Add("@KoumokuCode", SqlDbType.NVarChar, KoumokuCodeLength).Value = VesselKoumokuCode
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        Dim text As String = Me.ToText(reader("contents1"))
                        ' 名称が入っている行だけを選択肢にする
                        If text.Length > 0 Then
                            results.Add(text)
                        End If
                    End While
                End Using
            End Using

            Return results
        End Function
#End Region

#Region "参照系"
        ''' <summary>
        ''' INVOICE が登録済みかどうかを判定します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <returns>判定結果。</returns>
        Public Function SelectExistence(connection As SqlConnection, transaction As SqlTransaction,
                                        invoiceNoMain As String) As ShanghaiInvoiceExistence
            Using command As New SqlCommand("dbo.usp_shanghai_invoice_sonzai_check", connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@invoice_no", SqlDbType.NVarChar, InvoiceNoMainLength).Value = invoiceNoMain

                Dim headerCount As SqlParameter = Me.AddOutputInteger(command, "@invoice_header_count")
                Dim subCount As SqlParameter = Me.AddOutputInteger(command, "@invoice_sub_count")
                Dim jyufukuCount As SqlParameter = Me.AddOutputInteger(command, "@jyufuku_count")
                command.ExecuteNonQuery()

                Dim result As New ShanghaiInvoiceExistence()
                result.HeaderCount = Me.ToInteger(headerCount.Value)
                result.SubCount = Me.ToInteger(subCount.Value)
                result.DateAndMethodKinds = Me.ToInteger(jyufukuCount.Value)
                Return result
            End Using
        End Function

        ''' <summary>
        ''' INVOICE のヘッダを取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <returns>ヘッダ。見つからないときは Nothing。</returns>
        Public Function SelectHeader(connection As SqlConnection, transaction As SqlTransaction,
                                     invoiceNoMain As String) As ShanghaiInvoiceHeader
            Using command As New SqlCommand(ShanghaiInvoiceHeaderSql.CreateSelectHeaderSql(), connection, transaction)
                command.Parameters.Add("@InvoiceNoMain", SqlDbType.NVarChar, InvoiceNoMainLength).Value = invoiceNoMain

                Using reader As SqlDataReader = command.ExecuteReader()
                    ' 該当が無いとき
                    If Not reader.Read() Then
                        Return Nothing
                    End If

                    Dim header As New ShanghaiInvoiceHeader()
                    header.InvoiceNoMain = Me.ToText(reader("pk_invoice_no_main"))
                    header.BlDate = Me.ToNullableDate(reader("bl_date"))
                    header.SyukkaHoho = Me.ToNullableInteger(reader("syukka_hoho"))
                    header.CartonQty = Me.ToNullableInteger(reader("carton_qty"))
                    header.Vessel = Me.ToText(reader("vessel"))
                    header.ReturnPallet = Me.ToNullableInteger(reader("return_pallet"))
                    header.RefNo = Me.ToText(reader("ref_no"))
                    header.OtsunakaNo = Me.ToNullableInteger(reader("otsunaka_no"))
                    header.ShiiresakiCode = Me.ToText(reader("shiiresaki_code"))
                    header.UpdateYmd = Me.ToNullableDate(reader("update_ymd"))
                    header.UpdateLogin = Me.ToText(reader("update_login"))
                    Return header
                End Using
            End Using
        End Function

        ''' <summary>
        ''' INVOICE 明細から集計した合計情報を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <returns>集計した情報。</returns>
        Public Function SelectTotal(connection As SqlConnection, transaction As SqlTransaction,
                                    invoiceNoMain As String) As ShanghaiInvoiceTotal
            Using command As New SqlCommand("dbo.usp_shanghai_invoice_gokei_get", connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@invoice_no", SqlDbType.NVarChar, InvoiceNoMainLength).Value = invoiceNoMain

                Dim blDate As SqlParameter = Me.AddOutputParameter(command, "@bl_date", SqlDbType.DateTime, 0)
                Dim syukkaHoho As SqlParameter = Me.AddOutputInteger(command, "@syukka_hoho")
                Dim totalCarton As SqlParameter = Me.AddOutputInteger(command, "@total_carton")
                Dim totalAmount As SqlParameter = Me.AddOutputParameter(command, "@total_amount", SqlDbType.Float, 0)
                Dim totalSuryo As SqlParameter = Me.AddOutputParameter(command, "@total_suryo", SqlDbType.Float, 0)
                Dim containerText As SqlParameter = Me.AddOutputParameter(command, "@container_text", SqlDbType.NVarChar, ContainerTypeTextLength)
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
        ''' 作業用の一時データから明細を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>明細の一覧。</returns>
        Public Function SelectDetails(connection As SqlConnection, transaction As SqlTransaction,
                                      loginCode As String) As IList(Of ShanghaiInvoiceHeaderSub)
            Dim results As New List(Of ShanghaiInvoiceHeaderSub)()

            Using command As New SqlCommand(ShanghaiInvoiceHeaderSql.CreateSelectDetailTempSql(), connection, transaction)
                Me.AddLoginCode(command, loginCode)
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        Dim detail As New ShanghaiInvoiceHeaderSub()
                        detail.InvoiceNoMain = Me.ToText(reader("pk_invoice_no_main"))
                        detail.InvoiceNo = Me.ToText(reader("pk_invoice_no"))
                        detail.HaisoType = Me.ToNullableInteger(reader("haiso_type"))
                        detail.ContainerType = Me.ToNullableInteger(reader("container_type"))
                        detail.HaisoYoteiYmd = Me.ToNullableDate(reader("haiso_yotei_ymd"))
                        detail.ContainerNo = Me.ToText(reader("container_no"))
                        detail.SealNo = Me.ToText(reader("seal_no"))
                        detail.CartonQty = Me.ToNullableInteger(reader("carton_qty"))
                        detail.InvoiceNoPrint = Me.ToText(reader("invoice_no_print"))
                        detail.DelFlg = Me.ToInteger(reader("del_flg"))
                        results.Add(detail)
                    End While
                End Using
            End Using

            Return results
        End Function
#End Region

#Region "更新系"
        ''' <summary>
        ''' 作業用の一時データを作り直します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="shanghaiChotatsu">担当区分。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <param name="isNew">PACKING から新規に作るときは True。既存から作るときは False。</param>
        ''' <returns>作成した件数。</returns>
        ''' <remarks>
        ''' このストアドは利用者単位で一時データを削除してから作り直します。
        ''' </remarks>
        Public Function CreateDetailTemp(connection As SqlConnection, transaction As SqlTransaction,
                                         loginCode As String, shanghaiChotatsu As Integer,
                                         invoiceNoMain As String, isNew As Boolean) As Integer
            Using command As New SqlCommand("dbo.usp_shanghai_invoice_header_sub_temp_create", connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@login_code", SqlDbType.NVarChar, LoginCodeLength).Value = loginCode
                command.Parameters.Add("@shanghai_chotatsu", SqlDbType.Int).Value = shanghaiChotatsu
                command.Parameters.Add("@invoice_no", SqlDbType.NVarChar, InvoiceNoMainLength).Value = invoiceNoMain
                command.Parameters.Add("@shinki_flg", SqlDbType.Int).Value = If(isNew, CreateFromPacking, CreateFromExisting)

                Dim invoiceCount As SqlParameter = Me.AddOutputInteger(command, "@invoice_count")
                command.ExecuteNonQuery()
                Return Me.ToInteger(invoiceCount.Value)
            End Using
        End Function

        ''' <summary>
        ''' 作業用の一時データの 1 行を更新します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="detail">更新する内容。</param>
        ''' <returns>更新した件数。</returns>
        Public Function UpdateDetailTemp(connection As SqlConnection, transaction As SqlTransaction,
                                         loginCode As String, detail As ShanghaiInvoiceHeaderSub) As Integer
            Using command As New SqlCommand(ShanghaiInvoiceHeaderSql.CreateUpdateDetailTempSql(), connection, transaction)
                Me.AddLoginCode(command, loginCode)
                command.Parameters.Add("@InvoiceNoMain", SqlDbType.NVarChar, InvoiceNoMainLength).Value = detail.InvoiceNoMain
                command.Parameters.Add("@InvoiceNo", SqlDbType.NVarChar, InvoiceNoLength).Value = detail.InvoiceNo
                command.Parameters.Add("@HaisoType", SqlDbType.Int).Value = Me.ToDbValue(detail.HaisoType)
                command.Parameters.Add("@ContainerType", SqlDbType.Int).Value = Me.ToDbValue(detail.ContainerType)
                command.Parameters.Add("@HaisoYoteiYmd", SqlDbType.DateTime).Value = Me.ToDbValue(detail.HaisoYoteiYmd)
                command.Parameters.Add("@ContainerNo", SqlDbType.NVarChar, ContainerTextLength).Value = Me.ToDbText(detail.ContainerNo)
                command.Parameters.Add("@SealNo", SqlDbType.NVarChar, ContainerTextLength).Value = Me.ToDbText(detail.SealNo)
                command.Parameters.Add("@CartonQty", SqlDbType.Int).Value = Me.ToDbValue(detail.CartonQty)
                Return command.ExecuteNonQuery()
            End Using
        End Function
#End Region

#Region "登録系"
        ''' <summary>
        ''' INVOICE のヘッダを登録します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="header">登録する内容。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>登録した件数。</returns>
        Public Function InsertHeader(connection As SqlConnection, transaction As SqlTransaction,
                                     header As ShanghaiInvoiceHeader, loginCode As String) As Integer
            Using command As New SqlCommand(ShanghaiInvoiceHeaderSql.CreateInsertHeaderSql(), connection, transaction)
                Me.AddHeaderParameters(command, header, loginCode)
                Return command.ExecuteNonQuery()
            End Using
        End Function

        ''' <summary>
        ''' INVOICE のヘッダを更新します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="header">更新する内容。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>更新した件数。</returns>
        Public Function UpdateHeader(connection As SqlConnection, transaction As SqlTransaction,
                                     header As ShanghaiInvoiceHeader, loginCode As String) As Integer
            Using command As New SqlCommand(ShanghaiInvoiceHeaderSql.CreateUpdateHeaderSql(), connection, transaction)
                Me.AddHeaderParameters(command, header, loginCode)
                Return command.ExecuteNonQuery()
            End Using
        End Function

        ''' <summary>
        ''' 作業用の一時データを本体の明細へ反映します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <remarks>
        ''' ストアド側で、その INVOICE の明細を削除してから一時データを挿入します。
        ''' </remarks>
        Public Sub ApplyDetailTemp(connection As SqlConnection, transaction As SqlTransaction, loginCode As String)
            Using command As New SqlCommand("dbo.usp_shanghai_invoice_header_sub_update", connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@login_code", SqlDbType.NVarChar, LoginCodeLength).Value = loginCode
                command.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>
        ''' コンテナ番号の件数を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <returns>件数。</returns>
        Public Function SelectContainerCount(connection As SqlConnection, transaction As SqlTransaction,
                                             invoiceNoMain As String) As Integer
            Using command As New SqlCommand("dbo.usp_shanghai_invoice_container_count", connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@invoice_no", SqlDbType.NVarChar, InvoiceNoMainLength).Value = invoiceNoMain
                Dim containerCount As SqlParameter = Me.AddOutputInteger(command, "@container_count")
                command.ExecuteNonQuery()
                Return Me.ToInteger(containerCount.Value)
            End Using
        End Function

        ''' <summary>
        ''' 印刷用 INVOICE_NO を組み立てる対象の明細を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <returns>対象の一覧。コンテナ区分・コンテナNO・INVOICE_NO の順。</returns>
        Public Function SelectPrintTargets(connection As SqlConnection, transaction As SqlTransaction,
                                           invoiceNoMain As String) As IList(Of ShanghaiInvoicePrintTarget)
            Dim results As New List(Of ShanghaiInvoicePrintTarget)()

            Using command As New SqlCommand(ShanghaiInvoiceHeaderSql.CreateSelectPrintTargetsSql(), connection, transaction)
                command.Parameters.Add("@InvoiceNoMain", SqlDbType.NVarChar, InvoiceNoMainLength).Value = invoiceNoMain
                command.Parameters.Add("@ExcludedHaisoType", SqlDbType.Int).Value = NisshinHaisoType
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        Dim target As New ShanghaiInvoicePrintTarget()
                        target.InvoiceNo = Me.ToText(reader("pk_invoice_no"))
                        target.ContainerType = Me.ToNullableInteger(reader("container_type"))
                        target.ContainerNo = Me.ToText(reader("container_no"))
                        target.HaisoYoteiYmd = Me.ToNullableDate(reader("haiso_yotei_ymd"))
                        results.Add(target)
                    End While
                End Using
            End Using

            Return results
        End Function

        ''' <summary>
        ''' 印刷用 INVOICE_NO を更新します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <param name="containerNo">コンテナ番号。</param>
        ''' <param name="invoiceNoPrint">設定する印刷用 INVOICE_NO。</param>
        ''' <returns>更新した件数。</returns>
        Public Function UpdateInvoiceNoPrint(connection As SqlConnection, transaction As SqlTransaction,
                                             invoiceNoMain As String, containerNo As String,
                                             invoiceNoPrint As String) As Integer
            Using command As New SqlCommand(ShanghaiInvoiceHeaderSql.CreateUpdateInvoiceNoPrintSql(), connection, transaction)
                command.Parameters.Add("@InvoiceNoPrint", SqlDbType.NVarChar, InvoiceNoLength).Value = Me.ToDbText(invoiceNoPrint)
                command.Parameters.Add("@InvoiceNoMain", SqlDbType.NVarChar, InvoiceNoMainLength).Value = invoiceNoMain
                command.Parameters.Add("@ContainerNo", SqlDbType.NVarChar, ContainerTextLength).Value = containerNo
                Return command.ExecuteNonQuery()
            End Using
        End Function
        ''' <summary>
        ''' 支給包装材の納入期日を、INVOICE ヘッダサブの配送予定日で更新します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <remarks>
        ''' 対象は購買発注明細のうち `hosozai_flg = 'Z'` の行です。
        ''' 支給包装材は購買発注データが元々無く、システムで作っているため、納期をここで更新します。
        ''' 明細を本体へ反映した後に実行してください。
        ''' </remarks>
        Public Sub UpdateShikyuHosozaiNouki(connection As SqlConnection, transaction As SqlTransaction,
                                            invoiceNoMain As String)
            Using command As New SqlCommand("dbo.usp_shanghai_invoice_shikyu_hosozai_nouki_update", connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@invoice_no", SqlDbType.NVarChar, InvoiceNoMainLength).Value = invoiceNoMain
                command.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>
        ''' POEM の巻芯入出庫実績へ入庫情報を登録します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <param name="procedureName">実行するストアド名。環境ごとに向き先が異なります。</param>
        ''' <remarks>
        ''' **別データベース（POEM）を更新します。**
        ''' ストアド側で、その INVOICE の入庫分を削除してから登録し直します。
        ''' 同一インスタンス内の別データベースのため、同じトランザクションで扱えます。
        '''
        ''' ストアドは中でデータベース名を直接指定しており接続先に追従しないため、
        ''' 呼び出すストアド名を設定で切り替えます。名前は呼び出し側で検査済みのものを渡してください。
        ''' </remarks>
        Public Sub UpdatePoemMakishin(connection As SqlConnection, transaction As SqlTransaction,
                                      loginCode As String, invoiceNoMain As String, procedureName As String)
            Using command As New SqlCommand("dbo." & procedureName, connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@login_code", SqlDbType.NVarChar, LoginCodeLength).Value = loginCode
                command.Parameters.Add("@invoice_no", SqlDbType.NVarChar, InvoiceNoMainLength).Value = invoiceNoMain
                command.ExecuteNonQuery()
            End Using
        End Sub
        ''' <summary>
        ''' 同じ枝番でコンテナ番号が複数ある件数を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>該当する件数。該当が無いときは 0。</returns>
        ''' <remarks>
        ''' ストアドは GROUP BY と HAVING で絞っているため、**該当が無いと行が返らず
        ''' OUTPUT 引数は NULL のまま**になります。NULL は「該当なし」として 0 を返します。
        ''' </remarks>
        Public Function SelectDuplicateContainerCount(connection As SqlConnection, transaction As SqlTransaction,
                                                      loginCode As String) As Integer
            Using command As New SqlCommand("dbo.usp_shanghai_container_type_check", connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@login_code", SqlDbType.NVarChar, LoginCodeLength).Value = loginCode
                Dim containerCount As SqlParameter = Me.AddOutputInteger(command, "@container_count")
                command.ExecuteNonQuery()
                Return Me.ToInteger(containerCount.Value)
            End Using
        End Function

        ''' <summary>
        ''' 明細のコンテナ番号とシール番号を、作業用の一時データの内容で一括更新します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <remarks>
        ''' `@syori_flg = 0`（ヘッダサブの入力情報から更新）で実行します。
        ''' **`@syori_flg = 1` は使いません。**そちらは menuNo 102 向けの分岐で、
        ''' ストアド内の結合先のテーブル名が食い違っており、暗黙の直積になる恐れがあります。
        ''' </remarks>
        Public Sub UpdateContainerNo(connection As SqlConnection, transaction As SqlTransaction, loginCode As String)
            Using command As New SqlCommand("dbo.usp_shanghai_container_no_update", connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@login_code", SqlDbType.NVarChar, LoginCodeLength).Value = loginCode
                command.Parameters.Add("@syori_flg", SqlDbType.Int).Value = UpdateFromHeaderSub
                command.ExecuteNonQuery()
            End Using
        End Sub
#End Region

#Region "内部処理"
        ''' <summary>
        ''' ヘッダの引数をまとめて追加します。
        ''' </summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="header">登録・更新する内容。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        Private Sub AddHeaderParameters(command As SqlCommand, header As ShanghaiInvoiceHeader, loginCode As String)
            command.Parameters.Add("@InvoiceNoMain", SqlDbType.NVarChar, InvoiceNoMainLength).Value = header.InvoiceNoMain
            command.Parameters.Add("@RefNo", SqlDbType.NVarChar, RefNoLength).Value = Me.ToDbText(header.RefNo)
            command.Parameters.Add("@BlDate", SqlDbType.DateTime).Value = Me.ToDbValue(header.BlDate)
            command.Parameters.Add("@Vessel", SqlDbType.NVarChar, VesselLength).Value = Me.ToDbText(header.Vessel)
            command.Parameters.Add("@SyukkaHoho", SqlDbType.Int).Value = Me.ToDbValue(header.SyukkaHoho)
            command.Parameters.Add("@CartonQty", SqlDbType.Int).Value = Me.ToDbValue(header.CartonQty)
            command.Parameters.Add("@ReturnPallet", SqlDbType.Int).Value = Me.ToDbValue(header.ReturnPallet)
            command.Parameters.Add("@OtsunakaNo", SqlDbType.Int).Value = Me.ToDbValue(header.OtsunakaNo)
            command.Parameters.Add("@LoginCode", SqlDbType.NVarChar, LoginCodeLength).Value = loginCode
        End Sub

        ''' <summary>
        ''' 「値と表示名」の組を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="sql">実行する SELECT 文。</param>
        ''' <param name="valueColumn">値の列名。</param>
        ''' <param name="textColumn">表示名の列名。</param>
        ''' <returns>選択肢の一覧。</returns>
        ''' <remarks>
        ''' 列名は本クラス内の固定値のみを渡します。画面からの値は受け取りません。
        ''' </remarks>
        Private Function SelectCodeOptions(connection As SqlConnection, transaction As SqlTransaction,
                                           sql As String, valueColumn As String, textColumn As String) As IList(Of CodeOption)
            Dim results As New List(Of CodeOption)()

            Using command As New SqlCommand(sql, connection, transaction)
                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        Dim value As Integer? = Me.ToNullableInteger(reader(valueColumn))
                        ' 値が無い行は選択肢にしない
                        If value.HasValue Then
                            results.Add(New CodeOption(value.Value, Me.ToText(reader(textColumn))))
                        End If
                    End While
                End Using
            End Using

            Return results
        End Function

        ''' <summary>
        ''' ログインコードの引数を追加します。
        ''' </summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        Private Sub AddLoginCode(command As SqlCommand, loginCode As String)
            command.Parameters.Add("@LoginCode", SqlDbType.NVarChar, LoginCodeLength).Value = loginCode
        End Sub

        ''' <summary>
        ''' OUTPUT の引数を追加します。
        ''' </summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="parameterName">引数名。</param>
        ''' <param name="dbType">型。</param>
        ''' <param name="size">長さ。長さの指定が要らない型では 0 を渡します。</param>
        ''' <returns>追加した引数。</returns>
        Private Function AddOutputParameter(command As SqlCommand, parameterName As String,
                                            dbType As SqlDbType, size As Integer) As SqlParameter
            Dim parameter As SqlParameter
            ' 長さの指定が要る型のとき
            If size > 0 Then
                parameter = command.Parameters.Add(parameterName, dbType, size)
            Else
                parameter = command.Parameters.Add(parameterName, dbType)
            End If
            parameter.Direction = ParameterDirection.Output
            Return parameter
        End Function

        ''' <summary>
        ''' 整数の OUTPUT 引数を追加します。
        ''' </summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="parameterName">引数名。</param>
        ''' <returns>追加した引数。</returns>
        Private Function AddOutputInteger(command As SqlCommand, parameterName As String) As SqlParameter
            Return Me.AddOutputParameter(command, parameterName, SqlDbType.Int, 0)
        End Function

        ''' <summary>
        ''' 整数を DB へ渡せる形にします。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>値。未設定のときは DBNull。</returns>
        Private Function ToDbValue(value As Integer?) As Object
            ' 値が無いとき
            If Not value.HasValue Then
                Return DBNull.Value
            End If
            Return value.Value
        End Function

        ''' <summary>
        ''' 日付を DB へ渡せる形にします。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>値。未設定のときは DBNull。</returns>
        Private Function ToDbValue(value As Date?) As Object
            ' 値が無いとき
            If Not value.HasValue Then
                Return DBNull.Value
            End If
            Return value.Value
        End Function

        ''' <summary>
        ''' 文字列を DB へ渡せる形にします。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>値。空のときは DBNull。</returns>
        Private Function ToDbText(value As String) As Object
            ' 値が無いとき
            If String.IsNullOrEmpty(value) Then
                Return DBNull.Value
            End If
            Return value
        End Function

        ''' <summary>
        ''' 値を整数に変換します。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>変換した値。変換できないときは 0。</returns>
        Private Function ToInteger(value As Object) As Integer
            Dim converted As Integer? = Me.ToNullableInteger(value)
            Return If(converted.HasValue, converted.Value, 0)
        End Function

        ''' <summary>
        ''' 値を整数に変換します。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>変換した値。NULL のときは Nothing。</returns>
        Private Function ToNullableInteger(value As Object) As Integer?
            ' NULL のとき
            If value Is Nothing OrElse Convert.IsDBNull(value) Then
                Return Nothing
            End If
            Return Convert.ToInt32(value)
        End Function

        ''' <summary>
        ''' 値を実数に変換します。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>変換した値。NULL のときは Nothing。</returns>
        Private Function ToNullableDouble(value As Object) As Double?
            ' NULL のとき
            If value Is Nothing OrElse Convert.IsDBNull(value) Then
                Return Nothing
            End If
            Return Convert.ToDouble(value)
        End Function

        ''' <summary>
        ''' 値を日付に変換します。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>変換した値。NULL のときは Nothing。</returns>
        Private Function ToNullableDate(value As Object) As Date?
            ' NULL のとき
            If value Is Nothing OrElse Convert.IsDBNull(value) Then
                Return Nothing
            End If
            Return Convert.ToDateTime(value)
        End Function

        ''' <summary>
        ''' 値を文字列に変換します。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>変換した値。NULL のときは空文字。</returns>
        Private Function ToText(value As Object) As String
            ' NULL のとき
            If value Is Nothing OrElse Convert.IsDBNull(value) Then
                Return String.Empty
            End If
            Return Convert.ToString(value)
        End Function
#End Region

    End Class
End Namespace
