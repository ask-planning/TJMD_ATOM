Option Strict On
Option Explicit On
Option Infer Off

Imports System.Data
Imports System.Data.SqlClient
Imports Atom.Core.Entities
Imports Atom.Data.Sql

Namespace Repositories

    ''' <summary>
    ''' 輸入進捗のデータアクセスを担います。旧 Access 版の共通処理 schedule_update に相当します。
    ''' </summary>
    ''' <remarks>
    ''' 現在実装しているのは処理NO 5（上海連絡・ヘッダ）／担当区分 1（上海）の経路だけです。
    ''' 他の処理NO は該当画面の実装時に追加してください。
    ''' カレンダー自動更新は処理NO 10（配送連絡→ロジ）専用のため、ここには含めていません。
    ''' </remarks>
    Public Class ImportProgressRepository

#Region "定数"
        ''' <summary>INVOICE 番号（メイン）の桁数。</summary>
        Private Const InvoiceNoMainLength As Integer = 20

        ''' <summary>進捗テーブルの INVOICE 番号の桁数。</summary>
        Private Const ProgressInvoiceNoLength As Integer = 50

        ''' <summary>ログインコードの桁数。</summary>
        Private Const LoginCodeLength As Integer = 6

        ''' <summary>担当区分のうち上海を表す値。</summary>
        Private Const ShanghaiKubun As Integer = 1

        ''' <summary>種類が取得できないときの既定値。</summary>
        Private Const DefaultSyurui As Integer = 1

        ''' <summary>RMB 建ての種類。</summary>
        Private Const RmbSyurui As Integer = 2

        ''' <summary>特殊な INVOICE の種類。</summary>
        Private Const SpecialSyurui As Integer = 99
#End Region

#Region "参照系"
        ''' <summary>
        ''' INVOICE ヘッダから出荷方法と BL DATE を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <returns>出荷方法と BL DATE。取得できないときは各項目が Nothing。</returns>
        Public Function SelectHeaderInfo(connection As SqlConnection, transaction As SqlTransaction,
                                         invoiceNoMain As String) As ShanghaiInvoiceTotal
            Dim result As New ShanghaiInvoiceTotal()

            Using command As New SqlCommand("dbo.usp_shanghai_invoice_header_pickup", connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@invoice_no", SqlDbType.NVarChar, InvoiceNoMainLength).Value = invoiceNoMain

                Using reader As SqlDataReader = command.ExecuteReader()
                    ' 1 件目だけを使います（配送先ごとに複数行返ります）
                    If reader.Read() Then
                        result.SyukkaHoho = Me.ReadNullableInteger(reader, "syukka_hoho")
                        result.BlDate = Me.ReadNullableDate(reader, "bl_date")
                    End If
                End Using
            End Using

            Return result
        End Function

        ''' <summary>
        ''' INVOICE ヘッダから BL DATE と出荷方法を取り直します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <returns>出荷方法と BL DATE。取得できないときは各項目が Nothing。</returns>
        Public Function SelectHeaderBlDate(connection As SqlConnection, transaction As SqlTransaction,
                                           invoiceNoMain As String) As ShanghaiInvoiceTotal
            Using command As New SqlCommand("dbo.usp_shanghai_invoice_header_get", connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@shanghai_chotatsu", SqlDbType.Int).Value = ShanghaiKubun
                command.Parameters.Add("@invoice_no", SqlDbType.NVarChar, InvoiceNoMainLength).Value = invoiceNoMain

                Dim blDate As SqlParameter = command.Parameters.Add("@bl_date", SqlDbType.DateTime)
                blDate.Direction = ParameterDirection.Output
                Dim syukkaHoho As SqlParameter = command.Parameters.Add("@syukka_hoho", SqlDbType.Int)
                syukkaHoho.Direction = ParameterDirection.Output
                command.ExecuteNonQuery()

                Dim result As New ShanghaiInvoiceTotal()
                result.BlDate = Me.ToNullableDate(blDate.Value)
                result.SyukkaHoho = Me.ToNullableInteger(syukkaHoho.Value)
                Return result
            End Using
        End Function

        ''' <summary>
        ''' 進捗データの件数を取得します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。使わないときは Nothing。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <returns>登録済みの件数。</returns>
        Public Function SelectProgressCount(connection As SqlConnection, transaction As SqlTransaction,
                                            invoiceNoMain As String) As Integer
            Using command As New SqlCommand("dbo.usp_shinchoku_jyokyo_sonzai_check", connection, transaction)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add("@invoice_no", SqlDbType.NVarChar, InvoiceNoMainLength).Value = invoiceNoMain
                Dim count As SqlParameter = command.Parameters.Add("@touroku_count", SqlDbType.Int)
                count.Direction = ParameterDirection.Output
                command.ExecuteNonQuery()

                Dim result As Integer? = Me.ToNullableInteger(count.Value)
                ' 値が無いとき
                If Not result.HasValue Then
                    Return 0
                End If
                Return result.Value
            End Using
        End Function
#End Region

#Region "更新系"
        ''' <summary>
        ''' 上海連絡（ヘッダ）の進捗を更新します。進捗データが無いときは先に作成します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="syoriYmd">処理日。</param>
        Public Sub UpdateShanghaiHeaderProgress(connection As SqlConnection, transaction As SqlTransaction,
                                                invoiceNoMain As String, loginCode As String, syoriYmd As Date)
            ' INVOICE 番号が無いときは何もしません（現行仕様）
            If String.IsNullOrEmpty(invoiceNoMain) OrElse
               String.Equals(invoiceNoMain, "NULL", StringComparison.OrdinalIgnoreCase) Then
                Return
            End If

            Dim headerInfo As ShanghaiInvoiceTotal = Me.SelectHeaderInfo(connection, transaction, invoiceNoMain)
            Dim syurui As Integer = Me.ResolveSyurui(invoiceNoMain, headerInfo.SyukkaHoho)
            Dim blDate As Date? = headerInfo.BlDate

            ' ヘッダから BL DATE を取り直す（現行仕様）
            Dim headerBlDate As ShanghaiInvoiceTotal = Me.SelectHeaderBlDate(connection, transaction, invoiceNoMain)
            If headerBlDate.BlDate.HasValue Then
                blDate = headerBlDate.BlDate
            End If

            ' 進捗データが無いとき
            If Me.SelectProgressCount(connection, transaction, invoiceNoMain) = 0 Then
                Me.InsertProgress(connection, transaction, invoiceNoMain, syurui, blDate, loginCode)
            End If

            Using command As New SqlCommand(ImportProgressSql.CreateUpdateShanghaiHeaderProgressSql(),
                                            connection, transaction)
                Me.AddProgressInvoiceNo(command, invoiceNoMain)
                Me.AddNullableDate(command, "@SyukkoYmd", blDate)
                command.Parameters.Add("@SyoriYmd", SqlDbType.DateTime).Value = syoriYmd.Date
                command.Parameters.Add("@UpdateLogin", SqlDbType.NVarChar, LoginCodeLength).Value = loginCode
                command.ExecuteNonQuery()
            End Using
        End Sub
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 進捗データを 1 件登録します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <param name="syurui">進捗の種類。</param>
        ''' <param name="blDate">出港日。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        Private Sub InsertProgress(connection As SqlConnection, transaction As SqlTransaction,
                                   invoiceNoMain As String, syurui As Integer, blDate As Date?, loginCode As String)
            Using command As New SqlCommand(ImportProgressSql.CreateInsertProgressSql(), connection, transaction)
                Me.AddProgressInvoiceNo(command, invoiceNoMain)
                command.Parameters.Add("@Syurui", SqlDbType.Int).Value = syurui
                Me.AddNullableDate(command, "@SyukkoYmd", blDate)
                command.Parameters.Add("@UpdateLogin", SqlDbType.NVarChar, LoginCodeLength).Value = loginCode
                command.ExecuteNonQuery()
            End Using
        End Sub

        ''' <summary>
        ''' 進捗の種類を決めます。
        ''' </summary>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <param name="syukkaHoho">ヘッダから取得した出荷方法。</param>
        ''' <returns>進捗の種類。</returns>
        ''' <remarks>
        ''' 出荷方法が取得できたときは「出荷方法 + 1」。取得できないときは INVOICE 番号から判定します。
        ''' </remarks>
        Private Function ResolveSyurui(invoiceNoMain As String, syukkaHoho As Integer?) As Integer
            ' 出荷方法が取得できたとき
            If syukkaHoho.HasValue Then
                Return syukkaHoho.Value + 1
            End If

            ' SJ- かつ 5 文字目が 2 のとき
            If invoiceNoMain.StartsWith("SJ-", StringComparison.Ordinal) AndAlso
               invoiceNoMain.Length >= 5 AndAlso
               String.Equals(invoiceNoMain.Substring(4, 1), "2", StringComparison.Ordinal) Then
                Return SpecialSyurui
            End If
            ' 4 文字目が R のとき（RMB 建て）
            If invoiceNoMain.Length >= 4 AndAlso
               String.Equals(invoiceNoMain.Substring(3, 1), "R", StringComparison.OrdinalIgnoreCase) Then
                Return RmbSyurui
            End If
            ' 上記以外のとき
            Return DefaultSyurui
        End Function

        ''' <summary>進捗テーブルの INVOICE 番号パラメータを追加します。</summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        Private Sub AddProgressInvoiceNo(command As SqlCommand, invoiceNoMain As String)
            command.Parameters.Add("@InvoiceNoMain", SqlDbType.NVarChar, ProgressInvoiceNoLength).Value = invoiceNoMain
        End Sub

        ''' <summary>日付のパラメータを追加します。値が無いときは NULL にします。</summary>
        ''' <param name="command">対象のコマンド。</param>
        ''' <param name="parameterName">パラメータ名。</param>
        ''' <param name="value">設定する値。</param>
        Private Sub AddNullableDate(command As SqlCommand, parameterName As String, value As Date?)
            Dim parameter As SqlParameter = command.Parameters.Add(parameterName, SqlDbType.DateTime)
            ' 値があるとき
            If value.HasValue Then
                parameter.Value = value.Value
            Else
                parameter.Value = DBNull.Value
            End If
        End Sub

        ''' <summary>読み取り結果の整数を取得します。NULL のときは Nothing を返します。</summary>
        ''' <param name="reader">読み取り中のリーダー。</param>
        ''' <param name="columnName">列名。</param>
        ''' <returns>列の値。</returns>
        Private Function ReadNullableInteger(reader As SqlDataReader, columnName As String) As Integer?
            Dim index As Integer = reader.GetOrdinal(columnName)
            ' NULL のとき
            If reader.IsDBNull(index) Then
                Return Nothing
            End If
            Return Convert.ToInt32(reader.GetValue(index))
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
#End Region

    End Class
End Namespace
