Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports System.IO
Imports Atom.Core.Entities
Imports Atom.Core.Logging
Imports Atom.Core.Services
Imports Atom.Data.Repositories

Namespace Services

    ''' <summary>
    ''' 上海輸入 INVOICE / PACKING 取込（menuNo=101）の一連の手順を制御します。
    ''' </summary>
    ''' <remarks>
    ''' 画面は利用者への確認をはさみながらこのクラスのメソッドを順に呼び出します。
    ''' 1. LoadToTemp   Excel を解析して一時テーブルへ登録する
    ''' 2. Analyze      REF_NO・翌月以降の件数・登録済み件数を求める
    ''' 3. RunChecks    業務チェックを行い、エラーと警告の件数を求める
    ''' 4. Register     本登録する（置換・明細・ヘッダ・進捗をひとつのトランザクションで扱う）
    ''' </remarks>
    Public Class ShanghaiImportService

#Region "定数"
        ''' <summary>ログの発生元名。</summary>
        Private Const LogSource As String = "ShanghaiImportService"

        ''' <summary>発注先の項目コード。</summary>
        Public Const ChumonsakiKoumokuCode As String = "COS"

        ''' <summary>発注先区分のうち上海タジマを表す値。</summary>
        Private Const ShanghaiTajimaKubun As Integer = 1

        ''' <summary>上海タジマの RMB 建てを表す得意先区分。</summary>
        Private Const ShanghaiTajimaRmbKubun As Integer = 3

        ''' <summary>品目コードの列名。R3 品目コードと同時に更新する必要がある項目です。</summary>
        Private Const ShanghaiCodeColumnName As String = "shanghai_code"
#End Region

#Region "フィールド"
        ''' <summary>取込データのデータアクセス。</summary>
        Private ReadOnly _importRepository As ShanghaiImportRepository

        ''' <summary>進捗のデータアクセス。</summary>
        Private ReadOnly _progressRepository As ImportProgressRepository

        ''' <summary>Excel の解析。</summary>
        Private ReadOnly _parser As ShanghaiImportExcelParser

        ''' <summary>INVOICE 番号の編集。</summary>
        Private ReadOnly _formatter As ShanghaiInvoiceNoFormatter
#End Region

#Region "コンストラクタ"
        ''' <summary>既定の状態で初期化します。</summary>
        Public Sub New()
            Me._importRepository = New ShanghaiImportRepository()
            Me._progressRepository = New ImportProgressRepository()
            Me._parser = New ShanghaiImportExcelParser()
            Me._formatter = New ShanghaiInvoiceNoFormatter()
        End Sub
#End Region

#Region "参照系"
        ''' <summary>
        ''' 発注先の選択肢を取得します。
        ''' </summary>
        ''' <returns>項目マスタから取得した発注先の一覧。</returns>
        Public Function SelectChumonsakiList() As IList(Of Koumoku)
            Try
                Using connection As SqlConnection = Database.CreateOpenConnection()
                    Return Me._importRepository.SelectKoumokuList(connection, Nothing, ChumonsakiKoumokuCode)
                End Using
            Catch ex As SqlException
                Logger.WriteError(LogSource, "発注先の選択肢の取得に失敗しました。", ex, String.Empty)
                Throw
            End Try
        End Function

        ''' <summary>
        ''' 取込結果確認画面に表示する INVOICE の一覧を取得します。
        ''' </summary>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="invoiceNo">絞り込む INVOICE 番号。空のときは全件。</param>
        ''' <returns>取込結果の一覧。</returns>
        Public Function SelectInvoiceResults(loginCode As String, invoiceNo As String) As IList(Of ShanghaiInvoiceImportRow)
            Try
                Using connection As SqlConnection = Database.CreateOpenConnection()
                    Return Me._importRepository.SelectInvoiceResults(connection, Nothing, loginCode, invoiceNo)
                End Using
            Catch ex As SqlException
                Logger.WriteError(LogSource, "取込結果（INVOICE）の取得に失敗しました。", ex, loginCode)
                Throw
            End Try
        End Function

        ''' <summary>
        ''' 取込結果確認画面に表示する PACKING の一覧を取得します。
        ''' </summary>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="invoiceNo">絞り込む INVOICE 番号。空のときは全件。</param>
        ''' <returns>取込結果の一覧。</returns>
        Public Function SelectPackingResults(loginCode As String, invoiceNo As String) As IList(Of ShanghaiPackingImportRow)
            Try
                Using connection As SqlConnection = Database.CreateOpenConnection()
                    Return Me._importRepository.SelectPackingResults(connection, Nothing, loginCode, invoiceNo)
                End Using
            Catch ex As SqlException
                Logger.WriteError(LogSource, "取込結果（PACKING）の取得に失敗しました。", ex, loginCode)
                Throw
            End Try
        End Function

        ''' <summary>
        ''' 一時データに含まれる INVOICE 番号の一覧を取得します。取込結果確認画面の選択肢に使います。
        ''' </summary>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>INVOICE 番号の一覧。昇順。</returns>
        Public Function SelectInvoiceNoList(loginCode As String) As IList(Of String)
            Try
                Using connection As SqlConnection = Database.CreateOpenConnection()
                    Return Me._importRepository.SelectInvoiceNoList(connection, Nothing, loginCode)
                End Using
            Catch ex As SqlException
                Logger.WriteError(LogSource, "INVOICE 番号の一覧の取得に失敗しました。", ex, loginCode)
                Throw
            End Try
        End Function

        ''' <summary>
        ''' 一時データから取込前の情報を求めます。
        ''' </summary>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="shanghaiChotatsu">担当区分。REF_NO の編集に使います。</param>
        ''' <returns>取込前の情報。</returns>
        Public Function Analyze(loginCode As String, shanghaiChotatsu As Integer) As ShanghaiImportAnalysis
            Dim result As New ShanghaiImportAnalysis()

            Try
                Using connection As SqlConnection = Database.CreateOpenConnection()
                    result.HasTempData = Me._importRepository.SelectCheckTempCount(connection, Nothing, loginCode) > 0
                    ' 一時データが無いとき
                    If Not result.HasTempData Then
                        Return result
                    End If

                    Dim invoiceNos As IList(Of String) =
                        Me._importRepository.SelectInvoiceNoList(connection, Nothing, loginCode)
                    result.RefNo = Me._formatter.CreateRefNo(invoiceNos, shanghaiChotatsu)
                    ' INVOICE 番号があるとき
                    If invoiceNos.Count > 0 Then
                        result.InvoiceNoMain = Me._formatter.GetInvoiceNoMain(invoiceNos(0))
                    End If

                    result.FutureBlDateCount = Me._importRepository.SelectFutureBlDateCount(
                        connection, Nothing, loginCode, Me.GetNextMonthFirstDate())
                    result.Registered = Me._importRepository.SelectRegisteredCounts(connection, Nothing, loginCode)
                End Using
            Catch ex As SqlException
                Logger.WriteError(LogSource, "取込データの確認に失敗しました。", ex, loginCode)
                Throw
            End Try

            Return result
        End Function
#End Region

#Region "更新系"
        ''' <summary>
        ''' 取込結果確認画面での編集内容を、INVOICE の一時データへ反映します。
        ''' </summary>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="seqNo">更新する行の連番。</param>
        ''' <param name="field">更新する項目の定義。</param>
        ''' <param name="value">設定する値。Nothing のときは NULL を設定します。</param>
        ''' <remarks>
        ''' 更新するのは一時テーブルだけです。本登録済みの明細には反映されません。
        ''' またチェック処理はやり直さないため、ERR_CONTENTS は更新されません（旧 Access 版と同じ挙動）。
        ''' </remarks>
        Public Sub UpdateInvoiceResultValue(loginCode As String, seqNo As Integer,
                                            field As ShanghaiImportField, value As Object)
            Try
                Using connection As SqlConnection = Database.CreateOpenConnection()
                    ' 品目コードのとき（R3 品目コードも併せて更新する）
                    If IsShanghaiCodeField(field) Then
                        Me._importRepository.UpdateInvoiceShanghaiCode(connection, Nothing, loginCode, seqNo,
                                                                      ToCodeText(value))
                    ' それ以外の項目のとき
                    Else
                        Me._importRepository.UpdateInvoiceCheckTempValue(connection, Nothing, loginCode, seqNo, field, value)
                    End If
                End Using
            Catch ex As SqlException
                Logger.WriteError(LogSource, "取込結果（INVOICE）の更新に失敗しました。", ex, loginCode)
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' 取込結果確認画面での編集内容を、PACKING の一時データへ反映します。
        ''' </summary>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="seqNo">更新する行の連番。</param>
        ''' <param name="field">更新する項目の定義。</param>
        ''' <param name="value">設定する値。Nothing のときは NULL を設定します。</param>
        Public Sub UpdatePackingResultValue(loginCode As String, seqNo As Integer,
                                            field As ShanghaiImportField, value As Object)
            Try
                Using connection As SqlConnection = Database.CreateOpenConnection()
                    ' 品目コードのとき（R3 品目コードも併せて更新する）
                    If IsShanghaiCodeField(field) Then
                        Me._importRepository.UpdatePackingShanghaiCode(connection, Nothing, loginCode, seqNo,
                                                                      ToCodeText(value))
                    ' それ以外の項目のとき
                    Else
                        Me._importRepository.UpdatePackingCheckTempValue(connection, Nothing, loginCode, seqNo, field, value)
                    End If
                End Using
            Catch ex As SqlException
                Logger.WriteError(LogSource, "取込結果（PACKING）の更新に失敗しました。", ex, loginCode)
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' 更新対象が品目コードかどうかを判定します。
        ''' </summary>
        ''' <param name="field">更新する項目の定義。</param>
        ''' <returns>品目コードのとき True。</returns>
        ''' <remarks>
        ''' 品目コードは R3 品目コードの元になる値のため、単独では更新できません。
        ''' </remarks>
        Private Shared Function IsShanghaiCodeField(field As ShanghaiImportField) As Boolean
            ' 定義が無いとき
            If field Is Nothing Then
                Return False
            End If
            Return String.Equals(field.ColumnName, ShanghaiCodeColumnName, StringComparison.Ordinal)
        End Function

        ''' <summary>
        ''' 品目コードとして渡された値を文字列にします。
        ''' </summary>
        ''' <param name="value">画面から渡された値。</param>
        ''' <returns>文字列。値が無いときは空文字。</returns>
        Private Shared Function ToCodeText(value As Object) As String
            ' 値が無いとき
            If value Is Nothing Then
                Return String.Empty
            End If
            Return value.ToString()
        End Function

        ''' <summary>
        ''' Excel を解析して一時テーブルへ登録します。既存の一時データは削除します。
        ''' </summary>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="tsukaCode">画面で選択した通貨。JPY または RMB。</param>
        ''' <param name="invoiceStream">INVOICE の Excel。</param>
        ''' <param name="packingStream">PACKING の Excel。</param>
        Public Sub LoadToTemp(loginCode As String, tsukaCode As String,
                              invoiceStream As Stream, packingStream As Stream)
            ' Excel の解析はデータベースに触らないため、先に済ませます。
            Dim invoiceRows As IList(Of ShanghaiInvoiceImportRow) = Me._parser.ParseInvoice(invoiceStream, tsukaCode)
            Dim packingRows As IList(Of ShanghaiPackingImportRow) = Me._parser.ParsePacking(packingStream)

            Try
                Using connection As SqlConnection = Database.CreateOpenConnection()
                    Using transaction As SqlTransaction = connection.BeginTransaction()
                        Try
                            Me._importRepository.DeleteCheckTemp(connection, transaction, loginCode)
                            Me._importRepository.InsertInvoiceCheckTemp(connection, transaction, loginCode, invoiceRows)
                            Me._importRepository.InsertPackingCheckTemp(connection, transaction, loginCode, packingRows)
                            transaction.Commit()
                        Catch
                            transaction.Rollback()
                            Throw
                        End Try
                    End Using
                End Using
            Catch ex As SqlException
                Logger.WriteError(LogSource, "取込データの一時登録に失敗しました。", ex, loginCode)
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' 一時データに対して業務チェックを行います。
        ''' </summary>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="shanghaiChotatsu">担当区分。貿易のとき PO の重複チェックを行いません。</param>
        ''' <returns>チェック結果。</returns>
        Public Function RunChecks(loginCode As String, shanghaiChotatsu As Integer) As ImportCheckResult
            Dim result As New ImportCheckResult()

            Try
                Using connection As SqlConnection = Database.CreateOpenConnection()
                    Me._importRepository.UpdateErrorClear(connection, Nothing, loginCode)

                    ' 貿易以外のとき（1 つの PO が複数 INVOICE に分かれていないかを見る）
                    If shanghaiChotatsu <> ShanghaiInvoiceNoFormatter.BouekiKubun Then
                        Dim duplicated As Integer =
                            Me._importRepository.UpdateDuplicateCheck(connection, Nothing, loginCode)
                        ' 重複しているとき（以降のチェックは行わない）
                        If duplicated <> 0 Then
                            result.IsPoDuplicated = True
                            result.ErrorCount = duplicated
                            Return result
                        End If
                    End If

                    result.Merge(Me._importRepository.UpdateInvoiceCheck(connection, Nothing, loginCode))
                    result.Merge(Me._importRepository.UpdatePackingCheck(connection, Nothing, loginCode))

                    ' エラーが無いときだけ金額のチェックへ進む（現行仕様）
                    If result.ErrorCount = 0 Then
                        result.Merge(Me._importRepository.UpdateAmountCheck(connection, Nothing, loginCode))
                    End If
                End Using
            Catch ex As SqlException
                Logger.WriteError(LogSource, "取込データのチェックに失敗しました。", ex, loginCode)
                Throw
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 一時データを本テーブルへ登録します。削除から進捗更新までを 1 つのトランザクションで扱います。
        ''' </summary>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="shanghaiChotatsu">担当区分。</param>
        ''' <param name="customerKbn">発注先区分。</param>
        ''' <param name="tsukaCode">画面で選択した通貨。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <param name="isReplace">登録済みデータを置き換えるとき True。追加のときは False。</param>
        ''' <returns>登録された INVOICE と PACKING の件数。</returns>
        Public Function Register(loginCode As String, shanghaiChotatsu As Integer, customerKbn As Integer,
                                 tsukaCode As String, invoiceNoMain As String, isReplace As Boolean) As ShanghaiImportCounts
            Try
                Using connection As SqlConnection = Database.CreateOpenConnection()
                    Using transaction As SqlTransaction = connection.BeginTransaction()
                        Try
                            ' 置換のとき（登録済みのヘッダ・明細を削除する）
                            If isReplace Then
                                Me.DeleteRegistered(connection, transaction, loginCode, invoiceNoMain)
                            End If

                            Dim counts As ShanghaiImportCounts =
                                Me._importRepository.InsertMeisai(connection, transaction, loginCode, shanghaiChotatsu)
                            ' 明細が 1 件も登録されなかったとき
                            If counts.InvoiceCount = 0 OrElse counts.PackingCount = 0 Then
                                transaction.Rollback()
                                Return counts
                            End If

                            Dim total As ShanghaiInvoiceTotal =
                                Me._importRepository.SelectInvoiceTotal(connection, transaction, invoiceNoMain)

                            ' 追加でなく、かつヘッダが未登録のときだけヘッダを作る
                            If isReplace OrElse total.HeaderCount = 0 Then
                                Dim shiiresakiCode As String = Me._importRepository.SelectShiiresakiCode(
                                    connection, transaction, Me.ResolveCustomerKbn(customerKbn, tsukaCode))
                                Me._importRepository.InsertInvoiceHeader(connection, transaction, invoiceNoMain,
                                                                         total, shiiresakiCode, loginCode)
                            End If

                            Me._progressRepository.UpdateShanghaiHeaderProgress(connection, transaction,
                                                                                invoiceNoMain, loginCode, Date.Today)
                            transaction.Commit()
                            Return counts
                        Catch
                            transaction.Rollback()
                            Throw
                        End Try
                    End Using
                End Using
            Catch ex As SqlException
                Logger.WriteError(LogSource, "INVOICE / PACKING の本登録に失敗しました。", ex, loginCode)
                Throw
            End Try
        End Function
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 登録済みのヘッダ・ヘッダサブ・明細を削除します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        Private Sub DeleteRegistered(connection As SqlConnection, transaction As SqlTransaction,
                                     loginCode As String, invoiceNoMain As String)
            Dim invoiceNos As IList(Of String) =
                Me._importRepository.SelectRegisteredInvoiceNoList(connection, transaction, loginCode)
            Dim packingNos As IList(Of String) =
                Me._importRepository.SelectRegisteredPackingInvoiceNoList(connection, transaction, loginCode)

            Me._importRepository.DeleteInvoiceHeader(connection, transaction, invoiceNoMain)

            For Each invoiceNo As String In invoiceNos
                Me._importRepository.DeleteInvoiceHeaderSub(connection, transaction, invoiceNo)
                Me._importRepository.DeleteInvoiceMeisai(connection, transaction, invoiceNo)
            Next

            For Each invoiceNo As String In packingNos
                Me._importRepository.DeleteInvoiceHeaderSub(connection, transaction, invoiceNo)
                Me._importRepository.DeletePackingMeisai(connection, transaction, invoiceNo)
            Next
        End Sub

        ''' <summary>
        ''' 仕入先コードを引くための得意先区分を求めます。
        ''' </summary>
        ''' <param name="customerKbn">画面で選択した発注先区分。</param>
        ''' <param name="tsukaCode">画面で選択した通貨。</param>
        ''' <returns>得意先マスタの customer_kbn_num。</returns>
        ''' <remarks>
        ''' 得意先マスタは「得意先 × 通貨」で区分が分かれています。
        ''' 現行仕様に合わせ、上海タジマで RMB を選んだときだけ RMB 用の区分へ読み替えます。
        ''' 引数名を tsukaCode にしているのは、VB が大文字小文字を区別せず
        ''' 引数 tsuka が型 Tsuka を隠してしまうためです。
        ''' </remarks>
        Private Function ResolveCustomerKbn(customerKbn As Integer, tsukaCode As String) As Integer
            ' 上海タジマで RMB を選んだとき
            If customerKbn = ShanghaiTajimaKubun AndAlso
               String.Equals(tsukaCode, Tsuka.Rmb, StringComparison.Ordinal) Then
                Return ShanghaiTajimaRmbKubun
            End If
            Return customerKbn
        End Function

        ''' <summary>
        ''' 翌月 1 日を求めます。BL DATE の判定に使います。
        ''' </summary>
        ''' <returns>翌月 1 日。</returns>
        Private Function GetNextMonthFirstDate() As Date
            Dim today As Date = Date.Today
            Return New Date(today.Year, today.Month, 1).AddMonths(1)
        End Function
#End Region

    End Class
End Namespace
