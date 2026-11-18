Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports Atom.Core.Entities
Imports Atom.Core.Logging
Imports Atom.Data.Repositories

Namespace Services

    ''' <summary>
    ''' 入荷明細書（menuNo=103 §11）の発行を組み立てます。
    ''' </summary>
    ''' <remarks>
    ''' 設計書 §11.3 の手順です。
    ''' 1. 入荷明細を作り直す（利用者が「作り直す」を選んだときだけ）
    ''' 2. 印刷用一時テーブルへ写す
    ''' 3. 帳票に印字する内容を読み出す
    '''
    ''' 1 〜 3 をひとつのトランザクションで扱います。
    ''' 作成の途中で失敗したとき、中途半端な明細が残らないようにするためです。
    '''
    ''' **進捗状況の更新（§11.3 処理NO 10）は実装していません。**
    ''' 更新する列と条件が読み取れていないためです（設計書 §18 C-39）。
    ''' </remarks>
    Public Class ShanghaiNyukaMeisaiService

#Region "定数"
        ''' <summary>ログの発生元名。</summary>
        Private Const LogSource As String = "ShanghaiNyukaMeisaiService"

        ''' <summary>元INV_NO が未入力のときの文言。旧 Access 版を踏襲します。</summary>
        Public Const InvoiceNoRequiredMessage As String = "INVOICE_NOを入力してください。"

        ''' <summary>
        ''' 社内参照番号と納入先プラントの絞り込み。
        ''' </summary>
        ''' <remarks>
        ''' **103 は全件を対象にします。**ストアドが `LIKE` で受けるため、
        ''' すべてに当たる値を渡します（設計書 §18 C-50。`ATOM.adp` の p-code で確認）。
        ''' 105 は画面で指定した値を渡しますが、ここでは扱いません。
        ''' </remarks>
        Private Const AllFilter As String = "%%"
#End Region

#Region "フィールド"
        ''' <summary>データアクセス。</summary>
        Private ReadOnly _repository As ShanghaiNyukaMeisaiRepository
#End Region

#Region "コンストラクタ"
        ''' <summary>既定の構成で初期化します。</summary>
        Public Sub New()
            Me._repository = New ShanghaiNyukaMeisaiRepository()
        End Sub
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' 作成済みの入荷明細があるかを調べます。
        ''' </summary>
        ''' <param name="shanghaiChotatsu">担当区分。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <returns>調べた結果。</returns>
        ''' <remarks>
        ''' 設計書 §11.2。件数が 0 かどうかで、画面に出す確認の文言が変わります。
        ''' ここでは読み取るだけで、データは変えません。
        ''' </remarks>
        Public Function CheckExisting(shanghaiChotatsu As Integer,
                                      invoiceNoMain As String) As ShanghaiNyukaMeisaiCheckResult

            Dim result As New ShanghaiNyukaMeisaiCheckResult()
            result.ValidationMessage = Validate(invoiceNoMain)
            ' 入力に問題があるとき
            If Not String.IsNullOrEmpty(result.ValidationMessage) Then
                Return result
            End If

            Using connection As SqlConnection = Database.CreateOpenConnection()
                Using transaction As SqlTransaction = connection.BeginTransaction()
                    Try
                        result.Kensu = Me._repository.SelectKensu(connection, transaction,
                            shanghaiChotatsu, invoiceNoMain, AllFilter, AllFilter)
                        transaction.Commit()
                    Catch ex As Exception
                        transaction.Rollback()
                        Logger.WriteError(LogSource,
                            "入荷明細の件数を取得できませんでした。INVOICE " & invoiceNoMain, ex, String.Empty)
                        Throw
                    End Try
                End Using
            End Using

            Return result
        End Function

        ''' <summary>
        ''' 入荷明細書に印字する内容を作ります。
        ''' </summary>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="shanghaiChotatsu">担当区分。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <param name="recreate">入荷明細を作り直すときは True。</param>
        ''' <returns>発行の結果。</returns>
        ''' <remarks>
        ''' **`recreate` が True のときは、既存の入荷明細を削除して作り直します。**
        ''' 利用者が「作り直す」または「新規作成」を選んだときだけ True にしてください。
        ''' </remarks>
        Public Function CreateReport(loginCode As String, shanghaiChotatsu As Integer,
                                     invoiceNoMain As String,
                                     recreate As Boolean) As ShanghaiNyukaMeisaiResult

            Dim result As New ShanghaiNyukaMeisaiResult()
            result.ValidationMessage = Validate(invoiceNoMain)
            ' 入力に問題があるとき
            If Not String.IsNullOrEmpty(result.ValidationMessage) Then
                Return result
            End If

            Using connection As SqlConnection = Database.CreateOpenConnection()
                Using transaction As SqlTransaction = connection.BeginTransaction()
                    Try
                        ' 作り直すとき
                        If recreate Then
                            result.CreatedCount = Me._repository.ExecuteCreate(connection, transaction,
                                loginCode, shanghaiChotatsu, invoiceNoMain, AllFilter, AllFilter)
                        End If

                        result.PrintCount = Me._repository.ExecuteTempCreate(connection, transaction,
                            loginCode, shanghaiChotatsu, invoiceNoMain, AllFilter, AllFilter)

                        result.Rows = Me._repository.SelectReportRows(connection, transaction, loginCode)

                        Me.WarnOnLostRows(result, loginCode)

                        transaction.Commit()
                    Catch ex As Exception
                        transaction.Rollback()
                        Logger.WriteError(LogSource,
                            "入荷明細書の作成に失敗しました。INVOICE " & invoiceNoMain, ex, loginCode)
                        Throw
                    End Try
                End Using
            End Using

            Logger.WriteInformation(LogSource,
                "入荷明細書を作成しました。INVOICE " & invoiceNoMain &
                " 件数 " & result.Rows.Count.ToString(Globalization.CultureInfo.InvariantCulture), loginCode)
            Return result
        End Function
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 入力を確かめます。
        ''' </summary>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <returns>問題があるときはその文言。問題が無いときは空文字。</returns>
        ''' <remarks>
        ''' 設計書 §11.1。**担当区分での制限はありません。**
        ''' 配送手配書（§10.1）と違い、旧 Access 側にもその判定はありませんでした。
        ''' </remarks>
        Private Shared Function Validate(invoiceNoMain As String) As String
            ' 元INV_NO が未入力のとき
            If String.IsNullOrWhiteSpace(invoiceNoMain) Then
                Return InvoiceNoRequiredMessage
            End If
            Return String.Empty
        End Function

        ''' <summary>
        ''' 一時データを写したのに帳票の行が取れなかったときに、その旨をログへ残します。
        ''' </summary>
        ''' <param name="result">発行の結果。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <remarks>
        ''' **ビューは品目マスタと保管場所を外部結合しています。**
        ''' そのため写した件数と帳票の行数は本来そろいます。食い違うときは
        ''' ビューの結合条件が変わった可能性があるため、追えるように記録します。
        ''' </remarks>
        Private Sub WarnOnLostRows(result As ShanghaiNyukaMeisaiResult, loginCode As String)
            ' 写した件数と帳票の行数が合っているとき
            If result.PrintCount = result.Rows.Count Then
                Return
            End If

            Logger.WriteWarning(LogSource,
                "印刷用に写した件数と帳票の行数が合いません。写した件数 " &
                result.PrintCount.ToString(Globalization.CultureInfo.InvariantCulture) &
                " 帳票の行数 " & result.Rows.Count.ToString(Globalization.CultureInfo.InvariantCulture),
                loginCode)
        End Sub
#End Region

    End Class

    ''' <summary>
    ''' 作成済みの入荷明細を調べた結果です。
    ''' </summary>
    Public Class ShanghaiNyukaMeisaiCheckResult

        ''' <summary>入力に問題があるときの文言。問題が無いときは空文字。</summary>
        Public Property ValidationMessage As String = String.Empty

        ''' <summary>件数と最終更新日。</summary>
        Public Property Kensu As ShanghaiNyukaMeisaiKensu = New ShanghaiNyukaMeisaiKensu()

    End Class

    ''' <summary>
    ''' 入荷明細書の発行結果です。
    ''' </summary>
    Public Class ShanghaiNyukaMeisaiResult

        ''' <summary>入力に問題があるときの文言。問題が無いときは空文字。</summary>
        Public Property ValidationMessage As String = String.Empty

        ''' <summary>作り直したときの作成件数。</summary>
        Public Property CreatedCount As Integer

        ''' <summary>印刷用一時テーブルへ写した件数。</summary>
        Public Property PrintCount As Integer

        ''' <summary>帳票に印字する行。</summary>
        Public Property Rows As IList(Of ShanghaiNyukaMeisaiReportRow) = New List(Of ShanghaiNyukaMeisaiReportRow)()

    End Class

End Namespace
