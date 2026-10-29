Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports Atom.Core.Entities
Imports Atom.Core.Logging
Imports Atom.Core.Services
Imports Atom.Data.Repositories

Namespace Services

    ''' <summary>
    ''' 配送手配書（menuNo=103 §10）の発行を組み立てます。
    ''' </summary>
    ''' <remarks>
    ''' 設計書 §10.3 の手順です。
    ''' 1. 自分のログインコードの印刷用一時データを削除する
    ''' 2. 印刷対象をビューから拾う
    ''' 3. 配送パターンでまとめ、印刷用一時データへ登録する
    ''' 4. 帳票に印字する内容を読み出す
    '''
    ''' 1 〜 4 をひとつのトランザクションで扱います。
    ''' 途中で失敗したとき、作りかけの一時データが残らないようにするためです。
    '''
    ''' **進捗状況の更新（§10.6 処理NO 7）は実装していません。**
    ''' 更新する列と条件が読み取れていないためです（設計書 §18 C-39）。
    ''' **入荷スケジュールの同時出力（§10.5）も実装していません。**様式が未作成です。
    ''' </remarks>
    Public Class ShanghaiHaisoTehaiService

#Region "定数"
        ''' <summary>ログの発生元名。</summary>
        Private Const LogSource As String = "ShanghaiHaisoTehaiService"

        ''' <summary>担当区分のうち上海を表す値。</summary>
        Private Const ShanghaiKubun As Integer = 1

        ''' <summary>元INV_NO が未入力のときの文言。旧 Access 版を踏襲します。</summary>
        Public Const InvoiceNoRequiredMessage As String = "INVOICE NOを入力してください。"

        ''' <summary>担当区分が上海以外のときの文言。旧 Access 版を踏襲します。</summary>
        Public Const NotShanghaiMessage As String = "処理権限がありません。（国内発注用）"

        ''' <summary>印刷用一時テーブルの INVOICE 番号の文字数。</summary>
        Private Const InvoiceNoMaxLength As Integer = 50
#End Region

#Region "フィールド"
        ''' <summary>データアクセス。</summary>
        Private ReadOnly _repository As ShanghaiHaisoTehaiRepository

        ''' <summary>配送パターンごとのまとめ。</summary>
        Private ReadOnly _grouper As ShanghaiHaisoTehaiGrouper
#End Region

#Region "コンストラクタ"
        ''' <summary>既定の構成で初期化します。</summary>
        Public Sub New()
            Me._repository = New ShanghaiHaisoTehaiRepository()
            Me._grouper = New ShanghaiHaisoTehaiGrouper()
        End Sub
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' 配送手配書に印字する内容を作ります。
        ''' </summary>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="shanghaiChotatsu">担当区分。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <param name="executionDate">実行日。配送予定日が未入力のときに使います。</param>
        ''' <returns>発行の結果。</returns>
        Public Function CreateReport(loginCode As String, shanghaiChotatsu As Integer,
                                     invoiceNoMain As String,
                                     executionDate As Date) As ShanghaiHaisoTehaiResult

            Dim result As New ShanghaiHaisoTehaiResult()
            result.ValidationMessage = Me.Validate(shanghaiChotatsu, invoiceNoMain)
            ' 入力に問題があるとき
            If Not String.IsNullOrEmpty(result.ValidationMessage) Then
                Return result
            End If

            Using connection As SqlConnection = Database.CreateOpenConnection()
                Using transaction As SqlTransaction = connection.BeginTransaction()
                    Try
                        Me._repository.DeletePrintTemp(connection, transaction, loginCode)

                        Dim pickups As IList(Of ShanghaiHaisoTehaiPickup) =
                            Me._repository.SelectPickups(connection, transaction, invoiceNoMain)

                        Dim printRows As IList(Of ShanghaiHaisoTehaiPrintRow) =
                            Me._grouper.Group(pickups, executionDate)

                        For Each printRow As ShanghaiHaisoTehaiPrintRow In printRows
                            Me._repository.InsertPrintTemp(connection, transaction, loginCode, printRow)
                        Next

                        result.Rows = Me._repository.SelectReportRows(connection, transaction, loginCode)

                        Me.WarnOnLostRows(printRows, result.Rows, loginCode)
                        Me.WarnOnTooLongInvoiceNo(printRows, loginCode)

                        transaction.Commit()
                    Catch ex As Exception
                        transaction.Rollback()
                        Logger.WriteError(LogSource,
                            "配送手配書の作成に失敗しました。INVOICE " & invoiceNoMain, ex, loginCode)
                        Throw
                    End Try
                End Using
            End Using

            Logger.WriteInformation(LogSource,
                "配送手配書を作成しました。INVOICE " & invoiceNoMain &
                " 件数 " & result.Rows.Count.ToString(Globalization.CultureInfo.InvariantCulture), loginCode)
            Return result
        End Function
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 入力を確かめます。
        ''' </summary>
        ''' <param name="shanghaiChotatsu">担当区分。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <returns>問題があるときはその文言。問題が無いときは空文字。</returns>
        ''' <remarks>設計書 §10.1 の順で確かめます。</remarks>
        Private Function Validate(shanghaiChotatsu As Integer, invoiceNoMain As String) As String
            ' 元INV_NO が未入力のとき
            If String.IsNullOrWhiteSpace(invoiceNoMain) Then
                Return InvoiceNoRequiredMessage
            End If

            ' 担当が上海以外のとき
            If shanghaiChotatsu <> ShanghaiKubun Then
                Return NotShanghaiMessage
            End If

            Return String.Empty
        End Function

        ''' <summary>
        ''' 印刷用の行を作ったのに帳票の行が取れなかったときに、その旨をログへ残します。
        ''' </summary>
        ''' <param name="printRows">印刷用の行。</param>
        ''' <param name="reportRows">帳票の行。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <remarks>
        ''' **`View_Rpt_shanghai_haiso_tehai` は社員マスタと部門マスタを INNER JOIN しています。**
        ''' 発行者が `TJM_master.dbo.tbl_m_login_sys` に居ない、または部門が未設定だと
        ''' **1 行も返りません。**画面には「出力できる明細がありません。」としか出ないため、
        ''' 原因を追えるようここで記録します。
        ''' </remarks>
        Private Sub WarnOnLostRows(printRows As IList(Of ShanghaiHaisoTehaiPrintRow),
                                   reportRows As IList(Of ShanghaiHaisoTehaiReportRow),
                                   loginCode As String)

            ' 印刷用の行が無い、または帳票の行が取れているとき
            If printRows.Count = 0 OrElse reportRows.Count > 0 Then
                Return
            End If

            Logger.WriteWarning(LogSource,
                "印刷用の行を " & printRows.Count.ToString(Globalization.CultureInfo.InvariantCulture) &
                " 件作りましたが、帳票の行が 0 件です。" &
                "発行者が社員マスタに登録されているか、部門コードが設定されているかを確認してください。",
                loginCode)
        End Sub

        ''' <summary>
        ''' INVOICE 番号が長すぎる行があれば、その旨をログへ残します。
        ''' </summary>
        ''' <param name="printRows">印刷用の行。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <remarks>
        ''' 枝番をカンマで連結するため、1 つの配送パターンに枝番が多いと長くなります。
        ''' **一時テーブルの列は 50 文字**で、超えた分は ADO.NET が黙って切ります。
        ''' 帳票に出る番号が足りない状態になるため、気付けるよう記録します。
        ''' </remarks>
        Private Sub WarnOnTooLongInvoiceNo(printRows As IList(Of ShanghaiHaisoTehaiPrintRow), loginCode As String)
            For Each printRow As ShanghaiHaisoTehaiPrintRow In printRows
                ' 列の長さに収まっているとき
                If printRow.InvoiceNo.Length <= InvoiceNoMaxLength Then
                    Continue For
                End If

                Logger.WriteWarning(LogSource,
                    "INVOICE 番号が " & InvoiceNoMaxLength.ToString(Globalization.CultureInfo.InvariantCulture) &
                    " 文字を超えたため、途中までしか印字されません。" & printRow.InvoiceNo, loginCode)
            Next
        End Sub
#End Region

    End Class
End Namespace
