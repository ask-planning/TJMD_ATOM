Option Strict On
Option Explicit On
Option Infer Off

Imports System.Data.SqlClient
Imports Atom.Core.Entities
Imports Atom.Core.Logging
Imports Atom.Data.Repositories

Namespace Services

    ''' <summary>
    ''' 帳票の発行にあわせて輸入進捗を進めます。
    ''' </summary>
    ''' <remarks>
    ''' 旧 Access 版の共通処理 `schedule_update` のうち、帳票の発行で呼ぶ分です。
    ''' 設計書 §16 の対応表に従い、配送手配書・入荷スケジュール・入荷明細書から呼びます。
    '''
    ''' **帳票を出し終えてから呼びます。**データの準備と同じトランザクションに入れると、
    ''' 連携ファイルの配置に失敗したときにも進捗だけ進んでしまうためです。
    '''
    ''' **ここで失敗しても帳票の発行は取り消しません。**進捗は業務の目印であり、
    ''' 帳票そのものは既に出ているためです。失敗はログに残して先へ進めます。
    ''' </remarks>
    Public Class ImportProgressService

#Region "定数"
        ''' <summary>ログの発生元名。</summary>
        Private Const LogSource As String = "ImportProgressService"
#End Region

#Region "フィールド"
        ''' <summary>データアクセス。</summary>
        Private ReadOnly _repository As ImportProgressRepository
#End Region

#Region "コンストラクタ"
        ''' <summary>既定の構成で初期化します。</summary>
        Public Sub New()
            Me._repository = New ImportProgressRepository()
        End Sub
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' 進捗を進めます。
        ''' </summary>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <param name="syori">進める処理の種類。</param>
        ''' <param name="syoriYmd">処理日。</param>
        ''' <returns>進められたとき True。すでに日付が入っていたときは False。</returns>
        ''' <remarks>
        ''' **例外を外へ出しません。**呼び出し元は帳票の発行を終えた後であり、
        ''' ここで失敗しても利用者に見せる結果は変わらないためです。
        ''' </remarks>
        Public Function Update(loginCode As String, invoiceNoMain As String,
                               syori As ImportProgressSyori, syoriYmd As Date) As Boolean
            Try
                Using connection As SqlConnection = Database.CreateOpenConnection()
                    Using transaction As SqlTransaction = connection.BeginTransaction()
                        Try
                            Dim updated As Integer = Me._repository.UpdateReportProgress(
                                connection, transaction, invoiceNoMain, loginCode, syori, syoriYmd)
                            transaction.Commit()

                            ' すでに日付が入っていたとき（2 回目以降の発行）
                            If updated = 0 Then
                                Return False
                            End If

                            Logger.WriteInformation(LogSource,
                                "進捗を更新しました。INVOICE " & invoiceNoMain &
                                " 処理NO " & CInt(syori).ToString(Globalization.CultureInfo.InvariantCulture),
                                loginCode)
                            Return True
                        Catch
                            transaction.Rollback()
                            Throw
                        End Try
                    End Using
                End Using
            Catch ex As Exception
                Logger.WriteError(LogSource,
                    "進捗を更新できませんでした。帳票は出力済みです。INVOICE " & invoiceNoMain &
                    " 処理NO " & CInt(syori).ToString(Globalization.CultureInfo.InvariantCulture),
                    ex, loginCode)
                Return False
            End Try
        End Function
#End Region

    End Class
End Namespace
