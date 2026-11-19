Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports Atom.Core.Entities
Imports Atom.Core.Exceptions
Imports Atom.Core.Logging
Imports Atom.Data.Repositories

Namespace Services

    ''' <summary>
    ''' 輸入経費コード登録（menuNo=907）の業務処理です。
    ''' </summary>
    ''' <remarks>
    ''' **旧 Access と同じ挙動にする**方針です（設計書「実装方針の決定」）。
    ''' 旧は入力チェックを一切持たず、DB の制約だけで弾いていました。
    ''' そのためここでも、DB の制約と同じ条件だけを確かめ、旧より厳しくしません。
    '''
    ''' 基本方針で必須の項目（楽観ロック・サーバー側の再検証）だけは追加しています。
    ''' </remarks>
    Public Class ImportKeihiCodeService

#Region "定数"
        ''' <summary>ログの発生元名。</summary>
        Private Const LogSource As String = "ImportKeihiCodeService"

        ''' <summary>伝票コードの長さ。</summary>
        Private Const KeihiCodeLength As Integer = 4

        ''' <summary>経費名称の長さ。</summary>
        Private Const KeihiCodeTextLength As Integer = 50

        ''' <summary>REF の長さ。</summary>
        Private Const KeihiCodeRefLength As Integer = 50

        ''' <summary>区分の長さ。</summary>
        Private Const KeihiKbnLength As Integer = 3

        ''' <summary>伝票コードが未入力のときの文言。</summary>
        Public Const KeihiCodeRequiredMessage As String = "伝票コードを入力してください。"

        ''' <summary>伝票コードが長すぎるときの文言。</summary>
        Public Const KeihiCodeTooLongMessage As String = "伝票コードは4文字以内で入力してください。"

        ''' <summary>伝票コードが重複しているときの文言。</summary>
        Public Const KeihiCodeDuplicateMessage As String = "この伝票コードは既に登録されています。"

        ''' <summary>桁数を超えたときの文言。</summary>
        Public Const TooLongMessage As String = "入力できる文字数を超えています。"

        ''' <summary>保険対象が数値でないときの文言。</summary>
        Public Const HokenTaisyoFlgInvalidMessage As String = "保険対象は数値で入力してください。"

        ''' <summary>他の利用者が先に更新していたときの文言。</summary>
        Public Const ConcurrencyMessage As String =
            "他の利用者が更新しました。画面を読み直してから再度操作してください。"
#End Region

#Region "フィールド"
        ''' <summary>データアクセス。</summary>
        Private ReadOnly _repository As ImportKeihiCodeRepository
#End Region

#Region "コンストラクタ"
        ''' <summary>既定の構成で初期化します。</summary>
        Public Sub New()
            Me._repository = New ImportKeihiCodeRepository()
        End Sub
#End Region

#Region "参照系"
        ''' <summary>
        ''' 経費コードを全件取得します。
        ''' </summary>
        ''' <returns>伝票コードの昇順の一覧。</returns>
        Public Function SelectAll() As IList(Of ImportKeihiCode)
            Using connection As SqlConnection = Database.CreateOpenConnection()
                Return Me._repository.SelectAll(connection, Nothing)
            End Using
        End Function

        ''' <summary>
        ''' 経費区分の選択肢を取得します。
        ''' </summary>
        ''' <returns>選択肢の一覧。</returns>
        Public Function SelectKbnList() As IList(Of ImportKeihiKbn)
            Using connection As SqlConnection = Database.CreateOpenConnection()
                Return Me._repository.SelectKbnList(connection, Nothing)
            End Using
        End Function
#End Region

#Region "更新系"
        ''' <summary>
        ''' 経費コードを 1 件登録します。
        ''' </summary>
        ''' <param name="entity">登録する内容。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>入力に問題があるときはその文言。問題が無いときは空文字。</returns>
        Public Function Insert(entity As ImportKeihiCode, loginCode As String) As String
            Dim message As String = Validate(entity)
            ' 入力に問題があるとき
            If Not String.IsNullOrEmpty(message) Then
                Return message
            End If

            Using connection As SqlConnection = Database.CreateOpenConnection()
                Using transaction As SqlTransaction = connection.BeginTransaction()
                    Try
                        ' 同じ伝票コードが既にあるとき
                        If Me._repository.ExistsDuplicate(connection, transaction,
                                                          entity.KeihiCode, String.Empty) Then
                            transaction.Rollback()
                            Return KeihiCodeDuplicateMessage
                        End If

                        Me._repository.Insert(connection, transaction, entity, loginCode)
                        transaction.Commit()
                    Catch ex As Exception
                        transaction.Rollback()
                        Logger.WriteError(LogSource,
                            "経費コードの登録に失敗しました。コード " & entity.KeihiCode, ex, loginCode)
                        Throw
                    End Try
                End Using
            End Using

            Logger.WriteInformation(LogSource, "経費コードを登録しました。コード " & entity.KeihiCode, loginCode)
            Return String.Empty
        End Function

        ''' <summary>
        ''' 経費コードを 1 件更新します。
        ''' </summary>
        ''' <param name="entity">更新後の内容。</param>
        ''' <param name="originalKeihiCode">変更前の伝票コード。</param>
        ''' <param name="originalUpdateYmd">画面に表示した時点の更新日。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>入力に問題があるときはその文言。問題が無いときは空文字。</returns>
        ''' <remarks>
        ''' **更新日と更新者を書き換えるかどうかは、変わった項目で決まります。**
        ''' 伝票コード・経費名称・区分のいずれかが変わったときだけ書き換えます。
        ''' REF と保険対象だけの変更では書き換えません（設計書 C-11 / 既知の不具合 K-1）。
        ''' </remarks>
        Public Function Update(entity As ImportKeihiCode, originalKeihiCode As String,
                               originalUpdateYmd As Date?, loginCode As String) As String

            Dim message As String = Validate(entity)
            ' 入力に問題があるとき
            If Not String.IsNullOrEmpty(message) Then
                Return message
            End If

            Dim original As New ImportKeihiCode()
            original.KeihiCode = originalKeihiCode
            original.UpdateYmd = originalUpdateYmd

            Using connection As SqlConnection = Database.CreateOpenConnection()
                Using transaction As SqlTransaction = connection.BeginTransaction()
                    Try
                        ' 伝票コードを変えるとき、変更先が既に使われていないか確かめる
                        If Me._repository.ExistsDuplicate(connection, transaction,
                                                          entity.KeihiCode, originalKeihiCode) Then
                            transaction.Rollback()
                            Return KeihiCodeDuplicateMessage
                        End If

                        Dim stored As ImportKeihiCode =
                            Me._repository.SelectByCode(connection, transaction, originalKeihiCode)
                        ' 既に消えているとき
                        If stored Is Nothing Then
                            transaction.Rollback()
                            Return ConcurrencyMessage
                        End If

                        Dim updateStamp As Boolean = NeedsStamp(entity, stored)
                        Dim count As Integer = Me._repository.Update(connection, transaction,
                                                                     entity, original, loginCode, updateStamp)
                        ' 他の利用者が先に更新していたとき
                        If count = 0 Then
                            transaction.Rollback()
                            Return ConcurrencyMessage
                        End If

                        transaction.Commit()
                    Catch ex As ConcurrencyException
                        transaction.Rollback()
                        Throw
                    Catch ex As Exception
                        transaction.Rollback()
                        Logger.WriteError(LogSource,
                            "経費コードの更新に失敗しました。コード " & originalKeihiCode, ex, loginCode)
                        Throw
                    End Try
                End Using
            End Using

            Logger.WriteInformation(LogSource, "経費コードを更新しました。コード " & originalKeihiCode, loginCode)
            Return String.Empty
        End Function

        ''' <summary>
        ''' 経費コードを 1 件削除します。
        ''' </summary>
        ''' <param name="keihiCode">削除する伝票コード。</param>
        ''' <param name="originalUpdateYmd">画面に表示した時点の更新日。</param>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <returns>問題があるときはその文言。問題が無いときは空文字。</returns>
        ''' <remarks>
        ''' **物理削除です。**明細で使用中かどうかは見ません。
        ''' 旧にも止める仕組みが無かったためです（設計書 C-7 / 既知の不具合 K-4）。
        ''' </remarks>
        Public Function Delete(keihiCode As String, originalUpdateYmd As Date?, loginCode As String) As String
            Dim original As New ImportKeihiCode()
            original.KeihiCode = keihiCode
            original.UpdateYmd = originalUpdateYmd

            Using connection As SqlConnection = Database.CreateOpenConnection()
                Using transaction As SqlTransaction = connection.BeginTransaction()
                    Try
                        Dim count As Integer = Me._repository.Delete(connection, transaction, original)
                        ' 他の利用者が先に更新または削除していたとき
                        If count = 0 Then
                            transaction.Rollback()
                            Return ConcurrencyMessage
                        End If

                        transaction.Commit()
                    Catch ex As Exception
                        transaction.Rollback()
                        Logger.WriteError(LogSource,
                            "経費コードの削除に失敗しました。コード " & keihiCode, ex, loginCode)
                        Throw
                    End Try
                End Using
            End Using

            Logger.WriteInformation(LogSource, "経費コードを削除しました。コード " & keihiCode, loginCode)
            Return String.Empty
        End Function
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 入力を確かめます。
        ''' </summary>
        ''' <param name="entity">確かめる内容。</param>
        ''' <returns>問題があるときはその文言。問題が無いときは空文字。</returns>
        ''' <remarks>
        ''' **条件は DB の制約と同じにしています。**旧 Access に入力チェックが無かったため、
        ''' ここで独自の条件を足すと旧より厳しくなります（設計書 C-3 / C-24 / C-25）。
        ''' 伝票コードが数字かどうかは見ません。経費名称・区分・REF は未入力でも保存できます。
        ''' </remarks>
        Private Shared Function Validate(entity As ImportKeihiCode) As String
            ' 伝票コードが未入力のとき（DB が NOT NULL）
            If String.IsNullOrWhiteSpace(entity.KeihiCode) Then
                Return KeihiCodeRequiredMessage
            End If

            ' 伝票コードが長すぎるとき
            If entity.KeihiCode.Length > KeihiCodeLength Then
                Return KeihiCodeTooLongMessage
            End If

            ' ほかの項目が長すぎるとき
            If entity.KeihiCodeText.Length > KeihiCodeTextLength OrElse
               entity.KeihiCodeRef.Length > KeihiCodeRefLength OrElse
               entity.KeihiKbn.Length > KeihiKbnLength Then
                Return TooLongMessage
            End If

            Return String.Empty
        End Function

        ''' <summary>
        ''' 更新日と更新者を書き換えるかどうかを決めます。
        ''' </summary>
        ''' <param name="entity">更新後の内容。</param>
        ''' <param name="stored">更新前の内容。</param>
        ''' <returns>書き換えるとき True。</returns>
        ''' <remarks>
        ''' 旧 Access は伝票コード・経費名称・区分の 3 つにしか `AfterUpdate` を書いておらず、
        ''' REF と保険対象を直しても更新日が変わりませんでした。
        ''' **旧と同じ挙動にするため、ここでも同じ 3 つだけを見ます**（既知の不具合 K-1）。
        ''' </remarks>
        Private Shared Function NeedsStamp(entity As ImportKeihiCode, stored As ImportKeihiCode) As Boolean
            Return Not String.Equals(entity.KeihiCode, stored.KeihiCode, StringComparison.Ordinal) OrElse
                   Not String.Equals(entity.KeihiCodeText, stored.KeihiCodeText, StringComparison.Ordinal) OrElse
                   Not String.Equals(entity.KeihiKbn, stored.KeihiKbn, StringComparison.Ordinal)
        End Function
#End Region

    End Class
End Namespace
