Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System.Text
Imports Atom.Core.Entities
Imports Atom.Core.Logging
Imports Atom.Core.Services
Imports Atom.Data.Repositories

Namespace Services

    ''' <summary>
    ''' INVOICEヘッダー入力（menuNo=103）の一連の手順を制御します。
    ''' </summary>
    ''' <remarks>
    ''' 仕様は docs/specs/103_INVOICEヘッダー入力.md を正とします。
    ''' 現時点では表示までを担います。登録と帳票出力は未実装です。
    ''' </remarks>
    Public Class ShanghaiInvoiceHeaderService

#Region "定数"
        ''' <summary>ログの発生元名。</summary>
        Private Const LogSource As String = "ShanghaiInvoiceHeaderService"

        ''' <summary>
        ''' BL DATE と出荷方法が食い違っているときの文言です。
        ''' </summary>
        ''' <remarks>旧 Access 版の文言をそのまま踏襲します。</remarks>
        Public Const ConflictMessage As String =
            "同じINVOICE_NOで、BL_DATE、出荷方法が異なるPACKINGがあります。PACKINGデータを修正してください。"

        ''' <summary>元INV_NO が未選択のときの文言。</summary>
        Public Const InvoiceNoRequiredMessage As String = "INVOICE_NOを入力してください。"

        ''' <summary>船名が未入力のときの文言。</summary>
        Public Const VesselRequiredMessage As String = "船名を入力してください。"

        ''' <summary>乙仲が未入力のときの文言。</summary>
        Public Const OtsunakaRequiredMessage As String = "乙仲を入力してください。"

        ''' <summary>登録が終わったときの文言。</summary>
        Public Const RegisteredMessage As String = "登録完了しました。"

        ''' <summary>明細更新で元INV_NO が未選択のときの文言。</summary>
        Public Const InvoiceNoRequiredForUpdateMessage As String = "INVOICE NOを入力してください。"

        ''' <summary>同じサイズのコンテナが複数あるときの文言。</summary>
        Public Const ContainerDuplicatedMessage As String =
            "同じサイズのコンテナが複数ある場合、ここでの更新は出来ません。"

        ''' <summary>明細更新が終わったときの文言。</summary>
        Public Const ContainerUpdatedMessage As String = "更新完了しました。"

        ''' <summary>印刷用 INVOICE_NO の区切り文字。</summary>
        Private Const PrintNoSeparator As String = ","

        ''' <summary>POEM の更新を行うかを保持する設定キー。</summary>
        Private Const PoemUpdateEnabledKey As String = "Poem.UpdateEnabled"

        ''' <summary>巻芯入出庫を更新するストアド名を保持する設定キー。</summary>
        Private Const PoemProcedureNameKey As String = "Poem.MakishinProcedureName"

        ''' <summary>
        ''' ストアド名として認めるかたち。英数字とアンダースコアのみ。
        ''' </summary>
        ''' <remarks>
        ''' 設定値をそのままコマンド文へ組み込むため、識別子として妥当かを必ず確かめます。
        ''' </remarks>
        Private Const ProcedureNamePattern As String = "^[A-Za-z0-9_]+$"

        ''' <summary>POEM の更新を行わなかったときの文言。</summary>
        Public Const PoemSkippedMessage As String = "（POEM の巻芯入出庫は更新していません。）"
#End Region

#Region "フィールド"
        ''' <summary>データアクセス。</summary>
        Private ReadOnly _repository As ShanghaiInvoiceHeaderRepository

        ''' <summary>INVOICE 番号の編集。</summary>
        Private ReadOnly _invoiceNoFormatter As ShanghaiInvoiceNoFormatter
#End Region

#Region "コンストラクタ"
        ''' <summary>既定の状態で初期化します。</summary>
        Public Sub New()
            Me._repository = New ShanghaiInvoiceHeaderRepository()
            Me._invoiceNoFormatter = New ShanghaiInvoiceNoFormatter()
        End Sub
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' 画面の選択肢を取得します。
        ''' </summary>
        ''' <returns>選択肢の一式。</returns>
        Public Function LoadOptions() As ShanghaiInvoiceHeaderOptions
            Using connection As SqlConnection = Database.CreateOpenConnection()
                Dim options As New ShanghaiInvoiceHeaderOptions()
                options.InvoiceNoMains = Me._repository.SelectInvoiceNoMainOptions(connection, Nothing)
                options.Vessels = Me._repository.SelectVesselOptions(connection, Nothing)
                options.SyukkaHohos = Me._repository.SelectSyukkaHohoOptions(connection, Nothing)
                options.Otsunakas = Me._repository.SelectOtsunakaOptions(connection, Nothing)
                options.HaisoTypes = Me._repository.SelectHaisoTypeOptions(connection, Nothing)
                options.ContainerTypes = Me._repository.SelectContainerTypeOptions(connection, Nothing)
                Return options
            End Using
        End Function

        ''' <summary>
        ''' 元INV_NO を選んだときの表示内容を取得します。
        ''' </summary>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="shanghaiChotatsu">担当区分。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <returns>表示に必要な一式。</returns>
        ''' <remarks>
        ''' 設計書 §6.2 の手順に沿っています。
        ''' 1. 登録済みかを判定する
        ''' 2. BL DATE と出荷方法が食い違っていれば中断する
        ''' 3. ヘッダを読み込む（未登録のときは Nothing）
        ''' 4. 作業用の一時データを作り直す
        ''' 5. 明細から合計を取得する
        ''' 6. 明細を読み込む
        '''
        ''' 一時データの作り直しと読み込みを 1 つのトランザクションにまとめています。
        ''' 途中で失敗したときに、作りかけの一時データが残らないようにするためです。
        ''' </remarks>
        Public Function LoadView(loginCode As String, shanghaiChotatsu As Integer,
                                 invoiceNoMain As String) As ShanghaiInvoiceHeaderView

            Using connection As SqlConnection = Database.CreateOpenConnection()
                Using transaction As SqlTransaction = connection.BeginTransaction()
                    Try
                        Dim view As New ShanghaiInvoiceHeaderView()

                        view.Existence = Me._repository.SelectExistence(connection, transaction, invoiceNoMain)

                        ' BL DATE と出荷方法が食い違っているとき
                        If view.Existence.HasConflict Then
                            transaction.Commit()
                            Return view
                        End If

                        ' ヘッダが登録済みのとき
                        If Not view.Existence.IsNewHeader Then
                            view.Header = Me._repository.SelectHeader(connection, transaction, invoiceNoMain)
                        End If

                        ' ヘッダサブが未登録のときは PACKING から新規に作る
                        Dim isNew As Boolean = (view.Existence.SubCount = 0)
                        Me._repository.CreateDetailTemp(connection, transaction, loginCode, shanghaiChotatsu, invoiceNoMain, isNew)

                        view.Total = Me._repository.SelectTotal(connection, transaction, invoiceNoMain)
                        view.Details = Me._repository.SelectDetails(connection, transaction, loginCode)

                        transaction.Commit()
                        Return view
                    Catch ex As Exception
                        transaction.Rollback()
                        Logger.WriteError(LogSource, "表示内容の取得に失敗しました。INVOICE " & invoiceNoMain, ex, loginCode)
                        Throw
                    End Try
                End Using
            End Using
        End Function

        ''' <summary>
        ''' ヘッダと明細を登録します。
        ''' </summary>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="header">画面のヘッダ内容。</param>
        ''' <param name="details">画面の明細内容。</param>
        ''' <returns>登録の結果。</returns>
        ''' <remarks>
        ''' 設計書 §8 の手順です。
        ''' 1. 入力チェック（§8.1）
        ''' 2. 画面の明細を作業用の一時データへ反映する
        ''' 3. ヘッダを登録または更新する（§8.2）
        ''' 4. 一時データを本体の明細へ反映する（§8.3）
        ''' 5. 印刷用 INVOICE_NO を組み立てる（§8.4）
        ''' 6. 支給包装材の納入期日を更新する（§8.6）
        ''' 7. POEM の巻芯入出庫実績を更新する（§8.6）。**`Poem.UpdateEnabled` が true のときだけ**
        '''
        ''' 2 〜 7 をひとつのトランザクションで扱います。
        ''' 途中で失敗したとき、ヘッダだけ更新されて明細が古いままになるのを防ぐためです。
        ''' 7 は別データベース（POEM）を更新しますが、同一インスタンス内のため同じトランザクションで扱えます。
        '''
        ''' **外部システムへの通知（§8.5）だけが未実装です。**要否は確認事項 C-5 で確認します。
        ''' </remarks>
        Public Function Register(loginCode As String, header As ShanghaiInvoiceHeader,
                                 details As IList(Of ShanghaiInvoiceHeaderSub)) As ShanghaiInvoiceRegisterResult

            Dim result As New ShanghaiInvoiceRegisterResult()
            result.ValidationMessage = Me.Validate(header)
            ' 入力に問題があるとき
            If Not result.IsSuccess Then
                Return result
            End If

            Using connection As SqlConnection = Database.CreateOpenConnection()
                Using transaction As SqlTransaction = connection.BeginTransaction()
                    Try
                        ' 画面の明細を一時データへ反映する
                        For Each detail As ShanghaiInvoiceHeaderSub In details
                            Me._repository.UpdateDetailTemp(connection, transaction, loginCode, detail)
                        Next

                        Dim existence As ShanghaiInvoiceExistence =
                            Me._repository.SelectExistence(connection, transaction, header.InvoiceNoMain)

                        ' ヘッダが未登録のとき
                        If existence.IsNewHeader Then
                            Me._repository.InsertHeader(connection, transaction, header, loginCode)
                        Else
                            Me._repository.UpdateHeader(connection, transaction, header, loginCode)
                        End If

                        Me._repository.ApplyDetailTemp(connection, transaction, loginCode)

                        result.HasHaisoYotei = Me.UpdateInvoiceNoPrint(connection, transaction, header.InvoiceNoMain)

                        ' 後続処理（§8.6）。明細を本体へ反映した後に実行する
                        Me._repository.UpdateShikyuHosozaiNouki(connection, transaction, header.InvoiceNoMain)

                        Dim poemProcedureName As String = GetPoemProcedureName()

                        ' POEM を更新する設定で、ストアド名も正しく設定されているとき
                        If IsPoemUpdateEnabled() AndAlso poemProcedureName.Length > 0 Then
                            Me._repository.UpdatePoemMakishin(connection, transaction, loginCode,
                                                              header.InvoiceNoMain, poemProcedureName)
                        Else
                            result.IsPoemSkipped = True
                            Logger.WriteWarning(LogSource,
                                "設定により POEM の更新を行いませんでした。INVOICE " & header.InvoiceNoMain, loginCode)
                        End If

                        transaction.Commit()
                    Catch ex As Exception
                        transaction.Rollback()
                        Logger.WriteError(LogSource, "登録に失敗しました。INVOICE " & header.InvoiceNoMain, ex, loginCode)
                        Throw
                    End Try
                End Using
            End Using

            Logger.WriteInformation(LogSource, "登録しました。INVOICE " & header.InvoiceNoMain, loginCode)
            Return result
        End Function

        ''' <summary>
        ''' 明細のコンテナ番号を一括更新します。
        ''' </summary>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <param name="details">画面の明細内容。</param>
        ''' <returns>利用者に見せる文言。</returns>
        ''' <remarks>
        ''' 設計書 §9 の手順です。
        ''' 1. 元INV_NO が選ばれているか確かめる
        ''' 2. 画面の明細を作業用の一時データへ反映する
        ''' 3. 同じサイズのコンテナが複数あるか判定する。複数なら更新しない
        ''' 4. 明細のコンテナ番号とシール番号を一括更新する
        '''
        ''' **1 コンテナサイズ・1 コンテナNO のときだけ使える機能**です。
        ''' 2 〜 4 をひとつのトランザクションで扱います。
        ''' </remarks>
        Public Function UpdateContainerNo(loginCode As String, invoiceNoMain As String,
                                          details As IList(Of ShanghaiInvoiceHeaderSub)) As String

            ' 元INV_NO が未選択のとき
            If String.IsNullOrEmpty(invoiceNoMain) Then
                Return InvoiceNoRequiredForUpdateMessage
            End If

            Using connection As SqlConnection = Database.CreateOpenConnection()
                Using transaction As SqlTransaction = connection.BeginTransaction()
                    Try
                        ' 画面の明細を一時データへ反映する
                        For Each detail As ShanghaiInvoiceHeaderSub In details
                            Me._repository.UpdateDetailTemp(connection, transaction, loginCode, detail)
                        Next

                        Dim duplicated As Integer =
                            Me._repository.SelectDuplicateContainerCount(connection, transaction, loginCode)

                        ' 同じサイズのコンテナが複数あるとき
                        If duplicated > 1 Then
                            transaction.Commit()
                            Return ContainerDuplicatedMessage
                        End If

                        Me._repository.UpdateContainerNo(connection, transaction, loginCode)
                        transaction.Commit()
                    Catch ex As Exception
                        transaction.Rollback()
                        Logger.WriteError(LogSource, "明細更新に失敗しました。INVOICE " & invoiceNoMain, ex, loginCode)
                        Throw
                    End Try
                End Using
            End Using

            Logger.WriteInformation(LogSource, "明細のコンテナ番号を更新しました。INVOICE " & invoiceNoMain, loginCode)
            Return ContainerUpdatedMessage
        End Function

        ''' <summary>
        ''' 明細の 1 行を作業用の一時データへ反映します。
        ''' </summary>
        ''' <param name="loginCode">利用者のログインコード。</param>
        ''' <param name="detail">更新する内容。</param>
        ''' <returns>更新した件数。</returns>
        Public Function UpdateDetail(loginCode As String, detail As ShanghaiInvoiceHeaderSub) As Integer
            Using connection As SqlConnection = Database.CreateOpenConnection()
                Return Me._repository.UpdateDetailTemp(connection, Nothing, loginCode, detail)
            End Using
        End Function
#End Region

#Region "内部処理"
        ''' <summary>
        ''' POEM の更新を行う設定かどうかを返します。
        ''' </summary>
        ''' <returns>行うときは True。</returns>
        ''' <remarks>
        ''' `usp_poem_makishin_nyu_syukko_update` は、ストアドの中で `POEM.dbo.…` と
        ''' データベース名を直接指定しています。**どの接続先から呼んでも本番の POEM を書き換えます。**
        ''' 先頭で対象の入庫分を削除するため、検証環境から誤って実行すると本番データが失われます。
        ''' そのため既定を「行わない」とし、`Web.config` で明示したときだけ実行します。
        ''' </remarks>
        Private Shared Function IsPoemUpdateEnabled() As Boolean
            Dim value As String = ConfigurationManager.AppSettings(PoemUpdateEnabledKey)
            ' 設定が無いときは安全側に倒す
            If String.IsNullOrEmpty(value) Then
                Return False
            End If

            Dim enabled As Boolean
            ' 真偽値として読めないとき
            If Not Boolean.TryParse(value, enabled) Then
                Return False
            End If
            Return enabled
        End Function

        ''' <summary>
        ''' 巻芯入出庫を更新するストアド名を取得します。
        ''' </summary>
        ''' <returns>ストアド名。未設定または不正なかたちのときは空文字。</returns>
        ''' <remarks>
        ''' ストアドは中でデータベース名を直接指定しており、接続先に追従しません。
        ''' 環境ごとに向き先の違うストアドを用意し、ここで切り替えます。
        ''' 設定値をコマンド文へ組み込むため、**英数字とアンダースコアだけ**を認めます。
        ''' </remarks>
        Private Shared Function GetPoemProcedureName() As String
            Dim value As String = ConfigurationManager.AppSettings(PoemProcedureNameKey)
            ' 設定が無いとき
            If String.IsNullOrWhiteSpace(value) Then
                Return String.Empty
            End If

            Dim trimmed As String = value.Trim()
            ' 識別子として認められないかたちのとき
            If Not Regex.IsMatch(trimmed, ProcedureNamePattern) Then
                Logger.WriteWarning(LogSource,
                    "設定 " & PoemProcedureNameKey & " の値がストアド名として正しくありません。", String.Empty)
                Return String.Empty
            End If

            Return trimmed
        End Function

        ''' <summary>
        ''' 入力内容を確かめます。
        ''' </summary>
        ''' <param name="header">画面のヘッダ内容。</param>
        ''' <returns>問題があるときはその文言。無いときは空文字。</returns>
        ''' <remarks>
        ''' 設計書 §8.1。B/L DATE・出荷方法・カートン数・返却パレットは必須ではありません。
        ''' </remarks>
        Private Function Validate(header As ShanghaiInvoiceHeader) As String
            ' 元INV_NO が未選択のとき
            If header Is Nothing OrElse String.IsNullOrEmpty(header.InvoiceNoMain) Then
                Return InvoiceNoRequiredMessage
            End If
            ' 船名が未入力のとき
            If String.IsNullOrEmpty(header.Vessel) Then
                Return VesselRequiredMessage
            End If
            ' 乙仲が未選択のとき
            If Not header.OtsunakaNo.HasValue Then
                Return OtsunakaRequiredMessage
            End If
            Return String.Empty
        End Function

        ''' <summary>
        ''' 印刷用 INVOICE_NO を組み立てて更新します。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <returns>配送予定日が 1 件でも入っていたときは True。</returns>
        ''' <remarks>
        ''' 設計書 §8.4。コンテナ番号が変わるまでを 1 グループとし、枝番をカンマで連結します。
        ''' 一覧はコンテナ区分・コンテナNO・INVOICE_NO の順に並んでいることが前提です。
        ''' </remarks>
        Private Function UpdateInvoiceNoPrint(connection As SqlConnection, transaction As SqlTransaction,
                                              invoiceNoMain As String) As Boolean

            Dim targets As IList(Of ShanghaiInvoicePrintTarget) =
                Me._repository.SelectPrintTargets(connection, transaction, invoiceNoMain)

            Dim hasHaisoYotei As Boolean = False
            Dim currentContainerNo As String = String.Empty
            Dim currentInvoiceNo As String = String.Empty
            Dim builder As New StringBuilder()

            For Each target As ShanghaiInvoicePrintTarget In targets
                ' 配送予定日の有無を控える（対象外の行も数える。旧 Access と同じ）
                If target.HaisoYoteiYmd.HasValue Then
                    hasHaisoYotei = True
                End If

                ' 組み立ての対象外のとき
                If Not target.IsPrintTarget Then
                    Continue For
                End If

                Dim branchNo As String = Me._invoiceNoFormatter.GetBranchNo(target.InvoiceNo)

                ' コンテナ番号が変わったとき（ここまでを 1 グループとして確定する）
                If Not String.Equals(currentContainerNo, target.ContainerNo, StringComparison.Ordinal) Then
                    Me.FlushInvoiceNoPrint(connection, transaction, invoiceNoMain, currentContainerNo, builder)
                    currentContainerNo = target.ContainerNo
                    currentInvoiceNo = target.InvoiceNo
                    builder.Append(branchNo)
                    Continue For
                End If

                ' 同じ INVOICE 番号が続いたとき（重複して足さない）
                If String.Equals(currentInvoiceNo, target.InvoiceNo, StringComparison.Ordinal) Then
                    Continue For
                End If

                currentInvoiceNo = target.InvoiceNo
                builder.Append(PrintNoSeparator)
                builder.Append(branchNo)
            Next

            Me.FlushInvoiceNoPrint(connection, transaction, invoiceNoMain, currentContainerNo, builder)
            Return hasHaisoYotei
        End Function

        ''' <summary>
        ''' 組み立て中の印刷用 INVOICE_NO を書き込み、組み立てを始めからやり直せる状態にします。
        ''' </summary>
        ''' <param name="connection">開いている接続。</param>
        ''' <param name="transaction">トランザクション。</param>
        ''' <param name="invoiceNoMain">INVOICE 番号（メイン）。</param>
        ''' <param name="containerNo">コンテナ番号。空のときは何もしません。</param>
        ''' <param name="builder">組み立て中の文字列。書き込み後に空にします。</param>
        Private Sub FlushInvoiceNoPrint(connection As SqlConnection, transaction As SqlTransaction,
                                        invoiceNoMain As String, containerNo As String, builder As StringBuilder)
            ' 対象が無いとき
            If String.IsNullOrEmpty(containerNo) OrElse builder.Length = 0 Then
                builder.Length = 0
                Return
            End If

            Me._repository.UpdateInvoiceNoPrint(connection, transaction, invoiceNoMain, containerNo, builder.ToString())
            builder.Length = 0
        End Sub
#End Region

    End Class

    ''' <summary>
    ''' INVOICEヘッダー入力画面で使う選択肢の一式です。
    ''' </summary>
    Public NotInheritable Class ShanghaiInvoiceHeaderOptions

#Region "コンストラクタ"
        ''' <summary>既定の状態で初期化します。</summary>
        Public Sub New()
            Me.InvoiceNoMains = New List(Of String)()
            Me.Vessels = New List(Of String)()
            Me.SyukkaHohos = New List(Of CodeOption)()
            Me.Otsunakas = New List(Of CodeOption)()
            Me.HaisoTypes = New List(Of CodeOption)()
            Me.ContainerTypes = New List(Of CodeOption)()
        End Sub
#End Region

#Region "プロパティ"
        ''' <summary>元INV_NO の選択肢。B/L DATE の新しい順に並びます。</summary>
        Public Property InvoiceNoMains As IList(Of String)

        ''' <summary>船名／AIR便名の選択肢。</summary>
        Public Property Vessels As IList(Of String)

        ''' <summary>出荷方法の選択肢。</summary>
        Public Property SyukkaHohos As IList(Of CodeOption)

        ''' <summary>乙仲の選択肢。</summary>
        Public Property Otsunakas As IList(Of CodeOption)

        ''' <summary>配送タイプの選択肢。</summary>
        Public Property HaisoTypes As IList(Of CodeOption)

        ''' <summary>コンテナタイプの選択肢。</summary>
        Public Property ContainerTypes As IList(Of CodeOption)
#End Region

    End Class
End Namespace
