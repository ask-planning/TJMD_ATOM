Option Strict On
Option Explicit On
Option Infer Off

Namespace Entities

    ''' <summary>
    ''' 取込データのチェック結果の区分です。
    ''' </summary>
    Public Enum ImportCheckLevel

        ''' <summary>エラーも警告も無い。</summary>
        None = 0

        ''' <summary>取込不可のエラーがある。</summary>
        [Error] = 1

        ''' <summary>利用者の確認があれば取込できる警告がある。</summary>
        Warning = 2
    End Enum

    ''' <summary>
    ''' 取込データのチェック結果です。各チェックストアドが返す件数を集約します。
    ''' </summary>
    Public NotInheritable Class ImportCheckResult

#Region "プロパティ"
        ''' <summary>エラー（err_status = 1）の件数。</summary>
        Public Property ErrorCount As Integer

        ''' <summary>警告（err_status = 2）の件数。</summary>
        Public Property WarningCount As Integer

        ''' <summary>
        ''' 1 つの PO が複数 INVOICE に分かれているかどうか。
        ''' True のときは他のチェックを行わず取込を中止します。
        ''' </summary>
        Public Property IsPoDuplicated As Boolean

        ''' <summary>
        ''' チェック結果の区分。エラーがあればエラー、無ければ警告の有無で判断します。
        ''' </summary>
        Public ReadOnly Property Level As ImportCheckLevel
            Get
                ' エラーがあるとき
                If Me.ErrorCount > 0 Then
                    Return ImportCheckLevel.Error
                End If
                ' 警告があるとき
                If Me.WarningCount > 0 Then
                    Return ImportCheckLevel.Warning
                End If
                ' どちらも無いとき
                Return ImportCheckLevel.None
            End Get
        End Property
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' 別のチェック結果を取り込みます。エラーと警告の件数を足し合わせます。
        ''' </summary>
        ''' <param name="other">取り込むチェック結果。</param>
        Public Sub Merge(other As ImportCheckResult)
            ' 取り込む対象が無いとき
            If other Is Nothing Then
                Return
            End If
            Me.ErrorCount += other.ErrorCount
            Me.WarningCount += other.WarningCount
        End Sub
#End Region

    End Class

    ''' <summary>
    ''' 一時データを読み取って得た取込前の情報です。
    ''' </summary>
    Public NotInheritable Class ShanghaiImportAnalysis

#Region "プロパティ"
        ''' <summary>画面に表示する REF_NO。</summary>
        Public Property RefNo As String = String.Empty

        ''' <summary>INVOICE 番号（メイン）。ヘッダの作成と進捗更新に使います。</summary>
        Public Property InvoiceNoMain As String = String.Empty

        ''' <summary>翌月以降の BL DATE を持つ件数。0 以外のとき利用者へ確認します。</summary>
        Public Property FutureBlDateCount As Integer

        ''' <summary>登録済みの件数。</summary>
        Public Property Registered As ShanghaiImportCounts

        ''' <summary>一時データが存在するかどうか。</summary>
        Public Property HasTempData As Boolean
#End Region

#Region "コンストラクタ"
        ''' <summary>既定の状態で初期化します。</summary>
        Public Sub New()
            Me.Registered = New ShanghaiImportCounts()
        End Sub
#End Region

    End Class

    ''' <summary>
    ''' INVOICE と PACKING の件数の組です。登録済み件数や登録結果の件数を保持します。
    ''' </summary>
    Public NotInheritable Class ShanghaiImportCounts

#Region "プロパティ"
        ''' <summary>INVOICE の件数。</summary>
        Public Property InvoiceCount As Integer

        ''' <summary>PACKING の件数。</summary>
        Public Property PackingCount As Integer

        ''' <summary>どちらかに件数があるかどうか。</summary>
        Public ReadOnly Property HasAny As Boolean
            Get
                Return Me.InvoiceCount > 0 OrElse Me.PackingCount > 0
            End Get
        End Property
#End Region

    End Class

    ''' <summary>
    ''' INVOICE 明細から集計したヘッダ情報です。usp_shanghai_invoice_gokei_get の戻りを保持します。
    ''' </summary>
    Public NotInheritable Class ShanghaiInvoiceTotal

#Region "プロパティ"
        ''' <summary>BL DATE。</summary>
        Public Property BlDate As Date?

        ''' <summary>出荷方法。数値に変換済みの値。</summary>
        Public Property SyukkaHoho As Integer?

        ''' <summary>カートン数の合計。</summary>
        Public Property TotalCarton As Integer?

        ''' <summary>金額の合計。</summary>
        Public Property TotalAmount As Double?

        ''' <summary>数量の合計。</summary>
        Public Property TotalSuryo As Double?

        ''' <summary>コンテナタイプの表示名。</summary>
        Public Property ContainerText As String = String.Empty

        ''' <summary>登録済みヘッダの件数。0 以外のときはヘッダを作成しません。</summary>
        Public Property HeaderCount As Integer
#End Region

    End Class
End Namespace
