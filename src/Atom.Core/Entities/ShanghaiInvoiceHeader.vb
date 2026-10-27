Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic

Namespace Entities

    ''' <summary>
    ''' 上海輸入 INVOICE のヘッダ（dbo.tbl_t_shanghai_invoice_header）の 1 件分を表します。
    ''' </summary>
    ''' <remarks>
    ''' カートン数・金額・数量の合計は本テーブルの列ではありません。
    ''' 画面に表示する合計は <see cref="ShanghaiInvoiceTotal"/> を使います。
    ''' （`carton_qty` 列は存在しますが、旧 Access では明細の集計値を表示しています）
    ''' </remarks>
    Public Class ShanghaiInvoiceHeader

#Region "プロパティ"
        ''' <summary>INVOICE 番号（メイン）。主キー。</summary>
        Public Property InvoiceNoMain As String = String.Empty

        ''' <summary>B/L DATE。未設定のときは Nothing。</summary>
        Public Property BlDate As Date? = Nothing

        ''' <summary>出荷方法。未設定のときは Nothing。</summary>
        Public Property SyukkaHoho As Integer? = Nothing

        ''' <summary>カートン数。未設定のときは Nothing。</summary>
        Public Property CartonQty As Integer? = Nothing

        ''' <summary>船名または AIR 便名。</summary>
        Public Property Vessel As String = String.Empty

        ''' <summary>返却パレット数。未設定のときは Nothing。</summary>
        Public Property ReturnPallet As Integer? = Nothing

        ''' <summary>社内参照番号（REF_NO）。</summary>
        Public Property RefNo As String = String.Empty

        ''' <summary>乙仲（通関業者）。未設定のときは Nothing。</summary>
        Public Property OtsunakaNo As Integer? = Nothing

        ''' <summary>仕入先コード。</summary>
        Public Property ShiiresakiCode As String = String.Empty

        ''' <summary>更新日時。排他制御に使います。未設定のときは Nothing。</summary>
        Public Property UpdateYmd As Date? = Nothing

        ''' <summary>更新者のログインコード。</summary>
        Public Property UpdateLogin As String = String.Empty
#End Region

    End Class

    ''' <summary>
    ''' 上海輸入 INVOICE のヘッダサブ（明細）の 1 件分を表します。
    ''' </summary>
    ''' <remarks>
    ''' 画面で編集する間は作業用の一時テーブル
    ''' （dbo.tbl_t_shanghai_invoice_header_sub_temp）に保持します。
    ''' 登録時に usp_shanghai_invoice_header_sub_update が本体へ反映します。
    ''' 明細行は INVOICE の枝番から自動生成されるため、画面から追加・削除はしません。
    ''' </remarks>
    Public Class ShanghaiInvoiceHeaderSub

#Region "プロパティ"
        ''' <summary>INVOICE 番号（メイン）。</summary>
        Public Property InvoiceNoMain As String = String.Empty

        ''' <summary>INVOICE 番号（枝番付き）。行を識別するキーです。</summary>
        Public Property InvoiceNo As String = String.Empty

        ''' <summary>配送タイプ。未設定のときは Nothing。</summary>
        Public Property HaisoType As Integer? = Nothing

        ''' <summary>コンテナサイズの区分。未設定のときは Nothing。</summary>
        Public Property ContainerType As Integer? = Nothing

        ''' <summary>配送予定日。未設定のときは Nothing。</summary>
        Public Property HaisoYoteiYmd As Date? = Nothing

        ''' <summary>コンテナ番号。航空便のときは商品群が入ります。</summary>
        Public Property ContainerNo As String = String.Empty

        ''' <summary>シール番号。</summary>
        Public Property SealNo As String = String.Empty

        ''' <summary>カートン数。未設定のときは Nothing。</summary>
        Public Property CartonQty As Integer? = Nothing

        ''' <summary>印刷用 INVOICE_NO。登録時に自動で編集されます。</summary>
        Public Property InvoiceNoPrint As String = String.Empty

        ''' <summary>削除区分。</summary>
        Public Property DelFlg As Integer = 0
#End Region

    End Class

    ''' <summary>
    ''' INVOICE が登録済みかどうかの判定結果です。
    ''' ストアドプロシージャ usp_shanghai_invoice_sonzai_check の戻りを保持します。
    ''' </summary>
    Public NotInheritable Class ShanghaiInvoiceExistence

#Region "プロパティ"
        ''' <summary>ヘッダの登録件数。0 のときは未登録です。</summary>
        Public Property HeaderCount As Integer

        ''' <summary>ヘッダサブの登録件数。0 のときは PACKING から新規に作ります。</summary>
        Public Property SubCount As Integer

        ''' <summary>
        ''' BL DATE と出荷方法の種類数。
        ''' 2 以上のときは PACKING のデータに食い違いがあるため処理を中断します。
        ''' </summary>
        Public Property DateAndMethodKinds As Integer
#End Region

#Region "参照プロパティ"
        ''' <summary>ヘッダが未登録かどうか。</summary>
        Public ReadOnly Property IsNewHeader As Boolean
            Get
                Return Me.HeaderCount = 0
            End Get
        End Property

        ''' <summary>BL DATE と出荷方法に食い違いがあるかどうか。</summary>
        Public ReadOnly Property HasConflict As Boolean
            Get
                Return Me.DateAndMethodKinds >= 2
            End Get
        End Property
#End Region

    End Class

    ''' <summary>
    ''' 選択肢に使う「値と表示名」の組です。
    ''' 項目マスタ由来のビュー（View__syukka_hoho など）の 1 行に対応します。
    ''' </summary>
    Public NotInheritable Class CodeOption

#Region "コンストラクタ"
        ''' <summary>
        ''' 選択肢を初期化します。
        ''' </summary>
        ''' <param name="value">値。</param>
        ''' <param name="text">表示名。</param>
        Public Sub New(value As Integer, text As String)
            Me.Value = value
            Me.Text = text
        End Sub
#End Region

#Region "プロパティ"
        ''' <summary>値。</summary>
        Public ReadOnly Property Value As Integer

        ''' <summary>表示名。</summary>
        Public ReadOnly Property Text As String
#End Region

    End Class

    ''' <summary>
    ''' 印刷用 INVOICE_NO を組み立てるための明細 1 件分です。
    ''' </summary>
    Public NotInheritable Class ShanghaiInvoicePrintTarget

#Region "プロパティ"
        ''' <summary>INVOICE 番号（枝番付き）。</summary>
        Public Property InvoiceNo As String = String.Empty

        ''' <summary>コンテナサイズの区分。未設定のときは Nothing。</summary>
        Public Property ContainerType As Integer? = Nothing

        ''' <summary>コンテナ番号。</summary>
        Public Property ContainerNo As String = String.Empty

        ''' <summary>配送予定日。未設定のときは Nothing。</summary>
        Public Property HaisoYoteiYmd As Date? = Nothing
#End Region

#Region "参照プロパティ"
        ''' <summary>
        ''' 印刷用 INVOICE_NO の組み立て対象かどうか。
        ''' </summary>
        ''' <remarks>
        ''' コンテナ区分が未設定または 0、あるいはコンテナ番号が空の行は対象外です（旧 Access と同じ）。
        ''' </remarks>
        Public ReadOnly Property IsPrintTarget As Boolean
            Get
                ' コンテナ区分が未設定または 0 のとき
                If Not Me.ContainerType.HasValue OrElse Me.ContainerType.Value = 0 Then
                    Return False
                End If
                Return Not String.IsNullOrEmpty(Me.ContainerNo)
            End Get
        End Property
#End Region

    End Class

    ''' <summary>
    ''' 登録処理の結果です。
    ''' </summary>
    Public NotInheritable Class ShanghaiInvoiceRegisterResult

#Region "プロパティ"
        ''' <summary>入力チェックに引っかかったときの文言。問題が無いときは空文字。</summary>
        Public Property ValidationMessage As String = String.Empty

        ''' <summary>
        ''' 明細に配送予定日が 1 件でも入っていたかどうか。
        ''' </summary>
        ''' <remarks>
        ''' 旧 Access はこのとき外部システムへ HTTP で通知しています（設計書 §8.5）。
        ''' Web 版での扱いは未決のため、判定結果だけを返します。
        ''' </remarks>
        Public Property HasHaisoYotei As Boolean

        ''' <summary>
        ''' 設定により POEM の更新を行わなかったかどうか。
        ''' </summary>
        ''' <remarks>
        ''' POEM の更新は本番のデータベースを直接書き換えるため、既定では行いません。
        ''' 行わなかったことを画面に出して、利用者が気付けるようにします。
        ''' </remarks>
        Public Property IsPoemSkipped As Boolean
#End Region

#Region "参照プロパティ"
        ''' <summary>登録できたかどうか。</summary>
        Public ReadOnly Property IsSuccess As Boolean
            Get
                Return String.IsNullOrEmpty(Me.ValidationMessage)
            End Get
        End Property
#End Region

    End Class

    ''' <summary>
    ''' INVOICEヘッダー入力画面（menuNo=103）の表示に必要な一式です。
    ''' </summary>
    Public NotInheritable Class ShanghaiInvoiceHeaderView

#Region "コンストラクタ"
        ''' <summary>既定の状態で初期化します。</summary>
        Public Sub New()
            Me.Details = New List(Of ShanghaiInvoiceHeaderSub)()
        End Sub
#End Region

#Region "プロパティ"
        ''' <summary>登録状況の判定結果。</summary>
        Public Property Existence As ShanghaiInvoiceExistence

        ''' <summary>ヘッダ。未登録のときは Nothing。</summary>
        Public Property Header As ShanghaiInvoiceHeader

        ''' <summary>明細から集計した合計情報。</summary>
        Public Property Total As ShanghaiInvoiceTotal

        ''' <summary>明細。</summary>
        Public Property Details As IList(Of ShanghaiInvoiceHeaderSub)
#End Region

    End Class
End Namespace
