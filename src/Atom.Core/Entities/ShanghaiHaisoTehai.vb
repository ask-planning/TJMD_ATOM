Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic

Namespace Entities

#Region "印刷対象（ビューから拾った 1 行）"
    ''' <summary>
    ''' 配送手配書の印刷対象の 1 行です。dbo.View_shanghai_haiso_tehai_pickup の 1 行に対応します。
    ''' </summary>
    ''' <remarks>
    ''' INVOICE の枝番ごとに 1 行です。これを配送パターンでまとめて印刷用の 1 行にします
    ''' （設計書 §10.3）。
    ''' </remarks>
    Public Class ShanghaiHaisoTehaiPickup

        ''' <summary>INVOICE 番号（メイン）。</summary>
        Public Property InvoiceNoMain As String = String.Empty

        ''' <summary>INVOICE 番号（枝番まで含む）。</summary>
        Public Property InvoiceNo As String = String.Empty

        ''' <summary>B/L DATE。</summary>
        Public Property BlDate As Date?

        ''' <summary>船名。</summary>
        Public Property Vessel As String = String.Empty

        ''' <summary>乙仲番号。</summary>
        Public Property OtsunakaNo As Integer?

        ''' <summary>配送予定日。</summary>
        Public Property HaisoYoteiYmd As Date?

        ''' <summary>配送タイプ。</summary>
        Public Property HaisoType As Integer?

        ''' <summary>配送タイプ名。</summary>
        Public Property HaisoTypeText As String = String.Empty

        ''' <summary>コンテナ区分。</summary>
        Public Property ContainerType As Integer?

        ''' <summary>コンテナ区分名。帳票にはこちらを印字します。</summary>
        Public Property ContainerTypeText As String = String.Empty

        ''' <summary>コンテナ番号。</summary>
        Public Property ContainerNo As String = String.Empty

        ''' <summary>シール番号。</summary>
        Public Property SealNo As String = String.Empty

        ''' <summary>カートン数。</summary>
        Public Property CartonQty As Integer?

        ''' <summary>数量の合計。</summary>
        Public Property TotalSuryo As Integer?

        ''' <summary>配送先 1 箇所目。</summary>
        Public Property HaisoSaki1 As Integer?

        ''' <summary>配送先 2 箇所目。</summary>
        Public Property HaisoSaki2 As Integer?

        ''' <summary>配送先 3 箇所目。</summary>
        Public Property HaisoSaki3 As Integer?

    End Class
#End Region

#Region "印刷用の 1 行"
    ''' <summary>
    ''' 印刷用一時テーブル（dbo.tbl_t_shanghai_haiso_tehai_print_temp）へ登録する 1 行です。
    ''' </summary>
    ''' <remarks>
    ''' 配送パターンごとに 1 行です。枝番はカンマで連結し、カートン数と数量は合計します。
    ''' </remarks>
    Public Class ShanghaiHaisoTehaiPrintRow

        ''' <summary>INVOICE 番号（メイン）。</summary>
        Public Property InvoiceNoMain As String = String.Empty

        ''' <summary>印字する INVOICE 番号。先頭の番号に 2 件目以降の枝番をカンマで足したもの。</summary>
        Public Property InvoiceNo As String = String.Empty

        ''' <summary>船名。</summary>
        Public Property Vessel As String = String.Empty

        ''' <summary>B/L DATE。</summary>
        Public Property BlDate As Date?

        ''' <summary>配送タイプ。</summary>
        Public Property HaisoType As Integer?

        ''' <summary>コンテナ区分名。</summary>
        Public Property ContainerType As String = String.Empty

        ''' <summary>コンテナ番号。</summary>
        Public Property ContainerNo As String = String.Empty

        ''' <summary>シール番号。</summary>
        Public Property SealNo As String = String.Empty

        ''' <summary>カートン数の合計。</summary>
        Public Property TotalCarton As Integer

        ''' <summary>数量の合計。</summary>
        Public Property TotalQuantity As Integer

        ''' <summary>配送予定日。未入力のときは実行日が入ります（設計書 §10.3 / C-2）。</summary>
        Public Property HaisoYoteiYmd As Date?

        ''' <summary>配送先 1 箇所目。</summary>
        Public Property HaisoSaki1 As Integer?

        ''' <summary>配送先 2 箇所目。</summary>
        Public Property HaisoSaki2 As Integer?

        ''' <summary>配送先 3 箇所目。</summary>
        Public Property HaisoSaki3 As Integer?

        ''' <summary>乙仲番号。</summary>
        Public Property OtsunakaNo As Integer?

    End Class
#End Region

#Region "帳票の 1 行"
    ''' <summary>
    ''' 配送手配書に印字する 1 行です。dbo.View_Rpt_shanghai_haiso_tehai の 1 行に対応します。
    ''' </summary>
    ''' <remarks>
    ''' 帳票に出さない列（`total_quantity` や `atesaki_nisshin` など）は持ちません。
    ''' 必要になった時点で足してください。
    ''' </remarks>
    Public Class ShanghaiHaisoTehaiReportRow

        ''' <summary>参照番号。</summary>
        Public Property RefNo As String = String.Empty

        ''' <summary>印字する INVOICE 番号。</summary>
        Public Property InvoiceNo As String = String.Empty

        ''' <summary>船名。</summary>
        Public Property Vessel As String = String.Empty

        ''' <summary>B/L DATE。</summary>
        Public Property BlDate As Date?

        ''' <summary>配送タイプ。</summary>
        Public Property HaisoType As Integer?

        ''' <summary>コンテナ区分名。</summary>
        Public Property ContainerType As String = String.Empty

        ''' <summary>コンテナ番号。</summary>
        Public Property ContainerNo As String = String.Empty

        ''' <summary>シール番号。</summary>
        Public Property SealNo As String = String.Empty

        ''' <summary>カートン数の合計。</summary>
        Public Property TotalCarton As Integer?

        ''' <summary>配送予定日。</summary>
        Public Property HaisoYoteiYmd As Date?

        ''' <summary>乙仲名。</summary>
        Public Property OtsunakaNoText As String = String.Empty

        ''' <summary>乙仲の担当 1。</summary>
        Public Property OtsunakaTanto1 As String = String.Empty

        ''' <summary>乙仲の担当 2。</summary>
        Public Property OtsunakaTanto2 As String = String.Empty

        ''' <summary>配送先 1 箇所目のコード。</summary>
        Public Property HaisoSaki1No As Integer?

        ''' <summary>配送先 2 箇所目のコード。</summary>
        Public Property HaisoSaki2No As Integer?

        ''' <summary>配送先 3 箇所目のコード。</summary>
        Public Property HaisoSaki3No As Integer?

        ''' <summary>配送先 1 箇所目。</summary>
        Public Property HaisoSaki1 As ShanghaiHaisoSaki = New ShanghaiHaisoSaki()

        ''' <summary>配送先 2 箇所目。</summary>
        Public Property HaisoSaki2 As ShanghaiHaisoSaki = New ShanghaiHaisoSaki()

        ''' <summary>配送先 3 箇所目。</summary>
        Public Property HaisoSaki3 As ShanghaiHaisoSaki = New ShanghaiHaisoSaki()

        ''' <summary>発行者の氏名。</summary>
        Public Property DisplayName As String = String.Empty

        ''' <summary>発行者の部門名。</summary>
        Public Property BumonCodeText As String = String.Empty

        ''' <summary>発行者の電話番号。</summary>
        Public Property PhoneNo As String = String.Empty

        ''' <summary>発行者の FAX 番号。</summary>
        Public Property FaxNo As String = String.Empty

        ''' <summary>
        ''' 埋まっている配送先の箇所数です。
        ''' </summary>
        ''' <returns>1 〜 3。</returns>
        ''' <remarks>
        ''' この数で様式（定期便 / 定期便2 / 定期便3）が決まります（設計書 §10.4）。
        ''' 利用者が指定するものではなく、配送タイプに紐づいた固定値から決まります。
        '''
        ''' **判定は配送先名ではなくコードで行います。**設計書 §10.4 の条件が
        ''' 「NULL でなく 0 でもない」ためです。名前で判定すると、マスタに未登録の配送先が
        ''' あったときに箇所数が減り、2 件目以降が帳票から落ちます。
        ''' </remarks>
        Public ReadOnly Property HaisoSakiCount As Integer
            Get
                ' 3 箇所目まで設定されているとき
                If HasHaisoSaki(Me.HaisoSaki3No) Then
                    Return 3
                End If
                ' 2 箇所目まで設定されているとき
                If HasHaisoSaki(Me.HaisoSaki2No) Then
                    Return 2
                End If
                Return 1
            End Get
        End Property

        ''' <summary>
        ''' 配送先コードが設定されているかを判定します。
        ''' </summary>
        ''' <param name="haisoSakiNo">配送先コード。</param>
        ''' <returns>NULL でなく 0 でもないとき True。</returns>
        Private Shared Function HasHaisoSaki(haisoSakiNo As Integer?) As Boolean
            Return haisoSakiNo.HasValue AndAlso haisoSakiNo.Value <> 0
        End Function

    End Class

    ''' <summary>
    ''' 配送先 1 箇所分です。
    ''' </summary>
    Public Class ShanghaiHaisoSaki

        ''' <summary>配送先名。</summary>
        Public Property Name As String = String.Empty

        ''' <summary>住所。</summary>
        Public Property Address As String = String.Empty

        ''' <summary>電話番号。</summary>
        Public Property PhoneNo As String = String.Empty

        ''' <summary>担当者。</summary>
        Public Property Tanto As String = String.Empty

        ''' <summary>
        ''' 配送先が設定されているかどうかです。
        ''' </summary>
        ''' <returns>配送先名が入っているとき True。</returns>
        Public ReadOnly Property HasValue As Boolean
            Get
                Return Not String.IsNullOrWhiteSpace(Me.Name)
            End Get
        End Property

    End Class
#End Region

#Region "発行結果"
    ''' <summary>
    ''' 配送手配書の発行結果です。
    ''' </summary>
    Public Class ShanghaiHaisoTehaiResult

        ''' <summary>入力に問題があったときの文言。問題が無いときは空文字。</summary>
        Public Property ValidationMessage As String = String.Empty

        ''' <summary>発行した帳票の行。</summary>
        Public Property Rows As IList(Of ShanghaiHaisoTehaiReportRow) = New List(Of ShanghaiHaisoTehaiReportRow)()

        ''' <summary>
        ''' 発行できたかどうかです。
        ''' </summary>
        ''' <returns>入力に問題が無く、対象が 1 件以上あったとき True。</returns>
        Public ReadOnly Property IsSuccess As Boolean
            Get
                Return String.IsNullOrEmpty(Me.ValidationMessage) AndAlso Me.Rows.Count > 0
            End Get
        End Property

    End Class
#End Region

End Namespace
