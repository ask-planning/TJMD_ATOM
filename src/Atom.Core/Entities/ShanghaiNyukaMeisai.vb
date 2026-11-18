Option Strict On
Option Explicit On
Option Infer Off

Namespace Entities

    ''' <summary>
    ''' 入荷明細書（menuNo=103 §11）に印字する 1 行です。
    ''' </summary>
    ''' <remarks>
    ''' dbo.View_Rpt_shanghai_nyuka_meisai の 1 行に、枝番ごとのカートン数
    ''' （dbo.tbl_t_shanghai_invoice_header_sub.carton_qty）を足したものです。
    '''
    ''' **カートン数を明細から集計しない**のは、旧 ATOM がヘッダー側の値を出しているためです。
    ''' 実データで照合したところ 6,557 件中 138 件がずれました（設計書 §14.1.1.3）。
    ''' </remarks>
    Public Class ShanghaiNyukaMeisaiReportRow

#Region "並び順"
        ''' <summary>並び順。印刷用一時テーブルへ入れた順です。</summary>
        Public Property SortSeq As Integer?
#End Region

#Region "コンテナ単位の項目"
        ''' <summary>INVOICE 番号（枝番なし）。帳票の左上に大きく出ます。</summary>
        Public Property InvoiceNoMain As String = String.Empty

        ''' <summary>B/L DATE。</summary>
        Public Property BlDate As Date?

        ''' <summary>コンテナ区分名。「コンテナ」を付けて見出しにします。</summary>
        Public Property ContainerTypeText As String = String.Empty

        ''' <summary>コンテナ番号。</summary>
        Public Property ContainerNo As String = String.Empty

        ''' <summary>シール番号。</summary>
        Public Property SealNo As String = String.Empty

        ''' <summary>配送予定日。</summary>
        Public Property HaisoYoteiYmd As Date?

        ''' <summary>印字用の INVOICE 番号。枝番を連結したものです。</summary>
        Public Property InvoiceNoPrint As String = String.Empty

        ''' <summary>REMARKS。コンテナ単位の備考です。</summary>
        Public Property Remarks As String = String.Empty
#End Region

#Region "枝番単位の項目"
        ''' <summary>INVOICE 番号（枝番つき）。グループの見出しで、改ページの単位です。</summary>
        Public Property InvoiceNo As String = String.Empty

        ''' <summary>その枝番のカートン数。合計欄に出ます。</summary>
        Public Property TotalCarton As Integer?
#End Region

#Region "明細の項目"
        ''' <summary>納入先のプラントコード。グループの切れ目の判定に使います。</summary>
        Public Property NounyuPlant As String = String.Empty

        ''' <summary>納入先の名称。</summary>
        Public Property HokanBasyoText As String = String.Empty

        ''' <summary>在庫転送NO。</summary>
        Public Property ZaikoTensoNo As String = String.Empty

        ''' <summary>出荷伝票NO。</summary>
        Public Property SyukkaDenpyoNo As String = String.Empty

        ''' <summary>抜取検査の区分。抜取印と抜取箱数の元になります。</summary>
        Public Property NukitoriKensaFlg As Integer?

        ''' <summary>品目コード。</summary>
        Public Property R3HinmokuCode As String = String.Empty

        ''' <summary>CARTON_NO の開始。</summary>
        Public Property CartonNoFrom As Integer?

        ''' <summary>CARTON_NO の終了。</summary>
        Public Property CartonNoTo As Integer?

        ''' <summary>入庫箱数。</summary>
        Public Property CartonQty As Integer?

        ''' <summary>納入数量。</summary>
        Public Property NounyuSuryo As Decimal?
#End Region

    End Class

    ''' <summary>
    ''' 作成済みの入荷明細の件数と最終更新日です。
    ''' </summary>
    ''' <remarks>
    ''' `usp_shanghai_nyuka_meisai_kensu` の出力引数をまとめたものです（設計書 §11.2）。
    ''' 件数が 0 かどうかで、画面に出す確認の文言が変わります。
    ''' </remarks>
    Public Class ShanghaiNyukaMeisaiKensu

        ''' <summary>作成済みの件数。</summary>
        Public Property Count As Integer

        ''' <summary>最終更新日。1 件も無いときは Nothing。</summary>
        Public Property UpdatedAt As Date?

    End Class

End Namespace
