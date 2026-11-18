Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports Atom.Core.Entities

''' <summary>
''' 入荷明細書（menuNo=103 §11）の連携ファイルを組み立てて配置します。
''' </summary>
''' <remarks>
''' **これまでの帳票と形が違います。**配送手配書は 1 行＝1 枚ですが、入荷明細書は
''' 明細が繰り返す帳票で、SVF の「ヘッダー明細別レコード形式」を使います。
'''
''' **ヘッダー行は 1 枚につき 1 組しか持てません**（実機で確認）。
''' そのため、1 枚の中で変わりうる項目はすべて明細側の列に持たせ、
''' ヘッダー行には**元INV_NO の中で変わらないものだけ**を残します。
'''
''' 旧 ATOM はページヘッダーに B/L DATE やコンテナNO を**フィールドとして**置いており、
''' ページごとにそのページのレコードの値を出していました。1 つの元INV_NO が複数の
''' コンテナにまたがるのは実データで珍しくないため、**同じ出力になるようコンテナの項目も
''' 明細側へ移しています**（設計書 §14.1.1.3 / C-54）。
''' 枝番（`pk_invoice_no`）は様式のキー番号で改ページさせます。
''' </remarks>
Public Class NyukaMeisaiReportWriter

#Region "定数"
    ''' <summary>ヘッダー行の目印。見出しの 1 列目の名前でもあります。</summary>
    Private Const HeaderMarker As String = "header"

    ''' <summary>見出しのうち、区分列の次から数えたヘッダー項目の個数。</summary>
    Private Const HeaderItemCount As Integer = 5

    ''' <summary>ファイル名に付ける日時の書式。</summary>
    Private Const TimestampFormat As String = "yyyyMMdd_HHmmss"

    ''' <summary>日付の書式。</summary>
    Private Const DateFormat As String = "yyyy/MM/dd"

    ''' <summary>発行日の書式。</summary>
    Private Const HakkoYmdFormat As String = "yyyy年M月d日"

    ''' <summary>整数の書式。</summary>
    Private Const IntegerFormat As String = "#,##0"

    ''' <summary>小数を含む数の書式。割り切れるときは小数を出しません。</summary>
    Private Const DecimalFormat As String = "#,##0.##"

    ''' <summary>上海タジマの INVOICE 番号の接頭辞。</summary>
    Private Const ShanghaiTajimaPrefix As String = "SJ"

    ''' <summary>上海タジマのときのタイトル。</summary>
    Private Const TitleShanghaiTajima As String = "上海タジマ入荷明細"

    ''' <summary>上海タジマ以外のときのタイトル。</summary>
    Private Const TitleOther As String = "上海庸助貿易入荷明細"

    ''' <summary>コンテナ区分名に付ける語。</summary>
    Private Const ContainerSuffix As String = "コンテナ"

    ''' <summary>抜取検査があるときに付ける印。</summary>
    Private Const NukitoriMark As String = "※"

    ''' <summary>抜取検査の区分。1 のときは 1 箱だけを抜き取ります。</summary>
    Private Const NukitoriKensaOneBox As Integer = 1
#End Region

#Region "フィールド"
    ''' <summary>連携ファイルの書き出しを担います。</summary>
    Private ReadOnly _writer As New SvfDataFileWriter()
#End Region

#Region "公開メソッド"
    ''' <summary>
    ''' 入荷明細書の連携ファイルを配置します。
    ''' </summary>
    ''' <param name="settings">SVF の設定。</param>
    ''' <param name="rows">印字する行。SQL 側で並べ替えた順のまま渡します。</param>
    ''' <param name="loginCode">利用者のログインコード。</param>
    ''' <param name="hakkoYmd">発行日時。ファイル名と発行日に使います。</param>
    ''' <returns>配置の結果。</returns>
    ''' <remarks>
    ''' **並び順は呼び出し側の責任です。**コンテナ → 枝番 → 納入先 → 在庫転送NO →
    ''' 出荷伝票の順に並んでいる前提で、切れ目を見ながら重複を空にします。
    ''' </remarks>
    Public Function Write(settings As SvfReportSettings,
                          rows As IList(Of ShanghaiNyukaMeisaiReportRow),
                          loginCode As String,
                          hakkoYmd As Date) As NyukaMeisaiWriteResult

        Dim result As New NyukaMeisaiWriteResult()
        ' 対象が無いとき
        If rows Is Nothing OrElse rows.Count = 0 Then
            Return result
        End If

        Dim form As HaisoTehaiForm = HaisoTehaiForm.NyukaMeisai()
        result.FilePaths.Add(Me.WriteOneFile(settings, form, rows, loginCode, hakkoYmd))
        Return result
    End Function

    ''' <summary>
    ''' 連携ファイルの見出し行を返します。
    ''' </summary>
    ''' <returns>26 列の項目名。</returns>
    ''' <remarks>
    ''' `tools/svf/定義AT06入荷明細書.txt` および `schema.ini` の
    ''' `[定義AT06入荷明細書.txt]` と**同じ並び・同じ名前**でなければなりません。
    ''' 1 列目が区分列、続く 5 個がヘッダー項目、残りの 20 個が明細項目です。
    ''' </remarks>
    Public Shared Function CreateHeaders() As IList(Of String)
        Dim headers As New List(Of String)()

        ' 区分列
        headers.Add(HeaderMarker)

        ' ヘッダー項目（5 個）。元INV_NO の中で変わらないものだけ
        headers.Add("帳票区分")
        headers.Add("login_code")
        headers.Add("title")
        headers.Add("hakko_ymd")
        headers.Add("invoice_no_main")

        ' 明細項目（20 個）。枝番とコンテナの項目もここに入ります
        headers.Add("invoice_no")
        headers.Add("total_carton_d")
        headers.Add("bl_date")
        headers.Add("container_type_label")
        headers.Add("invoice_no_print")
        headers.Add("container_no")
        headers.Add("seal_no")
        headers.Add("haiso_yotei_ymd")
        headers.Add("remarks")
        headers.Add("hokan_basyo_text")
        headers.Add("zaiko_tenso_no")
        headers.Add("syukka_denpyo_no")
        headers.Add("nukitori_mark")
        headers.Add("R3_hinmoku_code")
        headers.Add("carton_no_from")
        headers.Add("carton_no_to")
        headers.Add("nukitori_hakosu")
        headers.Add("carton_qty")
        headers.Add("nounyu_suryo")
        headers.Add("irisu")

        Return headers
    End Function
#End Region

#Region "内部処理"
    ''' <summary>
    ''' 連携ファイルを 1 本配置します。
    ''' </summary>
    ''' <param name="settings">SVF の設定。</param>
    ''' <param name="form">様式。</param>
    ''' <param name="rows">印字する行。</param>
    ''' <param name="loginCode">利用者のログインコード。</param>
    ''' <param name="hakkoYmd">発行日時。</param>
    ''' <returns>配置したファイルのフルパス。</returns>
    ''' <remarks>
    ''' ヘッダー行は 1 本だけです。**元INV_NO の中で変わる項目は明細側に持たせている**ため、
    ''' コンテナや枝番がいくつあっても 1 ファイルで足ります。
    ''' </remarks>
    Private Function WriteOneFile(settings As SvfReportSettings,
                                  form As HaisoTehaiForm,
                                  rows As IList(Of ShanghaiNyukaMeisaiReportRow),
                                  loginCode As String,
                                  hakkoYmd As Date) As String

        Dim controlLines As New List(Of String)()
        controlLines.Add(SvfDataFileWriter.CreateQuotedControlLine("VrSetOutputPrinter", settings.OutputPrinter))
        controlLines.Add(SvfDataFileWriter.CreateControlLine("VrSetUseReportDirector", "1"))
        controlLines.Add(SvfDataFileWriter.CreateQuotedControlLine("VrSetDocName2", form.DisplayName))

        Dim group As New SvfHeaderDetailGroup()
        For Each value As String In Me.CreateHeaderValues(form, rows(0), loginCode, hakkoYmd)
            group.HeaderValues.Add(value)
        Next
        For Each detail As IList(Of String) In Me.CreateDetailRows(rows)
            group.DetailRows.Add(detail)
        Next

        Dim groups As New List(Of SvfHeaderDetailGroup)()
        groups.Add(group)

        Dim fileName As String = Me.CreateFileName(form, rows(0), hakkoYmd)
        Return Me._writer.WriteHeaderDetail(settings.WatchFolderPath, fileName, controlLines,
                                            CreateHeaders(), HeaderItemCount, groups, loginCode)
    End Function

    ''' <summary>
    ''' ヘッダー行の値を組み立てます。
    ''' </summary>
    ''' <param name="form">様式。</param>
    ''' <param name="row">先頭の行。</param>
    ''' <param name="loginCode">利用者のログインコード。</param>
    ''' <param name="hakkoYmd">発行日。</param>
    ''' <returns>区分列を含む 6 列の値。</returns>
    Private Function CreateHeaderValues(form As HaisoTehaiForm,
                                        row As ShanghaiNyukaMeisaiReportRow,
                                        loginCode As String,
                                        hakkoYmd As Date) As IList(Of String)

        Dim values As New List(Of String)()
        values.Add(HeaderMarker)
        values.Add(form.ReportKubun)
        values.Add(loginCode)
        values.Add(Me.CreateTitle(row))
        values.Add(hakkoYmd.ToString(HakkoYmdFormat, CultureInfo.InvariantCulture))
        values.Add(row.InvoiceNoMain)
        Return values
    End Function

    ''' <summary>
    ''' 明細行をまとめて組み立てます。
    ''' </summary>
    ''' <param name="rows">印字する行。</param>
    ''' <returns>明細行の一覧。各行は 20 列です。</returns>
    ''' <remarks>
    ''' **コンテナの項目は毎行入れます。**様式側はヘッダーレコードで 1 ページに 1 回だけ
    ''' 印字するため、同じ値が行ごとに並んでも重複して出ることはありません。
    '''
    ''' **重複した値は ATOM 側で空にします**（設計書 §14.1.1.3）。
    ''' 納入先・在庫転送NO・出荷伝票は、直前の行と同じなら空文字で渡し、
    ''' 様式側では値があるときだけ枠を出します。
    '''
    ''' カートン数は枝番ごとの合計で、**枝番の先頭行にだけ**入れます。
    ''' 様式の合計レコードが `SUM(total_carton_d)` で拾うため、
    ''' 全行に入れると枝番の行数ぶん足されてしまいます。
    ''' </remarks>
    Private Function CreateDetailRows(rows As IList(Of ShanghaiNyukaMeisaiReportRow)) As IList(Of IList(Of String))
        Dim details As New List(Of IList(Of String))()

        Dim previousInvoiceNo As String = Nothing
        Dim previousBasyo As String = Nothing
        Dim previousTenso As String = Nothing
        Dim previousDenpyo As String = Nothing

        For Each row As ShanghaiNyukaMeisaiReportRow In rows
            Dim isNewInvoice As Boolean =
                Not String.Equals(row.InvoiceNo, previousInvoiceNo, StringComparison.Ordinal)

            ' 枝番が変わったら、重複の判定もやり直します（改ページで先頭行になるため）
            If isNewInvoice Then
                previousBasyo = Nothing
                previousTenso = Nothing
                previousDenpyo = Nothing
            End If

            Dim values As New List(Of String)()
            values.Add(row.InvoiceNo)
            values.Add(If(isNewInvoice, FormatInteger(row.TotalCarton), String.Empty))
            values.Add(FormatDate(row.BlDate))
            values.Add(Me.CreateContainerTypeLabel(row))
            values.Add(row.InvoiceNoPrint)
            values.Add(row.ContainerNo)
            values.Add(row.SealNo)
            values.Add(FormatDate(row.HaisoYoteiYmd))
            values.Add(row.Remarks)
            values.Add(BlankWhenSame(row.HokanBasyoText, previousBasyo))
            values.Add(BlankWhenSame(row.ZaikoTensoNo, previousTenso))
            values.Add(BlankWhenSame(row.SyukkaDenpyoNo, previousDenpyo))
            values.Add(Me.CreateNukitoriMark(row))
            values.Add(row.R3HinmokuCode)
            values.Add(Me.CreateCartonNoFrom(row))
            values.Add(FormatInteger(row.CartonNoTo))
            values.Add(Me.CreateNukitoriHakosu(row))
            values.Add(FormatInteger(row.CartonQty))
            values.Add(FormatDecimal(row.NounyuSuryo))
            values.Add(Me.CreateIrisu(row))
            details.Add(values)

            previousInvoiceNo = row.InvoiceNo
            previousBasyo = row.HokanBasyoText
            previousTenso = row.ZaikoTensoNo
            previousDenpyo = row.SyukkaDenpyoNo
        Next

        Return details
    End Function

    ''' <summary>
    ''' 帳票のタイトルを決めます。
    ''' </summary>
    ''' <param name="row">帳票の行。</param>
    ''' <returns>タイトル。</returns>
    ''' <remarks>旧 Access の式 `=IIf(Left([pk_invoice_no_main],2)="SJ",…)` に合わせています。</remarks>
    Private Function CreateTitle(row As ShanghaiNyukaMeisaiReportRow) As String
        ' INVOICE 番号が 2 文字に満たないとき
        If row.InvoiceNoMain Is Nothing OrElse row.InvoiceNoMain.Length < ShanghaiTajimaPrefix.Length Then
            Return TitleOther
        End If

        Dim prefix As String = row.InvoiceNoMain.Substring(0, ShanghaiTajimaPrefix.Length)
        ' 上海タジマのとき
        If String.Equals(prefix, ShanghaiTajimaPrefix, StringComparison.OrdinalIgnoreCase) Then
            Return TitleShanghaiTajima
        End If
        Return TitleOther
    End Function

    ''' <summary>
    ''' コンテナ区分の見出しを組み立てます。
    ''' </summary>
    ''' <param name="row">帳票の行。</param>
    ''' <returns>「20Ftコンテナ」のような見出し。</returns>
    ''' <remarks>旧 Access の式 `=[container_type_text] & "コンテナ"` に合わせています。</remarks>
    Private Function CreateContainerTypeLabel(row As ShanghaiNyukaMeisaiReportRow) As String
        ' 区分名が無いとき
        If String.IsNullOrEmpty(row.ContainerTypeText) Then
            Return String.Empty
        End If
        Return row.ContainerTypeText & ContainerSuffix
    End Function

    ''' <summary>
    ''' 抜取印を決めます。
    ''' </summary>
    ''' <param name="row">帳票の行。</param>
    ''' <returns>抜取検査があるとき「※」、無いとき空文字。</returns>
    Private Function CreateNukitoriMark(row As ShanghaiNyukaMeisaiReportRow) As String
        ' 抜取検査が無いとき
        If Not HasNukitoriKensa(row) Then
            Return String.Empty
        End If
        Return NukitoriMark
    End Function

    ''' <summary>
    ''' 抜取箱数を決めます。
    ''' </summary>
    ''' <param name="row">帳票の行。</param>
    ''' <returns>抜取箱数。抜取検査が無いときは空文字。</returns>
    ''' <remarks>
    ''' 旧 Access の式に合わせ、区分が 1 のときは 1 箱、それ以外は CARTON_NO の範囲ぶんです。
    ''' </remarks>
    Private Function CreateNukitoriHakosu(row As ShanghaiNyukaMeisaiReportRow) As String
        ' 抜取検査が無いとき
        If Not HasNukitoriKensa(row) Then
            Return String.Empty
        End If

        ' 1 箱だけ抜き取るとき
        If row.NukitoriKensaFlg.Value = NukitoriKensaOneBox Then
            Return NukitoriKensaOneBox.ToString(CultureInfo.InvariantCulture)
        End If

        Dim count As Integer? = CartonCount(row)
        ' 範囲が取れないとき
        If Not count.HasValue Then
            Return String.Empty
        End If
        Return count.Value.ToString(IntegerFormat, CultureInfo.InvariantCulture)
    End Function

    ''' <summary>
    ''' CARTON_NO の開始を決めます。
    ''' </summary>
    ''' <param name="row">帳票の行。</param>
    ''' <returns>開始の番号。終了が無いときは空文字。</returns>
    ''' <remarks>
    ''' 旧 Access の式 `=IIf([carton_no_to] Is Null, Null, [carton_no_from])` に合わせています。
    ''' 終了が無い行は範囲として成り立たないため、開始も出しません。
    ''' </remarks>
    Private Function CreateCartonNoFrom(row As ShanghaiNyukaMeisaiReportRow) As String
        ' 終了が無いとき
        If Not row.CartonNoTo.HasValue Then
            Return String.Empty
        End If
        Return FormatInteger(row.CartonNoFrom)
    End Function

    ''' <summary>
    ''' 入数を求めます。
    ''' </summary>
    ''' <param name="row">帳票の行。</param>
    ''' <returns>1 箱あたりの数量。求められないときは空文字。</returns>
    ''' <remarks>
    ''' 旧 Access の式 `=[nounyu_suryo]/([carton_no_to]-[carton_no_from]+1)` に合わせています。
    ''' **0 除算に注意**が必要なため、範囲が 0 以下のときは空にします（設計書 §14.1.1.3）。
    ''' </remarks>
    Private Function CreateIrisu(row As ShanghaiNyukaMeisaiReportRow) As String
        Dim count As Integer? = CartonCount(row)
        ' 数量か範囲が取れないとき
        If Not row.NounyuSuryo.HasValue OrElse Not count.HasValue OrElse count.Value <= 0 Then
            Return String.Empty
        End If

        Dim irisu As Decimal = row.NounyuSuryo.Value / CDec(count.Value)
        Return irisu.ToString(DecimalFormat, CultureInfo.InvariantCulture)
    End Function

    ''' <summary>
    ''' 連携ファイルのファイル名を組み立てます。
    ''' </summary>
    ''' <param name="form">様式。</param>
    ''' <param name="row">先頭の行。</param>
    ''' <param name="hakkoYmd">発行日時。</param>
    ''' <returns>「接頭辞_INVOICE番号_日時.txt」の形のファイル名。</returns>
    Private Function CreateFileName(form As HaisoTehaiForm,
                                    row As ShanghaiNyukaMeisaiReportRow,
                                    hakkoYmd As Date) As String

        Dim key As String = SanitizeFileNamePart(row.InvoiceNoMain)
        Dim timestamp As String = hakkoYmd.ToString(TimestampFormat, CultureInfo.InvariantCulture)
        Return form.FilePrefix & "_" & key & "_" & timestamp & ".txt"
    End Function

    ''' <summary>
    ''' 抜取検査があるかどうかを判定します。
    ''' </summary>
    ''' <param name="row">帳票の行。</param>
    ''' <returns>抜取検査があるとき True。</returns>
    Private Shared Function HasNukitoriKensa(row As ShanghaiNyukaMeisaiReportRow) As Boolean
        ' 区分が無いとき
        If Not row.NukitoriKensaFlg.HasValue Then
            Return False
        End If
        Return row.NukitoriKensaFlg.Value <> 0
    End Function

    ''' <summary>
    ''' CARTON_NO の範囲の箱数を求めます。
    ''' </summary>
    ''' <param name="row">帳票の行。</param>
    ''' <returns>箱数。範囲が取れないときは Nothing。</returns>
    Private Shared Function CartonCount(row As ShanghaiNyukaMeisaiReportRow) As Integer?
        ' 開始か終了が無いとき
        If Not row.CartonNoFrom.HasValue OrElse Not row.CartonNoTo.HasValue Then
            Return Nothing
        End If
        Return row.CartonNoTo.Value - row.CartonNoFrom.Value + 1
    End Function

    ''' <summary>
    ''' 直前と同じ値のときは空文字にします。
    ''' </summary>
    ''' <param name="value">今回の値。</param>
    ''' <param name="previous">直前の値。</param>
    ''' <returns>直前と違うときは今回の値、同じときは空文字。</returns>
    Private Shared Function BlankWhenSame(value As String, previous As String) As String
        ' 直前と同じとき
        If String.Equals(value, previous, StringComparison.Ordinal) Then
            Return String.Empty
        End If
        Return If(value, String.Empty)
    End Function

    ''' <summary>
    ''' 日付を帳票の書式にします。
    ''' </summary>
    ''' <param name="value">日付。</param>
    ''' <returns>「yyyy/MM/dd」の形。値が無いときは空文字。</returns>
    Private Shared Function FormatDate(value As Date?) As String
        ' 値が無いとき
        If Not value.HasValue Then
            Return String.Empty
        End If
        Return value.Value.ToString(DateFormat, CultureInfo.InvariantCulture)
    End Function

    ''' <summary>
    ''' 整数を帳票の書式にします。
    ''' </summary>
    ''' <param name="value">整数。</param>
    ''' <returns>3 桁区切りの文字列。値が無いときは空文字。</returns>
    Private Shared Function FormatInteger(value As Integer?) As String
        ' 値が無いとき
        If Not value.HasValue Then
            Return String.Empty
        End If
        Return value.Value.ToString(IntegerFormat, CultureInfo.InvariantCulture)
    End Function

    ''' <summary>
    ''' 小数を含む数を帳票の書式にします。
    ''' </summary>
    ''' <param name="value">数値。</param>
    ''' <returns>3 桁区切りの文字列。値が無いときは空文字。</returns>
    Private Shared Function FormatDecimal(value As Decimal?) As String
        ' 値が無いとき
        If Not value.HasValue Then
            Return String.Empty
        End If
        Return value.Value.ToString(DecimalFormat, CultureInfo.InvariantCulture)
    End Function

    ''' <summary>
    ''' ファイル名に使えない文字を取り除きます。
    ''' </summary>
    ''' <param name="value">元の値。</param>
    ''' <returns>ファイル名に使える文字だけにした値。</returns>
    ''' <remarks>
    ''' INVOICE 番号には `/` が入ることがあり、そのままではパスとして解釈されます。
    ''' 値から名前を組み立てるため、ここで必ず落とします。
    ''' </remarks>
    Private Shared Function SanitizeFileNamePart(value As String) As String
        ' 値が無いとき
        If String.IsNullOrEmpty(value) Then
            Return String.Empty
        End If

        Dim invalid As Char() = Path.GetInvalidFileNameChars()
        Dim builder As New StringBuilder(value.Length)
        For Each c As Char In value
            ' 使えない文字のとき
            If Array.IndexOf(invalid, c) >= 0 Then
                Continue For
            End If
            builder.Append(c)
        Next
        Return builder.ToString()
    End Function
#End Region

End Class

''' <summary>
''' 入荷明細書の配置結果です。
''' </summary>
Public Class NyukaMeisaiWriteResult

    ''' <summary>配置した連携ファイルのフルパス。元INV_NO につき 1 本です。</summary>
    Public ReadOnly Property FilePaths As IList(Of String) = New List(Of String)()

    ''' <summary>
    ''' 1 本でも配置したかどうかです。
    ''' </summary>
    ''' <returns>配置したとき True。</returns>
    Public ReadOnly Property HasOutput As Boolean
        Get
            Return Me.FilePaths.Count > 0
        End Get
    End Property

End Class
