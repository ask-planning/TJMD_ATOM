Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports Atom.Core.Entities
Imports Atom.Core.Exceptions
Imports NPOI.SS.UserModel

Namespace Services

    ''' <summary>
    ''' INVOICE / PACKING の Excel を解析して取込行の一覧を作ります。
    ''' 旧 Access 版の DoCmd.TransferSpreadsheet に相当します。
    ''' </summary>
    ''' <remarks>
    ''' .xls（HSSF）と .xlsx（XSSF）の両方に対応します。
    ''' 1 行目はヘッダー行として読み飛ばし、先頭列が空の行は明細の終わりとして扱わず読み飛ばします。
    ''' 列の位置は固定です。対応は docs/specs/101_INVOICE・PACKING取込.md を参照してください。
    ''' </remarks>
    Public NotInheritable Class ShanghaiImportExcelParser

#Region "定数"
        ''' <summary>ヘッダー行の行数。</summary>
        Private Const HeaderRowCount As Integer = 1

        ''' <summary>R3 購買伝票番号が未設定のときに使う値。</summary>
        Private Const NoPurchaseSlipNo As String = "9999999999"

        ''' <summary>INVOICE の形式が違うときの文言。</summary>
        Private Const InvoiceFormatMessage As String = "INVOICEのフォーマットが違います。"

        ''' <summary>PACKING の形式が違うときの文言。</summary>
        Private Const PackingFormatMessage As String = "PACKINGのフォーマットが違います。"

        ''' <summary>INVOICE の読み込みに失敗したときの文言。</summary>
        Private Const InvoiceReadMessage As String = "INVOICEのEXCELの取り込み処理のエラーです。EXCELの内容等を確認してください。"

        ''' <summary>PACKING の読み込みに失敗したときの文言。</summary>
        Private Const PackingReadMessage As String = "PACKINGのEXCELの取り込み処理のエラーです。EXCELの内容等を確認してください。"

        ''' <summary>INVOICE の列位置。</summary>
        Private Const InvoiceLineSeqColumn As Integer = 0
        Private Const InvoiceSoNoColumn As Integer = 1
        Private Const InvoiceSoNoSeqColumn As Integer = 2
        Private Const InvoiceShanghaiCodeColumn As Integer = 3
        Private Const InvoiceShanghaiCodeTextColumn As Integer = 4
        Private Const InvoiceKikakuColumn As Integer = 5
        Private Const InvoicePartSortColumn As Integer = 6
        Private Const InvoiceSuryoTaniColumn As Integer = 7
        Private Const InvoiceSuryoColumn As Integer = 8
        Private Const InvoiceUnitPriceColumn As Integer = 9
        Private Const InvoiceAmountColumn As Integer = 10
        Private Const InvoiceDummyNo1Column As Integer = 11
        Private Const InvoiceDummyNo2Column As Integer = 12
        Private Const InvoiceTajimaPoNoColumn As Integer = 13
        Private Const InvoiceR3KoubaiDenpyoNoColumn As Integer = 14
        Private Const InvoiceInvoiceNoColumn As Integer = 15
        Private Const InvoiceBlDateColumn As Integer = 16
        Private Const InvoiceSyukkaHohoColumn As Integer = 17

        ''' <summary>PACKING の列位置。</summary>
        Private Const PackingTrueOrFalseColumn As Integer = 0
        Private Const PackingInvoiceNoColumn As Integer = 1
        Private Const PackingTajimaPoNoColumn As Integer = 2
        Private Const PackingShanghaiCodeColumn As Integer = 3
        Private Const PackingSuryoColumn As Integer = 4
        Private Const PackingUnitPriceColumn As Integer = 5
        Private Const PackingAmountColumn As Integer = 6
        Private Const PackingCartonNoFromColumn As Integer = 7
        Private Const PackingCartonNoToColumn As Integer = 8
        Private Const PackingNetWeightColumn As Integer = 9
        Private Const PackingGrossWeightColumn As Integer = 10
        Private Const PackingM3Column As Integer = 11
        Private Const PackingBlDateColumn As Integer = 12
        Private Const PackingSyukkaHohoColumn As Integer = 13
#End Region

#Region "フィールド"
        ''' <summary>日付として認める書式。Excel が文字列で日付を持っている場合に使います。</summary>
        Private Shared ReadOnly _dateFormats As String() = {
            "yy/MM/dd", "yyyy/MM/dd", "yy/M/d", "yyyy/M/d",
            "yyyy-MM-dd", "yyyyMMdd", "yy.MM.dd", "yyyy.MM.dd"
        }
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' INVOICE の Excel を解析します。
        ''' </summary>
        ''' <param name="stream">Excel ファイルのストリーム。</param>
        ''' <param name="tsuka">画面で選択した通貨。JPY または RMB。</param>
        ''' <returns>取込行の一覧。1 件も無いときは空の一覧。</returns>
        Public Function ParseInvoice(stream As Stream, tsuka As String) As IList(Of ShanghaiInvoiceImportRow)
            Dim results As New List(Of ShanghaiInvoiceImportRow)()
            Dim sheet As ISheet = Me.OpenFirstSheet(stream, InvoiceReadMessage)

            Try
                For rowIndex As Integer = HeaderRowCount To sheet.LastRowNum
                    Dim row As IRow = sheet.GetRow(rowIndex)
                    ' 行が存在しないとき
                    If row Is Nothing Then
                        Continue For
                    End If
                    ' 先頭列が空のとき（明細ではない行）
                    If Me.GetString(row, InvoiceLineSeqColumn).Length = 0 Then
                        Continue For
                    End If

                    results.Add(Me.CreateInvoiceRow(row, tsuka))
                Next
            Catch ex As ImportFormatException
                Throw
            Catch ex As Exception
                Throw New ImportFormatException(InvoiceFormatMessage, ex)
            End Try

            Return results
        End Function

        ''' <summary>
        ''' PACKING の Excel を解析します。カートン番号の補完も行います。
        ''' </summary>
        ''' <param name="stream">Excel ファイルのストリーム。</param>
        ''' <returns>取込行の一覧。1 件も無いときは空の一覧。</returns>
        Public Function ParsePacking(stream As Stream) As IList(Of ShanghaiPackingImportRow)
            Dim results As New List(Of ShanghaiPackingImportRow)()
            Dim sheet As ISheet = Me.OpenFirstSheet(stream, PackingReadMessage)

            ' 直前の行のカートン番号。空欄の補完に使います。
            Dim previousCartonNoFrom As String = String.Empty
            Dim previousCartonNoTo As String = String.Empty

            Try
                For rowIndex As Integer = HeaderRowCount To sheet.LastRowNum
                    Dim row As IRow = sheet.GetRow(rowIndex)
                    ' 行が存在しないとき
                    If row Is Nothing Then
                        Continue For
                    End If
                    ' 先頭列が空のとき（明細ではない行）
                    If Me.GetString(row, PackingTrueOrFalseColumn).Length = 0 Then
                        Continue For
                    End If

                    Dim entity As ShanghaiPackingImportRow = Me.CreatePackingRow(row)
                    Dim cartonNoFrom As String = Me.GetString(row, PackingCartonNoFromColumn)
                    Dim cartonNoTo As String = Me.GetString(row, PackingCartonNoToColumn)

                    entity.CartonNoFrom = Me.ResolveCartonNoFrom(cartonNoFrom, previousCartonNoFrom)
                    entity.CartonNoTo = Me.ResolveCartonNoTo(cartonNoTo, previousCartonNoTo)
                    results.Add(entity)

                    ' 値が入っていた行だけを次の行の引き継ぎ元にする
                    If cartonNoFrom.Length > 0 Then
                        previousCartonNoFrom = cartonNoFrom
                    End If
                    If cartonNoTo.Length > 0 Then
                        previousCartonNoTo = cartonNoTo
                    End If
                Next
            Catch ex As ImportFormatException
                Throw
            Catch ex As Exception
                Throw New ImportFormatException(PackingFormatMessage, ex)
            End Try

            Return results
        End Function
#End Region

#Region "内部処理"
        ''' <summary>
        ''' Excel を開いて先頭のシートを返します。
        ''' </summary>
        ''' <param name="stream">Excel ファイルのストリーム。</param>
        ''' <param name="errorMessage">開けなかったときに表示する文言。</param>
        ''' <returns>先頭のシート。</returns>
        Private Function OpenFirstSheet(stream As Stream, errorMessage As String) As ISheet
            ' ストリームが無いとき
            If stream Is Nothing Then
                Throw New ImportFormatException(errorMessage)
            End If

            Dim workbook As IWorkbook
            Try
                ' .xls と .xlsx のどちらでも開けるよう形式を自動判定させる
                workbook = WorkbookFactory.Create(stream)
            Catch ex As Exception
                Throw New ImportFormatException(errorMessage, ex)
            End Try

            ' シートが 1 つも無いとき
            If workbook.NumberOfSheets <= 0 Then
                Throw New ImportFormatException(errorMessage)
            End If

            Dim sheet As ISheet = workbook.GetSheetAt(0)
            ' 明細行が無いとき（ヘッダー行しかない、または空）
            If sheet Is Nothing OrElse sheet.LastRowNum < HeaderRowCount Then
                Throw New ImportFormatException(errorMessage)
            End If
            Return sheet
        End Function

        ''' <summary>
        ''' INVOICE の 1 行を作成します。
        ''' </summary>
        ''' <param name="row">Excel の行。</param>
        ''' <param name="tsuka">画面で選択した通貨。</param>
        ''' <returns>作成した取込行。</returns>
        Private Function CreateInvoiceRow(row As IRow, tsuka As String) As ShanghaiInvoiceImportRow
            Dim entity As New ShanghaiInvoiceImportRow()
            Dim shanghaiCode As String = Me.GetString(row, InvoiceShanghaiCodeColumn)

            entity.LineSeq = Me.GetNullableDouble(row, InvoiceLineSeqColumn)
            entity.SoNo = Me.GetString(row, InvoiceSoNoColumn)
            entity.SoNoSeq = Me.GetNullableInteger(row, InvoiceSoNoSeqColumn)
            entity.ShanghaiCode = Me.FormatShanghaiCode(shanghaiCode)
            entity.ShanghaiCodeText = Me.GetString(row, InvoiceShanghaiCodeTextColumn)
            entity.Kikaku = Me.GetString(row, InvoiceKikakuColumn)
            entity.PartSort = Me.GetString(row, InvoicePartSortColumn)
            entity.SuryoTani = Me.GetString(row, InvoiceSuryoTaniColumn)
            entity.Suryo = Me.GetNullableDecimal(row, InvoiceSuryoColumn)
            entity.UnitPrice = Me.GetNullableDouble(row, InvoiceUnitPriceColumn)
            entity.Amount = Me.GetNullableDouble(row, InvoiceAmountColumn)
            entity.DummyNo1 = Me.GetString(row, InvoiceDummyNo1Column)
            entity.DummyNo2 = Me.GetString(row, InvoiceDummyNo2Column)
            entity.TajimaPoNo = Me.GetString(row, InvoiceTajimaPoNoColumn)
            entity.R3KoubaiDenpyoNo = Me.ResolvePurchaseSlipNo(Me.GetString(row, InvoiceR3KoubaiDenpyoNoColumn))
            entity.InvoiceNo = Me.GetString(row, InvoiceInvoiceNoColumn)
            entity.BlDate = Me.GetNullableDate(row, InvoiceBlDateColumn)
            entity.SyukkaHoho = Me.GetString(row, InvoiceSyukkaHohoColumn)

            entity.R3HinmokuCode = Me.FormatR3HinmokuCode(shanghaiCode)
            entity.TajimaPoNoOriginal = entity.TajimaPoNo
            entity.AmountText = Me.GetString(row, InvoiceAmountColumn)
            entity.Tsuka = tsuka
            Return entity
        End Function

        ''' <summary>
        ''' PACKING の 1 行を作成します。カートン番号は呼び出し側で設定します。
        ''' </summary>
        ''' <param name="row">Excel の行。</param>
        ''' <returns>作成した取込行。</returns>
        Private Function CreatePackingRow(row As IRow) As ShanghaiPackingImportRow
            Dim entity As New ShanghaiPackingImportRow()
            Dim shanghaiCode As String = Me.GetString(row, PackingShanghaiCodeColumn)

            entity.TrueOrFalse = Me.GetString(row, PackingTrueOrFalseColumn)
            entity.InvoiceNo = Me.GetString(row, PackingInvoiceNoColumn)
            entity.TajimaPoNo = Me.GetString(row, PackingTajimaPoNoColumn)
            entity.ShanghaiCode = Me.FormatShanghaiCode(shanghaiCode)
            entity.Suryo = Me.GetNullableDecimal(row, PackingSuryoColumn)
            entity.UnitPrice = Me.GetNullableDouble(row, PackingUnitPriceColumn)
            entity.Amount = Me.GetNullableDouble(row, PackingAmountColumn)
            entity.NetWeight = Me.GetNullableDouble(row, PackingNetWeightColumn)
            entity.GrossWeight = Me.GetNullableDouble(row, PackingGrossWeightColumn)
            entity.M3 = Me.GetNullableDouble(row, PackingM3Column)
            entity.BlDate = Me.GetNullableDate(row, PackingBlDateColumn)
            entity.SyukkaHoho = Me.GetString(row, PackingSyukkaHohoColumn)

            entity.R3HinmokuCode = Me.FormatR3HinmokuCode(shanghaiCode)
            entity.TajimaPoNoOriginal = entity.TajimaPoNo
            Return entity
        End Function

        ''' <summary>
        ''' 上海コードを整えます。数値のときは 5 桁になるようゼロ埋めします。
        ''' </summary>
        ''' <param name="value">Excel から読み取った値。</param>
        ''' <returns>整えた上海コード。</returns>
        ''' <remarks>
        ''' 取込結果確認画面での編集時にも同じ規則が必要なため、
        ''' 実体は ShanghaiCodeFormatter へ集約しています。
        ''' </remarks>
        Private Function FormatShanghaiCode(value As String) As String
            Return ShanghaiCodeFormatter.FormatShanghaiCode(value)
        End Function

        ''' <summary>
        ''' R3 品目コードを作ります。上海コードが数値のときは 18 桁になるようゼロ埋めします。
        ''' </summary>
        ''' <param name="shanghaiCode">上海コード。</param>
        ''' <returns>R3 品目コード。</returns>
        ''' <remarks>
        ''' 取込結果確認画面での編集時にも同じ規則が必要なため、
        ''' 実体は ShanghaiCodeFormatter へ集約しています。
        ''' </remarks>
        Private Function FormatR3HinmokuCode(shanghaiCode As String) As String
            Return ShanghaiCodeFormatter.FormatR3HinmokuCode(shanghaiCode)
        End Function

        ''' <summary>
        ''' R3 購買伝票番号を決めます。未設定のときは固定値を返します。
        ''' </summary>
        ''' <param name="value">Excel から読み取った値。</param>
        ''' <returns>R3 購買伝票番号。</returns>
        Private Function ResolvePurchaseSlipNo(value As String) As String
            ' 未設定のとき（R3 未発注品）
            If value.Length = 0 Then
                Return NoPurchaseSlipNo
            End If
            Return value
        End Function

        ''' <summary>
        ''' 開始カートン番号を決めます。空欄のときは直前の行の値を引き継ぎます。
        ''' </summary>
        ''' <param name="value">Excel から読み取った値。</param>
        ''' <param name="previousValue">直前の行の開始カートン番号。</param>
        ''' <returns>開始カートン番号。引き継ぐ値も無いときは空文字。</returns>
        Private Function ResolveCartonNoFrom(value As String, previousValue As String) As String
            ' 値が入っているとき
            If value.Length > 0 Then
                Return value
            End If
            ' 空欄のとき（直前の値を引き継ぐ）
            Return previousValue
        End Function

        ''' <summary>
        ''' 終了カートン番号を決めます。直前の行と同じ値のときは設定しません。
        ''' </summary>
        ''' <param name="value">Excel から読み取った値。</param>
        ''' <param name="previousValue">直前の行の終了カートン番号。</param>
        ''' <returns>終了カートン番号。設定しないときは空文字。</returns>
        Private Function ResolveCartonNoTo(value As String, previousValue As String) As String
            ' 空欄のとき
            If value.Length = 0 Then
                Return String.Empty
            End If
            ' 直前と同じ値のとき（重複値として扱わない）
            If String.Equals(value, previousValue, StringComparison.Ordinal) Then
                Return String.Empty
            End If
            Return value
        End Function

        ''' <summary>
        ''' セルの値を文字列で取得します。前後の空白は除きます。
        ''' </summary>
        ''' <param name="row">Excel の行。</param>
        ''' <param name="columnIndex">列位置。</param>
        ''' <returns>セルの値。空のときは空文字。</returns>
        Private Function GetString(row As IRow, columnIndex As Integer) As String
            Dim cell As ICell = row.GetCell(columnIndex)
            ' セルが無いとき
            If cell Is Nothing Then
                Return String.Empty
            End If

            Select Case cell.CellType
                ' 空白のとき
                Case CellType.Blank
                    Return String.Empty
                ' 数値のとき
                Case CellType.Numeric
                    ' 日付書式のとき
                    If DateUtil.IsCellDateFormatted(cell) Then
                        Return cell.DateCellValue.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture)
                    End If
                    Return cell.NumericCellValue.ToString("0.############", CultureInfo.InvariantCulture)
                ' 真偽値のとき
                Case CellType.Boolean
                    Return cell.BooleanCellValue.ToString()
                ' 数式のとき
                Case CellType.Formula
                    Return Me.GetFormulaString(cell)
                ' 上記以外（文字列・エラー）のとき
                Case Else
                    Dim text As String = cell.ToString()
                    ' 値が無いとき
                    If text Is Nothing Then
                        Return String.Empty
                    End If
                    Return text.Trim()
            End Select
        End Function

        ''' <summary>
        ''' 数式セルの計算結果を文字列で取得します。
        ''' </summary>
        ''' <param name="cell">対象のセル。</param>
        ''' <returns>計算結果の文字列。取得できないときは空文字。</returns>
        Private Function GetFormulaString(cell As ICell) As String
            Select Case cell.CachedFormulaResultType
                ' 数値のとき
                Case CellType.Numeric
                    ' 日付書式のとき
                    If DateUtil.IsCellDateFormatted(cell) Then
                        Return cell.DateCellValue.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture)
                    End If
                    Return cell.NumericCellValue.ToString("0.############", CultureInfo.InvariantCulture)
                ' 文字列のとき
                Case CellType.String
                    Dim text As String = cell.StringCellValue
                    ' 値が無いとき
                    If text Is Nothing Then
                        Return String.Empty
                    End If
                    Return text.Trim()
                ' 上記以外のとき
                Case Else
                    Return String.Empty
            End Select
        End Function

        ''' <summary>
        ''' セルの値を数値で取得します。
        ''' </summary>
        ''' <param name="row">Excel の行。</param>
        ''' <param name="columnIndex">列位置。</param>
        ''' <returns>セルの値。数値として読めないときは Nothing。</returns>
        Private Function GetNullableDouble(row As IRow, columnIndex As Integer) As Double?
            Dim text As String = Me.GetString(row, columnIndex)
            ' 空のとき
            If text.Length = 0 Then
                Return Nothing
            End If

            Dim parsed As Double = 0
            ' 数値として読めたとき
            If Double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, parsed) Then
                Return parsed
            End If
            ' 読めないとき
            Return Nothing
        End Function

        ''' <summary>
        ''' セルの値を 10 進数で取得します。
        ''' </summary>
        ''' <param name="row">Excel の行。</param>
        ''' <param name="columnIndex">列位置。</param>
        ''' <returns>セルの値。数値として読めないときは Nothing。</returns>
        Private Function GetNullableDecimal(row As IRow, columnIndex As Integer) As Decimal?
            Dim text As String = Me.GetString(row, columnIndex)
            ' 空のとき
            If text.Length = 0 Then
                Return Nothing
            End If

            Dim parsed As Decimal = 0
            ' 数値として読めたとき
            If Decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, parsed) Then
                Return parsed
            End If
            ' 読めないとき
            Return Nothing
        End Function

        ''' <summary>
        ''' セルの値を整数で取得します。
        ''' </summary>
        ''' <param name="row">Excel の行。</param>
        ''' <param name="columnIndex">列位置。</param>
        ''' <returns>セルの値。数値として読めないときは Nothing。</returns>
        Private Function GetNullableInteger(row As IRow, columnIndex As Integer) As Integer?
            Dim value As Double? = Me.GetNullableDouble(row, columnIndex)
            ' 数値として読めないとき
            If Not value.HasValue Then
                Return Nothing
            End If
            Return CInt(Math.Truncate(value.Value))
        End Function

        ''' <summary>
        ''' セルの値を日付で取得します。日付書式のセルと文字列のどちらにも対応します。
        ''' </summary>
        ''' <param name="row">Excel の行。</param>
        ''' <param name="columnIndex">列位置。</param>
        ''' <returns>セルの値。日付として読めないときは Nothing。</returns>
        Private Function GetNullableDate(row As IRow, columnIndex As Integer) As Date?
            Dim cell As ICell = row.GetCell(columnIndex)
            ' セルが無いとき
            If cell Is Nothing Then
                Return Nothing
            End If

            ' 日付書式の数値セルのとき
            If cell.CellType = CellType.Numeric AndAlso DateUtil.IsCellDateFormatted(cell) Then
                Return cell.DateCellValue
            End If

            Dim text As String = Me.GetString(row, columnIndex)
            ' 空のとき
            If text.Length = 0 Then
                Return Nothing
            End If

            Dim parsed As Date = Date.MinValue
            ' 想定した書式で読めたとき
            If Date.TryParseExact(text, _dateFormats, CultureInfo.InvariantCulture,
                                  DateTimeStyles.None, parsed) Then
                Return parsed
            End If
            ' 想定外の書式のとき（環境の書式で読めるかを試す）
            If Date.TryParse(text, parsed) Then
                Return parsed
            End If
            ' 読めないとき
            Return Nothing
        End Function
#End Region

    End Class
End Namespace
