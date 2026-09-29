Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text

Namespace Services

    ''' <summary>
    ''' 上海輸入の INVOICE 番号の編集を行います。
    ''' </summary>
    ''' <remarks>
    ''' 画面表示用の REF_NO と、明細テーブルの pk_invoice_no_main は別のルールで切り出します。
    ''' 詳細は docs/specs/101_INVOICE・PACKING取込.md を参照してください。
    ''' </remarks>
    Public NotInheritable Class ShanghaiInvoiceNoFormatter

#Region "定数"
        ''' <summary>担当区分のうち貿易を表す値。</summary>
        Public Const BouekiKubun As Integer = 3

        ''' <summary>商品群の区切り文字。</summary>
        Private Const Separator As String = "/"

        ''' <summary>通貨記号が無いときの共通部分の桁数。</summary>
        Private Const ShortKeyLength As Integer = 8

        ''' <summary>通貨記号があるときの共通部分の桁数。</summary>
        Private Const LongKeyLength As Integer = 9

        ''' <summary>共通部分の桁数を判定する位置（1 始まり）。</summary>
        Private Const RefNoCheckPosition As Integer = 5
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' 明細テーブルへ登録する INVOICE 番号（メイン）を求めます。
        ''' </summary>
        ''' <param name="invoiceNo">INVOICE 番号。</param>
        ''' <returns>INVOICE 番号（メイン）。判定できないときは元の値。</returns>
        ''' <remarks>
        ''' 本登録のストアドと同じ判定（9 文字目が数値なら先頭 9 文字、そうでなければ先頭 8 文字）にします。
        ''' </remarks>
        Public Function GetInvoiceNoMain(invoiceNo As String) As String
            ' 桁数が足りないとき
            If String.IsNullOrEmpty(invoiceNo) OrElse invoiceNo.Length < ShortKeyLength Then
                Return If(invoiceNo, String.Empty)
            End If
            ' 9 文字に届かないとき
            If invoiceNo.Length < LongKeyLength Then
                Return invoiceNo.Substring(0, ShortKeyLength)
            End If

            ' 9 文字目が数値のとき
            If Me.IsDigit(invoiceNo.Substring(LongKeyLength - 1, 1)) Then
                Return invoiceNo.Substring(0, LongKeyLength)
            End If
            ' 9 文字目が数値でないとき
            Return invoiceNo.Substring(0, ShortKeyLength)
        End Function

        ''' <summary>
        ''' 画面に表示する REF_NO を作ります。
        ''' </summary>
        ''' <param name="invoiceNos">一時データに含まれる INVOICE 番号の一覧。昇順であること。</param>
        ''' <param name="shanghaiChotatsu">担当区分。3 のとき貿易として編集します。</param>
        ''' <returns>表示用の文字列。対象が無いときは空文字。</returns>
        Public Function CreateRefNo(invoiceNos As IList(Of String), shanghaiChotatsu As Integer) As String
            ' 対象が無いとき
            If invoiceNos Is Nothing OrElse invoiceNos.Count = 0 Then
                Return String.Empty
            End If

            Dim isBoueki As Boolean = (shanghaiChotatsu = BouekiKubun)
            Dim keyLength As Integer = Me.GetKeyLength(invoiceNos(0), isBoueki)
            ' 共通部分を切り出せないとき
            If keyLength <= 0 Then
                Return invoiceNos(0)
            End If

            Dim builder As New StringBuilder()
            Dim currentKey As String = String.Empty

            For Each invoiceNo As String In invoiceNos
                ' 値が無いとき
                If String.IsNullOrEmpty(invoiceNo) Then
                    Continue For
                End If

                Dim key As String = Me.SafeSubstring(invoiceNo, 0, keyLength)

                ' 1 件目のとき
                If builder.Length = 0 Then
                    builder.Append(invoiceNo)
                    currentKey = key
                    Continue For
                End If

                ' 共通部分が変わったとき（別の INVOICE として区切る）
                If Not String.Equals(key, currentKey, StringComparison.Ordinal) Then
                    builder.Append(Separator)
                    builder.Append(invoiceNo)
                    currentKey = key
                    Continue For
                End If

                ' 共通部分が同じとき（商品群だけを足す）
                Dim suffix As String = String.Empty
                If invoiceNo.Length > keyLength Then
                    suffix = invoiceNo.Substring(keyLength)
                End If
                ' 貿易は区切り文字を付けず、それ以外は付ける
                If Not isBoueki Then
                    builder.Append(Separator)
                End If
                builder.Append(suffix)
            Next

            Return builder.ToString()
        End Function
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 共通部分として扱う桁数を求めます。
        ''' </summary>
        ''' <param name="invoiceNo">先頭の INVOICE 番号。</param>
        ''' <param name="isBoueki">貿易のとき True。</param>
        ''' <returns>共通部分の桁数。求められないときは 0。</returns>
        Private Function GetKeyLength(invoiceNo As String, isBoueki As Boolean) As Integer
            ' 値が無いとき
            If String.IsNullOrEmpty(invoiceNo) Then
                Return 0
            End If

            ' 貿易のとき（商品群が区切り文字から始まる）
            If isBoueki Then
                Dim separatorIndex As Integer = invoiceNo.IndexOf(Separator, StringComparison.Ordinal)
                ' 区切り文字が無いとき
                If separatorIndex <= 0 Then
                    Return invoiceNo.Length
                End If
                Return separatorIndex
            End If

            ' 貿易以外のとき（5 文字目が数値なら通貨記号なしとみなす）
            If invoiceNo.Length < RefNoCheckPosition Then
                Return invoiceNo.Length
            End If
            If Me.IsDigit(invoiceNo.Substring(RefNoCheckPosition - 1, 1)) Then
                Return ShortKeyLength
            End If
            Return LongKeyLength
        End Function

        ''' <summary>
        ''' 1 文字が数字かどうかを返します。
        ''' </summary>
        ''' <param name="value">判定する 1 文字の文字列。</param>
        ''' <returns>数字のとき True。</returns>
        Private Function IsDigit(value As String) As Boolean
            Dim parsed As Integer = 0
            Return Integer.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, parsed)
        End Function

        ''' <summary>
        ''' 桁数を超えないように部分文字列を取り出します。
        ''' </summary>
        ''' <param name="value">元の文字列。</param>
        ''' <param name="startIndex">開始位置。</param>
        ''' <param name="length">取り出す長さ。</param>
        ''' <returns>取り出した文字列。</returns>
        Private Function SafeSubstring(value As String, startIndex As Integer, length As Integer) As String
            ' 元の文字列が短いとき
            If value.Length <= startIndex Then
                Return String.Empty
            End If
            ' 指定より短いとき
            If value.Length < startIndex + length Then
                Return value.Substring(startIndex)
            End If
            Return value.Substring(startIndex, length)
        End Function
#End Region

    End Class
End Namespace
