Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text
Imports Atom.Core.Entities

''' <summary>
''' 配送手配書（menuNo=103 §10 / §14）の連携ファイルを組み立てて配置します。
''' </summary>
''' <remarks>
''' 配送先の箇所数で様式が決まります（設計書 §10.4）。
''' 様式ごとに連携ファイルを 1 本ずつ作り、同じ様式の行はまとめて 1 本に入れます。
''' SVF 側は 1 行につき 1 ページを出力します。
'''
''' 列の並びと名前は `tools/svf/定義AT01配送手配書.txt` と一致させてください。
''' ここが 1 つでもずれると、SVF 側では原因が分かりません。
''' </remarks>
Public Class HaisoTehaiReportWriter

#Region "定数"
    ''' <summary>ファイル名に使う日時の書式。</summary>
    Private Const TimestampFormat As String = "yyyyMMdd_HHmmss"

    ''' <summary>発行日の書式。</summary>
    Private Const HakkoYmdFormat As String = "yyyy年M月d日"

    ''' <summary>B/L DATE の書式。</summary>
    Private Const BlDateFormat As String = "yyyy/MM/dd"

    ''' <summary>配送予定日の書式。曜日は別に足します。</summary>
    Private Const HaisoYoteiFormat As String = "M/d"

    ''' <summary>氏名などをつなぐ全角空白。</summary>
    Private Const WideSpace As String = "　"

    ''' <summary>SVF へ渡すドキュメント名。</summary>
    Private Const DocumentName As String = "配送手配書"

    ''' <summary>日本語の書式に使う地域。曜日を日本語で出すために指定します。</summary>
    Private Shared ReadOnly _japaneseCulture As CultureInfo = CultureInfo.GetCultureInfo("ja-JP")

    ''' <summary>
    ''' AIR の配送タイプです。設計書 §10.7 の一覧から取りました。
    ''' </summary>
    ''' <remarks>
    ''' 旧 Access は `haiso_type >= 10` で絞り込んでおり、**非 AIR が混ざる不具合**がありました
    ''' （設計書 §10.8）。Web 版では値を明示して判定します。
    ''' </remarks>
    Private Shared ReadOnly _airHaisoTypes As New HashSet(Of Integer)() From {20, 21, 22, 23, 24, 25, 27, 28}
#End Region

#Region "フィールド"
    ''' <summary>連携ファイルの書き出し。</summary>
    Private ReadOnly _writer As New SvfDataFileWriter()
#End Region

#Region "公開メソッド"
    ''' <summary>
    ''' 連携ファイルを組み立てて監視フォルダへ配置します。
    ''' </summary>
    ''' <param name="settings">SVF の設定。</param>
    ''' <param name="rows">帳票に印字する行。</param>
    ''' <param name="loginCode">利用者のログインコード。</param>
    ''' <param name="hakkoYmd">発行日。</param>
    ''' <returns>配置した結果。</returns>
    Public Function Write(settings As SvfReportSettings,
                          rows As IList(Of ShanghaiHaisoTehaiReportRow),
                          loginCode As String,
                          hakkoYmd As Date) As HaisoTehaiWriteResult

        Dim result As New HaisoTehaiWriteResult()
        ' 対象が無いとき
        If rows Is Nothing OrElse rows.Count = 0 Then
            Return result
        End If

        ' 配送先の箇所数でまとめる。箇所数の昇順で出したいため SortedDictionary を使う
        Dim rowsByCount As New SortedDictionary(Of Integer, List(Of ShanghaiHaisoTehaiReportRow))()

        For Each row As ShanghaiHaisoTehaiReportRow In rows
            ' AIR のとき（様式が未作成のため、出さずに控えるだけ）
            If Me.IsAir(row) Then
                result.SkippedAirCount += 1
                Continue For
            End If

            Dim count As Integer = row.HaisoSakiCount
            ' その箇所数が初めて出たとき
            If Not rowsByCount.ContainsKey(count) Then
                rowsByCount.Add(count, New List(Of ShanghaiHaisoTehaiReportRow)())
            End If
            rowsByCount(count).Add(row)
        Next

        For Each pair As KeyValuePair(Of Integer, List(Of ShanghaiHaisoTehaiReportRow)) In rowsByCount
            Dim form As HaisoTehaiForm = HaisoTehaiForm.FromHaisoSakiCount(pair.Key)
            Dim filePath As String = Me.WriteOneForm(settings, form, pair.Value, loginCode, hakkoYmd)
            result.Forms.Add(form)
            result.FilePaths.Add(filePath)
        Next

        Return result
    End Function

    ''' <summary>
    ''' 連携ファイルの見出し行を返します。
    ''' </summary>
    ''' <returns>27 列の項目名。</returns>
    ''' <remarks>
    ''' `tools/svf/定義AT01配送手配書.txt` および `schema.ini` の `[定義AT01配送手配書.txt]` と
    ''' **同じ並び・同じ名前**でなければなりません。
    ''' </remarks>
    Public Shared Function CreateHeaders() As IList(Of String)
        Dim headers As New List(Of String)()
        headers.Add("帳票区分")
        headers.Add("login_code")
        headers.Add("hakko_ymd")
        headers.Add("atesaki")
        headers.Add("phone_fax")
        headers.Add("ref_no")
        headers.Add("invoice_no")
        headers.Add("vessel")
        headers.Add("bl_date")
        headers.Add("container_type")
        headers.Add("container_no")
        headers.Add("seal_no")
        headers.Add("total_carton")
        headers.Add("haiso_yotei")
        For number As Integer = 1 To 3
            Dim prefix As String = "haiso_saki" & number.ToString(CultureInfo.InvariantCulture) & "_"
            headers.Add(prefix & "name")
            headers.Add(prefix & "address")
            headers.Add(prefix & "phone_no")
            headers.Add(prefix & "tanto")
        Next
        headers.Add("tantosha")
        Return headers
    End Function
#End Region

#Region "内部処理"
    ''' <summary>
    ''' 様式 1 つ分の連携ファイルを配置します。
    ''' </summary>
    ''' <param name="settings">SVF の設定。</param>
    ''' <param name="form">様式。</param>
    ''' <param name="rows">その様式で出す行。</param>
    ''' <param name="loginCode">利用者のログインコード。</param>
    ''' <param name="hakkoYmd">発行日。</param>
    ''' <returns>配置したファイルのフルパス。</returns>
    Private Function WriteOneForm(settings As SvfReportSettings,
                                  form As HaisoTehaiForm,
                                  rows As IList(Of ShanghaiHaisoTehaiReportRow),
                                  loginCode As String,
                                  hakkoYmd As Date) As String

        Dim controlLines As New List(Of String)()
        controlLines.Add(SvfDataFileWriter.CreateQuotedControlLine("VrSetOutputPrinter", settings.OutputPrinter))
        controlLines.Add(SvfDataFileWriter.CreateControlLine("VrSetUseReportDirector", "1"))
        controlLines.Add(SvfDataFileWriter.CreateQuotedControlLine("VrSetDocName2", DocumentName))

        Dim dataRows As New List(Of IList(Of String))()
        For Each row As ShanghaiHaisoTehaiReportRow In rows
            dataRows.Add(Me.CreateDataRow(form, row, loginCode, hakkoYmd))
        Next

        Dim fileName As String = Me.CreateFileName(form, rows(0), hakkoYmd)
        Return Me._writer.Write(settings.WatchFolderPath, fileName, controlLines,
                                CreateHeaders(), dataRows, loginCode)
    End Function

    ''' <summary>
    ''' データ行 1 行分を組み立てます。
    ''' </summary>
    ''' <param name="form">様式。</param>
    ''' <param name="row">帳票の行。</param>
    ''' <param name="loginCode">利用者のログインコード。</param>
    ''' <param name="hakkoYmd">発行日。</param>
    ''' <returns>27 列の値。見出しと同じ並びです。</returns>
    Private Function CreateDataRow(form As HaisoTehaiForm,
                                   row As ShanghaiHaisoTehaiReportRow,
                                   loginCode As String,
                                   hakkoYmd As Date) As IList(Of String)

        Dim values As New List(Of String)()
        values.Add(form.ReportKubun)
        values.Add(loginCode)
        values.Add(hakkoYmd.ToString(HakkoYmdFormat, _japaneseCulture))
        values.Add(Me.CreateAtesaki(row))
        values.Add(Me.CreatePhoneFax(row))
        values.Add(row.RefNo)
        values.Add(row.InvoiceNo)
        values.Add(row.Vessel)
        values.Add(FormatDate(row.BlDate, BlDateFormat))
        values.Add(row.ContainerType)
        values.Add(row.ContainerNo)
        values.Add(row.SealNo)
        values.Add(FormatNumber(row.TotalCarton))
        values.Add(Me.CreateHaisoYotei(row))
        Me.AddHaisoSaki(values, row.HaisoSaki1)
        Me.AddHaisoSaki(values, row.HaisoSaki2)
        Me.AddHaisoSaki(values, row.HaisoSaki3)
        values.Add(Me.CreateTantosha(row))
        Return values
    End Function

    ''' <summary>
    ''' 配送先 1 箇所分の 4 列を足します。
    ''' </summary>
    ''' <param name="values">組み立て中の値。</param>
    ''' <param name="saki">配送先。</param>
    Private Sub AddHaisoSaki(values As IList(Of String), saki As ShanghaiHaisoSaki)
        values.Add(saki.Name)
        values.Add(saki.Address)
        values.Add(saki.PhoneNo)
        values.Add(saki.Tanto)
    End Sub

    ''' <summary>
    ''' 宛名を組み立てます。
    ''' </summary>
    ''' <param name="row">帳票の行。</param>
    ''' <returns>乙仲名と担当を全角空白でつないだ文字列。</returns>
    ''' <remarks>
    ''' 旧 Access は帳票側で 3 つの枠を並べていました。Web 版は 1 つの枠に収めるため
    ''' ここで連結します（設計書 §14.1.1）。**空の項目は区切りごと詰めます。**
    ''' </remarks>
    Private Function CreateAtesaki(row As ShanghaiHaisoTehaiReportRow) As String
        Return JoinNonEmpty(WideSpace, row.OtsunakaNoText, row.OtsunakaTanto1, row.OtsunakaTanto2)
    End Function

    ''' <summary>
    ''' 電話番号と FAX 番号を組み立てます。
    ''' </summary>
    ''' <param name="row">帳票の行。</param>
    ''' <returns>ラベルを含む文字列。</returns>
    ''' <remarks>ラベルも帳票の枠ではなくこちらに含めます（設計書 §14.1.1）。</remarks>
    Private Function CreatePhoneFax(row As ShanghaiHaisoTehaiReportRow) As String
        Return "PHONE: " & row.PhoneNo & " FAX : " & row.FaxNo
    End Function

    ''' <summary>
    ''' 配送予定日を組み立てます。
    ''' </summary>
    ''' <param name="row">帳票の行。</param>
    ''' <returns>「7/26(月)」の形の文字列。日付が無いときは空文字。</returns>
    ''' <remarks>
    ''' 曜日はビューの `weekday_no_text` を使わず、ここで日本語の曜日を作ります。
    ''' `week_day` の採番規則が読み取れていないためです（設計書 §18 C-38）。
    '''
    ''' **かっこは半角です。**日付部分が半角のため、全角だと並びが揃いません。
    ''' </remarks>
    Private Function CreateHaisoYotei(row As ShanghaiHaisoTehaiReportRow) As String
        ' 配送予定日が無いとき
        If Not row.HaisoYoteiYmd.HasValue Then
            Return String.Empty
        End If

        Dim value As Date = row.HaisoYoteiYmd.Value
        Return value.ToString(HaisoYoteiFormat, _japaneseCulture) &
               "(" & value.ToString("ddd", _japaneseCulture) & ")"
    End Function

    ''' <summary>
    ''' 発行者を組み立てます。
    ''' </summary>
    ''' <param name="row">帳票の行。</param>
    ''' <returns>部門名と氏名を全角空白でつないだ文字列。</returns>
    Private Function CreateTantosha(row As ShanghaiHaisoTehaiReportRow) As String
        Return JoinNonEmpty(WideSpace, row.BumonCodeText, row.DisplayName)
    End Function

    ''' <summary>
    ''' 連携ファイルのファイル名を組み立てます。
    ''' </summary>
    ''' <param name="form">様式。</param>
    ''' <param name="row">先頭の行。INVOICE 番号をファイル名に使います。</param>
    ''' <param name="hakkoYmd">発行日時。</param>
    ''' <returns>「接頭辞_参照番号_日時.txt」の形のファイル名。</returns>
    Private Function CreateFileName(form As HaisoTehaiForm,
                                    row As ShanghaiHaisoTehaiReportRow,
                                    hakkoYmd As Date) As String

        Dim key As String = SanitizeFileNamePart(row.InvoiceNo)
        Dim timestamp As String = hakkoYmd.ToString(TimestampFormat, CultureInfo.InvariantCulture)
        Return form.FilePrefix & "_" & key & "_" & timestamp & ".txt"
    End Function

    ''' <summary>
    ''' AIR かどうかを判定します。
    ''' </summary>
    ''' <param name="row">帳票の行。</param>
    ''' <returns>AIR のとき True。</returns>
    Private Function IsAir(row As ShanghaiHaisoTehaiReportRow) As Boolean
        ' 配送タイプが無いとき
        If Not row.HaisoType.HasValue Then
            Return False
        End If
        Return _airHaisoTypes.Contains(row.HaisoType.Value)
    End Function

    ''' <summary>
    ''' 空でない値だけをつなぎます。
    ''' </summary>
    ''' <param name="separator">区切り文字。</param>
    ''' <param name="values">つなぐ値。</param>
    ''' <returns>つないだ文字列。</returns>
    Private Shared Function JoinNonEmpty(separator As String, ParamArray values As String()) As String
        Dim builder As New StringBuilder()
        For Each value As String In values
            ' 値が無いとき
            If String.IsNullOrWhiteSpace(value) Then
                Continue For
            End If
            ' 2 件目以降のとき
            If builder.Length > 0 Then
                builder.Append(separator)
            End If
            builder.Append(value.Trim())
        Next
        Return builder.ToString()
    End Function

    ''' <summary>
    ''' 日付を書式化します。
    ''' </summary>
    ''' <param name="value">元の値。</param>
    ''' <param name="format">書式。</param>
    ''' <returns>書式化した文字列。値が無いときは空文字。</returns>
    Private Shared Function FormatDate(value As Date?, format As String) As String
        ' 値が無いとき
        If Not value.HasValue Then
            Return String.Empty
        End If
        Return value.Value.ToString(format, CultureInfo.InvariantCulture)
    End Function

    ''' <summary>
    ''' 数値を書式化します。
    ''' </summary>
    ''' <param name="value">元の値。</param>
    ''' <returns>書式化した文字列。値が無いときは空文字。</returns>
    Private Shared Function FormatNumber(value As Integer?) As String
        ' 値が無いとき
        If Not value.HasValue Then
            Return String.Empty
        End If
        Return value.Value.ToString(CultureInfo.InvariantCulture)
    End Function

    ''' <summary>
    ''' ファイル名に使えない文字を取り除きます。
    ''' </summary>
    ''' <param name="value">元の値。</param>
    ''' <returns>ファイル名に使える文字列。空になるときは「NONE」。</returns>
    ''' <remarks>
    ''' INVOICE 番号には `/` が入ることがあります。そのままだとパスとして解釈されるため落とします。
    ''' </remarks>
    Private Shared Function SanitizeFileNamePart(value As String) As String
        ' 値が無いとき
        If String.IsNullOrWhiteSpace(value) Then
            Return "NONE"
        End If

        Dim builder As New StringBuilder()
        Dim invalidChars As Char() = IO.Path.GetInvalidFileNameChars()
        For Each letter As Char In value
            ' 使えない文字のとき
            If Array.IndexOf(invalidChars, letter) >= 0 Then
                Continue For
            End If
            builder.Append(letter)
        Next

        ' 取り除いた結果が空になったとき
        If builder.Length = 0 Then
            Return "NONE"
        End If
        Return builder.ToString()
    End Function
#End Region

End Class
