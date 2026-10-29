Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text
Imports Atom.Core.Entities

Namespace Services

    ''' <summary>
    ''' 配送手配書の印刷対象を配送パターンごとにまとめます。
    ''' </summary>
    ''' <remarks>
    ''' 設計書 §10.3。INVOICE の枝番ごとに拾った行を、配送パターンで 1 行にまとめます。
    ''' カートン数と数量は合計し、INVOICE 番号は枝番をカンマで連結します。
    '''
    ''' データベースに依存しない処理のため、ここ（Atom.Core）に置いています。
    ''' </remarks>
    Public NotInheritable Class ShanghaiHaisoTehaiGrouper

#Region "定数"
        ''' <summary>枝番をつなぐ区切り文字。</summary>
        Private Const BranchSeparator As String = ","
#End Region

#Region "フィールド"
        ''' <summary>INVOICE 番号の編集。</summary>
        Private ReadOnly _invoiceNoFormatter As New ShanghaiInvoiceNoFormatter()
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' 印刷対象を配送パターンごとにまとめます。
        ''' </summary>
        ''' <param name="pickups">ビューから拾った行。並び順は呼び出し側で整えてください。</param>
        ''' <param name="executionDate">実行日。配送予定日が未入力のときに使います。</param>
        ''' <returns>印刷用の行。元の並び順を保ちます。</returns>
        ''' <remarks>
        ''' **配送パターンの単位は、帳票に印字する項目のうち合計しないものすべて**とします。
        ''' すなわち B/L DATE・船名・乙仲・配送予定日・配送タイプ・コンテナ区分・
        ''' コンテナ番号・シール番号・配送先 1〜3 です。
        ''' 旧 Access の VBA は「配送パターン」とだけ書かれており、**具体的な単位は確認中**です
        ''' （設計書 §18 C-36）。1 つでも違えば別の手配書になるため、この単位にしています。
        ''' </remarks>
        Public Function Group(pickups As IList(Of ShanghaiHaisoTehaiPickup),
                              executionDate As Date) As IList(Of ShanghaiHaisoTehaiPrintRow)

            Dim results As New List(Of ShanghaiHaisoTehaiPrintRow)()
            ' 対象が無いとき
            If pickups Is Nothing OrElse pickups.Count = 0 Then
                Return results
            End If

            Dim indexByKey As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
            Dim branchByKey As New Dictionary(Of String, StringBuilder)(StringComparer.Ordinal)
            Dim invoiceNosByKey As New Dictionary(Of String, HashSet(Of String))(StringComparer.Ordinal)

            For Each pickup As ShanghaiHaisoTehaiPickup In pickups
                Dim haisoYoteiYmd As Date = If(pickup.HaisoYoteiYmd, executionDate)
                Dim key As String = Me.CreateKey(pickup, haisoYoteiYmd)

                ' そのパターンが初めて出たとき
                If Not indexByKey.ContainsKey(key) Then
                    indexByKey.Add(key, results.Count)
                    branchByKey.Add(key, New StringBuilder(pickup.InvoiceNo))
                    invoiceNosByKey.Add(key, New HashSet(Of String)(StringComparer.Ordinal) From {pickup.InvoiceNo})
                    results.Add(Me.CreateRow(pickup, haisoYoteiYmd))
                    Continue For
                End If

                ' 同じ INVOICE 番号が再び出たとき（合計も連結も二重に行わない）
                If Not invoiceNosByKey(key).Add(pickup.InvoiceNo) Then
                    Continue For
                End If

                Dim row As ShanghaiHaisoTehaiPrintRow = results(indexByKey(key))
                row.TotalCarton += If(pickup.CartonQty, 0)
                row.TotalQuantity += If(pickup.TotalSuryo, 0)

                branchByKey(key).Append(BranchSeparator)
                branchByKey(key).Append(Me._invoiceNoFormatter.GetBranchNo(pickup.InvoiceNo))
            Next

            ' 連結した INVOICE 番号を書き戻す
            For Each pair As KeyValuePair(Of String, Integer) In indexByKey
                results(pair.Value).InvoiceNo = branchByKey(pair.Key).ToString()
            Next

            Return results
        End Function
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 配送パターンを表す文字列を組み立てます。
        ''' </summary>
        ''' <param name="pickup">拾った行。</param>
        ''' <param name="haisoYoteiYmd">配送予定日。未入力を補った後の値。</param>
        ''' <returns>比較に使う文字列。</returns>
        Private Function CreateKey(pickup As ShanghaiHaisoTehaiPickup, haisoYoteiYmd As Date) As String
            Dim builder As New StringBuilder()
            builder.Append(Me.KeyPart(pickup.BlDate))
            builder.Append(Me.KeyPart(pickup.Vessel))
            builder.Append(Me.KeyPart(pickup.OtsunakaNo))
            builder.Append(Me.KeyPart(CType(haisoYoteiYmd, Date?)))
            builder.Append(Me.KeyPart(pickup.HaisoType))
            builder.Append(Me.KeyPart(pickup.ContainerTypeText))
            builder.Append(Me.KeyPart(pickup.ContainerNo))
            builder.Append(Me.KeyPart(pickup.SealNo))
            builder.Append(Me.KeyPart(pickup.HaisoSaki1))
            builder.Append(Me.KeyPart(pickup.HaisoSaki2))
            builder.Append(Me.KeyPart(pickup.HaisoSaki3))
            Return builder.ToString()
        End Function

        ''' <summary>
        ''' 印刷用の行を作ります。
        ''' </summary>
        ''' <param name="pickup">拾った行。</param>
        ''' <param name="haisoYoteiYmd">配送予定日。未入力を補った後の値。</param>
        ''' <returns>印刷用の行。</returns>
        Private Function CreateRow(pickup As ShanghaiHaisoTehaiPickup,
                                   haisoYoteiYmd As Date) As ShanghaiHaisoTehaiPrintRow

            Dim row As New ShanghaiHaisoTehaiPrintRow()
            row.InvoiceNoMain = pickup.InvoiceNoMain
            row.InvoiceNo = pickup.InvoiceNo
            row.Vessel = pickup.Vessel
            row.BlDate = pickup.BlDate
            row.HaisoType = pickup.HaisoType
            row.ContainerType = pickup.ContainerTypeText
            row.ContainerNo = pickup.ContainerNo
            row.SealNo = pickup.SealNo
            row.TotalCarton = If(pickup.CartonQty, 0)
            row.TotalQuantity = If(pickup.TotalSuryo, 0)
            row.HaisoYoteiYmd = haisoYoteiYmd
            row.HaisoSaki1 = pickup.HaisoSaki1
            row.HaisoSaki2 = pickup.HaisoSaki2
            row.HaisoSaki3 = pickup.HaisoSaki3
            row.OtsunakaNo = pickup.OtsunakaNo
            Return row
        End Function

        ''' <summary>
        ''' 比較用の文字列を作ります。区切りを付けて値どうしが混ざらないようにします。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>比較用の文字列。</returns>
        Private Function KeyPart(value As String) As String
            Return If(value, String.Empty) & vbTab
        End Function

        ''' <summary>
        ''' 比較用の文字列を作ります。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>比較用の文字列。</returns>
        Private Function KeyPart(value As Integer?) As String
            ' 値が無いとき
            If Not value.HasValue Then
                Return vbTab
            End If
            Return value.Value.ToString(CultureInfo.InvariantCulture) & vbTab
        End Function

        ''' <summary>
        ''' 比較用の文字列を作ります。時刻は無視し、日付だけで比べます。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>比較用の文字列。</returns>
        Private Function KeyPart(value As Date?) As String
            ' 値が無いとき
            If Not value.HasValue Then
                Return vbTab
            End If
            Return value.Value.ToString("yyyyMMdd", CultureInfo.InvariantCulture) & vbTab
        End Function
#End Region

    End Class
End Namespace
