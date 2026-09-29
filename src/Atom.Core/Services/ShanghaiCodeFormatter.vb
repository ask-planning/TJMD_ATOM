Option Strict On
Option Explicit On
Option Infer Off

Imports System.Globalization

Namespace Services

    ''' <summary>
    ''' 上海コードと R3 品目コードの書式をそろえます。
    ''' </summary>
    ''' <remarks>
    ''' 取込時（ShanghaiImportExcelParser）と、取込結果確認画面での編集時の両方から使います。
    ''' 同じ規則を 2 か所に書かないよう、ここへ集約しています。
    ''' すでにゼロ埋めされた値を渡しても結果は変わりません（何度実行しても同じ値になります）。
    ''' </remarks>
    Public NotInheritable Class ShanghaiCodeFormatter

#Region "定数"
        ''' <summary>上海コードの桁数。</summary>
        Private Const ShanghaiCodeLength As Integer = 5

        ''' <summary>R3 品目コードの桁数。</summary>
        Private Const R3HinmokuCodeLength As Integer = 18
#End Region

#Region "コンストラクタ"
        ''' <summary>インスタンス化を禁止します。</summary>
        Private Sub New()
        End Sub
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' 上海コードを整えます。数値のときは 5 桁になるようゼロ埋めします。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>整えた上海コード。</returns>
        Public Shared Function FormatShanghaiCode(value As String) As String
            Return Pad(value, ShanghaiCodeLength)
        End Function

        ''' <summary>
        ''' R3 品目コードを作ります。上海コードが数値のときは 18 桁になるようゼロ埋めします。
        ''' </summary>
        ''' <param name="shanghaiCode">上海コード。</param>
        ''' <returns>R3 品目コード。</returns>
        Public Shared Function FormatR3HinmokuCode(shanghaiCode As String) As String
            Return Pad(shanghaiCode, R3HinmokuCodeLength)
        End Function
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 数値のときだけ指定の桁数までゼロ埋めします。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <param name="length">ゼロ埋め後の桁数。</param>
        ''' <returns>整えた値。</returns>
        Private Shared Function Pad(value As String, length As Integer) As String
            ' 値が無いとき
            If String.IsNullOrEmpty(value) Then
                Return String.Empty
            End If
            ' 数値のとき
            If IsNumericCode(value) Then
                Return value.PadLeft(length, "0"c)
            End If
            ' 数値でないとき
            Return value
        End Function

        ''' <summary>
        ''' 数値として読める値かどうかを判定します。
        ''' </summary>
        ''' <param name="value">判定する値。</param>
        ''' <returns>数値として読めるとき True。</returns>
        Private Shared Function IsNumericCode(value As String) As Boolean
            Dim parsed As Double = 0
            Return Double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, parsed)
        End Function
#End Region

    End Class
End Namespace
