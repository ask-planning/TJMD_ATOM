\xEF\xBB\xBFNamespace Common
    ''' <summary>丸め共通関数（現行 四捨五入 / 切り捨て に対応）。</summary>
    Public Module NumberHelper
        ''' <summary>四捨五入する。</summary>
        Public Function RoundHalfUp(value As Decimal, digits As Integer) As Decimal
            Return Math.Round(value, digits, MidpointRounding.AwayFromZero)
        End Function

        ''' <summary>指定桁で切り捨てる。</summary>
        Public Function Truncate(value As Decimal, digits As Integer) As Decimal
            Dim factor As Decimal = CDec(Math.Pow(10, digits))
            Return Math.Truncate(value * factor) / factor
        End Function
    End Module
End Namespace
