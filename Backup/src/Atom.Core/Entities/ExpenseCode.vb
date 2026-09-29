Option Strict On
Option Explicit On
Option Infer Off

Namespace Entities

    ''' <summary>
    ''' 経費コード（menuNo=907）の 1 件分を表します。
    ''' 旧 ATOM の F_M_経費コード登録_SUB に対応します。
    ''' </summary>
    Public Class ExpenseCode

#Region "プロパティ"
        ''' <summary>経費コード。主キーです（pk_keihi_code）。</summary>
        Public Property KeihiCode As String = String.Empty

        ''' <summary>経費名称（keihi_code_text）。</summary>
        Public Property KeihiCodeText As String = String.Empty

        ''' <summary>経費区分（keihi_kbn）。</summary>
        Public Property KeihiKbn As String = String.Empty

        ''' <summary>更新日時（update_ymd）。未設定のときは Nothing。</summary>
        Public Property UpdateYmd As Date? = Nothing

        ''' <summary>更新者のログインコード（update_login）。</summary>
        Public Property UpdateLogin As String = String.Empty
#End Region

    End Class
End Namespace
