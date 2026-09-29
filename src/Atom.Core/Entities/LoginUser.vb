Option Strict On
Option Explicit On
Option Infer Off

Namespace Entities

    ''' <summary>
    ''' ログインした社員の情報です。TJM_master.dbo.tbl_m_login_sys から取得します。
    ''' </summary>
    ''' <remarks>
    ''' ModeAtom / ShanghaiChotatsu / BumonCode / MailAddress は現時点で用途が確定していません。
    ''' 旧 Access 版の後続処理で参照されている可能性があるため、値の保持のみ行い、
    ''' これらによる処理の分岐は実装しません。
    ''' </remarks>
    Public NotInheritable Class LoginUser

#Region "プロパティ"
        ''' <summary>社員コード（6桁）。</summary>
        Public Property LoginCode As String = String.Empty

        ''' <summary>社員名。ヘッダーに表示します。</summary>
        Public Property LoginName As String = String.Empty

        ''' <summary>部門コード。</summary>
        Public Property BumonCode As String = String.Empty

        ''' <summary>
        ''' ATOM の利用権限（login_mode_atom）。0 / 1 / 2 / 9 の値を取ります。
        ''' 値の意味が未確定のため、保持のみ行います。
        ''' </summary>
        Public Property ModeAtom As Integer?

        ''' <summary>上海調達の区分（shanghai_chotatsu）。用途が未確定のため保持のみ行います。</summary>
        Public Property ShanghaiChotatsu As Integer?

        ''' <summary>メールアドレス。旧 Access 版では未使用です。</summary>
        Public Property MailAddress As String = String.Empty
#End Region

    End Class
End Namespace
