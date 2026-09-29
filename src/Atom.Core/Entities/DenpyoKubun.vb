Option Strict On
Option Explicit On
Option Infer Off

Namespace Entities

    ''' <summary>
    ''' 伝票区分の 1 件分を表します。輸入経費コード登録の伝票コード選択肢に使います。
    ''' </summary>
    Public Class DenpyoKubun

#Region "定数"
        ''' <summary>伝票コードの桁数。表示時にこの桁数までゼロ埋めします。</summary>
        Public Const CodeLength As Integer = 4

        ''' <summary>コードと名称の区切り文字（全角スペース）。</summary>
        Public Const CodeNameSeparator As String = "　"
#End Region

#Region "プロパティ"
        ''' <summary>伝票コード。</summary>
        Public Property Code As String = String.Empty

        ''' <summary>伝票区分の名称。</summary>
        Public Property Name As String = String.Empty

        ''' <summary>4 桁までゼロ埋めした伝票コード。</summary>
        Public ReadOnly Property PaddedCode As String
            Get
                Return Me.Code.PadLeft(CodeLength, "0"c)
            End Get
        End Property

        ''' <summary>選択肢に表示する「コード　名称」形式の文言。</summary>
        Public ReadOnly Property DisplayText As String
            Get
                Return Me.PaddedCode & CodeNameSeparator & Me.Name
            End Get
        End Property
#End Region

    End Class
End Namespace
