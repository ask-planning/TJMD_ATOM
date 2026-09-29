Option Strict On
Option Explicit On
Option Infer Off

Namespace Entities

    ''' <summary>
    ''' 項目マスタ（dbo.tbl_m_koumoku）の 1 件分を表します。
    ''' 主キーは 項目コード ＋ 連番 の複合キーです。
    ''' </summary>
    Public Class Koumoku

#Region "プロパティ"
        ''' <summary>項目コード。</summary>
        Public Property KoumokuCode As String = String.Empty

        ''' <summary>連番。</summary>
        Public Property SeqNo As Integer = 0

        ''' <summary>内容1。</summary>
        Public Property Contents1 As String = String.Empty

        ''' <summary>内容2。</summary>
        Public Property Contents2 As String = String.Empty

        ''' <summary>内容3。</summary>
        Public Property Contents3 As String = String.Empty

        ''' <summary>内容4。</summary>
        Public Property Contents4 As String = String.Empty

        ''' <summary>数値。未設定のときは Nothing。</summary>
        Public Property ContentsNum As Decimal? = Nothing

        ''' <summary>並び順。未設定のときは Nothing。</summary>
        Public Property SortSeq As Integer? = Nothing

        ''' <summary>更新日時。未設定のときは Nothing。</summary>
        Public Property UpdateYmd As Date? = Nothing

        ''' <summary>更新者のログインコード。</summary>
        Public Property UpdateLogin As String = String.Empty
#End Region

    End Class

    ''' <summary>
    ''' 項目マスタの内容欄に表示する見出しです。
    ''' ストアドプロシージャ usp_koumoku_midashi_get の戻り値を保持します。
    ''' </summary>
    Public Class KoumokuHeading

#Region "プロパティ"
        ''' <summary>内容1の見出し。未指定のときは空文字。</summary>
        Public Property Heading1 As String = String.Empty

        ''' <summary>内容2の見出し。未指定のときは空文字。</summary>
        Public Property Heading2 As String = String.Empty

        ''' <summary>内容3の見出し。未指定のときは空文字。</summary>
        Public Property Heading3 As String = String.Empty

        ''' <summary>内容4の見出し。未指定のときは空文字。</summary>
        Public Property Heading4 As String = String.Empty
#End Region

    End Class

    ''' <summary>
    ''' 項目マスタで選択できる項目コードの 1 件分を表します。
    ''' 項目コードの一覧は、項目マスタ自身の pk_koumoku_code = 'KMK' の行に登録されています。
    ''' </summary>
    Public Class KoumokuKubun

#Region "プロパティ"
        ''' <summary>'KMK' 側の連番。専用画面名の取得に使います。</summary>
        Public Property SeqNo As Integer = 0

        ''' <summary>項目コード。</summary>
        Public Property Code As String = String.Empty

        ''' <summary>項目コードの名称。</summary>
        Public Property Name As String = String.Empty

        ''' <summary>選択肢に表示する「コード　名称」形式の文言。</summary>
        Public ReadOnly Property DisplayText As String
            Get
                Return Me.Code & "　" & Me.Name
            End Get
        End Property
#End Region

    End Class

    ''' <summary>
    ''' 項目マスタの更新可否を表します。
    ''' ストアドプロシージャ usp_koumoku_koushin_flg_get の戻り値に対応します。
    ''' </summary>
    Public Enum KoumokuUpdateMode

        ''' <summary>編集できます。</summary>
        Editable = 0

        ''' <summary>変更できません。</summary>
        [ReadOnly] = 1

        ''' <summary>追加のみできます。</summary>
        AddOnly = 2

    End Enum
End Namespace
