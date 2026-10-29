Option Strict On
Option Explicit On
Option Infer Off

Imports System.Text

Namespace Services

    ''' <summary>
    ''' 入力された文字列を揃えます。
    ''' </summary>
    ''' <remarks>
    ''' 利用者の入力は全角と半角が混ざります。コードや番号を全角のまま扱うと、
    ''' 検索で一致しない、帳票のファイル名に全角が入るといった不具合になります。
    ''' **保持する前にここで半角へ揃えます。**
    '''
    ''' `Microsoft.VisualBasic.Strings.StrConv` は使っていません。
    ''' あちらは全角カナまで半角カナへ変換するため、対象が広すぎます。
    ''' ここでは**英数字と空白だけ**を狭めます。
    ''' </remarks>
    Public NotInheritable Class TextNormalizer

#Region "定数"
        ''' <summary>全角の数字「０」。</summary>
        Private Const WideZero As Char = "０"c

        ''' <summary>全角の数字「９」。</summary>
        Private Const WideNine As Char = "９"c

        ''' <summary>全角の英大文字「Ａ」。</summary>
        Private Const WideUpperA As Char = "Ａ"c

        ''' <summary>全角の英大文字「Ｚ」。</summary>
        Private Const WideUpperZ As Char = "Ｚ"c

        ''' <summary>全角の英小文字「ａ」。</summary>
        Private Const WideLowerA As Char = "ａ"c

        ''' <summary>全角の英小文字「ｚ」。</summary>
        Private Const WideLowerZ As Char = "ｚ"c

        ''' <summary>全角と半角の文字コードの差。</summary>
        Private Const WideToNarrowOffset As Integer = &HFEE0

        ''' <summary>全角の空白。</summary>
        Private Const WideSpace As Char = "　"c

        ''' <summary>半角の空白。</summary>
        Private Const NarrowSpace As Char = " "c
#End Region

#Region "コンストラクタ"
        ''' <summary>インスタンス化を禁止します。</summary>
        Private Sub New()
        End Sub
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' 全角の英数字と空白を半角へ揃えます。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>半角へ揃えた値。元が空のときは空文字。</returns>
        ''' <remarks>
        ''' 英数字と空白以外（かな・漢字・記号）はそのまま残します。
        ''' </remarks>
        Public Shared Function ToHalfWidth(value As String) As String
            ' 値が無いとき
            If String.IsNullOrEmpty(value) Then
                Return String.Empty
            End If

            Dim builder As New StringBuilder(value.Length)
            For Each letter As Char In value
                builder.Append(NarrowChar(letter))
            Next
            Return builder.ToString()
        End Function

        ''' <summary>
        ''' 全角の英数字と空白を半角へ揃え、前後の空白を落とします。
        ''' </summary>
        ''' <param name="value">元の値。</param>
        ''' <returns>揃えた値。元が空のときは空文字。</returns>
        ''' <remarks>
        ''' 全角空白は `Trim` では落ちないため、**半角へ揃えてから**落とします。
        ''' </remarks>
        Public Shared Function ToHalfWidthAndTrim(value As String) As String
            Return ToHalfWidth(value).Trim()
        End Function
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 1 文字を半角へ揃えます。
        ''' </summary>
        ''' <param name="letter">元の文字。</param>
        ''' <returns>半角へ揃えた文字。対象外のときは元の文字。</returns>
        Private Shared Function NarrowChar(letter As Char) As Char
            ' 全角の空白のとき
            If letter = WideSpace Then
                Return NarrowSpace
            End If

            ' 全角の数字のとき
            If letter >= WideZero AndAlso letter <= WideNine Then
                Return ToNarrow(letter)
            End If

            ' 全角の英大文字のとき
            If letter >= WideUpperA AndAlso letter <= WideUpperZ Then
                Return ToNarrow(letter)
            End If

            ' 全角の英小文字のとき
            If letter >= WideLowerA AndAlso letter <= WideLowerZ Then
                Return ToNarrow(letter)
            End If

            Return letter
        End Function

        ''' <summary>
        ''' 全角の文字を半角の文字へ移します。
        ''' </summary>
        ''' <param name="letter">全角の文字。</param>
        ''' <returns>半角の文字。</returns>
        Private Shared Function ToNarrow(letter As Char) As Char
            Return ChrW(AscW(letter) - WideToNarrowOffset)
        End Function
#End Region

    End Class
End Namespace
