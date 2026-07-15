Imports System.Reflection
Imports System.Resources

Namespace Common
    ''' <summary>
    ''' 画面タイトルをリソース（Resources\ScreenTitles.resx）から取得する。
    ''' キーは menuNo。フォームにタイトル文字列をべた書きしないための共通取得口。
    ''' </summary>
    Public Module ScreenTitles

        ' タイトル定義リソース（埋め込み名 = ルート名前空間 + フォルダ + ファイル名）
        Private ReadOnly Titles As ResourceManager =
            New ResourceManager("Atom.App.Resources.ScreenTitles", Assembly.GetExecutingAssembly())

        ''' <summary>menuNo に対応する画面タイトルを返す（未定義・失敗時は空文字）。</summary>
        ''' <param name="menuNo">メニュー番号。</param>
        Public Function GetTitle(menuNo As Integer) As String
            Try
                ' リソースから menuNo をキーにタイトルを取得する
                Dim title As String = Titles.GetString(menuNo.ToString())
                If String.IsNullOrEmpty(title) Then
                    ' 未定義のとき：空文字を返す
                    Return ""
                End If
                Return title
            Catch
                ' 読込失敗のとき：空文字を返す
                Return ""
            End Try
        End Function
    End Module
End Namespace
