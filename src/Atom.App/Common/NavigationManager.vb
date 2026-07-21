Imports System.Windows.Forms

Namespace Common
    ''' <summary>画面遷移（現行 form_open_close 相当）。メニューから画面をモーダルで開き、閉じるとメニューへ戻る。</summary>
    Public Module NavigationManager
        ''' <summary>指定フォームをモーダルで開く（Nothing のときは何もしない）。</summary>
        ''' <param name="target">開く画面。</param>
        ''' <param name="owner">親ウィンドウ。</param>
        Public Sub OpenForm(target As Form, owner As IWin32Window)
            ' 未実装（Nothing）のとき：何もしない
            If target Is Nothing Then
                Return
            End If
            ' 画面をモーダル表示し、閉じたら破棄する
            Using target
                target.ShowDialog(owner)
            End Using
        End Sub
    End Module
End Namespace
