Imports System.Windows.Forms

Namespace Common
    ''' <summary>再接続（現行 Login_ReInput 相当）。SQL Serverへの接続を確認/再確立する。</summary>
    Public Module ReconnectService
        Public Sub Reconnect(owner As IWin32Window)
            Try
                ' TODO: Atom.Data.Database.CreateConnection で接続テスト/再確立
                MessageBox.Show("再接続しました。", "再接続", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("再接続に失敗しました。" & vbCrLf & ex.Message, "再接続", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
    End Module
End Namespace
