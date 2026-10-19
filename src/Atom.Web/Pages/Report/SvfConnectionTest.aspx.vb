Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Configuration
Imports System.Globalization
Imports System.IO
Imports System.Security.Principal
Imports Atom.Report
Imports Atom.Web.Common

Namespace Pages.Report

    ''' <summary>
    ''' SVF 連携確認画面です。
    ''' </summary>
    ''' <remarks>
    ''' Web サーバーから SVF の監視フォルダへ連携ファイルを配置できるかを確かめる、動作確認用の画面です。
    ''' 手作業で置けても、IIS のアプリケーションプールの実行アカウントでは権限が足りない場合があるため、
    ''' 実際にアプリケーションから書き込んで確かめます。
    ''' 本番の帳票画面を作る際に、本画面は削除して差し支えありません。
    ''' </remarks>
    Public Class SvfConnectionTest
        Inherits BasePage

#Region "定数"
        ''' <summary>この画面のメニュー番号。メニューには載せないため 0 とします。</summary>
        Private Const ThisMenuNo As Integer = 0

        ''' <summary>監視フォルダのパスを保持する設定キー。</summary>
        Private Const WatchFolderKey As String = "Svf.WatchFolderPath"

        ''' <summary>出力先の論理プリンタ名を保持する設定キー。</summary>
        Private Const OutputPrinterKey As String = "Svf.OutputPrinter"

        ''' <summary>連携ファイルの接頭辞。疎通確認用のジョブが拾います。</summary>
        Private Const FilePrefix As String = "AT00"

        ''' <summary>ファイル名に使う日時の書式。</summary>
        Private Const TimestampFormat As String = "yyyyMMdd_HHmmss"

        ''' <summary>確認用に渡す参照番号。</summary>
        Private Const TestRefNo As String = "TEST-0001"
#End Region

#Region "イベント"
        ''' <summary>
        ''' 画面の読み込み時の処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
            Me.MenuNo = ThisMenuNo

            ' 初回表示のとき
            If Not Me.IsPostBack Then
                Me.ShowEnvironment()
            End If
        End Sub

        ''' <summary>
        ''' 「連携ファイルを置く」を押したときの処理です。
        ''' </summary>
        ''' <param name="sender">イベントの発生元。</param>
        ''' <param name="e">イベント引数。</param>
        Private Sub BtnWrite_Click(sender As Object, e As EventArgs) Handles btnWrite.Click
            Me.ShowEnvironment()

            Dim watchFolderPath As String = GetSetting(WatchFolderKey)
            ' 監視フォルダが設定されていないとき
            If String.IsNullOrWhiteSpace(watchFolderPath) Then
                Me.ShowResult("Web.config の " & WatchFolderKey & " が設定されていません。")
                Return
            End If

            Try
                Dim filePath As String = Me.WriteDataFile(watchFolderPath)
                Me.ShowResult(
                    "配置しました。" & vbCrLf &
                    filePath & vbCrLf & vbCrLf &
                    "SVF の監視間隔は 15 秒です。しばらく待ってから、" &
                    "ファイルが消えていること と PDF が出力されていること を確認してください。")
            Catch ex As UnauthorizedAccessException
                Me.ShowResult("書き込みが拒否されました。実行アカウントにフォルダへの書き込み権限がありません。" &
                              vbCrLf & ex.Message)
            Catch ex As IOException
                Me.ShowResult("ファイルの入出力で失敗しました。" & vbCrLf & ex.Message)
            End Try
        End Sub
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 連携ファイルを組み立てて配置します。
        ''' </summary>
        ''' <param name="watchFolderPath">監視フォルダのパス。</param>
        ''' <returns>配置したファイルのフルパス。</returns>
        Private Function WriteDataFile(watchFolderPath As String) As String
            Dim outputPrinter As String = GetSetting(OutputPrinterKey)

            Dim controlLines As New List(Of String)()
            controlLines.Add(SvfDataFileWriter.CreateQuotedControlLine("VrSetOutputPrinter", outputPrinter))
            controlLines.Add(SvfDataFileWriter.CreateControlLine("VrSetUseReportDirector", "1"))
            controlLines.Add(SvfDataFileWriter.CreateQuotedControlLine("VrSetDocName2", "ATOM接続テスト"))

            Dim headers As New List(Of String)()
            headers.Add("REF_NO")
            headers.Add("MEMO")

            Dim row As New List(Of String)()
            row.Add(TestRefNo)
            row.Add("Web からの配置テスト " & DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture))

            Dim rows As New List(Of IList(Of String))()
            rows.Add(row)

            Dim writer As New SvfDataFileWriter()
            Return writer.Write(watchFolderPath, Me.BuildFileName(), controlLines, headers, rows, Me.CurrentLoginCode())
        End Function

        ''' <summary>
        ''' 連携ファイルのファイル名を組み立てます。
        ''' </summary>
        ''' <returns>「接頭辞_参照番号_日時.txt」の形のファイル名。</returns>
        Private Function BuildFileName() As String
            Dim timestamp As String = DateTime.Now.ToString(TimestampFormat, CultureInfo.InvariantCulture)
            Return FilePrefix & "_" & TestRefNo & "_" & timestamp & ".txt"
        End Function

        ''' <summary>
        ''' 実行環境の状態を画面に表示します。
        ''' </summary>
        Private Sub ShowEnvironment()
            Me.litAccount.Text = GetCurrentAccountName()

            Dim watchFolderPath As String = GetSetting(WatchFolderKey)
            Me.litWatchFolder.Text = If(String.IsNullOrWhiteSpace(watchFolderPath), "（未設定）", watchFolderPath)
            Me.litFolderState.Text = GetFolderState(watchFolderPath)
            Me.litFileName.Text = Me.BuildFileName()
        End Sub

        ''' <summary>
        ''' 結果を画面に表示します。
        ''' </summary>
        ''' <param name="message">表示する内容。</param>
        Private Sub ShowResult(message As String)
            Me.litResult.Text = message
            Me.pnlResult.Visible = True
        End Sub

        ''' <summary>
        ''' 操作している利用者のログインコードを取得します。
        ''' </summary>
        ''' <returns>ログインコード。</returns>
        Private Function CurrentLoginCode() As String
            Return SessionContext.Current.LoginCode
        End Function

        ''' <summary>
        ''' 処理を実行しているアカウント名を取得します。
        ''' </summary>
        ''' <returns>アカウント名。取得できないときはその旨の文言。</returns>
        ''' <remarks>
        ''' IIS のアプリケーションプールの ID で動くため、手作業で置けても書けない場合があります。
        ''' どのアカウントで書こうとしているかが分かると、権限設定の依頼がしやすくなります。
        ''' </remarks>
        Private Shared Function GetCurrentAccountName() As String
            Dim identity As WindowsIdentity = WindowsIdentity.GetCurrent()
            ' 取得できなかったとき
            If identity Is Nothing Then
                Return "（取得できません）"
            End If
            Return identity.Name
        End Function

        ''' <summary>
        ''' 監視フォルダの状態を調べます。
        ''' </summary>
        ''' <param name="folderPath">調べるフォルダのパス。</param>
        ''' <returns>状態を表す文言。</returns>
        Private Shared Function GetFolderState(folderPath As String) As String
            ' 設定されていないとき
            If String.IsNullOrWhiteSpace(folderPath) Then
                Return "（未設定）"
            End If

            Try
                ' 存在するとき
                If Directory.Exists(folderPath) Then
                    Return "参照できます。"
                End If
                Return "見つかりません。パスと共有の公開設定を確認してください。"
            Catch ex As IOException
                Return "参照できません。" & ex.Message
            Catch ex As UnauthorizedAccessException
                Return "参照が拒否されました。" & ex.Message
            End Try
        End Function

        ''' <summary>
        ''' 設定値を取得します。
        ''' </summary>
        ''' <param name="key">設定キー。</param>
        ''' <returns>設定値。未設定のときは空文字。</returns>
        Private Shared Function GetSetting(key As String) As String
            Dim value As String = ConfigurationManager.AppSettings(key)
            Return If(value, String.Empty)
        End Function
#End Region

    End Class
End Namespace
