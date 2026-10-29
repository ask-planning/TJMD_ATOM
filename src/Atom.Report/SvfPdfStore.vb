Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.IO
Imports Atom.Core.Logging

''' <summary>
''' SVF が出力した PDF を扱います。
''' </summary>
''' <remarks>
''' 画面の流れは次のとおりです（設計書 §10 / §19）。
''' 1. 出力ボタンを押したら、**これから出す帳票の PDF だけ**を先に削除する
''' 2. 連携ファイルを置く（SVF が拾って PDF を作る。監視間隔は 15 秒）
''' 3. 利用者が任意のタイミングで表示ボタンを押す
''' 4. PDF があれば表示し、まだ無ければ「印刷中」と伝える
'''
''' **ほかの帳票の PDF を消さない**ため、削除は帳票区分を指定して行います。
'''
''' PDF の名前は SVF 側のジョブが決めます（`帳票区分_ログインコード.pdf`）。
''' 名前の規則は `docs/04_SVF帳票連携_設定手順.md` §6 を正とします。
''' </remarks>
Public Class SvfPdfStore

#Region "定数"
    ''' <summary>ログの発生元名。</summary>
    Private Const LogSource As String = "SvfPdfStore"

    ''' <summary>PDF の拡張子。</summary>
    Private Const PdfExtension As String = ".pdf"
#End Region

#Region "公開メソッド"
    ''' <summary>
    ''' 指定した帳票区分の PDF を削除します。
    ''' </summary>
    ''' <param name="pdfFolderPath">PDF の出力先フォルダ。</param>
    ''' <param name="reportKubuns">削除する帳票区分。</param>
    ''' <param name="loginCode">利用者のログインコード。</param>
    ''' <remarks>
    ''' **古い PDF が残っていると、利用者が前回の内容を今回のものと取り違えます。**
    ''' 出力の前に必ず消します。消せなかったときは、そのことを伝えて処理を止めます。
    ''' </remarks>
    Public Sub Delete(pdfFolderPath As String, reportKubuns As IList(Of String), loginCode As String)
        For Each reportKubun As String In reportKubuns
            Dim filePath As String = Me.CreatePath(pdfFolderPath, reportKubun, loginCode)
            ' 残っていないとき
            If Not File.Exists(filePath) Then
                Continue For
            End If

            File.Delete(filePath)
            Logger.WriteInformation(LogSource, "出力前に PDF を削除しました。" & filePath, loginCode)
        Next
    End Sub

    ''' <summary>
    ''' PDF が出力されているかを調べます。
    ''' </summary>
    ''' <param name="pdfFolderPath">PDF の出力先フォルダ。</param>
    ''' <param name="reportKubun">帳票区分。</param>
    ''' <param name="loginCode">利用者のログインコード。</param>
    ''' <returns>出力されているとき True。</returns>
    Public Function Exists(pdfFolderPath As String, reportKubun As String, loginCode As String) As Boolean
        Return File.Exists(Me.CreatePath(pdfFolderPath, reportKubun, loginCode))
    End Function

    ''' <summary>
    ''' PDF のフルパスを組み立てます。
    ''' </summary>
    ''' <param name="pdfFolderPath">PDF の出力先フォルダ。</param>
    ''' <param name="reportKubun">帳票区分。</param>
    ''' <param name="loginCode">利用者のログインコード。</param>
    ''' <returns>PDF のフルパス。</returns>
    ''' <remarks>
    ''' 帳票区分は呼び出し側で <see cref="HaisoTehaiForm.FindByReportKubun"/> を通し、
    ''' **定義済みの値であることを確かめてから**渡してください。
    ''' ログインコードはセッションの値だけを渡し、画面からの入力を使わないでください。
    ''' </remarks>
    Public Function CreatePath(pdfFolderPath As String, reportKubun As String, loginCode As String) As String
        Dim fileName As String = reportKubun & "_" & loginCode & PdfExtension
        Return Path.Combine(pdfFolderPath, fileName)
    End Function
#End Region

End Class
