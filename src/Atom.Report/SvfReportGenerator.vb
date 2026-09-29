Option Strict On
Option Explicit On
Option Infer Off

Imports System.IO
Imports Atom.Core.Logging

''' <summary>
''' SVF による帳票出力を担います。
''' Web サーバー上で PDF を生成し、生成したファイルのパスを返します。
''' 呼び出し側は使用後に必ずファイルを削除してください。
''' </summary>
''' <remarks>
''' 要確認：SVF をサーバー側で実行する構成には、クライアント実行型とは別のライセンスが
''' 必要になる可能性があります。実装前に販売元へ確認してください。
''' </remarks>
Public Class SvfReportGenerator

#Region "定数"
    ''' <summary>ログの発生元名。</summary>
    Private Const LogSource As String = "SvfReportGenerator"
#End Region

#Region "公開メソッド"
    ''' <summary>
    ''' 帳票を PDF として生成します。
    ''' </summary>
    ''' <param name="formFilePath">SVF の帳票定義ファイルのパス。</param>
    ''' <param name="workFolderPath">一時ファイルを置くフォルダのパス。</param>
    ''' <param name="loginCode">操作している利用者のログインコード。</param>
    ''' <returns>生成した PDF ファイルのパス。</returns>
    Public Function CreatePdf(formFilePath As String, workFolderPath As String, loginCode As String) As String
        ' 同時実行で衝突しないよう一意なファイル名にする
        Dim outputPath As String = Path.Combine(workFolderPath, Guid.NewGuid().ToString("N") & ".pdf")

        Try
            ' TODO: SVF のランタイム（VrOut など）を用いて outputPath へ PDF を出力する
            Throw New NotImplementedException("SVF の帳票出力処理は未実装です。")
        Catch ex As Exception
            Logger.WriteError(LogSource, "帳票の生成に失敗しました。定義 " & formFilePath, ex, loginCode)
            Throw
        End Try

        Return outputPath
    End Function
#End Region

End Class
