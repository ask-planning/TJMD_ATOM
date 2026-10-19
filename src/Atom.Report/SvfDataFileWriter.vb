Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports Atom.Core.Logging

''' <summary>
''' SVF へ渡す連携ファイルを生成し、監視フォルダへ配置します。
''' </summary>
''' <remarks>
''' SVF（Universal Connect/X）は監視フォルダに置かれたテキストファイルを拾い、
''' 様式へ流し込んで PDF を出力します。本クラスはそのテキストファイルを作る役割だけを持ちます。
''' 仕様は docs/specs/103_INVOICEヘッダー入力.md §19 を正とします。
'''
''' ファイルの形式は次のとおりです。
''' <code>
''' &lt;start&gt;
''' VrSetOutputPrinter="（論理プリンタ名）"
''' VrSetUseReportDirector=1
''' VrSetDocName2="（ドキュメント名）"
''' &lt;end&gt;
''' 見出し行（タブ区切り）
''' データ行（タブ区切り）
''' </code>
'''
''' 文字コードは Shift_JIS（CP932）、改行は CRLF、BOM は付けません。
''' SVF 側の「SVF エンコーディング」が MS932 のため、これ以外では読めません。
''' </remarks>
Public Class SvfDataFileWriter

#Region "定数"
    ''' <summary>ログの発生元名。</summary>
    Private Const LogSource As String = "SvfDataFileWriter"

    ''' <summary>連携ファイルの文字コード（Shift_JIS）。</summary>
    Private Const CodePageShiftJis As Integer = 932

    ''' <summary>制御部の開始行。</summary>
    Private Const ControlStartLine As String = "<start>"

    ''' <summary>制御部の終了行。</summary>
    Private Const ControlEndLine As String = "<end>"

    ''' <summary>列の区切り文字。</summary>
    Private Const ColumnSeparator As String = vbTab

    ''' <summary>
    ''' 書き込み中のファイルに付ける拡張子です。
    ''' </summary>
    ''' <remarks>
    ''' SVF は一定間隔で監視フォルダを見ているため、書き込み途中のファイルを拾うと
    ''' 内容が欠けたまま処理されてしまいます。これを避けるため、いったん監視対象外の
    ''' 名前で書き出し、書き終えてから正しい名前へ変更します。
    ''' </remarks>
    Private Const WritingExtension As String = ".writing"
#End Region

#Region "公開メソッド"
    ''' <summary>
    ''' 連携ファイルを生成し、監視フォルダへ配置します。
    ''' </summary>
    ''' <param name="watchFolderPath">SVF の監視フォルダのパス。</param>
    ''' <param name="fileName">配置するファイル名。拡張子を含めます。</param>
    ''' <param name="controlLines">制御部に書く行。「名前=値」の形で渡します。</param>
    ''' <param name="headers">見出し行の項目名。</param>
    ''' <param name="rows">データ行。各行の列数は見出しと一致させます。</param>
    ''' <param name="loginCode">操作している利用者のログインコード。ログに記録します。</param>
    ''' <returns>配置したファイルのフルパス。</returns>
    Public Function Write(watchFolderPath As String,
                          fileName As String,
                          controlLines As IList(Of String),
                          headers As IList(Of String),
                          rows As IList(Of IList(Of String)),
                          loginCode As String) As String

        Me.Validate(watchFolderPath, fileName, headers, rows)

        Dim content As String = Me.BuildContent(controlLines, headers, rows)
        Dim finalPath As String = Path.Combine(watchFolderPath, fileName)
        Dim writingPath As String = finalPath & WritingExtension

        Try
            ' 監視対象外の名前で書き出す
            Dim encoding As Encoding = Encoding.GetEncoding(CodePageShiftJis)
            File.WriteAllText(writingPath, content, encoding)

            ' 同名が残っているときは先に消す（File.Move は上書きできないため）
            If File.Exists(finalPath) Then
                File.Delete(finalPath)
            End If

            ' 書き終えてから監視対象の名前へ変更する
            File.Move(writingPath, finalPath)
        Catch ex As Exception
            Me.DeleteQuietly(writingPath, loginCode)
            Logger.WriteError(LogSource, "連携ファイルの配置に失敗しました。配置先 " & finalPath, ex, loginCode)
            Throw
        End Try

        Logger.WriteInformation(LogSource, "連携ファイルを配置しました。" & finalPath, loginCode)
        Return finalPath
    End Function

    ''' <summary>
    ''' 制御部の 1 行を組み立てます。値をダブルクォートで囲みます。
    ''' </summary>
    ''' <param name="name">SVF の API 関数名。</param>
    ''' <param name="value">設定する値。</param>
    ''' <returns>「名前="値"」の形の文字列。</returns>
    Public Shared Function CreateQuotedControlLine(name As String, value As String) As String
        Return name & "=""" & NormalizeValue(value) & """"
    End Function

    ''' <summary>
    ''' 制御部の 1 行を組み立てます。値をそのまま書きます。数値の指定に使います。
    ''' </summary>
    ''' <param name="name">SVF の API 関数名。</param>
    ''' <param name="value">設定する値。</param>
    ''' <returns>「名前=値」の形の文字列。</returns>
    Public Shared Function CreateControlLine(name As String, value As String) As String
        Return name & "=" & NormalizeValue(value)
    End Function
#End Region

#Region "内部処理"
    ''' <summary>
    ''' 引数の内容を確かめます。
    ''' </summary>
    ''' <param name="watchFolderPath">監視フォルダのパス。</param>
    ''' <param name="fileName">ファイル名。</param>
    ''' <param name="headers">見出し行。</param>
    ''' <param name="rows">データ行。</param>
    Private Sub Validate(watchFolderPath As String,
                         fileName As String,
                         headers As IList(Of String),
                         rows As IList(Of IList(Of String)))

        If String.IsNullOrWhiteSpace(watchFolderPath) Then
            Throw New ArgumentException("監視フォルダのパスが設定されていません。", NameOf(watchFolderPath))
        End If

        If String.IsNullOrWhiteSpace(fileName) Then
            Throw New ArgumentException("ファイル名が指定されていません。", NameOf(fileName))
        End If

        If headers Is Nothing OrElse headers.Count = 0 Then
            Throw New ArgumentException("見出し行が指定されていません。", NameOf(headers))
        End If

        If rows Is Nothing Then
            Throw New ArgumentNullException(NameOf(rows))
        End If

        ' 列数が食い違うと SVF 側で列がずれるため、ここで止める
        Dim rowNo As Integer = 0
        For Each row As IList(Of String) In rows
            rowNo += 1
            If row Is Nothing OrElse row.Count <> headers.Count Then
                Dim actual As String = If(row Is Nothing, "0", row.Count.ToString())
                Throw New ArgumentException(
                    "データ行の列数が見出しと一致しません。" &
                    rowNo.ToString() & " 行目は " & actual & " 列、見出しは " & headers.Count.ToString() & " 列です。",
                    NameOf(rows))
            End If
        Next
    End Sub

    ''' <summary>
    ''' 連携ファイルの中身を組み立てます。
    ''' </summary>
    ''' <param name="controlLines">制御部に書く行。</param>
    ''' <param name="headers">見出し行。</param>
    ''' <param name="rows">データ行。</param>
    ''' <returns>ファイルに書き込む文字列。</returns>
    Private Function BuildContent(controlLines As IList(Of String),
                                  headers As IList(Of String),
                                  rows As IList(Of IList(Of String))) As String

        Dim builder As New StringBuilder()

        builder.Append(ControlStartLine).Append(vbCrLf)
        ' 制御部の指定があるとき
        If controlLines IsNot Nothing Then
            For Each line As String In controlLines
                builder.Append(NormalizeValue(line)).Append(vbCrLf)
            Next
        End If
        builder.Append(ControlEndLine).Append(vbCrLf)

        builder.Append(Me.BuildLine(headers)).Append(vbCrLf)

        For Each row As IList(Of String) In rows
            builder.Append(Me.BuildLine(row)).Append(vbCrLf)
        Next

        Return builder.ToString()
    End Function

    ''' <summary>
    ''' 1 行分をタブ区切りで組み立てます。
    ''' </summary>
    ''' <param name="values">列の値。</param>
    ''' <returns>タブ区切りの文字列。</returns>
    Private Function BuildLine(values As IList(Of String)) As String
        Dim cells As New List(Of String)(values.Count)
        For Each value As String In values
            cells.Add(NormalizeValue(value))
        Next
        Return String.Join(ColumnSeparator, cells.ToArray())
    End Function

    ''' <summary>
    ''' 値を 1 行 1 セルに収まる形へ整えます。
    ''' </summary>
    ''' <param name="value">元の値。</param>
    ''' <returns>タブと改行を半角空白に置き換え、前後の空白を落とした値。</returns>
    ''' <remarks>
    ''' 値にタブや改行が混ざると列や行がずれ、SVF 側では原因が分かりません。
    ''' 取引先名や住所に改行が入っている場合があるため、ここで必ず落とします。
    ''' </remarks>
    Private Shared Function NormalizeValue(value As String) As String
        ' 値が無いとき
        If value Is Nothing Then
            Return String.Empty
        End If

        Dim normalized As String = value.Replace(vbTab, " ")
        normalized = normalized.Replace(vbCrLf, " ")
        normalized = normalized.Replace(vbCr, " ")
        normalized = normalized.Replace(vbLf, " ")
        Return normalized.Trim()
    End Function

    ''' <summary>
    ''' 書きかけのファイルを消します。消せなくても例外にしません。
    ''' </summary>
    ''' <param name="filePath">消すファイルのパス。</param>
    ''' <param name="loginCode">操作している利用者のログインコード。</param>
    ''' <remarks>
    ''' 本来の例外を握りつぶさないよう、後始末の失敗はログに残して先へ進めます。
    ''' </remarks>
    Private Sub DeleteQuietly(filePath As String, loginCode As String)
        Try
            If File.Exists(filePath) Then
                File.Delete(filePath)
            End If
        Catch ex As IOException
            Logger.WriteWarning(LogSource, "書きかけのファイルを削除できませんでした。" & filePath & " " & ex.Message, loginCode)
        Catch ex As UnauthorizedAccessException
            Logger.WriteWarning(LogSource, "書きかけのファイルを削除できませんでした。" & filePath & " " & ex.Message, loginCode)
        End Try
    End Sub
#End Region

End Class
