Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic

''' <summary>
''' 配送手配書の様式です。
''' </summary>
''' <remarks>
''' 配送先の箇所数で決まります（設計書 §10.4）。
''' 3 つの様式は**同じ 27 列**を受け取り、様式に置かれていない列は無視されます。
''' 連携ファイルの接頭辞で SVF のジョブが振り分けられ、帳票区分で PDF の名前が分かれます。
''' </remarks>
Public NotInheritable Class HaisoTehaiForm

#Region "フィールド"
    ''' <summary>箇所数から様式を引くための一覧。</summary>
    Private Shared ReadOnly _formsByCount As Dictionary(Of Integer, HaisoTehaiForm) = CreateForms()
#End Region

#Region "コンストラクタ"
    ''' <summary>
    ''' 様式を初期化します。
    ''' </summary>
    ''' <param name="filePrefix">連携ファイルの接頭辞。</param>
    ''' <param name="reportKubun">帳票区分。PDF の名前に使われます。</param>
    ''' <param name="displayName">画面に出す名前。</param>
    Private Sub New(filePrefix As String, reportKubun As String, displayName As String)
        Me.FilePrefix = filePrefix
        Me.ReportKubun = reportKubun
        Me.DisplayName = displayName
    End Sub
#End Region

#Region "プロパティ"
    ''' <summary>連携ファイルの接頭辞。SVF のジョブがこれで拾います。</summary>
    Public ReadOnly Property FilePrefix As String

    ''' <summary>帳票区分。出力される PDF の名前に使われます。</summary>
    Public ReadOnly Property ReportKubun As String

    ''' <summary>画面に出す名前。</summary>
    Public ReadOnly Property DisplayName As String
#End Region

#Region "公開メソッド"
    ''' <summary>
    ''' 配送先の箇所数から様式を求めます。
    ''' </summary>
    ''' <param name="haisoSakiCount">配送先の箇所数。</param>
    ''' <returns>対応する様式。範囲外のときは 1 箇所の様式。</returns>
    Public Shared Function FromHaisoSakiCount(haisoSakiCount As Integer) As HaisoTehaiForm
        ' 一覧に無いとき
        If Not _formsByCount.ContainsKey(haisoSakiCount) Then
            Return _formsByCount(1)
        End If
        Return _formsByCount(haisoSakiCount)
    End Function

    ''' <summary>
    ''' 帳票区分から様式を求めます。
    ''' </summary>
    ''' <param name="reportKubun">帳票区分。</param>
    ''' <returns>対応する様式。見つからないときは Nothing。</returns>
    ''' <remarks>
    ''' PDF を表示するときに、画面から渡された値を確かめるために使います。
    ''' **一覧に無い値は受け付けません。**任意のパスを組み立てられないようにするためです。
    ''' </remarks>
    Public Shared Function FindByReportKubun(reportKubun As String) As HaisoTehaiForm
        For Each form As HaisoTehaiForm In _formsByCount.Values
            ' 帳票区分が一致したとき
            If String.Equals(form.ReportKubun, reportKubun, StringComparison.Ordinal) Then
                Return form
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' 配送手配書のすべての様式を返します。
    ''' </summary>
    ''' <returns>様式の一覧。</returns>
    Public Shared Function All() As IList(Of HaisoTehaiForm)
        Return New List(Of HaisoTehaiForm)(_formsByCount.Values)
    End Function
#End Region

#Region "内部処理"
    ''' <summary>
    ''' 様式の一覧を作ります。
    ''' </summary>
    ''' <returns>箇所数で引ける一覧。</returns>
    Private Shared Function CreateForms() As Dictionary(Of Integer, HaisoTehaiForm)
        Dim forms As New Dictionary(Of Integer, HaisoTehaiForm)()
        forms.Add(1, New HaisoTehaiForm("AT01", "TEIKI", "配送手配書（定期便）"))
        forms.Add(2, New HaisoTehaiForm("AT02", "TEIKI2", "配送手配書（定期便2）"))
        forms.Add(3, New HaisoTehaiForm("AT03", "TEIKI3", "配送手配書（定期便3）"))
        Return forms
    End Function
#End Region

End Class

''' <summary>
''' SVF の設定です。Web.config から読み取った値を渡します。
''' </summary>
''' <remarks>
''' パスや出力先をコードに直書きしないため、呼び出し側が設定を読んで渡します。
''' </remarks>
Public Class SvfReportSettings

    ''' <summary>監視フォルダのパス。連携ファイルを置く先です。</summary>
    Public Property WatchFolderPath As String = String.Empty

    ''' <summary>PDF の出力先フォルダのパス。</summary>
    Public Property PdfFolderPath As String = String.Empty

    ''' <summary>出力先の論理プリンタ名。</summary>
    Public Property OutputPrinter As String = String.Empty

End Class

''' <summary>
''' 配送手配書の配置結果です。
''' </summary>
Public Class HaisoTehaiWriteResult

    ''' <summary>配置した様式。</summary>
    Public ReadOnly Property Forms As IList(Of HaisoTehaiForm) = New List(Of HaisoTehaiForm)()

    ''' <summary>配置した連携ファイルのフルパス。</summary>
    Public ReadOnly Property FilePaths As IList(Of String) = New List(Of String)()

    ''' <summary>様式が無いため出さなかった AIR の件数。</summary>
    Public Property SkippedAirCount As Integer

    ''' <summary>
    ''' 1 本でも配置したかどうかです。
    ''' </summary>
    ''' <returns>配置したとき True。</returns>
    Public ReadOnly Property HasOutput As Boolean
        Get
            Return Me.FilePaths.Count > 0
        End Get
    End Property

End Class
