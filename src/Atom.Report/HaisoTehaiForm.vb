Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic

''' <summary>
''' 配送手配書の様式です。
''' </summary>
''' <remarks>
''' 配送タイプが AIR かどうかで分かれ、AIR でなければ配送先の箇所数で決まります（設計書 §10.4）。
''' 連携ファイルの接頭辞で SVF のジョブが振り分けられ、帳票区分で PDF の名前が分かれます。
'''
''' **定期便（AT01〜03）と AIR（AT04）は列が違います。**
''' 定期便の 3 つは同じ 27 列で、様式に置かれていない列は無視されます。
''' AIR は中身が「輸入通関依頼」で共通の列が少ないため、**18 列の専用の形**を使います
''' （設計書 §14.1.1.1）。どちらの形かは <see cref="IsAir"/> で見分けます。
''' </remarks>
Public NotInheritable Class HaisoTehaiForm

#Region "フィールド"
    ''' <summary>箇所数から様式を引くための一覧。定期便の 3 つだけが入ります。</summary>
    Private Shared ReadOnly _formsByCount As Dictionary(Of Integer, HaisoTehaiForm) = CreateForms()

    ''' <summary>AIR の様式。配送先の箇所数では決まらないため別に持ちます。</summary>
    Private Shared ReadOnly _air As New HaisoTehaiForm("AT04", "AIR", "配送手配書（AIR）", True, False)

    ''' <summary>日新入荷の様式。配送手配書に付随して出すため別に持ちます。</summary>
    Private Shared ReadOnly _nisshin As New HaisoTehaiForm("AT05", "NISSHIN", "日新入荷", False, True)

    ''' <summary>
    ''' 入荷明細書の様式（menuNo=103 §11）。
    ''' </summary>
    ''' <remarks>
    ''' 配送手配書とは別のボタンから出す帳票ですが、**様式の一覧はここに集めています。**
    ''' PDF を表示するときに帳票区分を照合する先が <see cref="FindByReportKubun"/> だけであり、
    ''' 一覧から漏れると表示できなくなるためです。
    ''' 連携ファイルの組み立ては <c>NyukaMeisaiReportWriter</c> が担います。
    ''' </remarks>
    Private Shared ReadOnly _nyukaMeisai As New HaisoTehaiForm("AT06", "NYUKA", "入荷明細書", False, False)
#End Region

#Region "コンストラクタ"
    ''' <summary>
    ''' 様式を初期化します。
    ''' </summary>
    ''' <param name="filePrefix">連携ファイルの接頭辞。</param>
    ''' <param name="reportKubun">帳票区分。PDF の名前に使われます。</param>
    ''' <param name="displayName">画面に出す名前。</param>
    ''' <param name="isAir">AIR の様式のときは True。</param>
    ''' <param name="isNisshin">日新入荷の様式のときは True。</param>
    Private Sub New(filePrefix As String, reportKubun As String, displayName As String,
                    isAir As Boolean, isNisshin As Boolean)
        Me.FilePrefix = filePrefix
        Me.ReportKubun = reportKubun
        Me.DisplayName = displayName
        Me.IsAir = isAir
        Me.IsNisshin = isNisshin
    End Sub
#End Region

#Region "プロパティ"
    ''' <summary>連携ファイルの接頭辞。SVF のジョブがこれで拾います。</summary>
    Public ReadOnly Property FilePrefix As String

    ''' <summary>帳票区分。出力される PDF の名前に使われます。</summary>
    Public ReadOnly Property ReportKubun As String

    ''' <summary>画面に出す名前。</summary>
    Public ReadOnly Property DisplayName As String

    ''' <summary>
    ''' AIR の様式かどうかです。
    ''' </summary>
    ''' <remarks>連携ファイルの列が定期便と違うため、組み立て方を切り替えるのに使います。</remarks>
    Public ReadOnly Property IsAir As Boolean

    ''' <summary>
    ''' 日新入荷の様式かどうかです。
    ''' </summary>
    ''' <remarks>
    ''' 列は定期便と同じですが、**宛名の中身だけが違います**（乙仲ではなく配送先）。
    ''' その切り替えに使います（設計書 §14.1.1.2）。
    ''' </remarks>
    Public ReadOnly Property IsNisshin As Boolean
#End Region

#Region "公開メソッド"
    ''' <summary>
    ''' 配送先の箇所数から様式を求めます。
    ''' </summary>
    ''' <param name="haisoSakiCount">配送先の箇所数。</param>
    ''' <returns>対応する様式。範囲外のときは 1 箇所の様式。</returns>
    ''' <remarks>AIR はここでは求められません。<see cref="Air"/> を使ってください。</remarks>
    Public Shared Function FromHaisoSakiCount(haisoSakiCount As Integer) As HaisoTehaiForm
        ' 一覧に無いとき
        If Not _formsByCount.ContainsKey(haisoSakiCount) Then
            Return _formsByCount(1)
        End If
        Return _formsByCount(haisoSakiCount)
    End Function

    ''' <summary>
    ''' AIR の様式を返します。
    ''' </summary>
    ''' <returns>AIR の様式。</returns>
    Public Shared Function Air() As HaisoTehaiForm
        Return _air
    End Function

    ''' <summary>
    ''' 日新入荷の様式を返します。
    ''' </summary>
    ''' <returns>日新入荷の様式。</returns>
    Public Shared Function Nisshin() As HaisoTehaiForm
        Return _nisshin
    End Function

    ''' <summary>
    ''' 入荷明細書の様式を返します。
    ''' </summary>
    ''' <returns>入荷明細書の様式。</returns>
    Public Shared Function NyukaMeisai() As HaisoTehaiForm
        Return _nyukaMeisai
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
        For Each form As HaisoTehaiForm In All()
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
    ''' <returns>様式の一覧。定期便 3 つと AIR と日新入荷と入荷明細書。</returns>
    ''' <remarks>
    ''' PDF の表示で帳票区分を照合するために使います。**出力前に PDF を消す用途には使わないでください。**
    ''' 別のボタンで出した帳票まで消えてしまいます。そちらは <see cref="HaisoTehaiAll"/> を使います。
    ''' </remarks>
    Public Shared Function All() As IList(Of HaisoTehaiForm)
        Dim forms As New List(Of HaisoTehaiForm)(_formsByCount.Values)
        forms.Add(_air)
        forms.Add(_nisshin)
        forms.Add(_nyukaMeisai)
        Return forms
    End Function

    ''' <summary>
    ''' 配送手配書のボタンで出す様式だけを返します。
    ''' </summary>
    ''' <returns>定期便 3 つと AIR と日新入荷。</returns>
    ''' <remarks>
    ''' 「配送手配書」を押したときに消してよい PDF の範囲です。
    ''' 入荷明細書は別のボタンで出すため、ここには含めません。
    ''' </remarks>
    Public Shared Function HaisoTehaiAll() As IList(Of HaisoTehaiForm)
        Dim forms As New List(Of HaisoTehaiForm)(_formsByCount.Values)
        forms.Add(_air)
        forms.Add(_nisshin)
        Return forms
    End Function
#End Region

#Region "内部処理"
    ''' <summary>
    ''' 定期便の様式の一覧を作ります。
    ''' </summary>
    ''' <returns>箇所数で引ける一覧。</returns>
    Private Shared Function CreateForms() As Dictionary(Of Integer, HaisoTehaiForm)
        Dim forms As New Dictionary(Of Integer, HaisoTehaiForm)()
        forms.Add(1, New HaisoTehaiForm("AT01", "TEIKI", "配送手配書（定期便）", False, False))
        forms.Add(2, New HaisoTehaiForm("AT02", "TEIKI2", "配送手配書（定期便2）", False, False))
        forms.Add(3, New HaisoTehaiForm("AT03", "TEIKI3", "配送手配書（定期便3）", False, False))
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
