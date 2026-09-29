Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic
Imports System.Configuration
Imports System.IO
Imports System.Text
Imports System.Web
Imports System.Web.Caching
Imports System.Web.Hosting
Imports System.Web.Script.Serialization

Namespace Common

    ''' <summary>
    ''' 画面定義（メニュー構成・遷移先 URL・ヘッダー表示）を保持します。
    ''' 定義は App_Data の JSON ファイルから読み込み、ファイルを更新すると再ビルドなしで反映されます。
    ''' </summary>
    Public Class ScreenCatalog

#Region "定数"
        ''' <summary>キャッシュに格納するときのキー。</summary>
        Private Const CacheKey As String = "Atom.ScreenCatalog"

        ''' <summary>ヘッダー背景色の最終的な既定値。定義ファイルに指定が無いときに使います。</summary>
        Private Const FallbackHeaderBackColor As String = "#400080"
#End Region

#Region "プロパティ"
        ''' <summary>ヘッダー背景色の既定値。</summary>
        Public Property DefaultHeaderBackColor As String = FallbackHeaderBackColor

        ''' <summary>メニューのブロック。表示順に並びます。</summary>
        Public Property Blocks As List(Of MenuBlock)

        ''' <summary>
        ''' 現在の画面定義を取得します。ファイルが更新されると自動で読み直します。
        ''' </summary>
        Public Shared ReadOnly Property Current As ScreenCatalog
            Get
                Dim cache As Cache = HttpRuntime.Cache
                Dim cached As ScreenCatalog = TryCast(cache(CacheKey), ScreenCatalog)
                ' 読み込み済みのとき
                If cached IsNot Nothing Then
                    Return cached
                End If

                Dim filePath As String = ResolveFilePath()
                Dim loaded As ScreenCatalog = Load(filePath)
                cache.Insert(CacheKey, loaded, New CacheDependency(filePath))
                Return loaded
            End Get
        End Property
#End Region

#Region "コンストラクタ"
        ''' <summary>既定の状態で初期化します。</summary>
        Public Sub New()
            Me.Blocks = New List(Of MenuBlock)()
        End Sub
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' メニュー番号に対応する画面定義を取得します。
        ''' </summary>
        ''' <param name="menuNo">メニュー番号。</param>
        ''' <returns>該当する画面定義。存在しないときは Nothing。</returns>
        Public Function FindScreen(menuNo As Integer) As ScreenDefinition
            For Each block As MenuBlock In Me.Blocks
                For Each screen As ScreenDefinition In block.Screens
                    ' 該当するメニュー番号のとき
                    If screen.MenuNo = menuNo Then
                        Return screen
                    End If
                Next
            Next
            ' 該当が無いとき
            Return Nothing
        End Function

        ''' <summary>
        ''' 指定した列に置くブロックを取得します。
        ''' </summary>
        ''' <param name="column">列の番号。1 から 3。</param>
        ''' <returns>その列に置くブロックを定義順に並べた一覧。該当が無いときは空の一覧。</returns>
        Public Function GetBlocksByColumn(column As Integer) As IList(Of MenuBlock)
            Dim results As New List(Of MenuBlock)()
            For Each block As MenuBlock In Me.Blocks
                ' 指定した列のブロックのとき
                If block.Column = column Then
                    results.Add(block)
                End If
            Next
            Return results
        End Function

        ''' <summary>
        ''' メニュー番号に対応するヘッダーの背景色を取得します。
        ''' </summary>
        ''' <param name="menuNo">メニュー番号。</param>
        ''' <returns>背景色の CSS 表記。画面固有の指定が無いときは既定色。</returns>
        Public Function GetHeaderBackColor(menuNo As Integer) As String
            Dim screen As ScreenDefinition = Me.FindScreen(menuNo)
            ' 画面固有の指定があるとき
            If screen IsNot Nothing AndAlso Not String.IsNullOrEmpty(screen.HeaderBackColor) Then
                Return screen.HeaderBackColor
            End If
            ' 指定が無いとき
            If String.IsNullOrEmpty(Me.DefaultHeaderBackColor) Then
                Return FallbackHeaderBackColor
            End If
            Return Me.DefaultHeaderBackColor
        End Function

        ''' <summary>
        ''' メニュー番号に対応する画面名を取得します。
        ''' </summary>
        ''' <param name="menuNo">メニュー番号。</param>
        ''' <returns>画面名。該当が無いときは空文字。</returns>
        Public Function GetTitle(menuNo As Integer) As String
            Dim screen As ScreenDefinition = Me.FindScreen(menuNo)
            ' 該当が無いとき
            If screen Is Nothing Then
                Return String.Empty
            End If
            Return screen.Title
        End Function
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 画面定義ファイルの物理パスを求めます。
        ''' </summary>
        ''' <returns>画面定義ファイルの物理パス。</returns>
        Private Shared Function ResolveFilePath() As String
            Dim virtualPath As String = AppSettings.ScreenCatalogPath
            Dim filePath As String = HostingEnvironment.MapPath(virtualPath)
            ' パスを解決できないとき
            If String.IsNullOrEmpty(filePath) Then
                Throw New ConfigurationErrorsException("画面定義ファイルのパスを解決できません。設定値: " & virtualPath)
            End If
            ' ファイルが無いとき
            If Not File.Exists(filePath) Then
                Throw New ConfigurationErrorsException("画面定義ファイルが見つかりません。設定値: " & virtualPath)
            End If
            Return filePath
        End Function

        ''' <summary>
        ''' 画面定義ファイルを読み込みます。
        ''' </summary>
        ''' <param name="filePath">画面定義ファイルの物理パス。</param>
        ''' <returns>読み込んだ画面定義。</returns>
        Private Shared Function Load(filePath As String) As ScreenCatalog
            Dim json As String = File.ReadAllText(filePath, Encoding.UTF8)
            Dim serializer As New JavaScriptSerializer()
            Dim catalog As ScreenCatalog = serializer.Deserialize(Of ScreenCatalog)(json)

            ' 内容が読み取れないとき
            If catalog Is Nothing Then
                Throw New ConfigurationErrorsException("画面定義ファイルの内容を読み取れません。パス: " & filePath)
            End If
            ' ブロックが 1 つも無いとき
            If catalog.Blocks Is Nothing Then
                catalog.Blocks = New List(Of MenuBlock)()
            End If
            Return catalog
        End Function
#End Region

    End Class
End Namespace
