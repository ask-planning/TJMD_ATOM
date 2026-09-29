Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic

Namespace Common

    ''' <summary>
    ''' 画面 1 つ分の定義です。App_Data の画面定義ファイルから読み込みます。
    ''' </summary>
    Public Class ScreenDefinition

#Region "定数"
        ''' <summary>権限が無いときにボタンへ添える文言。</summary>
        Private Const NoPermissionText As String = "権限なし"
#End Region

#Region "プロパティ"
        ''' <summary>メニュー番号。</summary>
        Public Property MenuNo As Integer

        ''' <summary>メニューおよびヘッダーに表示する画面名。</summary>
        Public Property Title As String = String.Empty

        ''' <summary>遷移先のアプリケーション相対 URL。未実装のときは空文字。</summary>
        Public Property Url As String = String.Empty

        ''' <summary>ヘッダーの背景色。空のときは既定色を使います。</summary>
        Public Property HeaderBackColor As String = String.Empty

        ''' <summary>
        ''' メニューに表示しない画面かどうか。
        ''' 他の画面からのみ開く画面（取込内容確認など）に True を設定します。
        ''' メニューには出ませんが、メニュー番号による遷移・権限判定・ヘッダー表示は通常どおり行います。
        ''' </summary>
        Public Property Hidden As Boolean

        ''' <summary>画面が実装済みかどうか。</summary>
        Public ReadOnly Property IsImplemented As Boolean
            Get
                Return Not String.IsNullOrEmpty(Me.Url)
            End Get
        End Property

        ''' <summary>この画面を開く権限があるかどうか。</summary>
        Public ReadOnly Property IsAllowed As Boolean
            Get
                Return PermissionChecker.IsAllowed(Me.MenuNo)
            End Get
        End Property

        ''' <summary>
        ''' メニューのボタンを押せるかどうか。
        ''' 未実装でも押下は許可し、押下後に「現在未実装です。」を表示します。
        ''' </summary>
        Public ReadOnly Property IsEnabled As Boolean
            Get
                Return Me.IsAllowed
            End Get
        End Property

        ''' <summary>ボタンに添える補足文言。押せるときは空文字。</summary>
        Public ReadOnly Property StatusText As String
            Get
                ' 権限が無いとき
                If Not Me.IsAllowed Then
                    Return NoPermissionText
                End If
                Return String.Empty
            End Get
        End Property
#End Region

    End Class

    ''' <summary>
    ''' メニューのブロック（見出し付きのボタンのまとまり）です。
    ''' </summary>
    Public Class MenuBlock

#Region "定数"
        ''' <summary>ブロックの色名として認めるもの（水色）。</summary>
        Private Const BlueColorName As String = "blue"

        ''' <summary>ブロックの色名として認めるもの（ピンク）。</summary>
        Private Const PinkColorName As String = "pink"

        ''' <summary>水色ブロックの CSS クラス。</summary>
        Private Const BlueCssClass As String = "atom-block atom-block-blue"

        ''' <summary>ピンクブロックの CSS クラス。</summary>
        Private Const PinkCssClass As String = "atom-block atom-block-pink"

#End Region

#Region "プロパティ"
        ''' <summary>ブロックの見出し。</summary>
        Public Property Title As String = String.Empty

        ''' <summary>ブロックの色名。blue または pink。</summary>
        Public Property ColorName As String = BlueColorName

        ''' <summary>ブロックを置く列。1 から 3。</summary>
        Public Property Column As Integer = 1

        ''' <summary>ブロックに含まれる画面。メニューに出さないものも含みます。</summary>
        Public Property Screens As List([Of] ScreenDefinition)

        ''' <summary>メニューに表示する画面だけを取り出したもの。</summary>
        Public ReadOnly Property VisibleScreens As IList([Of] ScreenDefinition)
            Get
                Dim results As New List([Of] ScreenDefinition)()
                For Each screen As ScreenDefinition In Me.Screens
                    ' メニューに表示する画面のとき
                    If Not screen.Hidden Then
                        results.Add(screen)
                    End If
                Next
                Return results
            End Get
        End Property

        ''' <summary>
        ''' ブロックに適用する CSS クラスです。
        ''' 外部ファイルの値をそのまま class 属性へ出さないよう、既知の値だけを対応付けます。
        ''' </summary>
        Public ReadOnly Property CssClass As String
            Get
                Return Me.ColorCssClass
            End Get
        End Property

        ''' <summary>色名に対応する CSS クラス。</summary>
        Private ReadOnly Property ColorCssClass As String
            Get
                ' ピンク指定のとき
                If String.Equals(Me.ColorName, PinkColorName, StringComparison.OrdinalIgnoreCase) Then
                    Return PinkCssClass
                End If
                ' 上記以外のとき（既定は水色）
                Return BlueCssClass
            End Get
        End Property

#End Region

#Region "コンストラクタ"
        ''' <summary>既定の状態で初期化します。</summary>
        Public Sub New()
            Me.Screens = New List([Of] ScreenDefinition)()
        End Sub
#End Region

    End Class
End Namespace
