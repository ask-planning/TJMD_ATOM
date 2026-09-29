Option Strict On
Option Explicit On
Option Infer Off

Imports System.Collections.Generic

Namespace Common

    ''' <summary>
    ''' メニューの 1 項目を表します。
    ''' </summary>
    Public NotInheritable Class MenuEntry

#Region "コンストラクタ"
        ''' <summary>メニュー項目を初期化します。</summary>
        ''' <param name="menuNo">メニュー番号。</param>
        ''' <param name="title">ボタンに表示する名称。</param>
        Public Sub New(menuNo As Integer, title As String)
            Me.MenuNo = menuNo
            Me.Title = title
        End Sub
#End Region

#Region "プロパティ"
        ''' <summary>メニュー番号。</summary>
        Public ReadOnly Property MenuNo As Integer

        ''' <summary>ボタンに表示する名称。</summary>
        Public ReadOnly Property Title As String

        ''' <summary>この項目を押せるかどうか。権限があり、かつ画面が実装済みのとき True。</summary>
        Public ReadOnly Property IsEnabled As Boolean
            Get
                Return PermissionChecker.IsAllowed(Me.MenuNo) AndAlso PageUrlMap.IsImplemented(Me.MenuNo)
            End Get
        End Property

        ''' <summary>遷移先の画面 URL。未実装のときは空文字。</summary>
        Public ReadOnly Property NavigateUrl As String
            Get
                ' 押せないときは遷移先を持たせない
                If Not Me.IsEnabled Then
                    Return String.Empty
                End If
                Return PageUrlMap.GetUrl(Me.MenuNo)
            End Get
        End Property

        ''' <summary>ボタンに添える補足文言。押せるときは空文字。</summary>
        Public ReadOnly Property StatusText As String
            Get
                ' 権限が無いとき
                If Not PermissionChecker.IsAllowed(Me.MenuNo) Then
                    Return "権限なし"
                End If
                ' 画面が未実装のとき
                If Not PageUrlMap.IsImplemented(Me.MenuNo) Then
                    Return "未実装"
                End If
                Return String.Empty
            End Get
        End Property
#End Region

    End Class

    ''' <summary>
    ''' メニューのブロック（見出し付きのボタンのまとまり）を表します。
    ''' </summary>
    Public NotInheritable Class MenuBlock

#Region "コンストラクタ"
        ''' <summary>メニューブロックを初期化します。</summary>
        ''' <param name="title">ブロックの見出し。</param>
        ''' <param name="entries">ブロックに含まれるメニュー項目。</param>
        Public Sub New(title As String, entries As IList(Of MenuEntry))
            Me.Title = title
            Me.Entries = entries
        End Sub
#End Region

#Region "プロパティ"
        ''' <summary>ブロックの見出し。</summary>
        Public ReadOnly Property Title As String

        ''' <summary>ブロックに含まれるメニュー項目。</summary>
        Public ReadOnly Property Entries As IList(Of MenuEntry)
#End Region

    End Class

    ''' <summary>
    ''' メインメニューの構成を定義します。現行の F_0_START_MENU に相当します。
    ''' </summary>
    Public NotInheritable Class MenuDefinition

#Region "コンストラクタ"
        ''' <summary>インスタンス化を禁止します。</summary>
        Private Sub New()
        End Sub
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' メニューの全ブロックを取得します。
        ''' </summary>
        ''' <returns>表示順に並んだメニューブロックの一覧。</returns>
        Public Shared Function CreateBlocks() As IList(Of MenuBlock)
            Dim blocks As New List(Of MenuBlock)()

            blocks.Add(New MenuBlock("上海輸入定期便＆AIR便", New List(Of MenuEntry) From {
                New MenuEntry(101, "INVOICE取込"),
                New MenuEntry(102, "INVOICE確認"),
                New MenuEntry(103, "INVOICEヘッダー入力"),
                New MenuEntry(104, "R3購買発注取込"),
                New MenuEntry(105, "R3購買発注確認"),
                New MenuEntry(106, "R3在庫転送作成"),
                New MenuEntry(107, "SAP伝票金額入力")
            }))

            blocks.Add(New MenuBlock("その他輸入関連", New List(Of MenuEntry) From {
                New MenuEntry(201, "調達R3購買発注取込"),
                New MenuEntry(202, "調達R3購買発注確認"),
                New MenuEntry(203, "調達配送手配書作成")
            }))

            blocks.Add(New MenuBlock("輸入経費入力", New List(Of MenuEntry) From {
                New MenuEntry(301, "諸経費入力")
            }))

            blocks.Add(New MenuBlock("マスタ関連", New List(Of MenuEntry) From {
                New MenuEntry(902, "社内参照番号発番"),
                New MenuEntry(903, "項目マスタ"),
                New MenuEntry(904, "カレンダー"),
                New MenuEntry(905, "仕入先マスタ登録"),
                New MenuEntry(906, "商品名称入力"),
                New MenuEntry(907, "輸入経費コード登録")
            }))

            Return blocks
        End Function
#End Region

    End Class
End Namespace
