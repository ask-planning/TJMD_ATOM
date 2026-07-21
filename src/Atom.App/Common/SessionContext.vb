Imports System.Collections.Generic

Namespace Common
    ''' <summary>ログイン中の利用者情報と権限（現行 PUBLIC変数 / login系 に対応）。</summary>
    Public NotInheritable Class SessionContext

#Region "フィールド"
        Private Shared _current As SessionContext
        Private ReadOnly _allowedMenus As New HashSet(Of Integer)()
#End Region

#Region "プロパティ"
        ''' <summary>現在のセッション（シングルトン）。</summary>
        Public Shared ReadOnly Property Current As SessionContext
            Get
                ' 未生成のとき：新しく作る
                If _current Is Nothing Then
                    _current = New SessionContext()
                End If
                Return _current
            End Get
        End Property

        ''' <summary>ログイン社員コード。</summary>
        Public Property LoginCode As String = ""
        ''' <summary>ログイン者の氏名。</summary>
        Public Property LoginName As String = ""
        ''' <summary>所属部門。</summary>
        Public Property Department As String = ""
#End Region

#Region "公開メソッド"
        ''' <summary>許可メニュー番号を設定する（ログイン時に権限取得結果を反映）。</summary>
        ''' <param name="menuNos">許可するメニュー番号の一覧。</param>
        Public Sub SetPermissions(menuNos As IEnumerable(Of Integer))
            ' 既存の許可を消してから設定し直す
            Me._allowedMenus.Clear()
            For Each menuNo As Integer In menuNos
                Me._allowedMenus.Add(menuNo)
            Next
        End Sub

        ''' <summary>指定メニューが許可されているかを返す。</summary>
        ''' <param name="menuNo">メニュー番号。</param>
        Public Function IsAllowed(menuNo As Integer) As Boolean
            Return Me._allowedMenus.Contains(menuNo)
        End Function
#End Region

    End Class
End Namespace
