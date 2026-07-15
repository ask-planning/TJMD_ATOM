\xEF\xBB\xBFNamespace Common
    ''' <summary>メニュー番号に対する権限判定（現行 管理者フラグ判定 相当）。</summary>
    Public Module PermissionChecker
        ''' <summary>指定メニューが許可されているかを返す。</summary>
        ''' <param name="menuNo">メニュー番号。</param>
        Public Function IsAllowed(menuNo As Integer) As Boolean
            Return SessionContext.Current.IsAllowed(menuNo)
        End Function
    End Module
End Namespace
