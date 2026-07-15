''' <summary>
''' SVF帳票出力の基盤（SVF Client = 旧 SVF for .NET Framework を利用）。
''' ※ SVF Client のアセンブリ参照は導入後に追加する（未参照のためここではIF雛形のみ）。
''' </summary>
Public Class ReportService
    ''' <summary>帳票フォームファイル(SVFX-Designer作成)とデータで帳票を出力する。</summary>
    Public Sub Print(formFile As String, data As System.Data.DataTable)
        ' TODO: SVF Client の API で formFile にデータをマージして出力/印刷。
        Throw New NotImplementedException("SVF Client 導入後に実装")
    End Sub
End Class
