Option Strict On
Option Explicit On
Option Infer Off

Imports System.Configuration

Namespace Common

    ''' <summary>
    ''' Web.config の appSettings から設定値を取得します。
    ''' 設定値をコードへ直接書かないため、参照はすべてこのクラスを通します。
    ''' </summary>
    Public NotInheritable Class AppSettings

#Region "定数"
        ''' <summary>システムバージョンのキー。</summary>
        Private Const SystemVersionKey As String = "Atom.SystemVersion"

        ''' <summary>マニュアル URL のキー。</summary>
        Private Const ManualUrlKey As String = "Atom.ManualUrl"

        ''' <summary>画面定義ファイルのキー。</summary>
        Private Const ScreenCatalogPathKey As String = "Atom.ScreenCatalogPath"
#End Region

#Region "コンストラクタ"
        ''' <summary>インスタンス化を禁止します。</summary>
        Private Sub New()
        End Sub
#End Region

#Region "プロパティ"
        ''' <summary>
        ''' システムバージョンです。ヘッダー表示とログイン履歴の version_no に使います。
        ''' </summary>
        Public Shared ReadOnly Property SystemVersion As String
            Get
                Return GetRequired(SystemVersionKey)
            End Get
        End Property

        ''' <summary>
        ''' マニュアルの URL です。未設定のときは空文字を返します。
        ''' </summary>
        Public Shared ReadOnly Property ManualUrl As String
            Get
                Dim value As String = ConfigurationManager.AppSettings(ManualUrlKey)
                ' 未設定のとき
                If value Is Nothing Then
                    Return String.Empty
                End If
                Return value.Trim()
            End Get
        End Property

        ''' <summary>
        ''' 画面定義ファイルのアプリケーション相対パスです。
        ''' </summary>
        Public Shared ReadOnly Property ScreenCatalogPath As String
            Get
                Return GetRequired(ScreenCatalogPathKey)
            End Get
        End Property
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 必須の設定値を取得します。未設定のときは例外を送出します。
        ''' </summary>
        ''' <param name="key">appSettings のキー。</param>
        ''' <returns>設定値。</returns>
        Private Shared Function GetRequired(key As String) As String
            Dim value As String = ConfigurationManager.AppSettings(key)
            ' 未設定または空のとき（既定値で黙って動かさず、起動時に気づけるようにする）
            If String.IsNullOrEmpty(value) OrElse value.Trim().Length = 0 Then
                Throw New ConfigurationErrorsException("Web.config の appSettings に " & key & " が設定されていません。")
            End If
            Return value.Trim()
        End Function
#End Region

    End Class
End Namespace
