Option Strict On
Option Explicit On
Option Infer Off

Imports System.Configuration

Namespace Configuration

    ''' <summary>
    ''' いまどの環境で動いているかを判定し、環境ごとの設定値を取り出します。
    ''' </summary>
    ''' <remarks>
    ''' **切り替えは Atom.Environment の 1 キーだけです。**
    ''' 環境ごとに値が違う設定は、キー名の後ろに @環境名 を付けて併記します。
    '''
    ''' 例: Atom.UploadFolderPath@Local / Atom.UploadFolderPath@Dev / Atom.UploadFolderPath@Prod
    '''
    ''' 取り出すときは @環境名 を付けた値を先に探し、無ければ環境名なしのキーを使います。
    ''' どの環境でも同じ値のものは、これまでどおり環境名なしで 1 行だけ書けば足ります。
    '''
    ''' **未設定のときは本番ではなく Local として扱います。**
    ''' 設定を書き忘れたときに本番として動き出すほうが危険なためです。
    ''' </remarks>
    Public NotInheritable Class AppEnvironment

#Region "定数"
        ''' <summary>環境名のキー。</summary>
        Public Const EnvironmentKey As String = "Atom.Environment"

        ''' <summary>環境ごとのキーを作るときの区切り。</summary>
        Private Const Separator As String = "@"

        ''' <summary>開発 PC の環境名。</summary>
        Public Const Local As String = "Local"

        ''' <summary>検証サーバーの環境名。</summary>
        Public Const Dev As String = "Dev"

        ''' <summary>本番サーバーの環境名。</summary>
        Public Const Prod As String = "Prod"
#End Region

#Region "コンストラクタ"
        ''' <summary>インスタンス化を禁止します。</summary>
        Private Sub New()
        End Sub
#End Region

#Region "プロパティ"
        ''' <summary>
        ''' いまの環境名です。Local / Dev / Prod のいずれかを返します。
        ''' </summary>
        ''' <remarks>
        ''' 未設定や想定外の値のときは Local を返します（安全側に倒すため）。
        ''' </remarks>
        Public Shared ReadOnly Property Current As String
            Get
                Dim value As String = ConfigurationManager.AppSettings(EnvironmentKey)
                ' 未設定のとき
                If String.IsNullOrEmpty(value) Then
                    Return Local
                End If

                Dim trimmed As String = value.Trim()
                ' 本番のとき
                If String.Equals(trimmed, Prod, StringComparison.OrdinalIgnoreCase) Then
                    Return Prod
                End If
                ' 検証サーバーのとき
                If String.Equals(trimmed, Dev, StringComparison.OrdinalIgnoreCase) Then
                    Return Dev
                End If
                ' 開発 PC のとき、または想定外の値のとき
                Return Local
            End Get
        End Property

        ''' <summary>
        ''' 本番環境で動いているときに True を返します。
        ''' </summary>
        Public Shared ReadOnly Property IsProduction As Boolean
            Get
                Return String.Equals(Current, Prod, StringComparison.Ordinal)
            End Get
        End Property

        ''' <summary>
        ''' 画面に出すための環境の表示名です。本番のときは空文字を返します。
        ''' </summary>
        Public Shared ReadOnly Property DisplayName As String
            Get
                ' 本番のとき（目印を出さない）
                If String.Equals(Current, Prod, StringComparison.Ordinal) Then
                    Return String.Empty
                End If
                ' 検証サーバーのとき
                If String.Equals(Current, Dev, StringComparison.Ordinal) Then
                    Return "検証環境"
                End If
                ' 開発 PC のとき
                Return "開発環境"
            End Get
        End Property
#End Region

#Region "公開メソッド"
        ''' <summary>
        ''' 環境を考慮して appSettings の値を取得します。
        ''' </summary>
        ''' <param name="key">環境名を付けていないキー。</param>
        ''' <returns>設定値。どちらのキーも無いときは Nothing。</returns>
        ''' <remarks>
        ''' キー@環境名 を先に探し、無ければ キー を探します。
        ''' </remarks>
        Public Shared Function GetAppSetting(key As String) As String
            Dim scoped As String = ConfigurationManager.AppSettings(ScopedKey(key))
            ' 環境ごとの値があるとき
            If scoped IsNot Nothing Then
                Return scoped
            End If
            Return ConfigurationManager.AppSettings(key)
        End Function

        ''' <summary>
        ''' 環境を考慮して appSettings の値を取得します。未設定のときは例外を送出します。
        ''' </summary>
        ''' <param name="key">環境名を付けていないキー。</param>
        ''' <returns>前後の空白を落とした設定値。</returns>
        ''' <remarks>
        ''' 既定値で黙って動かさず、起動時に気づけるようにします。
        ''' </remarks>
        Public Shared Function GetRequiredAppSetting(key As String) As String
            Dim value As String = GetAppSetting(key)
            ' 未設定または空のとき
            If String.IsNullOrEmpty(value) OrElse value.Trim().Length = 0 Then
                Throw New ConfigurationErrorsException(
                    "Web.config の appSettings に " & ScopedKey(key) & " または " & key & " が設定されていません。")
            End If
            Return value.Trim()
        End Function

        ''' <summary>
        ''' 環境を考慮して接続文字列を取得します。
        ''' </summary>
        ''' <param name="name">環境名を付けていない接続名。</param>
        ''' <returns>接続文字列。見つからないときは例外を送出します。</returns>
        ''' <remarks>
        ''' 実値は connectionStrings.config にあります（Git の管理対象外）。
        ''' </remarks>
        Public Shared Function GetConnectionString(name As String) As String
            Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings(ScopedKey(name))
            ' 環境ごとの接続が無いとき
            If setting Is Nothing Then
                setting = ConfigurationManager.ConnectionStrings(name)
            End If

            ' どちらも無いとき
            If setting Is Nothing OrElse String.IsNullOrEmpty(setting.ConnectionString) Then
                Throw New ConfigurationErrorsException(
                    "接続文字列 " & ScopedKey(name) & " または " & name & " が設定されていません。" &
                    "connectionStrings.config を確認してください。")
            End If
            Return setting.ConnectionString
        End Function

        ''' <summary>
        ''' 接続文字列からデータベース名を取り出します。
        ''' </summary>
        ''' <param name="connectionString">接続文字列。</param>
        ''' <returns>データベース名。取り出せないときは空文字。</returns>
        ''' <remarks>
        ''' 画面の帯に接続先を出すために使います。
        ''' **接続文字列そのものを画面へ出してはいけません。**データベース名だけを取り出します。
        ''' </remarks>
        Public Shared Function GetDatabaseName(connectionString As String) As String
            ' 空のとき
            If String.IsNullOrEmpty(connectionString) Then
                Return String.Empty
            End If

            Try
                Dim builder As New System.Data.SqlClient.SqlConnectionStringBuilder(connectionString)
                Return builder.InitialCatalog
            Catch ex As ArgumentException
                ' 解析できない接続文字列のとき（帯の表示だけのため、落とさず空で返す）
                Return String.Empty
            End Try
        End Function
#End Region

#Region "内部処理"
        ''' <summary>
        ''' 環境名を付けたキーを作ります。
        ''' </summary>
        ''' <param name="key">環境名を付けていないキー。</param>
        ''' <returns>キー@環境名 の形のキー。</returns>
        Private Shared Function ScopedKey(key As String) As String
            Return key & Separator & Current
        End Function
#End Region

    End Class
End Namespace
